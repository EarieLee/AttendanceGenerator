using System.Data;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;
using ExcelDataReader;
using AttendanceGenerator.Models;

namespace AttendanceGenerator.Services;

public sealed class AttendanceReader
{
    private readonly AttendanceRuleEngine _ruleEngine;
    private readonly Action<string> _log;

    public AttendanceReader(AttendanceRuleEngine ruleEngine, Action<string>? log = null)
    {
        _ruleEngine = ruleEngine;
        _log = log ?? (_ => { });
    }

    public IReadOnlyList<AttendanceRecord> Read(string filePath, YearMonth month, IList<AttendanceException> exceptions)
    {
        _log($"读取考勤报表：{filePath}");
        List<AttendanceRecord> records;
        try
        {
            using var workbook = new XLWorkbook(filePath);
            records = ReadMonthlySummary(workbook, month, exceptions);
            if (records.Count == 0)
            {
                _log("未在“月度汇总”识别到每日结果区，改用“每日统计”。");
                records = ReadDailySummary(workbook, month, exceptions);
            }
        }
        catch (Exception ex) when (ShouldTryTabularFallback(ex))
        {
            _log($"ClosedXML 读取考勤报表失败（{ex.Message}），改用兼容模式读取。");
            records = ReadWithExcelDataReader(filePath, month, exceptions);
        }

        foreach (var record in records)
        {
            record.NormalizedStatus = _ruleEngine.Normalize(record, exceptions);
        }

        _log($"考勤记录读取完成：{records.Count} 条。");
        return records;
    }

    private List<AttendanceRecord> ReadMonthlySummary(XLWorkbook workbook, YearMonth month, IList<AttendanceException> exceptions)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains("月度汇总", StringComparison.Ordinal))
            ?? workbook.Worksheets.FirstOrDefault(ws => ws.RowsUsed().Any(r => RowText(r).Contains("考勤结果", StringComparison.Ordinal)));
        if (sheet == null)
        {
            return [];
        }

        var used = sheet.RangeUsed();
        if (used == null)
        {
            return [];
        }

        var headerRow = sheet.RowsUsed().FirstOrDefault(r => RowText(r).Contains("姓名", StringComparison.Ordinal)
            && RowText(r).Contains("考勤结果", StringComparison.Ordinal));
        if (headerRow == null)
        {
            return [];
        }

        var nameCol = FindColumn(headerRow, "姓名");
        var deptCol = FindColumn(headerRow, "部门");
        var resultStartCol = FindColumn(headerRow, "考勤结果");
        if (nameCol == 0 || resultStartCol == 0)
        {
            return [];
        }

        var days = DateTime.DaysInMonth(month.Year, month.Month);
        var dataStart = headerRow.RowNumber() + 2;
        var result = new List<AttendanceRecord>();
        foreach (var row in sheet.Rows(dataStart, used.LastRow().RowNumber()))
        {
            var name = AttendanceRuleEngine.CleanName(CellText(row.Cell(nameCol)));
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var department = deptCol > 0 ? CellText(row.Cell(deptCol)) : string.Empty;
            for (var day = 1; day <= days; day++)
            {
                var raw = CellText(row.Cell(resultStartCol + day - 1));
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }

                result.Add(new AttendanceRecord
                {
                    Name = name,
                    Department = department,
                    Date = new DateOnly(month.Year, month.Month, day),
                    RawStatus = raw
                });
            }
        }

        _log($"从 Sheet“{sheet.Name}”识别月度每日结果区：姓名列 {nameCol}，日期起始列 {resultStartCol}。");
        return result;
    }

    private List<AttendanceRecord> ReadDailySummary(XLWorkbook workbook, YearMonth month, IList<AttendanceException> exceptions)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains("每日", StringComparison.Ordinal));
        if (sheet == null)
        {
            exceptions.Add(new AttendanceException { Reason = "考勤报表未找到“月度汇总”或“每日统计”Sheet" });
            return [];
        }

        var headerRow = sheet.RowsUsed().FirstOrDefault(r =>
        {
            var text = RowText(r);
            return text.Contains("姓名", StringComparison.Ordinal) && text.Contains("日期", StringComparison.Ordinal);
        });
        if (headerRow == null)
        {
            exceptions.Add(new AttendanceException { Reason = $"Sheet“{sheet.Name}”未找到姓名/日期表头" });
            return [];
        }

        var headers = BuildHeaderMap(headerRow, headerRow.RowBelow());
        var nameCol = FindHeader(headers, "姓名");
        var deptCol = FindHeader(headers, "部门");
        var dateCol = FindHeader(headers, "日期");
        var shiftCol = FindHeader(headers, "班次");
        if (nameCol == 0 || dateCol == 0)
        {
            return [];
        }

        var result = new List<AttendanceRecord>();
        var used = sheet.RangeUsed();
        if (used == null)
        {
            return result;
        }

        for (var rowNumber = headerRow.RowNumber() + 2; rowNumber <= used.LastRow().RowNumber(); rowNumber++)
        {
            var row = sheet.Row(rowNumber);
            var name = AttendanceRuleEngine.CleanName(CellText(row.Cell(nameCol)));
            var date = TryGetDate(row.Cell(dateCol));
            if (string.IsNullOrWhiteSpace(name) || date == null || date.Value.Year != month.Year || date.Value.Month != month.Month)
            {
                continue;
            }

            var rawParts = new List<string>();
            if (shiftCol > 0)
            {
                rawParts.Add(CellText(row.Cell(shiftCol)));
            }

            foreach (var pair in headers.Where(h => h.Value.Contains("打卡结果", StringComparison.Ordinal)
                         || h.Value.Contains("关联的审批单", StringComparison.Ordinal)
                         || h.Value.Contains("请假", StringComparison.Ordinal)
                         || h.Value.Contains("迟到", StringComparison.Ordinal)
                         || h.Value.Contains("早退", StringComparison.Ordinal)
                         || h.Value.Contains("缺卡", StringComparison.Ordinal)
                         || h.Value.Contains("旷工", StringComparison.Ordinal)))
            {
                rawParts.Add(CellText(row.Cell(pair.Key)));
            }

            var raw = string.Join(" ", rawParts.Where(s => !string.IsNullOrWhiteSpace(s)));
            result.Add(new AttendanceRecord
            {
                Name = name,
                Department = deptCol > 0 ? CellText(row.Cell(deptCol)) : string.Empty,
                Date = date.Value,
                RawStatus = raw
            });
        }

        return result;
    }

    private List<AttendanceRecord> ReadWithExcelDataReader(string filePath, YearMonth month, IList<AttendanceException> exceptions)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        DataSet dataSet;
        try
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                UseColumnDataType = false,
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = false
                }
            });
        }
        catch (Exception ex) when (ShouldTryMarkupFallback(ex))
        {
            _log($"ExcelDataReader 兼容读取失败（{ex.Message}），继续尝试 HTML/XML 表格兼容读取。");
            dataSet = ReadMarkupWorkbook(filePath);
        }

        var records = ReadMonthlySummary(dataSet, month);
        if (records.Count == 0)
        {
            _log("兼容模式未在“月度汇总”识别到每日结果区，改用“每日统计”。");
            records = ReadDailySummary(dataSet, month, exceptions);
        }

        return records;
    }

    private static DataSet ReadMarkupWorkbook(string filePath)
    {
        var bytes = File.ReadAllBytes(filePath);
        var text = DecodeText(bytes);
        var trimmed = text.TrimStart();
        if (trimmed.StartsWith("<", StringComparison.Ordinal))
        {
            if (trimmed.StartsWith("<!DOCTYPE html", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("<html", StringComparison.OrdinalIgnoreCase))
            {
                return ReadHtmlTables(text);
            }

            return ReadXmlSpreadsheet(text);
        }

        throw new InvalidDataException("考勤报表不是可识别的 xlsx、xls、HTML 表格或 XML Spreadsheet 文件。");
    }

    private static DataSet ReadXmlSpreadsheet(string text)
    {
        var document = XDocument.Parse(text, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        var dataSet = new DataSet();
        var worksheets = document.Descendants().Where(e => e.Name.LocalName == "Worksheet").ToList();
        foreach (var worksheet in worksheets)
        {
            var name = worksheet.Attributes().FirstOrDefault(a => a.Name.LocalName == "Name")?.Value
                ?? $"Sheet{dataSet.Tables.Count + 1}";
            var rows = worksheet.Descendants().Where(e => e.Name.LocalName == "Table")
                .Elements().Where(e => e.Name.LocalName == "Row")
                .Select(ReadXmlSpreadsheetRow)
                .ToList();
            AddTable(dataSet, name, rows);
        }

        if (dataSet.Tables.Count == 0)
        {
            throw new InvalidDataException("XML Spreadsheet 中未找到 Worksheet/Table 数据。");
        }

        return dataSet;
    }

    private static List<string> ReadXmlSpreadsheetRow(XElement row)
    {
        var values = new List<string>();
        var currentIndex = 1;
        foreach (var cell in row.Elements().Where(e => e.Name.LocalName == "Cell"))
        {
            var explicitIndex = cell.Attributes().FirstOrDefault(a => a.Name.LocalName == "Index")?.Value;
            if (int.TryParse(explicitIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index > currentIndex)
            {
                while (currentIndex < index)
                {
                    values.Add(string.Empty);
                    currentIndex++;
                }
            }

            var data = cell.Descendants().FirstOrDefault(e => e.Name.LocalName == "Data");
            values.Add(NormalizeCellText(data?.Value ?? cell.Value));
            currentIndex++;
        }

        return values;
    }

    private static DataSet ReadHtmlTables(string text)
    {
        var dataSet = new DataSet();
        var tableMatches = Regex.Matches(text, @"<table\b[^>]*>(?<table>.*?)</table>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        var tableIndex = 1;
        foreach (Match tableMatch in tableMatches)
        {
            var rows = new List<List<string>>();
            foreach (Match rowMatch in Regex.Matches(tableMatch.Groups["table"].Value, @"<tr\b[^>]*>(?<row>.*?)</tr>", RegexOptions.IgnoreCase | RegexOptions.Singleline))
            {
                var row = Regex.Matches(rowMatch.Groups["row"].Value, @"<t[dh]\b[^>]*>(?<cell>.*?)</t[dh]>", RegexOptions.IgnoreCase | RegexOptions.Singleline)
                    .Cast<Match>()
                    .Select(match => NormalizeCellText(WebUtility.HtmlDecode(Regex.Replace(match.Groups["cell"].Value, "<.*?>", string.Empty, RegexOptions.Singleline))))
                    .ToList();
                if (row.Any(cell => !string.IsNullOrWhiteSpace(cell)))
                {
                    rows.Add(row);
                }
            }

            AddTable(dataSet, $"Sheet{tableIndex++}", rows);
        }

        if (dataSet.Tables.Count == 0)
        {
            throw new InvalidDataException("HTML 文件中未找到 table 数据。");
        }

        return dataSet;
    }

    private static void AddTable(DataSet dataSet, string name, IReadOnlyList<IReadOnlyList<string>> rows)
    {
        var maxColumns = rows.Count == 0 ? 0 : rows.Max(row => row.Count);
        var table = new DataTable(GetUniqueTableName(dataSet, string.IsNullOrWhiteSpace(name) ? $"Sheet{dataSet.Tables.Count + 1}" : name));
        for (var column = 0; column < maxColumns; column++)
        {
            table.Columns.Add($"Column{column + 1}", typeof(string));
        }

        foreach (var row in rows)
        {
            var dataRow = table.NewRow();
            for (var column = 0; column < row.Count; column++)
            {
                dataRow[column] = row[column];
            }

            table.Rows.Add(dataRow);
        }

        dataSet.Tables.Add(table);
    }

    private static string GetUniqueTableName(DataSet dataSet, string desiredName)
    {
        var name = desiredName;
        var suffix = 1;
        while (dataSet.Tables.Contains(name))
        {
            name = $"{desiredName}_{suffix++}";
        }

        return name;
    }

    private static string DecodeText(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
        {
            return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
        {
            return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        }

        var utf8 = Encoding.UTF8.GetString(bytes);
        return utf8.Contains('\uFFFD', StringComparison.Ordinal)
            ? Encoding.GetEncoding("GB18030").GetString(bytes)
            : utf8;
    }

    private static string NormalizeCellText(string value)
    {
        return value.Replace("\r", " ").Replace("\n", " ").Trim();
    }

    private List<AttendanceRecord> ReadMonthlySummary(DataSet workbook, YearMonth month)
    {
        var sheet = workbook.Tables.Cast<DataTable>().FirstOrDefault(table => table.TableName.Contains("月度汇总", StringComparison.Ordinal))
            ?? workbook.Tables.Cast<DataTable>().FirstOrDefault(table => table.Rows.Cast<DataRow>().Any(row => RowText(row).Contains("考勤结果", StringComparison.Ordinal)));
        if (sheet == null)
        {
            return [];
        }

        var headerMatch = sheet.Rows.Cast<DataRow>().Select((row, index) => (row, index))
            .FirstOrDefault(item => RowText(item.row).Contains("姓名", StringComparison.Ordinal)
                && RowText(item.row).Contains("考勤结果", StringComparison.Ordinal));
        if (headerMatch.row == null)
        {
            return [];
        }

        var headerIndex = headerMatch.index;
        var headerRow = sheet.Rows[headerIndex];
        var nameCol = FindColumn(headerRow, "姓名");
        var deptCol = FindColumn(headerRow, "部门");
        var resultStartCol = FindColumn(headerRow, "考勤结果");
        if (nameCol < 0 || resultStartCol < 0)
        {
            return [];
        }

        var days = DateTime.DaysInMonth(month.Year, month.Month);
        var result = new List<AttendanceRecord>();
        for (var rowIndex = headerIndex + 2; rowIndex < sheet.Rows.Count; rowIndex++)
        {
            var row = sheet.Rows[rowIndex];
            var name = AttendanceRuleEngine.CleanName(CellText(row, nameCol));
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var department = deptCol >= 0 ? CellText(row, deptCol) : string.Empty;
            for (var day = 1; day <= days; day++)
            {
                var raw = CellText(row, resultStartCol + day - 1);
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }

                result.Add(new AttendanceRecord
                {
                    Name = name,
                    Department = department,
                    Date = new DateOnly(month.Year, month.Month, day),
                    RawStatus = raw
                });
            }
        }

        _log($"兼容模式从 Sheet“{sheet.TableName}”识别月度每日结果区：姓名列 {nameCol + 1}，日期起始列 {resultStartCol + 1}。");
        return result;
    }

    private List<AttendanceRecord> ReadDailySummary(DataSet workbook, YearMonth month, IList<AttendanceException> exceptions)
    {
        var sheet = workbook.Tables.Cast<DataTable>().FirstOrDefault(table => table.TableName.Contains("每日", StringComparison.Ordinal))
            ?? workbook.Tables.Cast<DataTable>().FirstOrDefault(table => table.Rows.Cast<DataRow>().Any(row =>
            {
                var text = RowText(row);
                return text.Contains("姓名", StringComparison.Ordinal) && text.Contains("日期", StringComparison.Ordinal);
            }));
        if (sheet == null)
        {
            exceptions.Add(new AttendanceException { Reason = "考勤报表未找到“月度汇总”或“每日统计”Sheet" });
            return [];
        }

        var headerIndex = -1;
        for (var index = 0; index < sheet.Rows.Count; index++)
        {
            var text = RowText(sheet.Rows[index]);
            if (text.Contains("姓名", StringComparison.Ordinal) && text.Contains("日期", StringComparison.Ordinal))
            {
                headerIndex = index;
                break;
            }
        }

        if (headerIndex < 0)
        {
            exceptions.Add(new AttendanceException { Reason = $"Sheet“{sheet.TableName}”未找到姓名/日期表头" });
            return [];
        }

        var headers = BuildHeaderMap(sheet.Rows[headerIndex], headerIndex + 1 < sheet.Rows.Count ? sheet.Rows[headerIndex + 1] : null);
        var nameCol = FindHeaderIndex(headers, "姓名");
        var deptCol = FindHeaderIndex(headers, "部门");
        var dateCol = FindHeaderIndex(headers, "日期");
        var shiftCol = FindHeaderIndex(headers, "班次");
        if (nameCol < 0 || dateCol < 0)
        {
            return [];
        }

        var result = new List<AttendanceRecord>();
        for (var rowIndex = headerIndex + 2; rowIndex < sheet.Rows.Count; rowIndex++)
        {
            var row = sheet.Rows[rowIndex];
            var name = AttendanceRuleEngine.CleanName(CellText(row, nameCol));
            var date = TryGetDate(row[dateCol]);
            if (string.IsNullOrWhiteSpace(name) || date == null || date.Value.Year != month.Year || date.Value.Month != month.Month)
            {
                continue;
            }

            var rawParts = new List<string>();
            if (shiftCol >= 0)
            {
                rawParts.Add(CellText(row, shiftCol));
            }

            foreach (var pair in headers.Where(h => h.Value.Contains("打卡结果", StringComparison.Ordinal)
                         || h.Value.Contains("关联的审批单", StringComparison.Ordinal)
                         || h.Value.Contains("请假", StringComparison.Ordinal)
                         || h.Value.Contains("迟到", StringComparison.Ordinal)
                         || h.Value.Contains("早退", StringComparison.Ordinal)
                         || h.Value.Contains("缺卡", StringComparison.Ordinal)
                         || h.Value.Contains("旷工", StringComparison.Ordinal)))
            {
                rawParts.Add(CellText(row, pair.Key));
            }

            result.Add(new AttendanceRecord
            {
                Name = name,
                Department = deptCol >= 0 ? CellText(row, deptCol) : string.Empty,
                Date = date.Value,
                RawStatus = string.Join(" ", rawParts.Where(s => !string.IsNullOrWhiteSpace(s)))
            });
        }

        return result;
    }

    private static Dictionary<int, string> BuildHeaderMap(DataRow row1, DataRow? row2 = null)
    {
        var map = new Dictionary<int, string>();
        for (var col = 0; col < row1.Table.Columns.Count; col++)
        {
            var text = string.Join(" ", new[] { CellText(row1, col), row2 == null ? string.Empty : CellText(row2, col) }
                .Where(s => !string.IsNullOrWhiteSpace(s))).Trim();
            if (!string.IsNullOrWhiteSpace(text))
            {
                map[col] = text;
            }
        }

        return map;
    }

    private static int FindHeaderIndex(Dictionary<int, string> headers, params string[] names)
    {
        foreach (var name in names)
        {
            var exact = headers.FirstOrDefault(h => h.Value.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (exact.Value != null)
            {
                return exact.Key;
            }

            var contains = headers.FirstOrDefault(h => h.Value.Contains(name, StringComparison.OrdinalIgnoreCase));
            if (contains.Value != null)
            {
                return contains.Key;
            }
        }

        return -1;
    }

    private static int FindColumn(DataRow row, string text)
    {
        for (var index = 0; index < row.Table.Columns.Count; index++)
        {
            var cellText = CellText(row, index);
            if (cellText.Equals(text, StringComparison.OrdinalIgnoreCase)
                || cellText.Contains(text, StringComparison.OrdinalIgnoreCase))
            {
                return index;
            }
        }

        return -1;
    }

    private static DateOnly? TryGetDate(object? value)
    {
        if (value == null || value == DBNull.Value)
        {
            return null;
        }

        if (value is DateTime dateTime)
        {
            return DateOnly.FromDateTime(dateTime);
        }

        if (value is double or float or decimal or int or long)
        {
            try
            {
                return DateOnly.FromDateTime(DateTime.FromOADate(Convert.ToDouble(value, CultureInfo.InvariantCulture)));
            }
            catch
            {
                return null;
            }
        }

        var text = CellText(value);
        if (DateOnly.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out var date)
            || DateOnly.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
        {
            return date;
        }

        if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTime)
            || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime))
        {
            return DateOnly.FromDateTime(dateTime);
        }

        return null;
    }

    private static bool ShouldTryMarkupFallback(Exception exception)
    {
        return exception is InvalidDataException
            || exception.Message.Contains("signature", StringComparison.OrdinalIgnoreCase)
            || exception.Message.Contains("corrupt", StringComparison.OrdinalIgnoreCase)
            || exception.Message.Contains("损坏", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ShouldTryTabularFallback(Exception exception)
    {
        return exception is InvalidDataException
            || exception.Message.Contains("corrupt", StringComparison.OrdinalIgnoreCase)
            || exception.Message.Contains("损坏", StringComparison.OrdinalIgnoreCase)
            || exception.GetType().FullName?.Contains("OpenXml", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static string RowText(DataRow row)
    {
        return string.Join(" ", row.ItemArray.Select(CellText).Where(s => !string.IsNullOrWhiteSpace(s)));
    }

    private static string CellText(DataRow row, int column)
    {
        return column >= 0 && column < row.Table.Columns.Count ? CellText(row[column]) : string.Empty;
    }

    private static string CellText(object? value)
    {
        if (value == null || value == DBNull.Value)
        {
            return string.Empty;
        }

        return Convert.ToString(value, CultureInfo.CurrentCulture)?.Replace("\r", " ").Replace("\n", " ").Trim() ?? string.Empty;
    }

    internal static Dictionary<int, string> BuildHeaderMap(IXLRow row1, IXLRow? row2 = null)
    {
        var used = row1.Worksheet.RangeUsed();
        var lastCol = used?.LastColumn().ColumnNumber() ?? row1.LastCellUsed()?.Address.ColumnNumber ?? 1;
        var map = new Dictionary<int, string>();
        for (var col = 1; col <= lastCol; col++)
        {
            var text = string.Join(" ", new[] { CellText(row1.Cell(col)), row2 == null ? string.Empty : CellText(row2.Cell(col)) }
                .Where(s => !string.IsNullOrWhiteSpace(s))).Trim();
            if (!string.IsNullOrWhiteSpace(text))
            {
                map[col] = text;
            }
        }

        return map;
    }

    internal static int FindHeader(Dictionary<int, string> headers, params string[] names)
    {
        foreach (var name in names)
        {
            var exact = headers.FirstOrDefault(h => h.Value.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (exact.Key > 0)
            {
                return exact.Key;
            }

            var contains = headers.FirstOrDefault(h => h.Value.Contains(name, StringComparison.OrdinalIgnoreCase));
            if (contains.Key > 0)
            {
                return contains.Key;
            }
        }

        return 0;
    }

    internal static int FindColumn(IXLRow row, string text)
    {
        return row.CellsUsed().FirstOrDefault(c => CellText(c).Equals(text, StringComparison.OrdinalIgnoreCase)
            || CellText(c).Contains(text, StringComparison.OrdinalIgnoreCase))?.Address.ColumnNumber ?? 0;
    }

    internal static DateOnly? TryGetDate(IXLCell cell)
    {
        if (cell.IsEmpty())
        {
            return null;
        }

        if (cell.DataType == XLDataType.DateTime)
        {
            return DateOnly.FromDateTime(cell.GetDateTime());
        }

        if (cell.DataType == XLDataType.Number)
        {
            try
            {
                return DateOnly.FromDateTime(DateTime.FromOADate(cell.GetDouble()));
            }
            catch
            {
                return null;
            }
        }

        var text = CellText(cell);
        if (DateOnly.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out var date)
            || DateOnly.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
        {
            return date;
        }

        if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out var dateTime)
            || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime))
        {
            return DateOnly.FromDateTime(dateTime);
        }

        return null;
    }

    internal static string CellText(IXLCell cell)
    {
        if (cell.IsEmpty())
        {
            return string.Empty;
        }

        return cell.GetFormattedString().Replace("\r", " ").Replace("\n", " ").Trim();
    }

    private static string RowText(IXLRow row)
    {
        return string.Join(" ", row.CellsUsed().Select(CellText));
    }
}

public readonly record struct YearMonth(int Year, int Month)
{
    public static YearMonth Parse(string value)
    {
        if (!DateTime.TryParseExact(value.Trim(), "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            throw new FormatException("月份格式必须为 yyyy-MM，例如 2026-05。");
        }

        return new YearMonth(date.Year, date.Month);
    }

    public override string ToString() => $"{Year:D4}-{Month:D2}";
}
