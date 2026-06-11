using System.Globalization;
using ClosedXML.Excel;
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
        using var workbook = new XLWorkbook(filePath);
        _log($"读取考勤报表：{filePath}");

        var records = ReadMonthlySummary(workbook, month, exceptions);
        if (records.Count == 0)
        {
            _log("未在“月度汇总”识别到每日结果区，改用“每日统计”。");
            records = ReadDailySummary(workbook, month, exceptions);
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
