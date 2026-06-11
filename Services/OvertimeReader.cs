using System.Globalization;
using ClosedXML.Excel;
using AttendanceGenerator.Models;

namespace AttendanceGenerator.Services;

public sealed class OvertimeReader
{
    private readonly Action<string> _log;

    public OvertimeReader(Action<string>? log = null)
    {
        _log = log ?? (_ => { });
    }

    public IReadOnlyList<OvertimeRecord> Read(string filePath, YearMonth month, IList<AttendanceException> exceptions)
    {
        using var workbook = new XLWorkbook(filePath);
        _log($"读取加班表：{filePath}");

        var records = new List<OvertimeRecord>();
        foreach (var sheet in workbook.Worksheets)
        {
            var headerRow = sheet.RowsUsed().Take(20).LastOrDefault(r =>
            {
                var text = string.Join(" ", r.CellsUsed().Select(AttendanceReader.CellText));
                return text.Contains("发起人姓名", StringComparison.Ordinal) && text.Contains("开始时间", StringComparison.Ordinal)
                    || text.Contains("加班人", StringComparison.Ordinal) && text.Contains("加班时间", StringComparison.Ordinal);
            });
            if (headerRow == null)
            {
                continue;
            }

            var headers = AttendanceReader.BuildHeaderMap(headerRow);
            var nameCol = AttendanceReader.FindHeader(headers, "发起人姓名", "加班人", "姓名");
            var dateCol = AttendanceReader.FindHeader(headers, "开始时间", "加班时间");
            var hoursCol = AttendanceReader.FindHeader(headers, "时长");
            var typeCol = AttendanceReader.FindHeader(headers, "加班类型", "是否法定假日", "加班补偿");
            var resultCol = AttendanceReader.FindHeader(headers, "审批结果");

            if (nameCol == 0 || dateCol == 0 || hoursCol == 0)
            {
                exceptions.Add(new AttendanceException
                {
                    Reason = $"Sheet“{sheet.Name}”疑似加班明细，但缺少姓名/日期/时长列"
                });
                continue;
            }

            var used = sheet.RangeUsed();
            if (used == null)
            {
                continue;
            }

            for (var rowNumber = headerRow.RowNumber() + 1; rowNumber <= used.LastRow().RowNumber(); rowNumber++)
            {
                var row = sheet.Row(rowNumber);
                var approvalResult = resultCol > 0 ? AttendanceReader.CellText(row.Cell(resultCol)) : string.Empty;
                if (approvalResult.Contains("拒绝", StringComparison.Ordinal) || approvalResult.Contains("撤销", StringComparison.Ordinal))
                {
                    continue;
                }

                var name = AttendanceRuleEngine.CleanName(AttendanceReader.CellText(row.Cell(nameCol)));
                var date = AttendanceReader.TryGetDate(row.Cell(dateCol));
                var hours = TryGetDecimal(row.Cell(hoursCol));
                if (string.IsNullOrWhiteSpace(name) || date == null || hours <= 0)
                {
                    continue;
                }

                if (date.Value.Year != month.Year || date.Value.Month != month.Month)
                {
                    continue;
                }

                var type = typeCol > 0 ? AttendanceReader.CellText(row.Cell(typeCol)) : string.Empty;
                records.Add(new OvertimeRecord
                {
                    Name = name,
                    Date = date.Value,
                    Hours = hours,
                    OvertimeType = type,
                    RawText = string.Join(" ", row.CellsUsed().Select(AttendanceReader.CellText))
                });
            }

            _log($"从 Sheet“{sheet.Name}”读取加班明细。");
        }

        _log($"加班记录读取完成：{records.Count} 条。");
        return records;
    }

    private static decimal TryGetDecimal(IXLCell cell)
    {
        if (cell.IsEmpty())
        {
            return 0;
        }

        if (cell.DataType == XLDataType.Number)
        {
            return Convert.ToDecimal(cell.GetDouble(), CultureInfo.InvariantCulture);
        }

        var text = AttendanceReader.CellText(cell);
        return decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out var value)
            || decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value)
            ? value
            : 0;
    }
}
