using ClosedXML.Excel;

namespace AttendanceGenerator.Services;

internal static class CommandLineRunner
{
    public static int Run(string[] args, Action<string> log, Action<string> error)
    {
        try
        {
            var command = args[0].Trim().ToLowerInvariant();
            return command switch
            {
                "generate" => Generate(args, log),
                "compare" => Compare(args, log),
                "help" or "--help" or "-h" => Usage(log),
                _ => Unknown(command, error)
            };
        }
        catch (Exception ex)
        {
            error(ex.ToString());
            return 1;
        }
    }

    private static int Generate(string[] args, Action<string> log)
    {
        if (args.Length < 5)
        {
            return Usage(log);
        }

        var holidays = args.Length >= 6 ? ParseDates(args[5]) : [];
        var templatePath = args.Length >= 7 ? args[6] : null;
        var service = new GenerationService(log);
        var output = service.Generate(args[1], args[2], args[3], args[4], holidays, templatePath);
        log(output);
        return 0;
    }

    private static int Compare(string[] args, Action<string> log)
    {
        if (args.Length < 3)
        {
            return Usage(log);
        }

        using var actual = new XLWorkbook(args[1]);
        using var expected = new XLWorkbook(args[2]);
        var actualSheet = FindAttendanceSheet(actual);
        var expectedSheet = FindAttendanceSheet(expected, actualSheet?.Name);
        if (actualSheet == null || expectedSheet == null)
        {
            throw new InvalidOperationException("未能在两个文件中找到考勤统计表 Sheet。");
        }

        var actualUsed = actualSheet.RangeUsed();
        var expectedUsed = expectedSheet.RangeUsed();
        var maxRow = Math.Max(actualUsed?.LastRow().RowNumber() ?? 1, expectedUsed?.LastRow().RowNumber() ?? 1);
        var maxCol = Math.Max(actualUsed?.LastColumn().ColumnNumber() ?? 1, expectedUsed?.LastColumn().ColumnNumber() ?? 1);
        var differences = 0;
        for (var row = 1; row <= maxRow; row++)
        {
            for (var col = 1; col <= maxCol; col++)
            {
                var actualText = Normalize(AttendanceReader.CellText(actualSheet.Cell(row, col)));
                var expectedText = Normalize(AttendanceReader.CellText(expectedSheet.Cell(row, col)));
                if (!ValuesEqual(actualText, expectedText))
                {
                    differences++;
                    if (differences <= 200)
                    {
                        log($"{actualSheet.Cell(row, col).Address}: actual=[{actualText}] expected=[{expectedText}]");
                    }
                }
            }
        }

        log($"差异单元格数：{differences}");
        return differences == 0 ? 0 : 2;
    }

    private static int Usage(Action<string> log)
    {
        log("用法：");
        log("  generate <考勤报表.xlsx> <加班表.xlsx> <yyyy-MM> <输出目录> [法定日期逗号分隔] [模板路径]");
        log("  compare <生成文件.xlsx> <手工原件.xlsx>");
        log("示例：");
        log("  generate 考勤报表.xlsx 加班表.xlsx 2026-05 ./out 2026-05-01,2026-05-02 Config/考勤统计模板.xlsx");
        log("  compare ./out/2026年05月考勤统计表_生成版.xlsx 5月手工做的.xlsx");
        return 1;
    }

    private static int Unknown(string command, Action<string> error)
    {
        error($"未知命令：{command}");
        return 1;
    }

    private static List<DateOnly> ParseDates(string value)
    {
        return value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(item => DateOnly.Parse(item, System.Globalization.CultureInfo.InvariantCulture))
            .ToList();
    }

    private static IXLWorksheet? FindAttendanceSheet(XLWorkbook workbook, string? preferredName = null)
    {
        if (!string.IsNullOrWhiteSpace(preferredName))
        {
            var exact = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals(preferredName, StringComparison.OrdinalIgnoreCase));
            if (exact != null)
            {
                return exact;
            }

            var month = System.Text.RegularExpressions.Regex.Match(preferredName, @"(?<month>\d{1,2})月");
            if (month.Success)
            {
                var sameMonth = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains($"{month.Groups["month"].Value}月", StringComparison.Ordinal)
                    && ws.Name.Contains("考勤统计", StringComparison.Ordinal));
                if (sameMonth != null)
                {
                    return sameMonth;
                }
            }
        }

        return workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains("考勤统计", StringComparison.Ordinal));
    }

    private static string Normalize(string value)
    {
        return value.Trim();
    }

    private static bool ValuesEqual(string actual, string expected)
    {
        if (actual == expected)
        {
            return true;
        }

        return decimal.TryParse(actual, out var actualNumber)
            && decimal.TryParse(expected, out var expectedNumber)
            && Math.Abs(actualNumber - expectedNumber) < 0.005m;
    }
}
