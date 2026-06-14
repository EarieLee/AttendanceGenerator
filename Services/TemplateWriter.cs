using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using AttendanceGenerator.Models;

namespace AttendanceGenerator.Services;

public sealed class TemplateWriter
{
    private readonly Action<string> _log;
    private readonly WorkCalendarSettings _workCalendar;

    private sealed record EmployeeTemplateRow(string Name, string Department, string Remark);
    private sealed record OvertimeTemplateRow(string Name, decimal Standard);

    public TemplateWriter(Action<string>? log = null, string? configDirectory = null)
    {
        _log = log ?? (_ => { });
        _workCalendar = LoadWorkCalendar(configDirectory);
    }

    public void Generate(
        string? templatePath,
        string outputPath,
        YearMonth month,
        IReadOnlySet<DateOnly> selectedHolidays,
        IReadOnlyList<AttendanceRecord> attendanceRecords,
        IReadOnlyList<OvertimeRecord> overtimeRecords,
        IReadOnlyList<AttendanceException> exceptions)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        using (var workbook = CreateWorkbook(templatePath, outputPath, month, selectedHolidays, attendanceRecords, overtimeRecords))
        {
            WriteAttendance(workbook, month, selectedHolidays, attendanceRecords, exceptions);
            WriteOvertime(workbook, month, overtimeRecords);
            RewriteNightShiftSheet(workbook, month);
            WriteExceptionSheet(workbook, exceptions);
            workbook.CalculateMode = XLCalculateMode.Auto;
            workbook.RecalculateAllFormulas();
            workbook.Save();
        }

        _log("Excel 生成完成。");
    }

    private XLWorkbook CreateWorkbook(
        string? templatePath,
        string outputPath,
        YearMonth month,
        IReadOnlySet<DateOnly> selectedHolidays,
        IReadOnlyList<AttendanceRecord> attendanceRecords,
        IReadOnlyList<OvertimeRecord> overtimeRecords)
    {
        var workbook = new XLWorkbook();
        BuildAttendanceTemplate(workbook, month, selectedHolidays, attendanceRecords);
        BuildOvertimeTemplate(workbook, month, overtimeRecords);
        BuildNightShiftTemplate(workbook, month, attendanceRecords);
        workbook.SaveAs(outputPath);
        _log($"已用代码生成模板到输出文件：{outputPath}");
        return workbook;
    }

    private void BuildAttendanceTemplate(
        XLWorkbook workbook,
        YearMonth month,
        IReadOnlySet<DateOnly> selectedHolidays,
        IReadOnlyList<AttendanceRecord> attendanceRecords)
    {
        var sheet = workbook.Worksheets.Add($"{month.Month}月考勤统计表 ");
        sheet.Range("A1:AU1").Merge();
        sheet.Cell("A1").Value = $"                   {month.Year}年{month.Month}月 Microtest考勤统计表      {BuildAttendanceTitleSummary(month, selectedHolidays)}";
        sheet.Cell("A2").Value = "            ";
        sheet.Range("A2:C3").Merge();

        var statHeaders = new[]
        {
            "合计", "出勤（天）", "休息", "法定（天）", "请假（天）", "年假（天）", "调休（天）",
            "陪产假（天）", "婚假（天）", "产假（天）", "其他", "扣工资合计（天数）", "备注"
        };
        for (var day = 1; day <= 31; day++)
        {
            var col = 3 + day;
            sheet.Cell(2, col).Value = new DateOnly(month.Year, month.Month, Math.Min(day, DateTime.DaysInMonth(month.Year, month.Month))).ToDateTime(TimeOnly.MinValue);
            sheet.Cell(3, col).Value = day <= DateTime.DaysInMonth(month.Year, month.Month)
                ? ToChineseWeekday(new DateOnly(month.Year, month.Month, day).DayOfWeek)
                : string.Empty;
        }

        for (var index = 0; index < statHeaders.Length; index++)
        {
            var col = 35 + index;
            sheet.Cell(2, col).Value = statHeaders[index];
            sheet.Range(2, col, 3, col).Merge();
        }

        var employees = GetEmployeeTemplateRows(month, attendanceRecords);

        var row = 4;
        foreach (var employee in employees)
        {
            sheet.Cell(row, 1).FormulaA1 = $"ROW()-3";
            sheet.Cell(row, 2).Value = employee.Name;
            sheet.Cell(row, 3).Value = employee.Department;
            sheet.Cell(row, 47).Value = employee.Remark;
            sheet.Row(row).Height = 34;
            row++;
        }

        var totalRow = row;
        for (var col = 35; col <= 46; col++)
        {
            sheet.Cell(totalRow, col).FormulaA1 = $"SUM({sheet.Cell(4, col).Address.ToStringRelative()}:"
                + $"{sheet.Cell(totalRow - 1, col).Address.ToStringRelative()})";
        }

        sheet.Cell(totalRow + 3, 1).Value = "编制： 肖玲玲                          审查：                                        审核：";
        sheet.Cell(totalRow + 6, 21).Value = "事假";
        sheet.Cell(totalRow + 6, 22).Value = "调休";
        sheet.Cell(totalRow + 6, 23).Value = "年假";
        ApplyAttendanceStyle(sheet, totalRow + 6);
    }


    private static List<EmployeeTemplateRow> GetEmployeeTemplateRows(YearMonth month, IReadOnlyList<AttendanceRecord> attendanceRecords)
    {
        return attendanceRecords
            .Select((record, index) => new { Record = record, Index = index })
            .GroupBy(item => AttendanceRuleEngine.CleanName(item.Record.Name))
            .OrderBy(g => g.Min(item => item.Index))
            .Select(g => g.First().Record)
            .GroupBy(r => AttendanceRuleEngine.CleanName(r.Name))
            .Select(g => new EmployeeTemplateRow(
                g.Key,
                g.Select(r => r.Department).FirstOrDefault(d => !string.IsNullOrWhiteSpace(d)) ?? string.Empty,
                string.Empty))
            .Where(e => !string.IsNullOrWhiteSpace(e.Name))
            .ToList();
    }

    private static void ApplyAttendanceStyle(IXLWorksheet sheet, int lastRow)
    {
        sheet.Row(1).Height = 31;
        sheet.Row(2).Height = 23;
        sheet.Row(3).Height = 39;
        SetSavedColumnWidth(sheet, 1, 3.54166666666667);
        SetSavedColumnWidth(sheet, 2, 7.60833333333333);
        SetSavedColumnWidth(sheet, 3, 13);
        for (var col = 4; col <= 34; col++)
        {
            SetSavedColumnWidth(sheet, col, col is 4 or 5 ? 5.25 : 13);
        }

        var statWidths = new Dictionary<int, double>
        {
            [35] = 7.34166666666667,
            [36] = 7.65833333333333,
            [37] = 6.71666666666667,
            [38] = 6.8,
            [39] = 8.275,
            [40] = 6.80833333333333,
            [41] = 6.1,
            [42] = 13,
            [43] = 6.51666666666667,
            [44] = 13,
            [45] = 6.525,
            [46] = 10.4083333333333,
            [47] = 18.5
        };
        foreach (var (col, width) in statWidths)
        {
            SetSavedColumnWidth(sheet, col, width);
        }

        var all = sheet.Range(1, 1, lastRow, 47);
        all.Style.Font.FontName = "宋体";
        all.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        all.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        all.Style.Alignment.WrapText = true;
        all.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        all.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        sheet.Cell(1, 1).Style.Font.Bold = false;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        sheet.Range(2, 4, 2, 34).Style.Fill.BackgroundColor = XLColor.FromHtml("#33CCCC");
        sheet.Range(2, 4, 2, 34).Style.DateFormat.Format = "d\"A\"";
        sheet.Range(3, 4, 3, 34).Style.Fill.BackgroundColor = XLColor.White;
        sheet.Range(2, 35, 3, 47).Style.Fill.BackgroundColor = XLColor.NoColor;
        sheet.SheetView.FreezeRows(3);
        sheet.SheetView.FreezeColumns(3);
    }

    private static void SetSavedColumnWidth(IXLWorksheet sheet, int column, double targetWidth)
    {
        sheet.Column(column).Width = Math.Max(0, targetWidth - 0.710625);
    }

    private void BuildOvertimeTemplate(XLWorkbook workbook, YearMonth month, IReadOnlyList<OvertimeRecord> overtimeRecords)
    {
        var sheet = workbook.Worksheets.Add("加班补贴表 ");
        sheet.Range("A1:AE1").Merge();
        sheet.Cell("A1").Value = $"{month.Year}年{month.Month}月加班统计表";
        sheet.Cell("A2").Value = "        日期\n\n姓名\n";
        for (var day = 1; day <= 31; day++)
        {
            sheet.Cell(2, day + 1).Value = new DateOnly(month.Year, month.Month, Math.Min(day, DateTime.DaysInMonth(month.Year, month.Month))).ToDateTime(TimeOnly.MinValue);
        }

        sheet.Cell(2, 33).Value = "加班时间";
        sheet.Cell(2, 34).Value = "标准";
        sheet.Cell(2, 35).Value = "加班补贴";
        sheet.Cell(2, 38).Value = "姓名";
        sheet.Cell(2, 39).Value = "平时加班";
        sheet.Cell(2, 41).Value = "姓名";
        sheet.Cell(2, 42).Value = "周末加班";
        sheet.Cell(2, 44).Value = "姓名";
        sheet.Cell(2, 45).Value = "法定加班";
        sheet.Cell(2, 48).Value = "平时验证";
        sheet.Cell(2, 50).Value = "周末验证";
        sheet.Cell(2, 52).Value = "法定验证";

        var rows = GetOvertimeTemplateRows(month, overtimeRecords);
        var row = 3;
        foreach (var item in rows)
        {
            if (item.Name == "合计")
            {
                AddOvertimeTotalRow(sheet, row++);
                continue;
            }

            AddOvertimeRow(sheet, row++, item.Name, item.Standard, item.Name.EndsWith('2'));
        }

        sheet.Cell(row, 44).FormulaA1 = $"AM{row - 1}+AP{row - 1}+AS{row - 1}";
        sheet.Cell(row, 50).FormulaA1 = $"AV{row - 1}+AX{row - 1}+AZ{row - 1}";
        sheet.Cell(row + 1, 1).Value = "  编制： 肖玲玲                                                                审核：";
        sheet.Cell(row + 1, 44).FormulaA1 = $"AR{row}-AI{row - 1}";
        sheet.Cell(row + 1, 48).FormulaA1 = $"(AV{row - 1}+AX{row - 1}+AZ{row - 1})+AI{row - 1}";
        ApplyOvertimeStyle(sheet, Math.Max(3, row + 1));
    }

    private static void AddOvertimeRow(IXLWorksheet sheet, int row, string name, decimal standard, bool restRow)
    {
        sheet.Cell(row, 1).Value = name;
        sheet.Cell(row, 33).FormulaA1 = $"SUM(B{row}:AF{row})";
        sheet.Cell(row, 34).Value = standard;
        sheet.Cell(row, 35).FormulaA1 = $"AG{row}*AH{row}";
        if (restRow)
        {
            sheet.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
            var baseName = name[..^1];
            sheet.Cell(row, 41).Value = baseName;
            sheet.Cell(row, 42).FormulaA1 = $"AI{row}";
            sheet.Cell(row, 50).FormulaA1 = $"AP{row}-AI{row}";
            sheet.Cell(row, 38).Style.Fill.BackgroundColor = XLColor.Yellow;
            sheet.Cell(row, 41).Style.Fill.BackgroundColor = XLColor.Yellow;
        }
        else
        {
            sheet.Cell(row, 38).Value = name;
            sheet.Cell(row, 39).FormulaA1 = $"AI{row}";
            sheet.Cell(row, 48).FormulaA1 = $"AM{row}-AI{row}";
        }
    }

    private static void AddOvertimeTotalRow(IXLWorksheet sheet, int row)
    {
        sheet.Cell(row, 1).Value = "合计";
        for (var col = 2; col <= 35; col++)
        {
            sheet.Cell(row, col).FormulaA1 = $"SUM({sheet.Cell(3, col).Address.ToStringRelative()}:{sheet.Cell(row - 1, col).Address.ToStringRelative()})";
        }

        sheet.Cell(row, 33).FormulaA1 = $"SUM(B{row}:AF{row})";
        sheet.Cell(row, 34).Value = "/";
        sheet.Cell(row, 38).Value = "合计";
        sheet.Cell(row, 39).FormulaA1 = $"SUM(AM3:AM{row - 1})";
        sheet.Cell(row, 41).Value = "合计";
        sheet.Cell(row, 42).FormulaA1 = $"SUM(AP3:AP{row - 1})";
        sheet.Cell(row, 44).Value = "合计";
        sheet.Cell(row, 45).FormulaA1 = $"SUM(AS3:AS{row - 1})";
        sheet.Cell(row, 48).FormulaA1 = $"SUM(AV3:AV{row - 1})";
        sheet.Cell(row, 50).FormulaA1 = $"SUM(AX3:AX{row - 1})";
        sheet.Cell(row, 52).FormulaA1 = $"SUM(AZ3:AZ{row - 1})";
    }

    private List<OvertimeTemplateRow> GetOvertimeTemplateRows(YearMonth month, IReadOnlyList<OvertimeRecord> overtimeRecords)
    {
        return overtimeRecords
            .Select((record, index) => new { Record = record, Index = index, Name = AttendanceRuleEngine.CleanName(record.Name) })
            .Where(item => !string.IsNullOrWhiteSpace(item.Name))
            .GroupBy(item => item.Name, StringComparer.Ordinal)
            .OrderBy(group => group.Min(item => item.Index))
            .SelectMany(group =>
            {
                var hasWorkdayOvertime = group.Any(item => !IsCalendarRestDay(item.Record.Date));
                var hasRestdayOvertime = group.Any(item => IsCalendarRestDay(item.Record.Date));
                var rows = new List<OvertimeTemplateRow>();
                if (hasWorkdayOvertime || !hasRestdayOvertime)
                {
                    rows.Add(new OvertimeTemplateRow(group.Key, 25));
                }

                if (hasRestdayOvertime)
                {
                    rows.Add(new OvertimeTemplateRow($"{group.Key}2", 40));
                }

                return rows;
            })
            .Append(new OvertimeTemplateRow("合计", 0))
            .ToList();
    }

    private static void ApplyOvertimeStyle(IXLWorksheet sheet, int lastRow)
    {
        sheet.Row(1).Height = 47;
        SetSavedColumnWidth(sheet, 1, 14.3666666666667);
        SetSavedColumnWidth(sheet, 2, 7.34166666666667);
        for (var col = 3; col <= 32; col++)
        {
            SetSavedColumnWidth(sheet, col, 13);
        }
        var widths = new Dictionary<int, double>
        {
            [33] = 10.1583333333333,
            [34] = 8.74166666666667,
            [35] = 15,
            [36] = 9,
            [37] = 13,
            [38] = 14.3666666666667,
            [39] = 11.8916666666667,
            [40] = 9,
            [41] = 14.3666666666667,
            [42] = 11.8833333333333,
            [43] = 9,
            [44] = 13.4333333333333,
            [45] = 10.3833333333333,
            [46] = 9.38333333333333,
            [47] = 9,
            [48] = 12.9583333333333,
            [49] = 9,
            [50] = 10.3833333333333,
            [51] = 9,
            [52] = 10.3833333333333,
            [53] = 9
        };
        foreach (var (col, width) in widths)
        {
            SetSavedColumnWidth(sheet, col, width);
        }

        for (var row = 3; row <= Math.Min(lastRow, 129); row++)
        {
            sheet.Row(row).Height = 44;
        }

        var all = sheet.Range(1, 1, lastRow, 45);
        all.Style.Font.FontName = "宋体";
        all.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        all.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        all.Style.Alignment.WrapText = true;
        all.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        all.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        sheet.Cell(1, 1).Style.Font.Bold = false;
        sheet.Cell(1, 1).Style.Font.FontSize = 22;
        foreach (var address in new[] { "A1", "AF1", "AG1", "AI1", "AL1", "AO1" })
        {
            sheet.Cell(address).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Background1);
        }
        sheet.Range(2, 2, 2, 32).Style.DateFormat.Format = "d";
        sheet.Range(2, 2, 2, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC000");
        sheet.Range(2, 11, 2, 11).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC000");
        sheet.Range(2, 17, 2, 18).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC000");
        sheet.Range(2, 24, 2, 25).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC000");
        sheet.Range(2, 31, 2, 32).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC000");
        foreach (var row in sheet.Rows(3, lastRow))
        {
            var name = AttendanceReader.CellText(row.Cell(1));
            if (name.EndsWith('2'))
            {
                row.Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                row.Cell(38).Style.Fill.BackgroundColor = XLColor.Yellow;
                row.Cell(41).Style.Fill.BackgroundColor = XLColor.Yellow;
            }
        }
        sheet.SheetView.FreezeRows(2);
        sheet.SheetView.FreezeColumns(1);
    }

    private static void BuildNightShiftTemplate(XLWorkbook workbook, YearMonth month, IReadOnlyList<AttendanceRecord> attendanceRecords)
    {
        var sheet = workbook.Worksheets.Add("夜班补贴");
        sheet.Range("A1:K1").Merge();
        sheet.Cell("A1").Value = $"{month.Year}年{month.Month}月中夜班补贴统计表";
        sheet.Cell("A2").Value = "         姓名\n     班次\n";
        var days = DateTime.DaysInMonth(month.Year, month.Month);
        for (var day = 1; day <= days; day++)
        {
            sheet.Cell(day + 2, 1).Value = new DateOnly(month.Year, month.Month, day).ToDateTime(TimeOnly.MinValue);
        }

        var nightRecords = attendanceRecords
            .Select((record, index) => new { Record = record, Index = index, Name = AttendanceRuleEngine.CleanName(record.Name) })
            .Where(item => !string.IsNullOrWhiteSpace(item.Name)
                && (item.Record.NormalizedStatus.Contains("中班", StringComparison.Ordinal)
                    || item.Record.NormalizedStatus.Contains("夜班", StringComparison.Ordinal)))
            .GroupBy(item => item.Name, StringComparer.Ordinal)
            .OrderBy(group => group.Min(item => item.Index))
            .ToList();

        for (var index = 0; index < nightRecords.Count; index++)
        {
            sheet.Cell(2, index + 2).Value = nightRecords[index].Key;
            foreach (var item in nightRecords[index])
            {
                if (item.Record.Date.Year == month.Year && item.Record.Date.Month == month.Month)
                {
                    sheet.Cell(item.Record.Date.Day + 2, index + 2).Value = item.Record.NormalizedStatus.Contains("夜班", StringComparison.Ordinal)
                        ? "夜班"
                        : "中班";
                }
            }
        }

        var summaryStartRow = days + 3;
        var lastSummaryColumn = Math.Max(11, nightRecords.Count + 1);
        for (var col = 2; col <= lastSummaryColumn; col++)
        {
            sheet.Cell(summaryStartRow, col).Value = 150;
            sheet.Cell(summaryStartRow + 1, col).FormulaA1 = $"COUNTIF({sheet.Cell(3, col).Address.ToStringRelative()}:{sheet.Cell(days + 2, col).Address.ToStringRelative()},\"夜班\")+"
                + $"COUNTIF({sheet.Cell(3, col).Address.ToStringRelative()}:{sheet.Cell(days + 2, col).Address.ToStringRelative()},\"中班\")";
            sheet.Cell(summaryStartRow + 2, col).FormulaA1 = $"{sheet.Cell(summaryStartRow, col).Address.ToStringRelative()}*{sheet.Cell(summaryStartRow + 1, col).Address.ToStringRelative()}";
        }

        sheet.Cell(summaryStartRow, 1).Value = "标准";
        sheet.Cell(summaryStartRow + 1, 1).Value = "合计天数";
        sheet.Cell(summaryStartRow + 2, 1).Value = "合计补贴";
        sheet.Cell(summaryStartRow + 3, 1).Value = "总计";
        sheet.Cell(summaryStartRow + 3, lastSummaryColumn).FormulaA1 = $"SUM(B{summaryStartRow + 2}:{sheet.Cell(summaryStartRow + 2, lastSummaryColumn).Address.ToStringRelative()})";
        sheet.Cell(summaryStartRow + 5, 1).Value = "  编制：肖玲玲                                                          审核：";

        var lastStyledRow = summaryStartRow + 5;
        var lastStyledColumn = Math.Max(40, nightRecords.Count + 1);
        sheet.Row(1).Height = 49;
        sheet.Row(2).Height = 69;
        SetSavedColumnWidth(sheet, 1, 22.1416666666667);
        for (var col = 2; col <= lastStyledColumn; col++)
        {
            SetSavedColumnWidth(sheet, col, 13);
        }
        SetSavedColumnWidth(sheet, 2, 16.0666666666667);
        SetSavedColumnWidth(sheet, 3, 16.0666666666667);
        SetSavedColumnWidth(sheet, 17, 18.3916666666667);
        SetSavedColumnWidth(sheet, 20, 16.9583333333333);

        var all = sheet.Range(1, 1, lastStyledRow, lastStyledColumn);
        all.Style.Font.FontName = "宋体";
        all.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        all.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        all.Style.Alignment.WrapText = true;
        all.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        all.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 28;
        sheet.Range(3, 1, days + 2, 1).Style.DateFormat.Format = "m/d";
        for (var day = 1; day <= days; day++)
        {
            var date = new DateOnly(month.Year, month.Month, day);
            if (date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday)
            {
                sheet.Cell(day + 2, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC000");
            }
        }
    }

    private void WriteAttendance(XLWorkbook workbook, YearMonth month, IReadOnlySet<DateOnly> selectedHolidays, IReadOnlyList<AttendanceRecord> records, IReadOnlyList<AttendanceException> exceptions)
    {
        var sheet = FindAttendanceSheet(workbook, month)
            ?? throw new InvalidOperationException("模板中未找到考勤统计表 Sheet。");
        var templateMonth = DetectTemplateMonth(sheet);
        var sameTemplateMonth = templateMonth == null || templateMonth.Value == month.Month;
        var layout = DetectAttendanceLayout(sheet)
            ?? throw new InvalidOperationException($"Sheet“{sheet.Name}”无法自动识别姓名列、部门列或日期列。");

        RewriteTitleAndDateHeaders(sheet, month, selectedHolidays, layout);
        RenameWorksheet(sheet, $"{month.Month}月考勤统计表 ");
        RemoveOtherAttendanceSheets(workbook, sheet);

        var byNameDate = records
            .GroupBy(r => (AttendanceRuleEngine.CleanName(r.Name), r.Date))
            .ToDictionary(g => g.Key, g => PickStatus(g.Select(r => r.NormalizedStatus)));
        var exceptionsByName = exceptions
            .Where(e => !string.IsNullOrWhiteSpace(e.Name))
            .GroupBy(e => AttendanceRuleEngine.CleanName(e.Name))
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.Ordinal);
        var used = sheet.RangeUsed()!;
        var days = DateTime.DaysInMonth(month.Year, month.Month);
        var written = 0;
        var employeeRows = new List<IXLRow>();
        for (var rowNumber = layout.DataStartRow; rowNumber <= used.LastRow().RowNumber(); rowNumber++)
        {
            var row = sheet.Row(rowNumber);
            var name = AttendanceRuleEngine.CleanName(AttendanceReader.CellText(row.Cell(layout.NameColumn)));
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            employeeRows.Add(row);
            if (layout.NameColumn > 1)
            {
                row.Cell(layout.NameColumn - 1).Value = employeeRows.Count;
            }

            var statuses = new Dictionary<int, string>();
            for (var day = 1; day <= layout.DateColumnCount; day++)
            {
                var cell = row.Cell(layout.DateStartColumn + day - 1);
                if (day > days)
                {
                    cell.Clear(XLClearOptions.Contents);
                    continue;
                }

                var date = new DateOnly(month.Year, month.Month, day);
                var existing = AttendanceReader.CellText(cell);
                var status = BuildBaselineStatus(date, selectedHolidays);
                if (sameTemplateMonth && IsReusableManualStatus(existing))
                {
                    status = existing;
                }
                else if (IsOutOfEmployment(row, layout, date))
                {
                    status = "/";
                }
                else if (byNameDate.TryGetValue((name, date), out var value))
                {
                    status = MergeSourceStatus(status, value);
                }

                cell.Value = status;
                ApplyAttendanceStatusFill(cell, status);
                statuses[day] = status;
                written++;
            }

            if (statuses.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
            {
                WriteStatistics(row, layout, month, selectedHolidays, statuses);
            }
            else
            {
                ClearStatistics(row, layout);
            }

            if (exceptionsByName.TryGetValue(name, out var rowExceptions))
            {
                WriteReviewNote(row, layout, rowExceptions);
            }
        }

        WriteTotalRow(sheet, layout, employeeRows);
        SetSavedColumnWidth(sheet, 4, 5.25);
        SetSavedColumnWidth(sheet, 5, 5.25);
        _log($"已写入考勤主表“{sheet.Name}”：{written} 个日期单元格。");
    }

    private static void WriteReviewNote(IXLRow row, AttendanceLayout layout, IReadOnlyList<AttendanceException> exceptions)
    {
        var remarkColumn = layout.StatColumns.FirstOrDefault(s => s.Key.Contains("备注", StringComparison.Ordinal)).Value;
        if (remarkColumn <= 0 || exceptions.Count == 0)
        {
            return;
        }

        var noteParts = exceptions
            .Take(2)
            .Select(e =>
            {
                var date = e.Date.HasValue ? $"{e.Date.Value.Month}/{e.Date.Value.Day} " : string.Empty;
                var raw = string.IsNullOrWhiteSpace(e.RawStatus) ? string.Empty : $"{e.RawStatus} ";
                return $"{date}{raw}{e.Reason}".Trim();
            })
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList();
        if (exceptions.Count > noteParts.Count)
        {
            noteParts.Add($"另{exceptions.Count - noteParts.Count}条见异常明细");
        }

        var existing = AttendanceReader.CellText(row.Cell(remarkColumn));
        var note = $"需复核：{string.Join("；", noteParts)}";
        row.Cell(remarkColumn).Value = string.IsNullOrWhiteSpace(existing)
            ? note
            : $"{existing}；{note}";
    }

    private void WriteOvertime(XLWorkbook workbook, YearMonth month, IReadOnlyList<OvertimeRecord> records)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains("加班", StringComparison.Ordinal));
        if (sheet == null)
        {
            _log("模板未找到加班 Sheet，跳过加班写入。");
            return;
        }

        RenameOvertimeSheetAndTitle(sheet, month);
        var headerRow = sheet.RowsUsed().FirstOrDefault(r =>
            r.CellsUsed().Count(c => AttendanceReader.TryGetDate(c).HasValue) >= 10);
        if (headerRow == null)
        {
            _log($"Sheet“{sheet.Name}”未识别到加班日期行，跳过。");
            return;
        }

        var dateColumns = headerRow.CellsUsed()
            .Select(c => (Column: c.Address.ColumnNumber, Date: AttendanceReader.TryGetDate(c)))
            .Where(x => x.Date.HasValue)
            .OrderBy(x => x.Column)
            .Take(31)
            .Select((x, index) => (Day: index + 1, x.Column))
            .ToDictionary(x => x.Day, x => x.Column);
        var used = sheet.RangeUsed();
        if (used == null)
        {
            return;
        }

        RewriteOvertimeDateHeaders(sheet, headerRow, dateColumns, month);
        var nameColumn = Math.Max(1, dateColumns.Values.Min() - 1);
        var byNameDate = records
            .GroupBy(r => (AttendanceRuleEngine.CleanName(r.Name), r.Date))
            .ToDictionary(g => g.Key, g => g.Sum(r => r.Hours));
        var namesWithRecords = byNameDate.Keys
            .Select(k => k.Item1)
            .ToHashSet(StringComparer.Ordinal);

        var written = 0;
        for (var rowNumber = headerRow.RowNumber() + 1; rowNumber <= used.LastRow().RowNumber(); rowNumber++)
        {
            var row = sheet.Row(rowNumber);
            var originalName = AttendanceReader.CellText(row.Cell(nameColumn));
            var name = AttendanceRuleEngine.CleanName(originalName);
            if (string.IsNullOrWhiteSpace(name) || !ContainsChinese(name) || IsOvertimeTotalRow(name))
            {
                continue;
            }

            foreach (var col in dateColumns.Values)
            {
                row.Cell(col).Clear(XLClearOptions.Contents);
            }

            if (!namesWithRecords.Contains(name))
            {
                continue;
            }

            for (var day = 1; day <= DateTime.DaysInMonth(month.Year, month.Month); day++)
            {
                if (!dateColumns.TryGetValue(day, out var col))
                {
                    continue;
                }

                var date = new DateOnly(month.Year, month.Month, day);
                var isRestRow = originalName.EndsWith('2');
                var isRestOvertimeDate = IsCalendarRestDay(date);
                if (byNameDate.TryGetValue((name, date), out var hours) && isRestRow == isRestOvertimeDate)
                {
                    var cell = row.Cell(col);
                    if (string.IsNullOrWhiteSpace(AttendanceReader.CellText(cell)))
                    {
                        cell.Value = hours;
                        written++;
                    }
                }
            }
        }

        _log($"已写入加班表“{sheet.Name}”：{written} 个单元格。");
    }

    private static bool IsOvertimeTotalRow(string name)
    {
        return name.Contains("合计", StringComparison.Ordinal) || name.Contains("总计", StringComparison.Ordinal);
    }

    private static IXLWorksheet? FindAttendanceSheet(XLWorkbook workbook, YearMonth month)
    {
        var monthText = $"{month.Month}月";
        return workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains(monthText, StringComparison.Ordinal)
                && ws.Name.Contains("考勤统计", StringComparison.Ordinal))
            ?? workbook.Worksheets
                .Where(ws => ws.Name.Contains("考勤统计", StringComparison.Ordinal))
                .OrderByDescending(UsedCellCount)
                .FirstOrDefault();
    }

    private static int UsedCellCount(IXLWorksheet worksheet)
    {
        var used = worksheet.RangeUsed();
        return used == null ? 0 : used.RowCount() * used.ColumnCount();
    }

    private static void RemoveOtherAttendanceSheets(XLWorkbook workbook, IXLWorksheet keptSheet)
    {
        foreach (var worksheet in workbook.Worksheets
                     .Where(ws => !ReferenceEquals(ws, keptSheet) && ws.Name.Contains("考勤统计", StringComparison.Ordinal))
                     .ToList())
        {
            worksheet.Delete();
        }
    }

    private static AttendanceLayout? DetectAttendanceLayout(IXLWorksheet sheet)
    {
        var used = sheet.RangeUsed();
        if (used == null)
        {
            return null;
        }

        IXLRow? dateHeaderRow = null;
        for (var rowNumber = used.FirstRow().RowNumber(); rowNumber <= Math.Min(used.LastRow().RowNumber(), 10); rowNumber++)
        {
            var row = sheet.Row(rowNumber);
            var dayLike = row.CellsUsed().Count(IsDateHeaderCell);
            if (dayLike >= 20)
            {
                dateHeaderRow = row;
                break;
            }
        }

        if (dateHeaderRow == null)
        {
            return null;
        }

        var dateCells = dateHeaderRow.CellsUsed()
            .Where(IsDateHeaderCell)
            .OrderBy(c => c.Address.ColumnNumber)
            .ToList();
        if (dateCells.Count == 0)
        {
            return null;
        }

        var dateStart = dateCells.First().Address.ColumnNumber;
        var dateEnd = dateCells.Last().Address.ColumnNumber;
        var stats = new Dictionary<string, int>();
        foreach (var cell in dateHeaderRow.CellsUsed().Where(c => c.Address.ColumnNumber > dateEnd))
        {
            var text = AttendanceReader.CellText(cell);
            if (!string.IsNullOrWhiteSpace(text))
            {
                stats[text] = cell.Address.ColumnNumber;
            }
        }

        return new AttendanceLayout(
            DateHeaderRow: dateHeaderRow.RowNumber(),
            WeekdayHeaderRow: dateHeaderRow.RowNumber() + 1,
            DataStartRow: dateHeaderRow.RowNumber() + 2,
            DateStartColumn: dateStart,
            DateColumnCount: dateCells.Count,
            NameColumn: dateStart - 2,
            DepartmentColumn: dateStart - 1,
            StatColumns: stats);
    }

    private static bool IsDateHeaderCell(IXLCell cell)
    {
        var date = AttendanceReader.TryGetDate(cell);
        if (date.HasValue)
        {
            return true;
        }

        var text = AttendanceReader.CellText(cell);
        return Regex.IsMatch(text, @"^\d{1,2}A?$");
    }

    private static bool ContainsChinese(string value)
    {
        return value.Any(ch => ch >= '\u4e00' && ch <= '\u9fff');
    }

    private void RewriteTitleAndDateHeaders(IXLWorksheet sheet, YearMonth month, IReadOnlySet<DateOnly> selectedHolidays, AttendanceLayout layout)
    {
        var titleCell = sheet.Cell(1, 1);
        var title = AttendanceReader.CellText(titleCell);
        if (!string.IsNullOrWhiteSpace(title))
        {
            title = Regex.Replace(title, @"\d{4}年\d{1,2}月", $"{month.Year}年{month.Month}月");
            titleCell.Value = Regex.Replace(title, @"（.*?）", BuildAttendanceTitleSummary(month, selectedHolidays));
        }

        var days = DateTime.DaysInMonth(month.Year, month.Month);
        SetSavedColumnWidth(sheet, layout.DateStartColumn, 5.25);
        SetSavedColumnWidth(sheet, layout.DateStartColumn + 1, 5.25);
        for (var day = 1; day <= layout.DateColumnCount; day++)
        {
            var headerCell = sheet.Cell(layout.DateHeaderRow, layout.DateStartColumn + day - 1);
            var weekdayCell = sheet.Cell(layout.WeekdayHeaderRow, layout.DateStartColumn + day - 1);
            if (day <= days)
            {
                var date = new DateOnly(month.Year, month.Month, day);
                headerCell.Value = date.ToDateTime(TimeOnly.MinValue);
                weekdayCell.Value = ToChineseWeekday(date.DayOfWeek);
            }
            else
            {
                headerCell.Clear(XLClearOptions.Contents);
                weekdayCell.Clear(XLClearOptions.Contents);
            }
        }
    }

    private string BuildAttendanceTitleSummary(YearMonth month, IReadOnlySet<DateOnly> selectedHolidays)
    {
        var days = DateTime.DaysInMonth(month.Year, month.Month);
        var dates = Enumerable.Range(1, days).Select(day => new DateOnly(month.Year, month.Month, day)).ToList();
        var legalDays = dates.Count(selectedHolidays.Contains);
        var restDays = dates.Count(date => !selectedHolidays.Contains(date) && IsCalendarRestDay(date));
        var workDays = days - legalDays - restDays;
        return $"（出勤{workDays}天 法定{legalDays}天  休{restDays}天 ）";
    }

    private static int? DetectTemplateMonth(IXLWorksheet sheet)
    {
        var text = $"{sheet.Name} {AttendanceReader.CellText(sheet.Cell(1, 1))}";
        var match = Regex.Match(text, @"(?<month>\d{1,2})月");
        return match.Success ? int.Parse(match.Groups["month"].Value, CultureInfo.InvariantCulture) : null;
    }

    private static void RenameWorksheet(IXLWorksheet sheet, string desiredName)
    {
        if (sheet.Name.Equals(desiredName, StringComparison.Ordinal))
        {
            return;
        }

        var workbook = sheet.Workbook;
        var name = desiredName;
        var suffix = 1;
        while (workbook.Worksheets.Any(ws => !ReferenceEquals(ws, sheet) && ws.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
        {
            name = $"{desiredName.Trim()}_{suffix++}";
        }

        sheet.Name = name;
    }

    private static void RenameOvertimeSheetAndTitle(IXLWorksheet sheet, YearMonth month)
    {
        RenameWorksheet(sheet, "加班补贴表 ");
        var titleCell = sheet.Cell(1, 1);
        var title = AttendanceReader.CellText(titleCell);
        if (!string.IsNullOrWhiteSpace(title))
        {
            titleCell.Value = Regex.Replace(title, @"\d{4}年\d{1,2}月", $"{month.Year}年{month.Month}月");
        }
    }

    private static void RewriteOvertimeDateHeaders(IXLWorksheet sheet, IXLRow headerRow, IReadOnlyDictionary<int, int> dateColumns, YearMonth month)
    {
        var days = DateTime.DaysInMonth(month.Year, month.Month);
        foreach (var pair in dateColumns)
        {
            var cell = sheet.Cell(headerRow.RowNumber(), pair.Value);
            if (pair.Key <= days)
            {
                cell.Value = new DateOnly(month.Year, month.Month, pair.Key).ToDateTime(TimeOnly.MinValue);
            }
            else
            {
                cell.Clear(XLClearOptions.Contents);
            }
        }
    }

    private static void RewriteNightShiftSheet(XLWorkbook workbook, YearMonth month)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains("夜班", StringComparison.Ordinal));
        if (sheet == null)
        {
            return;
        }

        var titleCell = sheet.Cell(1, 1);
        var title = AttendanceReader.CellText(titleCell);
        if (!string.IsNullOrWhiteSpace(title))
        {
            titleCell.Value = Regex.Replace(title, @"\d{4}年\d{1,2}月", $"{month.Year}年{month.Month}月");
        }

        SetSavedColumnWidth(sheet, 2, 16.0666666666667);
        SetSavedColumnWidth(sheet, 3, 16.0666666666667);

        var dateRows = sheet.RowsUsed()
            .Where(row => AttendanceReader.TryGetDate(row.Cell(1)).HasValue)
            .OrderBy(row => row.RowNumber())
            .Take(31)
            .ToList();
        var days = DateTime.DaysInMonth(month.Year, month.Month);
        for (var index = 0; index < dateRows.Count; index++)
        {
            var row = dateRows[index];
            if (index < days)
            {
                row.Cell(1).Value = new DateOnly(month.Year, month.Month, index + 1).ToDateTime(TimeOnly.MinValue);
            }
            else
            {
                row.Clear(XLClearOptions.Contents);
            }
        }
    }

    private void WriteStatistics(IXLRow row, AttendanceLayout layout, YearMonth month, IReadOnlySet<DateOnly> selectedHolidays, IReadOnlyDictionary<int, string> statuses)
    {
        var total = statuses.Count;
        var restDays = statuses.Count(pair => IsRestDay(month, selectedHolidays, pair.Key, pair.Value));
        var legalDays = statuses.Count(pair => IsLegalDay(month, selectedHolidays, pair.Key, pair.Value));
        var slashWorkdays = statuses.Count(pair => IsSlashWorkday(month, selectedHolidays, pair.Key, pair.Value));
        var holidayRestAbsences = statuses.Count(pair => IsHolidayRestAbsence(month, selectedHolidays, pair.Key, pair.Value));
        var leaveDays = RoundDays(statuses.Values.Sum(LeaveDays) + slashWorkdays + holidayRestAbsences);
        var attendanceDays = RoundDays(total - restDays - legalDays - leaveDays);

        SetStat(row, layout, "合计", total);
        SetStat(row, layout, "出勤", Math.Max(0, attendanceDays), writeZero: true);
        SetStat(row, layout, "休息", restDays);
        SetStat(row, layout, "法定", legalDays, writeZero: true);

        var approvedCompLeaveDays = IsApprovedCompLeave(row, layout)
            ? statuses.Values.Sum(v => LeaveDaysForType(v, "调休"))
            : 0;
        var unapprovedCompLeaveDays = statuses.Values.Sum(v => LeaveDaysForType(v, "调休")) - approvedCompLeaveDays;

        SetStat(row, layout, "请假", leaveDays);
        SetStat(row, layout, "年假", RoundDays(statuses.Values.Sum(v => LeaveDaysForType(v, "年假"))));
        SetStat(row, layout, "调休", RoundDays(approvedCompLeaveDays));
        SetStat(row, layout, "婚假", RoundDays(statuses.Values.Sum(v => LeaveDaysForType(v, "婚假"))));
        SetStat(row, layout, "产假", RoundDays(statuses.Values.Sum(v => LeaveDaysForType(v, "产假"))));
        SetStat(row, layout, "扣工资合计", RoundDays(slashWorkdays + holidayRestAbsences + unapprovedCompLeaveDays + statuses.Values.Sum(DeductibleLeaveDays)), writeZero: true);
    }

    private static bool IsApprovedCompLeave(IXLRow row, AttendanceLayout layout)
    {
        var remarkColumn = layout.StatColumns.FirstOrDefault(s => s.Key.Contains("备注", StringComparison.Ordinal)).Value;
        if (remarkColumn <= 0)
        {
            return false;
        }

        var remark = AttendanceReader.CellText(row.Cell(remarkColumn));
        return remark.Contains("调休", StringComparison.Ordinal);
    }

    private static void ClearStatistics(IXLRow row, AttendanceLayout layout)
    {
        foreach (var column in layout.StatColumns.Values.Distinct())
        {
            if (!row.Cell(column).HasFormula)
            {
                row.Cell(column).Clear(XLClearOptions.Contents);
            }
        }
    }

    private static void WriteTotalRow(IXLWorksheet sheet, AttendanceLayout layout, IReadOnlyList<IXLRow> employeeRows)
    {
        if (employeeRows.Count == 0)
        {
            return;
        }

        var used = sheet.RangeUsed();
        if (used == null)
        {
            return;
        }

        var totalRow = sheet.Rows(layout.DataStartRow, used.LastRow().RowNumber())
            .FirstOrDefault(row => string.IsNullOrWhiteSpace(AttendanceReader.CellText(row.Cell(layout.NameColumn)))
                && layout.StatColumns.Values.Any(col => row.Cell(col).HasFormula || !string.IsNullOrWhiteSpace(AttendanceReader.CellText(row.Cell(col)))));
        if (totalRow == null)
        {
            return;
        }

        foreach (var column in layout.StatColumns.Values.Distinct())
        {
            if (layout.StatColumns.Any(item => item.Value == column && item.Key.Contains("备注", StringComparison.Ordinal)))
            {
                continue;
            }

            var sum = employeeRows.Sum(row => TryReadDecimal(row.Cell(column)));
            totalRow.Cell(column).Value = sum == 0 ? 0 : sum;
        }
    }

    private static decimal TryReadDecimal(IXLCell cell)
    {
        var text = AttendanceReader.CellText(cell);
        return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var value)
            || decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value)
            ? value
            : 0;
    }

    private string BuildBaselineStatus(DateOnly date, IReadOnlySet<DateOnly> selectedHolidays)
    {
        if (selectedHolidays.Contains(date))
        {
            return "法定";
        }

        return IsCalendarRestDay(date) ? "休" : "公司出勤";
    }

    private static string MergeSourceStatus(string baseline, string source)
    {
        if (string.IsNullOrWhiteSpace(source))
        {
            return baseline;
        }

        if (source is "迟到" or "早退" or "缺卡" or "旷工")
        {
            return "公司出勤";
        }

        if ((baseline == "休" || baseline == "法定")
            && source is "公司出勤" or "旷工")
        {
            return baseline;
        }

        return source;
    }

    private static bool IsReusableManualStatus(string existing)
    {
        if (string.IsNullOrWhiteSpace(existing))
        {
            return false;
        }

        if (existing is "/" or "公司出勤" or "法定" or "迟到" or "早退" or "缺卡" or "旷工")
        {
            return false;
        }

        return existing == "休"
            || existing.Contains("中班", StringComparison.Ordinal)
            || existing.Contains("夜班", StringComparison.Ordinal)
            || existing.Contains("哺乳", StringComparison.Ordinal)
            || existing.Contains("请假", StringComparison.Ordinal)
            || existing.Contains("调休", StringComparison.Ordinal)
            || existing.Contains("事假", StringComparison.Ordinal)
            || existing.Contains("年假", StringComparison.Ordinal)
            || existing.Contains("病假", StringComparison.Ordinal)
            || existing.Contains("婚假", StringComparison.Ordinal)
            || existing.Contains("产假", StringComparison.Ordinal)
            || existing.Contains("丧假", StringComparison.Ordinal)
            || existing.Contains("陪产假", StringComparison.Ordinal)
            || existing.Contains("例假", StringComparison.Ordinal);
    }

    private static void ApplyAttendanceStatusFill(IXLCell cell, string status)
    {
        if (status == "/")
        {
            cell.Style.Fill.BackgroundColor = XLColor.NoColor;
            return;
        }

        if (status == "休" || status == "法定")
        {
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Background1);
            return;
        }

        if (status == "公司出勤")
        {
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent5);
            return;
        }

        if (status.Contains("中班", StringComparison.Ordinal) || status.Contains("夜班", StringComparison.Ordinal))
        {
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC000");
            return;
        }

        if (status.Contains("产假", StringComparison.Ordinal))
        {
            cell.Style.Fill.BackgroundColor = XLColor.Yellow;
            return;
        }

        if (status.Contains("年假", StringComparison.Ordinal))
        {
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1);
            return;
        }

        if (status.Contains("调休", StringComparison.Ordinal) && !status.Contains("请假", StringComparison.Ordinal))
        {
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent6);
            return;
        }

        if (status.Contains("假", StringComparison.Ordinal) || status.Contains("请假", StringComparison.Ordinal))
        {
            cell.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent2);
        }
    }

    private static bool IsOutOfEmployment(IXLRow row, AttendanceLayout layout, DateOnly date)
    {
        var remarkColumn = layout.StatColumns.FirstOrDefault(s => s.Key.Contains("备注", StringComparison.Ordinal)).Value;
        if (remarkColumn <= 0)
        {
            return false;
        }

        var remark = AttendanceReader.CellText(row.Cell(remarkColumn));
        if (string.IsNullOrWhiteSpace(remark))
        {
            return false;
        }

        var start = ParseRemarkDate(remark, "入职", date.Year);
        var end = ParseRemarkDate(remark, "离职", date.Year);
        var selfLeave = ParseRemarkDate(remark, "自离", date.Year);
        if (start.HasValue && date < start.Value)
        {
            return true;
        }

        if (end.HasValue && date > end.Value)
        {
            return true;
        }

        if (selfLeave.HasValue && date >= selfLeave.Value)
        {
            return true;
        }

        return false;
    }

    private static DateOnly? ParseRemarkDate(string remark, string keyword, int year)
    {
        var match = Regex.Match(remark, $@"(?<month>\d{{1,2}})\.(?<day>\d{{1,2}})\s*(?:新)?{Regex.Escape(keyword)}");
        if (!match.Success)
        {
            return null;
        }

        var month = int.Parse(match.Groups["month"].Value, CultureInfo.InvariantCulture);
        var day = int.Parse(match.Groups["day"].Value, CultureInfo.InvariantCulture);
        return new DateOnly(year, month, day);
    }

    private static bool IsLegalDay(YearMonth month, IReadOnlySet<DateOnly> selectedHolidays, int day, string status)
    {
        return status == "法定";
    }

    private bool IsRestDay(YearMonth month, IReadOnlySet<DateOnly> selectedHolidays, int day, string status)
    {
        var date = new DateOnly(month.Year, month.Month, day);
        return status == "休" && !selectedHolidays.Contains(date)
            || (status == "/" && !selectedHolidays.Contains(date) && IsCalendarRestDay(date));
    }

    private static bool IsHolidayRestAbsence(YearMonth month, IReadOnlySet<DateOnly> selectedHolidays, int day, string status)
    {
        var date = new DateOnly(month.Year, month.Month, day);
        return status == "休" && selectedHolidays.Contains(date);
    }

    private bool IsSlashWorkday(YearMonth month, IReadOnlySet<DateOnly> selectedHolidays, int day, string status)
    {
        if (status != "/")
        {
            return false;
        }

        var date = new DateOnly(month.Year, month.Month, day);
        return selectedHolidays.Contains(date) || !IsCalendarRestDay(date);
    }

    private bool IsCalendarRestDay(DateOnly date)
    {
        return _workCalendar.RestDays.Contains(date)
            || (date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday && !_workCalendar.Workdays.Contains(date));
    }

    private static decimal LeaveDays(string status)
    {
        var total = 0m;
        foreach (Match match in Regex.Matches(status, @"(?:请假|事假|调休|年假|病假|婚假|产假|丧假|陪产假|例假)(?<num>\d+(?:\.\d+)?)(?<unit>H|天)"))
        {
            var number = decimal.Parse(match.Groups["num"].Value, CultureInfo.InvariantCulture);
            total += match.Groups["unit"].Value == "天" ? number : number / 8;
        }

        if (total > 0)
        {
            return total;
        }

        return status is "旷工" or "产假" or "婚假" or "陪产假" ? 1 : 0;
    }

    private static decimal LeaveDaysForType(string status, string leaveType)
    {
        var total = 0m;
        foreach (Match match in Regex.Matches(status, $@"{Regex.Escape(leaveType)}(?<num>\d+(?:\.\d+)?)(?<unit>H|天)"))
        {
            var number = decimal.Parse(match.Groups["num"].Value, CultureInfo.InvariantCulture);
            total += match.Groups["unit"].Value == "天" ? number : number / 8;
        }

        if (total == 0 && status == leaveType)
        {
            total = 1;
        }

        return total;
    }

    private static decimal DeductibleLeaveDays(string status)
    {
        return TruncateDays(LeaveDaysForType(status, "请假"))
            + LeaveDaysForType(status, "事假")
            + LeaveDaysForType(status, "病假")
            + (status == "旷工" ? 1 : 0);
    }

    private static decimal RoundDays(decimal value)
    {
        return Math.Round(value, 2, MidpointRounding.AwayFromZero);
    }

    private static decimal TruncateDays(decimal value)
    {
        return Math.Truncate(value * 100) / 100;
    }

    private static void SetStat(IXLRow row, AttendanceLayout layout, string name, decimal value, bool writeZero = false)
    {
        var column = FindStatColumn(layout, name);
        if (column > 0)
        {
            row.Cell(column).Value = value == 0 && !writeZero ? string.Empty : value;
        }
    }

    private static int FindStatColumn(AttendanceLayout layout, string name)
    {
        var exact = layout.StatColumns.FirstOrDefault(s => NormalizeStatHeader(s.Key).Equals(name, StringComparison.Ordinal)).Value;
        if (exact > 0)
        {
            return exact;
        }

        var startsWith = layout.StatColumns.FirstOrDefault(s => NormalizeStatHeader(s.Key).StartsWith(name, StringComparison.Ordinal)).Value;
        if (startsWith > 0)
        {
            return startsWith;
        }

        return layout.StatColumns.FirstOrDefault(s => s.Key.Contains(name, StringComparison.Ordinal)).Value;
    }

    private static string NormalizeStatHeader(string header)
    {
        return Regex.Replace(header, @"（.*?）|\s", string.Empty);
    }

    private static WorkCalendarSettings LoadWorkCalendar(string? configDirectory)
    {
        if (string.IsNullOrWhiteSpace(configDirectory))
        {
            return new WorkCalendarSettings([], []);
        }

        Directory.CreateDirectory(configDirectory);
        var path = Path.Combine(configDirectory, "workCalendar.json");
        if (!File.Exists(path))
        {
            var defaults = new
            {
                restDays = new[] { "2026-04-06", "2026-05-04", "2026-05-05" },
                workdays = new[] { "2026-05-09" }
            };
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };
            File.WriteAllText(path, JsonSerializer.Serialize(defaults, options));
        }

        var json = File.ReadAllText(path);
        using var document = JsonDocument.Parse(json);
        return new WorkCalendarSettings(
            ReadDateSet(document.RootElement, "restDays"),
            ReadDateSet(document.RootElement, "workdays"));
    }

    private static HashSet<DateOnly> ReadDateSet(JsonElement root, string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out var element) || element.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        return element.EnumerateArray()
            .Select(item => DateOnly.TryParse(item.GetString(), out var date) ? date : (DateOnly?)null)
            .Where(date => date.HasValue)
            .Select(date => date!.Value)
            .ToHashSet();
    }

    private static string PickStatus(IEnumerable<string> statuses)
    {
        var list = statuses.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        if (list.Count == 0)
        {
            return string.Empty;
        }

        foreach (var priority in new[] { "/", "法定", "旷工", "缺卡", "迟到", "早退" })
        {
            var match = list.FirstOrDefault(s => s == priority);
            if (match != null)
            {
                return match;
            }
        }

        return list.FirstOrDefault(s => s.Contains("假", StringComparison.Ordinal) || s.StartsWith("调休", StringComparison.Ordinal))
            ?? list.First();
    }

    private static string ToChineseWeekday(DayOfWeek day)
    {
        return day switch
        {
            DayOfWeek.Monday => "一",
            DayOfWeek.Tuesday => "二",
            DayOfWeek.Wednesday => "三",
            DayOfWeek.Thursday => "四",
            DayOfWeek.Friday => "五",
            DayOfWeek.Saturday => "六",
            DayOfWeek.Sunday => "日",
            _ => string.Empty
        };
    }

    private static void WriteExceptionSheet(XLWorkbook workbook, IReadOnlyList<AttendanceException> exceptions)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "异常明细")
            ?? workbook.Worksheets.Add("异常明细");
        sheet.Clear();
        sheet.Cell(1, 1).Value = "姓名";
        sheet.Cell(1, 2).Value = "日期";
        sheet.Cell(1, 3).Value = "原始状态";
        sheet.Cell(1, 4).Value = "异常原因";
        sheet.Range(1, 1, 1, 4).Style.Font.Bold = true;

        for (var i = 0; i < exceptions.Count; i++)
        {
            var item = exceptions[i];
            var row = i + 2;
            sheet.Cell(row, 1).Value = item.Name;
            sheet.Cell(row, 2).Value = item.Date?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? string.Empty;
            sheet.Cell(row, 3).Value = item.RawStatus;
            sheet.Cell(row, 4).Value = item.Reason;
        }

        sheet.Columns(1, 4).AdjustToContents();
    }

    private sealed record AttendanceLayout(
        int DateHeaderRow,
        int WeekdayHeaderRow,
        int DataStartRow,
        int DateStartColumn,
        int DateColumnCount,
        int NameColumn,
        int DepartmentColumn,
        Dictionary<string, int> StatColumns);

    private sealed record WorkCalendarSettings(HashSet<DateOnly> RestDays, HashSet<DateOnly> Workdays);
}
