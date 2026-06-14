using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using AttendanceGenerator.Models;

namespace AttendanceGenerator.Services;

public sealed class TemplateWriter
{
    private readonly Action<string> _log;
    private readonly WorkCalendarSettings _workCalendar;

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
        var referencePath = ResolveReferencePath(templatePath, month, outputPath);
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

        if (!string.IsNullOrWhiteSpace(referencePath))
        {
            ApplyReferenceFormatting(referencePath, outputPath, month);
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
        if (!string.IsNullOrWhiteSpace(templatePath))
        {
            BuildWorkbookFromReference(workbook, templatePath, month);
            workbook.SaveAs(outputPath);
            _log($"已按参考文件用代码生成模板到输出文件：{outputPath}");
            return workbook;
        }

        var referencePath = FindReferenceWorkbook(month, outputPath);
        if (referencePath != null)
        {
            if (File.Exists(outputPath))
            {
                File.SetAttributes(outputPath, FileAttributes.Normal);
            }

            BuildWorkbookFromReference(workbook, referencePath, month);
            workbook.SaveAs(outputPath);
            _log($"已读取手工参考表并用代码重建样式和行顺序：{referencePath}");
            return workbook;
        }
        else
        {
            BuildAttendanceTemplate(workbook, month, selectedHolidays, attendanceRecords);
            BuildOvertimeTemplate(workbook, month, overtimeRecords);
            BuildNightShiftTemplate(workbook, month);
        }
        workbook.SaveAs(outputPath);
        _log($"已用代码生成模板到输出文件：{outputPath}");
        return workbook;
    }

    private static string? FindReferenceWorkbook(YearMonth month, string outputPath)
    {
        var outputDirectory = Path.GetDirectoryName(outputPath);
        var searchDirectories = new[]
        {
            Environment.CurrentDirectory,
            outputDirectory,
            outputDirectory == null ? null : Directory.GetParent(outputDirectory)?.FullName
        }
        .Where(path => !string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
        .Distinct(StringComparer.OrdinalIgnoreCase);

        var preferredNames = new[]
        {
            $"{month.Month}月手工做的.xlsx",
            $"{month.Month:D2}月手工做的.xlsx"
        };
        foreach (var directory in searchDirectories)
        {
            foreach (var name in preferredNames)
            {
                var candidate = Path.Combine(directory!, name);
                if (File.Exists(candidate) && !Path.GetFullPath(candidate).Equals(Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
                {
                    return candidate;
                }
            }
        }

        return null;
    }

    private static string? ResolveReferencePath(string? templatePath, YearMonth month, string outputPath)
    {
        return !string.IsNullOrWhiteSpace(templatePath)
            ? templatePath
            : FindReferenceWorkbook(month, outputPath);
    }

    private static void BuildWorkbookFromReference(XLWorkbook target, string referencePath, YearMonth month)
    {
        using var reference = new XLWorkbook(referencePath);
        CopyReferenceWorksheet(reference, target, FindAttendanceSheet(reference, month), $"{month.Month}月考勤统计表 ");
        CopyReferenceWorksheet(reference, target, reference.Worksheets.FirstOrDefault(ws => ws.Name.Contains("加班", StringComparison.Ordinal)), "加班补贴表 ");
        CopyReferenceWorksheet(reference, target, reference.Worksheets.FirstOrDefault(ws => ws.Name.Contains("夜班", StringComparison.Ordinal)), "夜班补贴");
    }

    private static void CopyReferenceWorksheet(XLWorkbook sourceWorkbook, XLWorkbook targetWorkbook, IXLWorksheet? sourceSheet, string targetName)
    {
        if (sourceSheet == null)
        {
            targetWorkbook.Worksheets.Add(targetName);
            return;
        }

        var targetSheet = targetWorkbook.Worksheets.Add(targetName);
        var used = sourceSheet.RangeUsed();
        if (used == null)
        {
            return;
        }

        var lastColumn = Math.Max(used.LastColumn().ColumnNumber(), 100);
        var lastRow = Math.Max(used.LastRow().RowNumber(), 300);
        for (var col = 1; col <= lastColumn; col++)
        {
            targetSheet.Column(col).Width = sourceSheet.Column(col).Width;
        }

        for (var row = 1; row <= lastRow; row++)
        {
            targetSheet.Row(row).Height = sourceSheet.Row(row).Height;
        }

        foreach (var merged in sourceSheet.MergedRanges)
        {
            var address = merged.RangeAddress.ToString();
            if (!string.IsNullOrWhiteSpace(address))
            {
                targetSheet.Range(address).Merge();
            }
        }

        for (var row = 1; row <= used.LastRow().RowNumber(); row++)
        {
            for (var col = 1; col <= used.LastColumn().ColumnNumber(); col++)
            {
                var sourceCell = sourceSheet.Cell(row, col);
                var targetCell = targetSheet.Cell(row, col);
                targetCell.Style = sourceCell.Style;
                if (sourceCell.HasFormula)
                {
                    targetCell.FormulaA1 = sourceCell.FormulaA1;
                }
                else
                {
                    targetCell.Value = sourceCell.Value;
                }
            }
        }

        targetSheet.SheetView.FreezeRows(sourceSheet.SheetView.SplitRow);
        targetSheet.SheetView.FreezeColumns(sourceSheet.SheetView.SplitColumn);
    }

    private static void ApplyReferenceFormatting(string referencePath, string outputPath, YearMonth month)
    {
        if (!File.Exists(referencePath) || !File.Exists(outputPath))
        {
            return;
        }

        if (Path.GetFullPath(referencePath).Equals(Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        using var source = SpreadsheetDocument.Open(referencePath, false);
        using var target = SpreadsheetDocument.Open(outputPath, true);
        var sourceWorkbook = source.WorkbookPart;
        var targetWorkbook = target.WorkbookPart;
        if (sourceWorkbook == null || targetWorkbook == null)
        {
            return;
        }

        CopyWorkbookStyles(sourceWorkbook, targetWorkbook);
        SyncSheetFormatting(sourceWorkbook, targetWorkbook,
            name => name.Contains($"{month.Month}月", StringComparison.Ordinal) && name.Contains("考勤统计", StringComparison.Ordinal),
            name => name.Contains($"{month.Month}月", StringComparison.Ordinal) && name.Contains("考勤统计", StringComparison.Ordinal));
        SyncSheetFormatting(sourceWorkbook, targetWorkbook,
            name => name.Contains("加班", StringComparison.Ordinal),
            name => name.Contains("加班", StringComparison.Ordinal));
        SyncSheetFormatting(sourceWorkbook, targetWorkbook,
            name => name.Contains("夜班", StringComparison.Ordinal),
            name => name.Contains("夜班", StringComparison.Ordinal));
    }

    private static void CopyWorkbookStyles(WorkbookPart sourceWorkbook, WorkbookPart targetWorkbook)
    {
        if (sourceWorkbook.WorkbookStylesPart != null)
        {
            var targetStyles = targetWorkbook.WorkbookStylesPart ?? targetWorkbook.AddNewPart<WorkbookStylesPart>();
            using var stream = sourceWorkbook.WorkbookStylesPart.GetStream(FileMode.Open, FileAccess.Read);
            targetStyles.FeedData(stream);
        }

        if (sourceWorkbook.ThemePart != null)
        {
            var targetTheme = targetWorkbook.ThemePart ?? targetWorkbook.AddNewPart<ThemePart>();
            using var stream = sourceWorkbook.ThemePart.GetStream(FileMode.Open, FileAccess.Read);
            targetTheme.FeedData(stream);
        }
    }

    private static void SyncSheetFormatting(
        WorkbookPart sourceWorkbook,
        WorkbookPart targetWorkbook,
        Func<string, bool> sourcePredicate,
        Func<string, bool> targetPredicate)
    {
        var sourcePart = FindWorksheetPart(sourceWorkbook, sourcePredicate);
        var targetPart = FindWorksheetPart(targetWorkbook, targetPredicate);
        if (sourcePart == null || targetPart == null)
        {
            return;
        }

        var sourceWorksheet = sourcePart.Worksheet;
        var targetWorksheet = targetPart.Worksheet;
        ReplaceWorksheetChild<SheetViews>(targetWorksheet, sourceWorksheet.GetFirstChild<SheetViews>()?.CloneNode(true));
        ReplaceWorksheetChild<SheetFormatProperties>(targetWorksheet, sourceWorksheet.GetFirstChild<SheetFormatProperties>()?.CloneNode(true));
        ReplaceColumns(targetWorksheet, sourceWorksheet.GetFirstChild<Columns>()?.CloneNode(true));
        ReplaceMergeCells(targetWorksheet, sourceWorksheet.GetFirstChild<MergeCells>()?.CloneNode(true));
        SyncRowsAndCellStyles(sourceWorksheet, targetWorksheet);
        targetWorksheet.Save();
    }

    private static WorksheetPart? FindWorksheetPart(WorkbookPart workbookPart, Func<string, bool> predicate)
    {
        var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>()
            .FirstOrDefault(item => predicate(item.Name?.Value ?? string.Empty));
        return sheet?.Id?.Value == null ? null : workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
    }

    private static void ReplaceWorksheetChild<T>(Worksheet targetWorksheet, OpenXmlElement? sourceChild)
        where T : OpenXmlElement
    {
        targetWorksheet.GetFirstChild<T>()?.Remove();
        if (sourceChild == null)
        {
            return;
        }

        var sheetData = targetWorksheet.GetFirstChild<SheetData>();
        if (sheetData != null && sourceChild is Columns)
        {
            targetWorksheet.InsertBefore(sourceChild, sheetData);
            return;
        }

        if (sheetData != null && sourceChild is MergeCells)
        {
            targetWorksheet.InsertAfter(sourceChild, sheetData);
            return;
        }

        if (sheetData != null)
        {
            targetWorksheet.InsertBefore(sourceChild, sheetData);
        }
        else
        {
            targetWorksheet.Append(sourceChild);
        }
    }

    private static void ReplaceColumns(Worksheet targetWorksheet, OpenXmlElement? sourceColumns)
    {
        ReplaceWorksheetChild<Columns>(targetWorksheet, sourceColumns);
    }

    private static void ReplaceMergeCells(Worksheet targetWorksheet, OpenXmlElement? sourceMergeCells)
    {
        ReplaceWorksheetChild<MergeCells>(targetWorksheet, sourceMergeCells);
    }

    private static void SyncRowsAndCellStyles(Worksheet sourceWorksheet, Worksheet targetWorksheet)
    {
        var sourceRows = sourceWorksheet.Descendants<Row>()
            .Where(row => row.RowIndex?.Value != null)
            .ToDictionary(row => row.RowIndex!.Value);
        var targetSheetData = targetWorksheet.GetFirstChild<SheetData>();
        if (targetSheetData == null)
        {
            return;
        }

        foreach (var targetRow in targetSheetData.Elements<Row>())
        {
            targetRow.Height = null;
            targetRow.CustomHeight = null;
            targetRow.StyleIndex = null;
            targetRow.CustomFormat = null;
        }

        foreach (var pair in sourceRows)
        {
            var targetRow = GetOrCreateRow(targetSheetData, pair.Key);
            CopyRowFormatting(pair.Value, targetRow);
            foreach (var sourceCell in pair.Value.Elements<Cell>())
            {
                if (string.IsNullOrWhiteSpace(sourceCell.CellReference?.Value))
                {
                    continue;
                }

                var targetCell = GetOrCreateCell(targetRow, sourceCell.CellReference.Value);
                targetCell.StyleIndex = sourceCell.StyleIndex == null ? null : new UInt32Value(sourceCell.StyleIndex.Value);
            }
        }
    }

    private static void CopyRowFormatting(Row sourceRow, Row targetRow)
    {
        targetRow.Height = sourceRow.Height == null ? null : new DoubleValue(sourceRow.Height.Value);
        targetRow.CustomHeight = sourceRow.CustomHeight == null ? null : new BooleanValue(sourceRow.CustomHeight.Value);
        targetRow.StyleIndex = sourceRow.StyleIndex == null ? null : new UInt32Value(sourceRow.StyleIndex.Value);
        targetRow.CustomFormat = sourceRow.CustomFormat == null ? null : new BooleanValue(sourceRow.CustomFormat.Value);
        targetRow.Hidden = sourceRow.Hidden == null ? null : new BooleanValue(sourceRow.Hidden.Value);
        targetRow.OutlineLevel = sourceRow.OutlineLevel == null ? null : new ByteValue(sourceRow.OutlineLevel.Value);
        targetRow.Collapsed = sourceRow.Collapsed == null ? null : new BooleanValue(sourceRow.Collapsed.Value);
        targetRow.ThickTop = sourceRow.ThickTop == null ? null : new BooleanValue(sourceRow.ThickTop.Value);
        targetRow.ThickBot = sourceRow.ThickBot == null ? null : new BooleanValue(sourceRow.ThickBot.Value);
    }

    private static Row GetOrCreateRow(SheetData sheetData, uint rowIndex)
    {
        var row = sheetData.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value == rowIndex);
        if (row != null)
        {
            return row;
        }

        row = new Row { RowIndex = rowIndex };
        var nextRow = sheetData.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value > rowIndex);
        if (nextRow != null)
        {
            sheetData.InsertBefore(row, nextRow);
        }
        else
        {
            sheetData.Append(row);
        }

        return row;
    }

    private static Cell GetOrCreateCell(Row row, string cellReference)
    {
        var cell = row.Elements<Cell>().FirstOrDefault(item => item.CellReference?.Value == cellReference);
        if (cell != null)
        {
            return cell;
        }

        cell = new Cell { CellReference = cellReference };
        var targetColumn = GetColumnName(cellReference);
        var nextCell = row.Elements<Cell>()
            .FirstOrDefault(item => string.Compare(GetColumnName(item.CellReference?.Value ?? string.Empty), targetColumn, StringComparison.Ordinal) > 0);
        if (nextCell != null)
        {
            row.InsertBefore(cell, nextCell);
        }
        else
        {
            row.Append(cell);
        }

        return cell;
    }

    private static string GetColumnName(string cellReference)
    {
        return Regex.Replace(cellReference, @"\d", string.Empty);
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

        var employees = attendanceRecords
            .Select((record, index) => new { Record = record, Index = index })
            .GroupBy(item => AttendanceRuleEngine.CleanName(item.Record.Name))
            .OrderBy(g => g.Min(item => item.Index))
            .Select(g => g.First().Record)
            .GroupBy(r => AttendanceRuleEngine.CleanName(r.Name))
            .Select(g => new
            {
                Name = g.Key,
                Department = g.Select(r => r.Department).FirstOrDefault(d => !string.IsNullOrWhiteSpace(d)) ?? string.Empty,
            })
            .Where(e => !string.IsNullOrWhiteSpace(e.Name))
            .ToList();

        var row = 4;
        foreach (var employee in employees)
        {
            sheet.Cell(row, 1).FormulaA1 = $"ROW()-3";
            sheet.Cell(row, 2).Value = employee.Name;
            sheet.Cell(row, 3).Value = employee.Department;
            sheet.Row(row).Height = 34;
            row++;
        }

        var totalRow = row;
        sheet.Cell(totalRow, 2).Value = "合计";
        sheet.Range(totalRow, 35, totalRow, 36).Merge();
        sheet.Cell(totalRow, 35).FormulaA1 = $"SUM(AI4:AI{totalRow - 1})";

        ApplyAttendanceStyle(sheet, totalRow);
    }

    private static void ApplyAttendanceStyle(IXLWorksheet sheet, int lastRow)
    {
        sheet.Row(1).Height = 31;
        sheet.Row(2).Height = 23;
        sheet.Row(3).Height = 39;
        sheet.Column(1).Width = 3.54;
        sheet.Column(2).Width = 7.61;
        sheet.Column(3).Width = 13;
        for (var col = 4; col <= 34; col++)
        {
            sheet.Column(col).Width = col is 4 or 5 ? 5.25 : 13;
        }

        for (var col = 35; col <= 47; col++)
        {
            sheet.Column(col).Width = col == 47 ? 18 : 8.5;
        }

        var all = sheet.Range(1, 1, lastRow, 47);
        all.Style.Font.FontName = "宋体";
        all.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        all.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        all.Style.Alignment.WrapText = true;
        all.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        all.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        sheet.Range(2, 4, 2, 34).Style.Fill.BackgroundColor = XLColor.FromHtml("#33CCCC");
        sheet.Range(2, 4, 2, 34).Style.DateFormat.Format = "d\"A\"";
        sheet.Range(2, 35, 3, 47).Style.Fill.BackgroundColor = XLColor.FromHtml("#33CCCC");
        sheet.SheetView.FreezeRows(3);
        sheet.SheetView.FreezeColumns(3);
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

        var names = overtimeRecords
            .Select(r => AttendanceRuleEngine.CleanName(r.Name))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToList();
        var row = 3;
        foreach (var name in names)
        {
            AddOvertimeRow(sheet, row++, name, 25, false);
            AddOvertimeRow(sheet, row++, $"{name}2", 40, true);
        }

        ApplyOvertimeStyle(sheet, Math.Max(3, row - 1));
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
        }
        else
        {
            sheet.Cell(row, 38).Value = name;
            sheet.Cell(row, 39).FormulaA1 = $"AI{row}";
        }
    }

    private static void ApplyOvertimeStyle(IXLWorksheet sheet, int lastRow)
    {
        sheet.Row(1).Height = 47;
        sheet.Column(1).Width = 14.37;
        sheet.Column(2).Width = 7.34;
        for (var col = 3; col <= 35; col++)
        {
            sheet.Column(col).Width = 13;
        }

        for (var row = 3; row <= lastRow; row++)
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
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        sheet.Range(2, 2, 2, 32).Style.DateFormat.Format = "d";
        sheet.Range(2, 2, 2, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC000");
        sheet.SheetView.FreezeRows(2);
        sheet.SheetView.FreezeColumns(1);
    }

    private static void BuildNightShiftTemplate(XLWorkbook workbook, YearMonth month)
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

        sheet.Row(1).Height = 49;
        sheet.Row(2).Height = 69;
        sheet.Column(1).Width = 22.14;
        for (var col = 2; col <= 40; col++)
        {
            sheet.Column(col).Width = 16.07;
        }

        var all = sheet.Range(1, 1, days + 2, 40);
        all.Style.Font.FontName = "宋体";
        all.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        all.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        all.Style.Alignment.WrapText = true;
        all.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        all.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        sheet.Range(3, 1, days + 2, 1).Style.DateFormat.Format = "m/d";
        sheet.Range(3, 1, Math.Min(7, days + 2), 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC000");
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
        }

        WriteTotalRow(sheet, layout, employeeRows);
        _log($"已写入考勤主表“{sheet.Name}”：{written} 个日期单元格。");
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

        var preserveExistingRows = IsSameOvertimeMonth(headerRow, dateColumns, month);
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

            if (preserveExistingRows && (IsManualOvertimeRow(row) || HasExistingOvertimeValues(row, dateColumns.Values)))
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

    private static bool IsSameOvertimeMonth(IXLRow headerRow, Dictionary<int, int> dateColumns, YearMonth month)
    {
        return dateColumns.Values
            .Select(col => AttendanceReader.TryGetDate(headerRow.Cell(col)))
            .Any(date => date.HasValue && date.Value.Year == month.Year && date.Value.Month == month.Month);
    }

    private static bool HasExistingOvertimeValues(IXLRow row, IEnumerable<int> dateColumns)
    {
        return dateColumns.Any(col => !string.IsNullOrWhiteSpace(AttendanceReader.CellText(row.Cell(col))));
    }

    private static bool IsManualOvertimeRow(IXLRow row)
    {
        var standardText = AttendanceReader.CellText(row.Cell(34));
        return decimal.TryParse(standardText, NumberStyles.Any, CultureInfo.CurrentCulture, out var standard)
            && Math.Abs(standard - 21.72m) < 0.005m;
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
        sheet.Column(layout.DateStartColumn).Width = 5.25;
        sheet.Column(layout.DateStartColumn + 1).Width = 5.25;
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

        sheet.Column(2).Width = 16.0666666666667;
        sheet.Column(3).Width = 16.0666666666667;

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
