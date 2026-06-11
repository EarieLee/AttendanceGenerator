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

    public TemplateWriter(Action<string>? log = null, string? configDirectory = null)
    {
        _log = log ?? (_ => { });
        _workCalendar = LoadWorkCalendar(configDirectory);
    }

    public void Generate(
        string templatePath,
        string outputPath,
        YearMonth month,
        IReadOnlySet<DateOnly> selectedHolidays,
        IReadOnlyList<AttendanceRecord> attendanceRecords,
        IReadOnlyList<OvertimeRecord> overtimeRecords,
        IReadOnlyList<AttendanceException> exceptions)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        File.Copy(templatePath, outputPath, overwrite: true);
        File.SetAttributes(outputPath, FileAttributes.Normal);
        _log($"已复制模板到输出文件：{outputPath}");

        using var workbook = new XLWorkbook(outputPath);
        WriteAttendance(workbook, month, selectedHolidays, attendanceRecords, exceptions);
        WriteOvertime(workbook, month, overtimeRecords);
        RewriteNightShiftSheet(workbook, month);
        WriteExceptionSheet(workbook, exceptions);
        workbook.CalculateMode = XLCalculateMode.Auto;
        workbook.RecalculateAllFormulas();
        workbook.Save();
        _log("Excel 生成完成。");
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
                if (IsOutOfEmployment(row, layout, date))
                {
                    status = "/";
                }
                else if (sameTemplateMonth && IsReusableManualStatus(existing))
                {
                    status = existing;
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

        RewriteOvertimeDateHeaders(sheet, headerRow, dateColumns, month);
        var nameColumn = Math.Max(1, dateColumns.Values.Min() - 1);
        var byNameDate = records
            .GroupBy(r => (AttendanceRuleEngine.CleanName(r.Name), r.Date))
            .ToDictionary(g => g.Key, g => g.Sum(r => r.Hours));

        var written = 0;
        for (var rowNumber = headerRow.RowNumber() + 1; rowNumber <= used.LastRow().RowNumber(); rowNumber++)
        {
            var row = sheet.Row(rowNumber);
            var originalName = AttendanceReader.CellText(row.Cell(nameColumn));
            var name = AttendanceRuleEngine.CleanName(originalName);
            if (string.IsNullOrWhiteSpace(name) || !ContainsChinese(name))
            {
                continue;
            }

            foreach (var col in dateColumns.Values)
            {
                row.Cell(col).Clear(XLClearOptions.Contents);
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

    private static IXLWorksheet? FindAttendanceSheet(XLWorkbook workbook, YearMonth month)
    {
        var monthText = $"{month.Month}月";
        return workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains(monthText, StringComparison.Ordinal)
                && ws.Name.Contains("考勤统计", StringComparison.Ordinal))
            ?? workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains("5月", StringComparison.Ordinal)
                && ws.Name.Contains("考勤统计", StringComparison.Ordinal))
            ?? workbook.Worksheets.FirstOrDefault(ws => ws.Name.Contains("考勤统计", StringComparison.Ordinal));
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

        if (source is "迟到" or "早退" or "缺卡")
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
        var end = ParseRemarkDate(remark, "离职", date.Year) ?? ParseRemarkDate(remark, "自离", date.Year);
        if (start.HasValue && date < start.Value)
        {
            return true;
        }

        if (end.HasValue && date >= end.Value)
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
        var column = layout.StatColumns.FirstOrDefault(s => s.Key.Contains(name, StringComparison.Ordinal)).Value;
        if (column > 0)
        {
            row.Cell(column).Value = value == 0 && !writeZero ? string.Empty : value;
        }
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
