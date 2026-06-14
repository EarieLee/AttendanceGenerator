using AttendanceGenerator.Services;

namespace AttendanceGenerator.UI;

public sealed partial class MainForm : Form
{
    public MainForm()
    {
        InitializeComponent();
        monthTextBox.Text = DateTime.Today.AddMonths(-1).ToString("yyyy-MM");
        RefreshHolidayList();
    }

    private void BrowseAttendanceButton_Click(object sender, EventArgs e)
    {
        BrowseFile(attendanceTextBox);
    }

    private void BrowseOvertimeButton_Click(object sender, EventArgs e)
    {
        BrowseFile(overtimeTextBox);
    }

    private async void GenerateButton_Click(object sender, EventArgs e)
    {
        generateButton.Enabled = false;
        logTextBox.Clear();

        try
        {
            var service = new GenerationService(AppendLog);
            var selectedHolidays = GetSelectedHolidays();
            var output = await Task.Run(() => service.Generate(
                attendanceTextBox.Text,
                overtimeTextBox.Text,
                monthTextBox.Text,
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                selectedHolidays));

            MessageBox.Show(this, $"生成成功：\n{output}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"生成失败：{ex}");
            MessageBox.Show(this, ex.Message, "生成失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            generateButton.Enabled = true;
        }
    }

    private void MonthTextBox_TextChanged(object sender, EventArgs e)
    {
        RefreshHolidayList();
    }

    private static void BrowseFile(TextBox target)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            target.Text = dialog.FileName;
        }
    }

    private void AppendLog(string message)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action<string>(AppendLog), message);
            return;
        }

        logTextBox.AppendText(message + Environment.NewLine);
    }

    private void RefreshHolidayList()
    {
        if (!TryParseYearMonth(monthTextBox.Text, out var month))
        {
            return;
        }

        holidayCheckedListBox.Items.Clear();
        var days = DateTime.DaysInMonth(month.Year, month.Month);
        var defaultHolidays = AttendanceRuleEngine.GetDefaultHolidayCandidates(month.Year);
        for (var day = 1; day <= days; day++)
        {
            var date = new DateOnly(month.Year, month.Month, day);
            holidayCheckedListBox.Items.Add(new HolidayListItem(date), defaultHolidays.Contains(date));
        }
    }

    private List<DateOnly> GetSelectedHolidays()
    {
        return holidayCheckedListBox.CheckedItems
            .OfType<HolidayListItem>()
            .Select(item => item.Date)
            .ToList();
    }

    private static bool TryParseYearMonth(string text, out YearMonth month)
    {
        try
        {
            month = YearMonth.Parse(text);
            return true;
        }
        catch
        {
            month = default;
            return false;
        }
    }

    private sealed class HolidayListItem
    {
        public HolidayListItem(DateOnly date)
        {
            Date = date;
        }

        public DateOnly Date { get; }

        public override string ToString()
        {
            return $"{Date:yyyy-MM-dd}  {ToChineseWeekday(Date.DayOfWeek)}";
        }

        private static string ToChineseWeekday(DayOfWeek day)
        {
            return day switch
            {
                DayOfWeek.Monday => "周一",
                DayOfWeek.Tuesday => "周二",
                DayOfWeek.Wednesday => "周三",
                DayOfWeek.Thursday => "周四",
                DayOfWeek.Friday => "周五",
                DayOfWeek.Saturday => "周六",
                DayOfWeek.Sunday => "周日",
                _ => string.Empty
            };
        }
    }

    private void titleLabel_Click(object sender, EventArgs e)
    {

    }
}
