namespace AttendanceGenerator.UI;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;
    private TableLayoutPanel rootLayout;
    private Panel headerPanel;
    private Label titleLabel;
    private Label subtitleLabel;
    private Panel inputPanel;
    private TableLayoutPanel inputLayout;
    private Label attendanceLabel;
    private TextBox attendanceTextBox;
    private Button browseAttendanceButton;
    private Label overtimeLabel;
    private TextBox overtimeTextBox;
    private Button browseOvertimeButton;
    private Label monthLabel;
    private TextBox monthTextBox;
    private Button generateButton;
    private GroupBox holidayGroupBox;
    private CheckedListBox holidayCheckedListBox;
    private GroupBox logGroupBox;
    private TextBox logTextBox;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }

        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        rootLayout = new TableLayoutPanel();
        headerPanel = new Panel();
        subtitleLabel = new Label();
        titleLabel = new Label();
        inputPanel = new Panel();
        inputLayout = new TableLayoutPanel();
        attendanceLabel = new Label();
        attendanceTextBox = new TextBox();
        browseAttendanceButton = new Button();
        overtimeLabel = new Label();
        overtimeTextBox = new TextBox();
        browseOvertimeButton = new Button();
        monthLabel = new Label();
        monthTextBox = new TextBox();
        generateButton = new Button();
        holidayGroupBox = new GroupBox();
        holidayCheckedListBox = new CheckedListBox();
        logGroupBox = new GroupBox();
        logTextBox = new TextBox();
        rootLayout.SuspendLayout();
        headerPanel.SuspendLayout();
        inputPanel.SuspendLayout();
        inputLayout.SuspendLayout();
        holidayGroupBox.SuspendLayout();
        logGroupBox.SuspendLayout();
        SuspendLayout();
        // 
        // rootLayout
        // 
        rootLayout.ColumnCount = 1;
        rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        rootLayout.Controls.Add(headerPanel, 0, 0);
        rootLayout.Controls.Add(inputPanel, 0, 1);
        rootLayout.Controls.Add(holidayGroupBox, 0, 2);
        rootLayout.Controls.Add(logGroupBox, 0, 3);
        rootLayout.Dock = DockStyle.Fill;
        rootLayout.Location = new Point(0, 0);
        rootLayout.Name = "rootLayout";
        rootLayout.Padding = new Padding(22);
        rootLayout.RowCount = 4;
        rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));
        rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 152F));
        rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 170F));
        rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        rootLayout.Size = new Size(960, 680);
        rootLayout.TabIndex = 0;
        // 
        // headerPanel
        // 
        headerPanel.Controls.Add(subtitleLabel);
        headerPanel.Controls.Add(titleLabel);
        headerPanel.Dock = DockStyle.Fill;
        headerPanel.Location = new Point(25, 25);
        headerPanel.Name = "headerPanel";
        headerPanel.Size = new Size(910, 66);
        headerPanel.TabIndex = 0;
        // 
        // subtitleLabel
        // 
        subtitleLabel.Dock = DockStyle.Top;
        subtitleLabel.ForeColor = Color.FromArgb(107, 114, 128);
        subtitleLabel.Location = new Point(0, 36);
        subtitleLabel.Name = "subtitleLabel";
        subtitleLabel.Size = new Size(910, 24);
        subtitleLabel.TabIndex = 1;
        subtitleLabel.Text = "选择两个业务文件，勾选法定节假日，生成统计表";
        subtitleLabel.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // titleLabel
        // 
        titleLabel.Dock = DockStyle.Top;
        titleLabel.Font = new Font("Microsoft YaHei UI", 16F, FontStyle.Bold);
        titleLabel.ForeColor = Color.FromArgb(31, 41, 55);
        titleLabel.Location = new Point(0, 0);
        titleLabel.Name = "titleLabel";
        titleLabel.Size = new Size(910, 36);
        titleLabel.TabIndex = 0;
        titleLabel.Text = "肖玲玲是🐖";
        titleLabel.TextAlign = ContentAlignment.MiddleLeft;
        titleLabel.Click += titleLabel_Click;
        // 
        // inputPanel
        // 
        inputPanel.BackColor = Color.White;
        inputPanel.Controls.Add(inputLayout);
        inputPanel.Dock = DockStyle.Fill;
        inputPanel.Location = new Point(25, 97);
        inputPanel.Name = "inputPanel";
        inputPanel.Padding = new Padding(16);
        inputPanel.Size = new Size(910, 146);
        inputPanel.TabIndex = 1;
        // 
        // inputLayout
        // 
        inputLayout.ColumnCount = 3;
        inputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 116F));
        inputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        inputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112F));
        inputLayout.Controls.Add(attendanceLabel, 0, 0);
        inputLayout.Controls.Add(attendanceTextBox, 1, 0);
        inputLayout.Controls.Add(browseAttendanceButton, 2, 0);
        inputLayout.Controls.Add(overtimeLabel, 0, 1);
        inputLayout.Controls.Add(overtimeTextBox, 1, 1);
        inputLayout.Controls.Add(browseOvertimeButton, 2, 1);
        inputLayout.Controls.Add(monthLabel, 0, 2);
        inputLayout.Controls.Add(monthTextBox, 1, 2);
        inputLayout.Controls.Add(generateButton, 2, 2);
        inputLayout.Dock = DockStyle.Fill;
        inputLayout.Location = new Point(16, 16);
        inputLayout.Name = "inputLayout";
        inputLayout.RowCount = 3;
        inputLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 33.3333321F));
        inputLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 33.3333321F));
        inputLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 33.3333321F));
        inputLayout.Size = new Size(878, 114);
        inputLayout.TabIndex = 0;
        // 
        // attendanceLabel
        // 
        attendanceLabel.Dock = DockStyle.Fill;
        attendanceLabel.Location = new Point(3, 0);
        attendanceLabel.Name = "attendanceLabel";
        attendanceLabel.Size = new Size(110, 37);
        attendanceLabel.TabIndex = 0;
        attendanceLabel.Text = "考勤报表";
        attendanceLabel.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // attendanceTextBox
        // 
        attendanceTextBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        attendanceTextBox.Location = new Point(119, 7);
        attendanceTextBox.Name = "attendanceTextBox";
        attendanceTextBox.Size = new Size(644, 23);
        attendanceTextBox.TabIndex = 1;
        // 
        // browseAttendanceButton
        // 
        browseAttendanceButton.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        browseAttendanceButton.Location = new Point(769, 4);
        browseAttendanceButton.Name = "browseAttendanceButton";
        browseAttendanceButton.Size = new Size(106, 29);
        browseAttendanceButton.TabIndex = 2;
        browseAttendanceButton.Text = "浏览...";
        browseAttendanceButton.UseVisualStyleBackColor = true;
        browseAttendanceButton.Click += BrowseAttendanceButton_Click;
        // 
        // overtimeLabel
        // 
        overtimeLabel.Dock = DockStyle.Fill;
        overtimeLabel.Location = new Point(3, 37);
        overtimeLabel.Name = "overtimeLabel";
        overtimeLabel.Size = new Size(110, 37);
        overtimeLabel.TabIndex = 3;
        overtimeLabel.Text = "加班表";
        overtimeLabel.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // overtimeTextBox
        // 
        overtimeTextBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        overtimeTextBox.Location = new Point(119, 44);
        overtimeTextBox.Name = "overtimeTextBox";
        overtimeTextBox.Size = new Size(644, 23);
        overtimeTextBox.TabIndex = 4;
        // 
        // browseOvertimeButton
        // 
        browseOvertimeButton.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        browseOvertimeButton.Location = new Point(769, 41);
        browseOvertimeButton.Name = "browseOvertimeButton";
        browseOvertimeButton.Size = new Size(106, 29);
        browseOvertimeButton.TabIndex = 5;
        browseOvertimeButton.Text = "浏览...";
        browseOvertimeButton.UseVisualStyleBackColor = true;
        browseOvertimeButton.Click += BrowseOvertimeButton_Click;
        // 
        // monthLabel
        // 
        monthLabel.Dock = DockStyle.Fill;
        monthLabel.Location = new Point(3, 74);
        monthLabel.Name = "monthLabel";
        monthLabel.Size = new Size(110, 40);
        monthLabel.TabIndex = 6;
        monthLabel.Text = "月份";
        monthLabel.TextAlign = ContentAlignment.MiddleLeft;
        // 
        // monthTextBox
        // 
        monthTextBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        monthTextBox.Location = new Point(119, 82);
        monthTextBox.Name = "monthTextBox";
        monthTextBox.Size = new Size(644, 23);
        monthTextBox.TabIndex = 7;
        monthTextBox.TextChanged += MonthTextBox_TextChanged;
        // 
        // generateButton
        // 
        generateButton.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        generateButton.BackColor = Color.FromArgb(37, 99, 235);
        generateButton.FlatAppearance.BorderSize = 0;
        generateButton.FlatStyle = FlatStyle.Flat;
        generateButton.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        generateButton.ForeColor = Color.White;
        generateButton.Location = new Point(769, 78);
        generateButton.Name = "generateButton";
        generateButton.Size = new Size(106, 31);
        generateButton.TabIndex = 8;
        generateButton.Text = "生成";
        generateButton.UseVisualStyleBackColor = false;
        generateButton.Click += GenerateButton_Click;
        // 
        // holidayGroupBox
        // 
        holidayGroupBox.BackColor = Color.White;
        holidayGroupBox.Controls.Add(holidayCheckedListBox);
        holidayGroupBox.Dock = DockStyle.Fill;
        holidayGroupBox.ForeColor = Color.FromArgb(31, 41, 55);
        holidayGroupBox.Location = new Point(25, 249);
        holidayGroupBox.Name = "holidayGroupBox";
        holidayGroupBox.Padding = new Padding(10);
        holidayGroupBox.Size = new Size(910, 164);
        holidayGroupBox.TabIndex = 2;
        holidayGroupBox.TabStop = false;
        holidayGroupBox.Text = "法定节假日（请按本次生成月份勾选）";
        // 
        // holidayCheckedListBox
        // 
        holidayCheckedListBox.BorderStyle = BorderStyle.None;
        holidayCheckedListBox.CheckOnClick = true;
        holidayCheckedListBox.ColumnWidth = 170;
        holidayCheckedListBox.Dock = DockStyle.Fill;
        holidayCheckedListBox.FormattingEnabled = true;
        holidayCheckedListBox.HorizontalScrollbar = true;
        holidayCheckedListBox.Location = new Point(10, 26);
        holidayCheckedListBox.MultiColumn = true;
        holidayCheckedListBox.Name = "holidayCheckedListBox";
        holidayCheckedListBox.Size = new Size(890, 128);
        holidayCheckedListBox.TabIndex = 0;
        // 
        // logGroupBox
        // 
        logGroupBox.Controls.Add(logTextBox);
        logGroupBox.Dock = DockStyle.Fill;
        logGroupBox.Location = new Point(25, 419);
        logGroupBox.Name = "logGroupBox";
        logGroupBox.Padding = new Padding(10);
        logGroupBox.Size = new Size(910, 236);
        logGroupBox.TabIndex = 3;
        logGroupBox.TabStop = false;
        logGroupBox.Text = "日志";
        // 
        // logTextBox
        // 
        logTextBox.BackColor = Color.FromArgb(17, 24, 39);
        logTextBox.BorderStyle = BorderStyle.None;
        logTextBox.Dock = DockStyle.Fill;
        logTextBox.Font = new Font("Consolas", 9F);
        logTextBox.ForeColor = Color.FromArgb(229, 231, 235);
        logTextBox.Location = new Point(10, 26);
        logTextBox.Multiline = true;
        logTextBox.Name = "logTextBox";
        logTextBox.ReadOnly = true;
        logTextBox.ScrollBars = ScrollBars.Vertical;
        logTextBox.Size = new Size(890, 200);
        logTextBox.TabIndex = 0;
        // 
        // MainForm
        // 
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        BackColor = Color.FromArgb(246, 248, 251);
        ClientSize = new Size(960, 680);
        Controls.Add(rootLayout);
        Font = new Font("Microsoft YaHei UI", 9F);
        MinimumSize = new Size(960, 680);
        Name = "MainForm";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "AttendanceGenerator - 考勤统计表生成器";
        rootLayout.ResumeLayout(false);
        headerPanel.ResumeLayout(false);
        inputPanel.ResumeLayout(false);
        inputLayout.ResumeLayout(false);
        inputLayout.PerformLayout();
        holidayGroupBox.ResumeLayout(false);
        logGroupBox.ResumeLayout(false);
        logGroupBox.PerformLayout();
        ResumeLayout(false);
    }
}
