using AttendanceGenerator.Models;

namespace AttendanceGenerator.Services;

public sealed class GenerationService
{
    private readonly Action<string> _log;

    public GenerationService(Action<string>? log = null)
    {
        _log = log ?? (_ => { });
    }

    public string Generate(
        string attendancePath,
        string overtimePath,
        string monthText,
        string outputDirectory,
        IEnumerable<DateOnly>? selectedHolidays = null,
        string? templatePath = null)
    {
        ValidateFile(attendancePath, "考勤报表文件");
        ValidateFile(overtimePath, "加班表文件");
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            throw new ArgumentException("请选择输出路径。");
        }

        Directory.CreateDirectory(outputDirectory);
        var month = YearMonth.Parse(monthText);
        var effectiveHolidays = selectedHolidays?.ToHashSet()
            ?? AttendanceRuleEngine.GetDefaultHolidayCandidates(month.Year)
                .Where(date => date.Year == month.Year && date.Month == month.Month)
                .ToHashSet();
        templatePath = string.IsNullOrWhiteSpace(templatePath) ? null : templatePath;
        var desiredOutputPath = Path.Combine(outputDirectory, $"{month.Year}年{month.Month:D2}月考勤统计表_生成版.xlsx");
        var outputPath = GetWritableOutputPath(desiredOutputPath);
        var appDirectory = AppContext.BaseDirectory;
        var configDirectory = Path.Combine(appDirectory, "Config");
        var logDirectory = Path.Combine(appDirectory, "Logs");
        Directory.CreateDirectory(logDirectory);

        using var fileLog = new StreamWriter(Path.Combine(logDirectory, $"run_{DateTime.Now:yyyyMMdd_HHmmss}.log"), append: false);
        void LogAll(string message)
        {
            var line = $"[{DateTime.Now:HH:mm:ss}] {message}";
            _log(line);
            fileLog.WriteLine(line);
            fileLog.Flush();
        }

        LogAll("开始生成考勤统计表。");
        if (!outputPath.Equals(desiredOutputPath, StringComparison.OrdinalIgnoreCase))
        {
            LogAll($"目标文件正在被占用，已改用新文件名：{outputPath}");
        }

        var exceptions = new List<AttendanceException>();
        var ruleEngine = new AttendanceRuleEngine(configDirectory, LogAll, effectiveHolidays);
        var attendanceReader = new AttendanceReader(ruleEngine, LogAll);
        var overtimeReader = new OvertimeReader(LogAll);
        var writer = new TemplateWriter(LogAll, configDirectory);

        var attendance = attendanceReader.Read(attendancePath, month, exceptions);
        var overtime = overtimeReader.Read(overtimePath, month, exceptions);
        writer.Generate(templatePath, outputPath, month, effectiveHolidays, attendance, overtime, exceptions);
        LogAll($"生成成功：{outputPath}");
        return outputPath;
    }

    private static string GetWritableOutputPath(string desiredPath)
    {
        if (!File.Exists(desiredPath))
        {
            return desiredPath;
        }

        try
        {
            File.SetAttributes(desiredPath, FileAttributes.Normal);
            using var stream = File.Open(desiredPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return desiredPath;
        }
        catch (IOException)
        {
            return WithTimestampSuffix(desiredPath);
        }
        catch (UnauthorizedAccessException)
        {
            return WithTimestampSuffix(desiredPath);
        }
    }

    private static string WithTimestampSuffix(string path)
    {
        var directory = Path.GetDirectoryName(path) ?? Environment.CurrentDirectory;
        var name = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);
        return Path.Combine(directory, $"{name}_{DateTime.Now:yyyyMMdd_HHmmss}{extension}");
    }

    private static void ValidateFile(string path, string displayName)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new FileNotFoundException($"请选择有效的{displayName}。", path);
        }
    }
}
