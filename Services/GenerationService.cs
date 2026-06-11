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
        templatePath = ResolveTemplatePath(templatePath, attendancePath, overtimePath, outputDirectory);
        ValidateFile(templatePath, "模板文件");
        var outputPath = Path.Combine(outputDirectory, $"{month.Year}年{month.Month:D2}月考勤统计表_生成版.xlsx");
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
        var exceptions = new List<AttendanceException>();
        var ruleEngine = new AttendanceRuleEngine(configDirectory, LogAll, selectedHolidays);
        var attendanceReader = new AttendanceReader(ruleEngine, LogAll);
        var overtimeReader = new OvertimeReader(LogAll);
        var writer = new TemplateWriter(LogAll, configDirectory);

        var attendance = attendanceReader.Read(attendancePath, month, exceptions);
        var overtime = overtimeReader.Read(overtimePath, month, exceptions);
        writer.Generate(templatePath, outputPath, month, selectedHolidays?.ToHashSet() ?? [], attendance, overtime, exceptions);
        LogAll($"生成成功：{outputPath}");
        return outputPath;
    }

    private static string ResolveTemplatePath(string? templatePath, string attendancePath, string overtimePath, string outputDirectory)
    {
        if (!string.IsNullOrWhiteSpace(templatePath))
        {
            return templatePath;
        }

        var searchDirectories = new[]
        {
            outputDirectory,
            Path.GetDirectoryName(attendancePath),
            Path.GetDirectoryName(overtimePath),
            AppContext.BaseDirectory,
            Environment.CurrentDirectory
        }
        .Where(p => !string.IsNullOrWhiteSpace(p) && Directory.Exists(p))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

        var fixedTemplate = Path.Combine(AppContext.BaseDirectory, "Config", "考勤统计模板.xlsx");
        if (File.Exists(fixedTemplate))
        {
            return fixedTemplate;
        }

        var preferredNames = new[] { "考勤统计模板.xlsx", "AttendanceTemplate.xlsx" };
        foreach (var directory in searchDirectories)
        {
            foreach (var name in preferredNames)
            {
                var candidate = Path.Combine(directory!, name);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        foreach (var directory in searchDirectories)
        {
            var candidate = Directory.EnumerateFiles(directory!, "*.xlsx")
                .Where(p => Path.GetFileName(p).Contains("考勤统计表", StringComparison.Ordinal)
                    || Path.GetFileName(p).Contains("考勤统计模板", StringComparison.Ordinal))
                .Where(p => !Path.GetFileName(p).Contains("生成版", StringComparison.Ordinal)
                    && !Path.GetFileName(p).Contains("微测检测_考勤报表", StringComparison.Ordinal))
                .OrderByDescending(File.GetLastWriteTime)
                .FirstOrDefault();
            if (candidate != null)
            {
                return candidate;
            }
        }

        throw new FileNotFoundException("未选择模板文件，且未能自动找到“考勤统计模板.xlsx”或类似“考勤统计表.xlsx”的模板。");
    }

    private static void ValidateFile(string path, string displayName)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new FileNotFoundException($"请选择有效的{displayName}。", path);
        }
    }
}
