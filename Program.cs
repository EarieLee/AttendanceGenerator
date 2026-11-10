namespace AttendanceGenerator;

using AttendanceGenerator.Services;
#if WINDOWS
using AttendanceGenerator.UI;
#endif

static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        if (args.Length > 0)
        {
            return CommandLineRunner.Run(args, Console.WriteLine, Console.Error.WriteLine);
        }

#if WINDOWS
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
        return 0;
#else
        Console.Error.WriteLine("请在 Windows 上直接启动界面，或使用命令行：generate <考勤报表.xlsx> <加班表.xlsx> <yyyy-MM> <输出目录> [法定日期逗号分隔] [模板路径]");
        return 1;
#endif
    }
}
