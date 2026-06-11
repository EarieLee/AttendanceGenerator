using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using AttendanceGenerator.Models;

namespace AttendanceGenerator.Services;

public sealed partial class AttendanceRuleEngine
{
    private readonly Dictionary<string, string> _mapping;
    private readonly HolidaySettings _holidaySettings;
    private readonly Action<string> _log;

    private static readonly string[] LeaveTypes =
    [
        "事假", "调休", "年假", "病假", "婚假", "产假", "丧假", "陪产假", "例假"
    ];

    public AttendanceRuleEngine(string configDirectory, Action<string>? log = null, IEnumerable<DateOnly>? selectedHolidays = null)
    {
        _log = log ?? (_ => { });
        Directory.CreateDirectory(configDirectory);
        var mappingPath = Path.Combine(configDirectory, "mapping.json");
        var holidayPath = Path.Combine(configDirectory, "holiday.json");

        EnsureDefaultFiles(mappingPath, holidayPath);
        _mapping = LoadMapping(mappingPath);
        _holidaySettings = selectedHolidays == null
            ? LoadHolidays(holidayPath)
            : new HolidaySettings("Manual", selectedHolidays.ToHashSet(), []);
    }

    public static IReadOnlySet<DateOnly> GetDefaultHolidayCandidates(int year)
    {
        return GetAutoChineseStatutoryHolidays(year);
    }

    public string Normalize(AttendanceRecord record, IList<AttendanceException> exceptions)
    {
        if (IsHoliday(record.Date))
        {
            return "法定";
        }

        var raw = NormalizeText(record.RawStatus);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return string.Empty;
        }

        if (raw == "/" || raw.Contains("未入职", StringComparison.Ordinal) || raw.Contains("离职", StringComparison.Ordinal))
        {
            return "/";
        }

        if (raw.StartsWith("休息", StringComparison.Ordinal))
        {
            return "休";
        }

        var leaveParts = ExtractLeaveParts(raw, record.Date).ToList();
        if (leaveParts.Count > 0)
        {
            return string.Join("     ", leaveParts);
        }

        foreach (var pair in _mapping)
        {
            if (raw.Contains(pair.Key, StringComparison.Ordinal))
            {
                return pair.Value;
            }
        }

        if (raw.Contains("未打卡", StringComparison.Ordinal) || raw.Contains("缺卡", StringComparison.Ordinal))
        {
            return "缺卡";
        }

        if (raw.Contains("迟到", StringComparison.Ordinal))
        {
            return "迟到";
        }

        if (raw.Contains("早退", StringComparison.Ordinal))
        {
            return "早退";
        }

        if (raw.Contains("旷工", StringComparison.Ordinal))
        {
            return "旷工";
        }

        if (raw.Contains("外出", StringComparison.Ordinal)
            || raw.Contains("出差", StringComparison.Ordinal))
        {
            return "公司出勤";
        }

        if (raw.Contains("正常", StringComparison.Ordinal)
            || raw.Contains("普通出勤", StringComparison.Ordinal)
            || raw.Contains("管理人员考勤", StringComparison.Ordinal)
            || raw.Contains("加班", StringComparison.Ordinal))
        {
            return "公司出勤";
        }

        exceptions.Add(new AttendanceException
        {
            Name = record.Name,
            Date = record.Date,
            RawStatus = record.RawStatus,
            Reason = "无法识别的考勤状态"
        });
        return raw;
    }

    public static string CleanName(string value)
    {
        var cleaned = NameNoiseRegex().Replace(value.Trim(), string.Empty);
        cleaned = cleaned.Trim();
        if (cleaned.EndsWith('2') && cleaned.Length > 1)
        {
            cleaned = cleaned[..^1];
        }

        return cleaned;
    }

    private static decimal? ExtractLeaveHours(string raw, string leaveType, DateOnly currentDate)
    {
        var match = Regex.Match(raw, $"{Regex.Escape(leaveType)}(?<start>\\d{{2}}-\\d{{2}})\\s+\\d{{1,2}}:\\d{{2}}到(?<end>\\d{{2}}-\\d{{2}})\\s+\\d{{1,2}}:\\d{{2}}\\s+(?<num>\\d+(?:\\.\\d+)?)(?<unit>小时|天|H|h)");
        if (match.Success)
        {
            var total = decimal.Parse(match.Groups["num"].Value, CultureInfo.InvariantCulture);
            if (match.Groups["unit"].Value.Contains("天", StringComparison.Ordinal))
            {
                total *= 8;
            }

            var start = ParseMonthDay(currentDate.Year, match.Groups["start"].Value);
            var end = ParseMonthDay(currentDate.Year, match.Groups["end"].Value);
            if (start.HasValue && end.HasValue && end.Value >= start.Value)
            {
                if (currentDate == end.Value && total > 8)
                {
                    var remainder = total % 8;
                    return remainder == 0 ? 8 : remainder;
                }

                if (currentDate > start.Value && currentDate < end.Value)
                {
                    return Math.Min(8, total);
                }

                if (currentDate == start.Value && total > 8)
                {
                    return 8;
                }
            }

            return total;
        }

        match = Regex.Match(raw, $"{Regex.Escape(leaveType)}.*?(?<num>\\d+(?:\\.\\d+)?)(?<unit>小时|天|H|h)");
        if (!match.Success)
        {
            return null;
        }

        var number = decimal.Parse(match.Groups["num"].Value, CultureInfo.InvariantCulture);
        var unit = match.Groups["unit"].Value;
        return unit.Contains("天", StringComparison.Ordinal) ? number * 8 : number;
    }

    private IEnumerable<string> ExtractLeaveParts(string raw, DateOnly currentDate)
    {
        var matches = LeaveTypes
            .Select(leaveType => new
            {
                LeaveType = leaveType,
                Index = raw.IndexOf(leaveType, StringComparison.Ordinal),
                Hours = ExtractLeaveHours(raw, leaveType, currentDate),
                IsDayUnit = Regex.IsMatch(raw, $@"{Regex.Escape(leaveType)}.*?\d+(?:\.\d+)?天")
            })
            .Where(item => item.Index >= 0)
            .OrderBy(item => item.Index);

        var matchList = matches.ToList();
        var hasMultipleLeaveParts = matchList.Count > 1;
        foreach (var item in matchList)
        {
            var mappedType = hasMultipleLeaveParts && item.LeaveType == "事假"
                ? "请假"
                : _mapping.TryGetValue(item.LeaveType, out var mapped) ? mapped : item.LeaveType;
            if (item.Hours > 0)
            {
                if (item.IsDayUnit && item.Hours.Value % 8 == 0)
                {
                    yield return $"{mappedType}{FormatHours(item.Hours.Value / 8)}天";
                }
                else
                {
                    yield return $"{mappedType}{FormatHours(item.Hours.Value)}H";
                }
            }
            else
            {
                yield return mappedType;
            }
        }
    }

    private static DateOnly? ParseMonthDay(int year, string value)
    {
        var parts = value.Split('-', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length == 2
            && int.TryParse(parts[0], out var month)
            && int.TryParse(parts[1], out var day)
            ? new DateOnly(year, month, day)
            : null;
    }

    private static string FormatHours(decimal hours)
    {
        return hours.ToString("0.##", CultureInfo.InvariantCulture);
    }

    private static string NormalizeText(string value)
    {
        return value.Replace("\r", " ").Replace("\n", " ").Trim();
    }

    private static Dictionary<string, string> LoadMapping(string path)
    {
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();
    }

    private bool IsHoliday(DateOnly date)
    {
        if (_holidaySettings.Workdays.Contains(date))
        {
            return false;
        }

        if (_holidaySettings.ExtraHolidays.Contains(date))
        {
            return true;
        }

        return _holidaySettings.Mode.Equals("AutoChineseStatutory", StringComparison.OrdinalIgnoreCase)
            && GetAutoChineseStatutoryHolidays(date.Year).Contains(date);
    }

    private static HolidaySettings LoadHolidays(string path)
    {
        var json = File.ReadAllText(path);
        using var document = JsonDocument.Parse(json);
        if (document.RootElement.ValueKind == JsonValueKind.Array)
        {
            return new HolidaySettings(
                "Manual",
                ReadDateArray(document.RootElement),
                []);
        }

        var mode = document.RootElement.TryGetProperty("mode", out var modeElement)
            ? modeElement.GetString() ?? "AutoChineseStatutory"
            : "AutoChineseStatutory";
        var extra = document.RootElement.TryGetProperty("extraHolidays", out var extraElement)
            ? ReadDateArray(extraElement)
            : [];
        var workdays = document.RootElement.TryGetProperty("workdays", out var workdayElement)
            ? ReadDateArray(workdayElement)
            : [];
        return new HolidaySettings(mode, extra, workdays);
    }

    private static HashSet<DateOnly> ReadDateArray(JsonElement element)
    {
        return element.ValueKind == JsonValueKind.Array
            ? element.EnumerateArray()
                .Select(v => v.GetString())
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Select(v => DateOnly.TryParse(v, out var date) ? date : (DateOnly?)null)
                .Where(d => d.HasValue)
                .Select(d => d!.Value)
                .ToHashSet()
            : [];
    }

    private static HashSet<DateOnly> GetAutoChineseStatutoryHolidays(int year)
    {
        var holidays = new HashSet<DateOnly>
        {
            new(year, 1, 1),
            new(year, 5, 1),
            new(year, 5, 2),
            new(year, 5, 3),
            new(year, 10, 1),
            new(year, 10, 2),
            new(year, 10, 3),
            QingMing(year)
        };

        AddLunarHoliday(holidays, year, 1, 1, 3);
        AddLunarHoliday(holidays, year, 5, 5, 1);
        AddLunarHoliday(holidays, year, 8, 15, 1);
        return holidays;
    }

    private static void AddLunarHoliday(HashSet<DateOnly> holidays, int year, int lunarMonth, int lunarDay, int durationDays)
    {
        try
        {
            var calendar = new ChineseLunisolarCalendar();
            var leapMonth = calendar.GetLeapMonth(year);
            var month = leapMonth > 0 && lunarMonth >= leapMonth ? lunarMonth + 1 : lunarMonth;
            var start = new DateOnly(calendar.ToDateTime(year, month, lunarDay, 0, 0, 0, 0).Year,
                calendar.ToDateTime(year, month, lunarDay, 0, 0, 0, 0).Month,
                calendar.ToDateTime(year, month, lunarDay, 0, 0, 0, 0).Day);
            for (var i = 0; i < durationDays; i++)
            {
                holidays.Add(start.AddDays(i));
            }
        }
        catch
        {
            // Lunar conversion can fail outside supported calendar ranges. In that case only config dates apply.
        }
    }

    private static DateOnly QingMing(int year)
    {
        var y = year % 100;
        var day = (int)(y * 0.2422 + 4.81) - y / 4;
        return new DateOnly(year, 4, day);
    }

    private void EnsureDefaultFiles(string mappingPath, string holidayPath)
    {
        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        if (!File.Exists(mappingPath))
        {
            var defaults = new Dictionary<string, string>
            {
                ["正常出勤"] = "公司出勤",
                ["普通出勤"] = "公司出勤",
                ["正常"] = "公司出勤",
                ["休息"] = "休",
                ["休"] = "休",
                ["法定节假日"] = "法定",
                ["法定"] = "法定",
                ["事假"] = "事假",
                ["调休"] = "调休",
                ["年假"] = "年假",
                ["病假"] = "病假",
                ["婚假"] = "婚假",
                ["产假"] = "产假",
                ["丧假"] = "丧假",
                ["旷工"] = "旷工",
                ["迟到"] = "迟到",
                ["早退"] = "早退",
                ["缺卡"] = "缺卡"
            };
            File.WriteAllText(mappingPath, JsonSerializer.Serialize(defaults, options));
            _log($"已创建默认映射配置：{mappingPath}");
        }

        if (!File.Exists(holidayPath))
        {
            var defaultHolidaySettings = new
            {
                mode = "AutoChineseStatutory",
                extraHolidays = Array.Empty<string>(),
                workdays = Array.Empty<string>()
            };
            File.WriteAllText(holidayPath, JsonSerializer.Serialize(defaultHolidaySettings, options));
            _log($"已创建默认节假日配置：{holidayPath}");
        }
    }

    [GeneratedRegex(@"（.*?）|\(.*?\)|\[.*?\]")]
    private static partial Regex NameNoiseRegex();

    private sealed record HolidaySettings(string Mode, HashSet<DateOnly> ExtraHolidays, HashSet<DateOnly> Workdays);
}
