namespace AttendanceGenerator.Models;

public sealed class AttendanceRecord
{
    public string Name { get; init; } = string.Empty;
    public string Department { get; init; } = string.Empty;
    public DateOnly Date { get; init; }
    public string RawStatus { get; init; } = string.Empty;
    public string NormalizedStatus { get; set; } = string.Empty;
    public string? LeaveType { get; init; }
    public decimal? LeaveHours { get; init; }
}

public sealed class AttendanceException
{
    public string Name { get; init; } = string.Empty;
    public DateOnly? Date { get; init; }
    public string RawStatus { get; init; } = string.Empty;
    public string Reason { get; init; } = string.Empty;
}
