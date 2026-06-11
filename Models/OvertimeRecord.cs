namespace AttendanceGenerator.Models;

public sealed class OvertimeRecord
{
    public string Name { get; init; } = string.Empty;
    public DateOnly Date { get; init; }
    public decimal Hours { get; init; }
    public string OvertimeType { get; init; } = string.Empty;
    public string RawText { get; init; } = string.Empty;
}
