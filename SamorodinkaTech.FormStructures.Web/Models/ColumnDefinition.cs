namespace SamorodinkaTech.FormStructures.Web.Models;

public sealed record ColumnDefinition
{
    public required int Index { get; init; }
    public required string Name { get; init; }
    public required string Path { get; init; }

    /// <summary>
    /// Optional alias to use as the CSV header for this column when exporting.
    /// If null/empty, <see cref="Name"/> is used.
    /// </summary>
    public string? CsvHeaderAlias { get; init; }

    // 1-based Excel column number of this leaf column in the original sheet (A=1, B=2, ...).
    // May be null for older stored schemas.
    public int? ExcelLeafColumn { get; init; }

    public string? ColumnNumber { get; init; }

    public ColumnType Type { get; init; } = ColumnType.String;
}
