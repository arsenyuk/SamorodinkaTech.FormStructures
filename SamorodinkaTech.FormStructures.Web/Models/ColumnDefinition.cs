namespace SamorodinkaTech.FormStructures.Web.Models;

public sealed record ColumnDefinition
{
    public required int Index { get; init; }
    public required string Name { get; init; }
    public required string Path { get; init; }

    public string? ColumnNumber { get; init; }

    public ColumnType Type { get; init; } = ColumnType.String;
}
