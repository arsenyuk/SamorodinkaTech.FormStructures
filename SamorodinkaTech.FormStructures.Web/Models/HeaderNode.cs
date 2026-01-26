namespace SamorodinkaTech.FormStructures.Web.Models;

public sealed record HeaderNode
{
    public required string Label { get; init; }
    public required int RowStart { get; init; }
    public required int RowEnd { get; init; }
    public required int ColStart { get; init; }
    public required int ColEnd { get; init; }

    public IReadOnlyList<HeaderNode> Children { get; init; } = Array.Empty<HeaderNode>();

    public int ColSpan => ColEnd - ColStart + 1;
}
