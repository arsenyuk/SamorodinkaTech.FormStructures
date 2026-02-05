namespace SamorodinkaTech.FormStructures.Web.Models;

public sealed record ReferenceBook
{
    public required string Id { get; init; }

    public required string Title { get; init; }

    public required string SourceFormula { get; init; }
    public string? SourceSheet { get; init; }
    public string? SourceRange { get; init; }

    public IReadOnlyList<string> AppliedTo { get; init; } = Array.Empty<string>();

    public required IReadOnlyList<string> Values { get; init; }
}
