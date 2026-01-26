namespace SamorodinkaTech.FormStructures.Web.Models;

public sealed record FormMeta
{
    public required string DisplayFormNumber { get; init; }
    public required string DisplayFormTitle { get; init; }

    public required DateTime UpdatedAtUtc { get; init; }
}
