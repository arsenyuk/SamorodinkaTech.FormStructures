using System.Text.Json.Serialization;

namespace SamorodinkaTech.FormStructures.Web.Models;

public sealed record FormStructure
{
    public required string FormNumber { get; init; }
    public string? TemplateFormNumber { get; init; }
    public required string FormTitle { get; init; }

    public required int Version { get; init; }
    public required DateTime UploadedAtUtc { get; init; }

    public required IReadOnlyList<HeaderNode> Header { get; init; }
    public required IReadOnlyList<ColumnDefinition> Columns { get; init; }

    public required string StructureHash { get; init; }

    [JsonIgnore]
    public string? SourceFileName { get; init; }
}
