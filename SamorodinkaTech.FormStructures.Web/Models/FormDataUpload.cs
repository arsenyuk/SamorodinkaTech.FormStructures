namespace SamorodinkaTech.FormStructures.Web.Models;

public sealed record FormDataRow
{
    public required int RowNumber { get; init; }
    public required IReadOnlyDictionary<string, string?> Values { get; init; }
}

public sealed record FormDataUpload
{
    public required string UploadId { get; init; }

    public required string FormNumber { get; init; }
    public required int FormVersion { get; init; }
    public required string StructureHash { get; init; }

    public required string OriginalFileName { get; init; }
    public required string FileSha256 { get; init; }

    public required DateTime UploadedAtUtc { get; init; }
    public required int RowCount { get; init; }
}

public sealed record FormDataFile
{
    public required FormDataUpload Upload { get; init; }
    public required IReadOnlyList<FormDataRow> Rows { get; init; }
}
