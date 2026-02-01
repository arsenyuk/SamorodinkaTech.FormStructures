using SamorodinkaTech.FormStructures.Web.Models;

namespace SamorodinkaTech.FormStructures.Web.OData;

public sealed class ODataUpload
{
    public required string Id { get; init; }

    public required string FormNumber { get; init; }
    public required int Version { get; init; }
    public required string UploadId { get; init; }

    public required string StructureHash { get; init; }

    public required string OriginalFileName { get; init; }
    public required string FileSha256 { get; init; }

    public required DateTime UploadedAtUtc { get; init; }
    public required int RowCount { get; init; }

    public ICollection<ODataRow> Rows { get; init; } = Array.Empty<ODataRow>();
    public ICollection<ODataColumn> Columns { get; init; } = Array.Empty<ODataColumn>();

    public static ODataUpload From(FormDataUpload upload)
    {
        return new ODataUpload
        {
            Id = ODataKeys.UploadKey(upload.FormNumber, upload.FormVersion, upload.UploadId),
            FormNumber = upload.FormNumber,
            Version = upload.FormVersion,
            UploadId = upload.UploadId,
            StructureHash = upload.StructureHash,
            OriginalFileName = upload.OriginalFileName,
            FileSha256 = upload.FileSha256,
            UploadedAtUtc = upload.UploadedAtUtc,
            RowCount = upload.RowCount,
        };
    }
}

public sealed class ODataRow
{
    public required string Id { get; init; }

    public required string UploadKey { get; init; }
    public required string FormNumber { get; init; }
    public required int Version { get; init; }
    public required string UploadId { get; init; }

    public required int RowNumber { get; init; }

    public IDictionary<string, object?> DynamicProperties { get; init; } = new Dictionary<string, object?>(StringComparer.Ordinal);
}

public sealed class ODataColumn
{
    public required string Id { get; init; }

    public required string UploadKey { get; init; }
    public required string FormNumber { get; init; }
    public required int Version { get; init; }

    public required int Index { get; init; }
    public required string ODataProperty { get; init; }

    public required string Name { get; init; }
    public required string Path { get; init; }
    public string? ColumnNumber { get; init; }
    public ColumnType Type { get; init; }
}
