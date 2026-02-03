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

public sealed class ODataReferenceBook
{
    public required string Id { get; init; }

    public required string FormNumber { get; init; }
    public required int Version { get; init; }

    public bool IsPending { get; init; }
    public string? PendingId { get; init; }

    public string? DisplayFormNumber { get; init; }
    public string? DisplayFormTitle { get; init; }

    public required string ReferenceBookId { get; init; }
    public required string Title { get; init; }

    public required string SourceFormula { get; init; }
    public string? SourceSheet { get; init; }
    public string? SourceRange { get; init; }

    public int ValueCount { get; init; }

    public ICollection<string> AppliedTo { get; init; } = Array.Empty<string>();
    public ICollection<string> Values { get; init; } = Array.Empty<string>();

    public static ODataReferenceBook FromCommitted(
        string formNumber,
        int version,
        FormMeta? meta,
        ReferenceBook book)
    {
        return new ODataReferenceBook
        {
            Id = ODataKeys.ReferenceBookKey(formNumber, version, book.Id),
            FormNumber = formNumber,
            Version = version,
            IsPending = false,
            PendingId = null,
            DisplayFormNumber = meta?.DisplayFormNumber,
            DisplayFormTitle = meta?.DisplayFormTitle,
            ReferenceBookId = book.Id,
            Title = book.Title,
            SourceFormula = book.SourceFormula,
            SourceSheet = book.SourceSheet,
            SourceRange = book.SourceRange,
            ValueCount = book.Values.Count,
            AppliedTo = book.AppliedTo.ToArray(),
            Values = book.Values.ToArray(),
        };
    }

    public static ODataReferenceBook FromPending(
        string formNumber,
        int intendedVersion,
        string pendingId,
        FormMeta? meta,
        ReferenceBook book)
    {
        return new ODataReferenceBook
        {
            Id = ODataKeys.PendingReferenceBookKey(formNumber, pendingId, book.Id),
            FormNumber = formNumber,
            Version = intendedVersion,
            IsPending = true,
            PendingId = pendingId,
            DisplayFormNumber = meta?.DisplayFormNumber,
            DisplayFormTitle = meta?.DisplayFormTitle,
            ReferenceBookId = book.Id,
            Title = book.Title,
            SourceFormula = book.SourceFormula,
            SourceSheet = book.SourceSheet,
            SourceRange = book.SourceRange,
            ValueCount = book.Values.Count,
            AppliedTo = book.AppliedTo.ToArray(),
            Values = book.Values.ToArray(),
        };
    }
}
