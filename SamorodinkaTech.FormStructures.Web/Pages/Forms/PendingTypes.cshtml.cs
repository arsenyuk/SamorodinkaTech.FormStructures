using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Forms;

public class PendingTypesModel : PageModel
{
    private readonly FormStorage _storage;
    private readonly ILogger<PendingTypesModel> _logger;

    public PendingTypesModel(FormStorage storage, ILogger<PendingTypesModel> logger)
    {
        _storage = storage;
        _logger = logger;
    }

    public string FormNumber { get; private set; } = string.Empty;
    public string PendingId { get; private set; } = string.Empty;
    public int IntendedVersion { get; private set; }

    public FormStorage.PendingUpload? PendingUpload { get; private set; }

    public IReadOnlyList<ReferenceBook> PendingReferenceBooks { get; private set; } = Array.Empty<ReferenceBook>();

    [BindProperty]
    public List<ColumnTypeEditRow> TypeEdits { get; set; } = new();

    public IActionResult OnGet(string formNumber, string pendingId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(pendingId))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        PendingId = pendingId;

        PendingUpload = _storage.TryLoadPending(FormNumber, PendingId);
        if (PendingUpload is null)
        {
            return NotFound();
        }

        PendingReferenceBooks = _storage.TryLoadPendingReferenceBooks(FormNumber, PendingId);

        IntendedVersion = PendingUpload.Meta.IntendedVersion;

        var inferredRefBookPaths = InferReferenceBookColumnPaths(PendingUpload.Structure, PendingReferenceBooks);

        TypeEdits = PendingUpload.Structure.Columns
            .OrderBy(c => c.Index)
            .Select(c => new ColumnTypeEditRow
            {
                Path = c.Path,
                Type = inferredRefBookPaths.Contains(c.Path)
                    ? ColumnType.ReferenceBook
                    : c.Type
            })
            .ToList();

        return Page();
    }

    private static HashSet<string> InferReferenceBookColumnPaths(FormStructure structure, IReadOnlyList<ReferenceBook> books)
    {
        var leafCols = ExtractLeafColumns(structure.Header);
        if (leafCols.Count != structure.Columns.Count)
        {
            return new HashSet<string>(StringComparer.Ordinal);
        }

        var refLeafCols = new HashSet<int>();

        foreach (var b in books)
        {
            foreach (var applied in b.AppliedTo)
            {
                if (!TryParseA1ColumnSpan(applied, out var colStart, out var colEnd))
                {
                    continue;
                }

                for (var i = 0; i < leafCols.Count; i++)
                {
                    var leafCol = leafCols[i];
                    if (leafCol >= colStart && leafCol <= colEnd)
                    {
                        refLeafCols.Add(leafCol);
                    }
                }
            }
        }

        var result = new HashSet<string>(StringComparer.Ordinal);
        for (var i = 0; i < structure.Columns.Count; i++)
        {
            if (refLeafCols.Contains(leafCols[i]))
            {
                result.Add(structure.Columns[i].Path);
            }
        }

        return result;
    }

    private static List<int> ExtractLeafColumns(IReadOnlyList<HeaderNode> roots)
    {
        var cols = new List<int>();

        void Walk(HeaderNode node)
        {
            if (node.Children.Count == 0)
            {
                cols.Add(node.ColStart);
                return;
            }

            foreach (var child in node.Children.OrderBy(c => c.ColStart).ThenBy(c => c.RowStart))
            {
                Walk(child);
            }
        }

        foreach (var root in roots.OrderBy(r => r.ColStart).ThenBy(r => r.RowStart))
        {
            Walk(root);
        }

        return cols;
    }

    private static bool TryParseA1ColumnSpan(string appliedTo, out int colStart, out int colEnd)
    {
        colStart = 0;
        colEnd = 0;

        if (string.IsNullOrWhiteSpace(appliedTo))
        {
            return false;
        }

        // Format examples: "Sheet1!$C$4:$C$100", "Sheet1!C4:D10".
        var s = appliedTo;
        var bang = s.IndexOf('!');
        if (bang >= 0)
        {
            s = s[(bang + 1)..];
        }

        var parts = s.Split(':', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0)
        {
            return false;
        }

        if (!TryParseA1Col(parts[0], out colStart))
        {
            return false;
        }

        colEnd = colStart;
        if (parts.Length >= 2 && TryParseA1Col(parts[1], out var end))
        {
            colEnd = end;
        }

        if (colEnd < colStart)
        {
            (colStart, colEnd) = (colEnd, colStart);
        }

        return colStart > 0 && colEnd > 0;
    }

    private static bool TryParseA1Col(string a1, out int col)
    {
        col = 0;

        if (string.IsNullOrWhiteSpace(a1))
        {
            return false;
        }

        var s = a1.Trim();
        if (s.StartsWith("$", StringComparison.Ordinal))
        {
            s = s[1..];
        }

        // Read leading column letters.
        var i = 0;
        while (i < s.Length && char.IsLetter(s[i]))
        {
            i++;
        }

        if (i == 0)
        {
            return false;
        }

        var letters = s[..i].ToUpperInvariant();
        var value = 0;
        foreach (var ch in letters)
        {
            if (ch < 'A' || ch > 'Z')
            {
                return false;
            }

            value = (value * 26) + (ch - 'A' + 1);
        }

        col = value;
        return col > 0;
    }

    public async Task<IActionResult> OnPostSaveTypesAsync(string formNumber, string pendingId, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(pendingId))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        PendingId = pendingId;

        PendingUpload = _storage.TryLoadPending(FormNumber, PendingId);
        if (PendingUpload is null)
        {
            return NotFound();
        }

        PendingReferenceBooks = _storage.TryLoadPendingReferenceBooks(FormNumber, PendingId);

        IntendedVersion = PendingUpload.Meta.IntendedVersion;

        if (TypeEdits.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "No column edits provided.");
            return Page();
        }

        var typeByPath = TypeEdits
            .Where(x => !string.IsNullOrWhiteSpace(x.Path))
            .GroupBy(x => x.Path, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.Last().Type, StringComparer.Ordinal);

        var updatedColumns = PendingUpload.Structure.Columns
            .Select(c => typeByPath.TryGetValue(c.Path, out var t) ? c with { Type = t } : c)
            .ToArray();

        var finalStructure = PendingUpload.Structure with { Columns = updatedColumns };

        try
        {
            await _storage.CommitPendingAsync(FormNumber, PendingId, finalStructure, ct);
            TempData["SaveMessage"] = "Column types saved. New schema version created.";
            return Redirect($"/forms/{Uri.EscapeDataString(FormNumber)}/v{finalStructure.Version}");
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Failed to commit pending upload (types) {FormNumber} ({PendingId}). ExceptionChain={ExceptionChain}",
                FormNumber,
                PendingId,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, $"Failed to create new schema version from pending upload. {ExceptionUtil.FormatExceptionChain(ex)}");
            return Page();
        }
    }

    public IActionResult OnPostCancel(string formNumber, string pendingId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(pendingId))
        {
            return NotFound();
        }

        _storage.DeletePending(formNumber, pendingId);
        TempData["UploadMessage"] = "Upload cancelled.";
        return Redirect("/settings/forms");
    }

    public sealed class ColumnTypeEditRow
    {
        public string Path { get; set; } = string.Empty;
        public ColumnType Type { get; set; } = ColumnType.String;
    }
}
