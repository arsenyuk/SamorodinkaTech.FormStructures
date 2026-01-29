using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Forms;

public class PendingModel : PageModel
{
    private readonly FormStorage _storage;
    private readonly ILogger<PendingModel> _logger;

    public PendingModel(FormStorage storage, ILogger<PendingModel> logger)
    {
        _storage = storage;
        _logger = logger;
    }

    public string FormNumber { get; private set; } = string.Empty;
    public string PendingId { get; private set; } = string.Empty;
    public int IntendedVersion { get; private set; }
    public int MapFromVersion { get; private set; }

    public FormStorage.PendingUpload? PendingUpload { get; private set; }
    public FormStructure? PreviousStructure { get; private set; }

    public string? MappingWarning { get; private set; }

    [BindProperty]
    public List<ColumnMapEditRow> MapEdits { get; set; } = new();

    public IActionResult OnGet(string formNumber, string pendingId, int mapFrom)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(pendingId) || mapFrom <= 0)
        {
            return NotFound();
        }

        FormNumber = formNumber;
        PendingId = pendingId;
        MapFromVersion = mapFrom;

        PendingUpload = _storage.TryLoadPending(FormNumber, PendingId);
        if (PendingUpload is null)
        {
            return NotFound();
        }

        IntendedVersion = PendingUpload.Meta.IntendedVersion;

        PreviousStructure = _storage.TryLoadStructure(FormNumber, MapFromVersion);
        if (PreviousStructure is null)
        {
            ModelState.AddModelError(string.Empty, $"Previous schema v{MapFromVersion} not found.");
            return Page();
        }

        var prevByPath = PreviousStructure.Columns.ToDictionary(c => c.Path, StringComparer.Ordinal);

        MapEdits = PendingUpload.Structure.Columns
            .OrderBy(c => c.Index)
            .Select(c => new ColumnMapEditRow
            {
                NewPath = c.Path,
                FromPath = prevByPath.ContainsKey(c.Path) ? c.Path : string.Empty,
                Type = c.Type
            })
            .ToList();

        SetWarningFromEdits();

        return Page();
    }

    public async Task<IActionResult> OnPostSaveMappingAsync(string formNumber, string pendingId, int mapFrom, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(pendingId) || mapFrom <= 0)
        {
            return NotFound();
        }

        FormNumber = formNumber;
        PendingId = pendingId;
        MapFromVersion = mapFrom;

        PendingUpload = _storage.TryLoadPending(FormNumber, PendingId);
        if (PendingUpload is null)
        {
            return NotFound();
        }

        IntendedVersion = PendingUpload.Meta.IntendedVersion;

        PreviousStructure = _storage.TryLoadStructure(FormNumber, MapFromVersion);
        if (PreviousStructure is null)
        {
            return NotFound();
        }

        var prevByPath = PreviousStructure.Columns.ToDictionary(c => c.Path, StringComparer.Ordinal);

        SetWarningFromEdits();

        var chosen = MapEdits
            .Select(e => e.FromPath)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .ToArray();

        var duplicates = chosen
            .GroupBy(p => p, StringComparer.Ordinal)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToArray();

        if (duplicates.Length > 0)
        {
            ModelState.AddModelError(string.Empty, "Each previous column can be mapped only once.");
        }

        foreach (var e in MapEdits)
        {
            if (!string.IsNullOrWhiteSpace(e.FromPath) && !prevByPath.ContainsKey(e.FromPath))
            {
                ModelState.AddModelError(string.Empty, $"Invalid mapping target: '{e.FromPath}'.");
                break;
            }
        }

        if (!ModelState.IsValid)
        {
            return Page();
        }

        var mapByNewPath = MapEdits
            .Where(e => !string.IsNullOrWhiteSpace(e.NewPath))
            .GroupBy(e => e.NewPath, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.Last(), StringComparer.Ordinal);

        var updatedColumns = PendingUpload.Structure.Columns
            .Select(c =>
            {
                if (!mapByNewPath.TryGetValue(c.Path, out var edit))
                {
                    return c;
                }

                if (!string.IsNullOrWhiteSpace(edit.FromPath) && prevByPath.TryGetValue(edit.FromPath, out var prev))
                {
                    return c with { Type = prev.Type };
                }

                return c with { Type = edit.Type };
            })
            .ToArray();

        var finalStructure = PendingUpload.Structure with { Columns = updatedColumns };

        try
        {
            await _storage.CommitPendingAsync(FormNumber, PendingId, finalStructure, ct);
            TempData["SaveMessage"] = "Column mapping saved. New schema version created.";
            return Redirect($"/forms/{Uri.EscapeDataString(FormNumber)}/v{finalStructure.Version}");
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Failed to commit pending upload {FormNumber} ({PendingId}). ExceptionChain={ExceptionChain}",
                FormNumber,
                PendingId,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, $"Failed to create new schema version from pending upload. {ExceptionUtil.FormatExceptionChain(ex)}");
            return Page();
        }
    }

    private void SetWarningFromEdits()
    {
        var mappedCount = MapEdits.Count(e => !string.IsNullOrWhiteSpace(e.FromPath));
        MappingWarning = mappedCount == 0
            ? "Warning: no columns are mapped from the previous version. You can still save, but all columns will use the selected 'Type (if unmapped)'."
            : null;
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

    public sealed class ColumnMapEditRow
    {
        public string NewPath { get; set; } = string.Empty;
        public string FromPath { get; set; } = string.Empty;
        public ColumnType Type { get; set; } = ColumnType.String;
    }
}
