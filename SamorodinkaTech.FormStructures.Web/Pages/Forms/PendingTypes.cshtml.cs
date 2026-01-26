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

        IntendedVersion = PendingUpload.Meta.IntendedVersion;

        TypeEdits = PendingUpload.Structure.Columns
            .OrderBy(c => c.Index)
            .Select(c => new ColumnTypeEditRow { Path = c.Path, Type = c.Type })
            .ToList();

        return Page();
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
            _logger.LogError(ex, "Failed to commit pending upload (types) {FormNumber} ({PendingId})", FormNumber, PendingId);
            ModelState.AddModelError(string.Empty, "Failed to create new schema version from pending upload.");
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
