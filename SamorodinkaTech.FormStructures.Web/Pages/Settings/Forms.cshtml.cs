using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Settings;

public class FormsModel : PageModel
{
    private readonly FormStorage _storage;
    private readonly FormDataStorage _dataStorage;
    private readonly ExcelFormParser _parser;
    private readonly ILogger<FormsModel> _logger;

    public FormsModel(FormStorage storage, FormDataStorage dataStorage, ExcelFormParser parser, ILogger<FormsModel> logger)
    {
        _storage = storage;
        _dataStorage = dataStorage;
        _parser = parser;
        _logger = logger;
    }

    public IReadOnlyList<FormRow> Forms { get; private set; } = Array.Empty<FormRow>();

    [BindProperty]
    public IFormFile? Upload { get; set; }

    public void OnGet()
    {
        LoadForms();
    }

    public async Task<IActionResult> OnPostAsync(CancellationToken ct)
    {
        if (Upload is null)
        {
            ModelState.AddModelError(nameof(Upload), "Please choose a .xlsx file.");
            LoadForms();
            return Page();
        }

        try
        {
            var result = await _storage.SaveAsync(Upload, _parser, ct);
            if (!result.IsNewVersion)
            {
                TempData["UploadMessage"] = $"No schema changes for {result.FormTitle} (#{result.FormNumber}); current version is v{result.Version}.";
            }
            else if (result.RequiresColumnMapping && result.PendingId is string pendingId)
            {
                TempData["UploadMessage"] = $"Upload staged for {result.FormTitle} (#{result.FormNumber}) v{result.Version}. Please confirm column mapping to create the new version.";
            }
            else
            {
                TempData["UploadMessage"] = $"Stored {result.FormTitle} (#{result.FormNumber}) v{result.Version}.";
            }

            if (!result.IsNewVersion)
            {
                return Redirect($"/forms/{Uri.EscapeDataString(result.FormNumber)}");
            }

            if (result.RequiresTypeSetup && result.PendingId is string typePendingId)
            {
                return Redirect($"/forms/{Uri.EscapeDataString(result.FormNumber)}/pending/{Uri.EscapeDataString(typePendingId)}/types");
            }

            if (result.RequiresColumnMapping && result.PreviousVersion is int prev && result.PendingId is string pendingId2)
            {
                return Redirect($"/forms/{Uri.EscapeDataString(result.FormNumber)}/pending/{Uri.EscapeDataString(pendingId2)}?mapFrom={prev}");
            }

            return Redirect($"/forms/{Uri.EscapeDataString(result.FormNumber)}");
        }
        catch (FormParseException ex)
        {
            _logger.LogWarning(ex, "Failed to parse uploaded file {FileName}", Upload.FileName);
            ModelState.AddModelError(string.Empty, ex.Message);
            LoadForms();
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error while processing uploaded file {FileName}", Upload.FileName);
            ModelState.AddModelError(string.Empty, "Unexpected error while processing the file. See logs for details.");
            LoadForms();
            return Page();
        }
    }

    private void LoadForms()
    {
        var latest = _storage.ListLatestForms();

        Forms = latest
            .Select(f =>
            {
                var lastUpload = _dataStorage.TryGetLatestUpload(f.FormNumber);
                return new FormRow(
                    FormNumber: f.FormNumber,
                    DisplayFormNumber: f.DisplayFormNumber,
                    DisplayFormTitle: f.DisplayFormTitle,
                    LatestVersion: f.Version,
                    LatestUploadedAtUtc: f.UploadedAtUtc,
                    LastDataUpload: lastUpload);
            })
            .OrderBy(f => f.DisplayFormTitle, StringComparer.OrdinalIgnoreCase)
            .ThenBy(f => f.DisplayFormNumber, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public sealed record FormRow(
        string FormNumber,
        string DisplayFormNumber,
        string DisplayFormTitle,
        int LatestVersion,
        DateTime LatestUploadedAtUtc,
        FormDataUpload? LastDataUpload);
}
