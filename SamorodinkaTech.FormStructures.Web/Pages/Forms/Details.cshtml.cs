using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Forms;

public class DetailsModel : PageModel
{
    private readonly FormStorage _storage;
    private readonly FormDataStorage _dataStorage;
    private readonly ExcelFormParser _parser;
    private readonly ILogger<DetailsModel> _logger;

    public DetailsModel(FormStorage storage, FormDataStorage dataStorage, ExcelFormParser parser, ILogger<DetailsModel> logger)
    {
        _storage = storage;
        _dataStorage = dataStorage;
        _parser = parser;
        _logger = logger;
    }

    public string FormNumber { get; private set; } = string.Empty;
    public FormMeta? Meta { get; private set; }
    public FormStructure? Latest { get; private set; }
    public IReadOnlyList<int> Versions { get; private set; } = Array.Empty<int>();
    public IReadOnlyList<FormDataUpload> LatestUploads { get; private set; } = Array.Empty<FormDataUpload>();

    public string DisplayFormNumber => Meta?.DisplayFormNumber ?? FormNumber;
    public string DisplayFormTitle => Meta?.DisplayFormTitle ?? (Latest?.FormTitle ?? FormNumber);

    public bool EditMeta { get; private set; }

    [BindProperty]
    public string? DisplayFormNumberEdit { get; set; }

    [BindProperty]
    public string? DisplayFormTitleEdit { get; set; }

    [BindProperty]
    public IFormFile? DataUpload { get; set; }

    [BindProperty]
    public IFormFile? SchemaUpload { get; set; }

    public IActionResult OnGet(string formNumber, bool editMeta = false)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Versions = _storage.ListVersions(FormNumber);
        Latest = _storage.TryGetLatestStructure(FormNumber);
        Meta = _storage.TryLoadFormMeta(FormNumber);
        EditMeta = editMeta;

        if (Latest is null)
        {
            return NotFound();
        }

        if (EditMeta)
        {
            DisplayFormNumberEdit = DisplayFormNumber;
            DisplayFormTitleEdit = DisplayFormTitle;
        }

        LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);

        return Page();
    }

    public IActionResult OnPostSaveMeta(string formNumber)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Versions = _storage.ListVersions(FormNumber);
        Latest = _storage.TryGetLatestStructure(FormNumber);
        Meta = _storage.TryLoadFormMeta(FormNumber);

        if (Latest is null)
        {
            return NotFound();
        }

        var newNumber = (DisplayFormNumberEdit ?? string.Empty).Trim();
        var newTitle = (DisplayFormTitleEdit ?? string.Empty).Trim();

        if (string.IsNullOrWhiteSpace(newNumber))
        {
            ModelState.AddModelError(nameof(DisplayFormNumberEdit), "Please provide a form number.");
        }

        if (string.IsNullOrWhiteSpace(newTitle))
        {
            ModelState.AddModelError(nameof(DisplayFormTitleEdit), "Please provide a form title.");
        }

        if (!ModelState.IsValid)
        {
            EditMeta = true;
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            return Page();
        }

        // Display form number/title are editable metadata and do not affect the form key/URL.
        _storage.SaveFormMeta(FormNumber, new FormMeta
        {
            DisplayFormNumber = newNumber,
            DisplayFormTitle = newTitle,
            UpdatedAtUtc = DateTime.UtcNow
        });

        TempData["MetaMessage"] = "Updated form info.";
        return Redirect($"/forms/{Uri.EscapeDataString(FormNumber)}");
    }

    public async Task<IActionResult> OnPostUploadDataAsync(string formNumber, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Versions = _storage.ListVersions(FormNumber);
        Latest = _storage.TryGetLatestStructure(FormNumber);
        Meta = _storage.TryLoadFormMeta(FormNumber);

        if (Latest is null)
        {
            return NotFound();
        }

        if (DataUpload is null)
        {
            ModelState.AddModelError(nameof(DataUpload), "Please choose a .xlsx file.");
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            return Page();
        }

        try
        {
            // Allow embedded form number to differ, but always store under this form index.
            var result = await _dataStorage.SaveAsync(
                DataUpload,
                _parser,
                ct,
                expectedFormNumber: null,
                targetFormNumber: FormNumber);
            TempData["UploadMessage"] = $"Stored data for #{result.FormNumber} v{result.Version}: {result.RowCount} rows.";
            return Redirect($"/forms/{Uri.EscapeDataString(result.FormNumber)}");
        }
        catch (FormParseException ex)
        {
            _logger.LogWarning(ex, "Failed to load data from uploaded file {FileName}", DataUpload.FileName);
            ModelState.AddModelError(string.Empty, ex.Message);
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error while processing uploaded data file {FileName}", DataUpload.FileName);
            ModelState.AddModelError(string.Empty, "Unexpected error while processing the file. See logs for details.");
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            return Page();
        }
    }

    public async Task<IActionResult> OnPostUploadSchemaAsync(string formNumber, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Versions = _storage.ListVersions(FormNumber);
        Latest = _storage.TryGetLatestStructure(FormNumber);
        Meta = _storage.TryLoadFormMeta(FormNumber);

        if (Latest is null)
        {
            return NotFound();
        }

        if (SchemaUpload is null)
        {
            ModelState.AddModelError(nameof(SchemaUpload), "Please choose a .xlsx file.");
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            return Page();
        }

        try
        {
            var result = await _storage.SaveAsync(SchemaUpload, _parser, ct, targetFormNumber: FormNumber);

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
            _logger.LogWarning(ex, "Failed to parse uploaded file {FileName}", SchemaUpload.FileName);
            ModelState.AddModelError(string.Empty, ex.Message);
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error while processing uploaded file {FileName}", SchemaUpload.FileName);
            ModelState.AddModelError(string.Empty, "Unexpected error while processing the file. See logs for details.");
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            return Page();
        }
    }

    public IActionResult OnGetDownloadData(string formNumber, int version, string uploadId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        var path = _dataStorage.GetOriginalFilePath(formNumber, version, uploadId);
        if (!System.IO.File.Exists(path))
        {
            return NotFound();
        }

        var downloadName = $"{formNumber}-v{version}-{uploadId}.xlsx";
        return PhysicalFile(path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", downloadName);
    }

    public IActionResult OnGetDataJson(string formNumber, int version, string uploadId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        var path = _dataStorage.GetDataJsonPath(formNumber, version, uploadId);
        if (!System.IO.File.Exists(path))
        {
            return NotFound();
        }

        var json = System.IO.File.ReadAllText(path);
        return Content(json, "application/json");
    }

    public IActionResult OnGetDownload(string formNumber, int version)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0)
        {
            return NotFound();
        }

        var path = _storage.GetOriginalFilePath(formNumber, version);
        if (!System.IO.File.Exists(path))
        {
            return NotFound();
        }

        var downloadName = $"{formNumber}-v{version}.xlsx";
        return PhysicalFile(path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", downloadName);
    }

    public IActionResult OnGetStructure(string formNumber, int version)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0)
        {
            return NotFound();
        }

        var structure = _storage.TryLoadStructure(formNumber, version);
        if (structure is null)
        {
            return NotFound();
        }

        var json = JsonUtil.ToStableJson(structure);
        return Content(json, "application/json");
    }

    public IActionResult OnPostDeleteVersion(string formNumber, int version)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            formNumber = RouteData.Values["formNumber"]?.ToString() ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0)
        {
            return NotFound();
        }

        // Delete uploads first (they reference the schema version).
        _dataStorage.DeleteVersion(formNumber, version);

        _storage.DeleteVersion(formNumber, version);

        TempData["DeleteMessage"] = $"Deleted schema v{version} (and its uploaded data).";

        // If that was the last version, the form page no longer exists.
        var latest = _storage.TryGetLatestStructure(formNumber);
        return latest is null
            ? Redirect("/")
            : RedirectToPage("/Forms/Details", new { formNumber });
    }

    public IActionResult OnPostDeleteForm(string formNumber)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            formNumber = RouteData.Values["formNumber"]?.ToString() ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return NotFound();
        }

        // Delete data uploads first, then schema.
        _dataStorage.DeleteForm(formNumber);

        // Treat missing folders as already deleted.
        _storage.DeleteForm(formNumber);

        TempData["DeleteMessage"] = $"Deleted form #{formNumber} (all versions and uploaded data).";
        return Redirect("/");
    }
}
