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

    [BindProperty]
    public IFormFile? CsvUpload { get; set; }

    [BindProperty]
    public string? CsvFormNumber { get; set; }

    [BindProperty]
    public string? CsvFormTitle { get; set; }

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
            else if (result.RequiresTypeSetup && result.PendingId is string typePendingIdForMsg)
            {
                TempData["UploadMessage"] = $"Upload staged for {result.FormTitle} (#{result.FormNumber}) v{result.Version}. Please confirm column types to create the form.";
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
            _logger.LogWarning(
                ex,
                "Failed to parse uploaded file {FileName}. ExceptionChain={ExceptionChain}",
                Upload.FileName,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
            LoadForms();
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Unexpected error while processing uploaded file {FileName}. ExceptionChain={ExceptionChain}",
                Upload.FileName,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
            LoadForms();
            return Page();
        }
    }

    public async Task<IActionResult> OnPostCsvAsync(CancellationToken ct)
    {
        var formNumber = (CsvFormNumber ?? string.Empty).Trim();
        var formTitle = (CsvFormTitle ?? string.Empty).Trim();

        if (string.IsNullOrWhiteSpace(formNumber))
        {
            ModelState.AddModelError(nameof(CsvFormNumber), "Please provide a form number.");
        }

        if (string.IsNullOrWhiteSpace(formTitle))
        {
            ModelState.AddModelError(nameof(CsvFormTitle), "Please provide a form title.");
        }

        if (CsvUpload is null)
        {
            ModelState.AddModelError(nameof(CsvUpload), "Please choose a .csv file.");
        }
        else
        {
            if (CsvUpload.Length == 0)
            {
                ModelState.AddModelError(nameof(CsvUpload), "Uploaded CSV file is empty.");
            }

            if (!string.Equals(Path.GetExtension(CsvUpload.FileName), ".csv", StringComparison.OrdinalIgnoreCase))
            {
                ModelState.AddModelError(nameof(CsvUpload), "Only .csv files are supported.");
            }
        }

        if (!ModelState.IsValid)
        {
            LoadForms();
            return Page();
        }

        // New-form only.
        if (_storage.ListVersions(formNumber).Count > 0)
        {
            ModelState.AddModelError(nameof(CsvFormNumber), $"Form #{formNumber} already exists. CSV import is only supported for creating new forms.");
            LoadForms();
            return Page();
        }

        try
        {
            await using var csvStream = CsvUpload!.OpenReadStream();
            var parsed = CsvHeaderParser.ParseHeaderRow(csvStream);

            var xlsxBytes = CsvSchemaTemplateBuilder.BuildXlsxTemplateBytes(formNumber, formTitle, parsed.Headers);

            var generatedFileName = $"{Path.GetFileNameWithoutExtension(CsvUpload.FileName)}.xlsx";
            var ms = new MemoryStream(xlsxBytes);
            var file = new FormFile(ms, 0, xlsxBytes.Length, "Upload", generatedFileName)
            {
                Headers = new HeaderDictionary(),
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            };

            var result = await _storage.SaveAsync(file, _parser, ct, targetFormNumber: formNumber);
            if (!result.IsNewVersion)
            {
                TempData["UploadMessage"] = $"No schema changes for {result.FormTitle} (#{result.FormNumber}); current version is v{result.Version}.";
            }
            else if (result.RequiresTypeSetup && result.PendingId is string typePendingIdForMsg)
            {
                TempData["UploadMessage"] = $"Upload staged for {result.FormTitle} (#{result.FormNumber}) v{result.Version}. Please confirm column types to create the form.";
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
            _logger.LogWarning(
                ex,
                "Failed to parse uploaded CSV file {FileName}. ExceptionChain={ExceptionChain}",
                CsvUpload!.FileName,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
            LoadForms();
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Unexpected error while processing uploaded CSV file {FileName}. ExceptionChain={ExceptionChain}",
                CsvUpload!.FileName,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
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
