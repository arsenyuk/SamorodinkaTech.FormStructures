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

    public IReadOnlyDictionary<string, ColumnFormulaInfo> LoadedFormulasByPath { get; private set; }
        = new Dictionary<string, ColumnFormulaInfo>(StringComparer.Ordinal);

    public IReadOnlyList<VersionSummary> VersionSummaries { get; private set; } = Array.Empty<VersionSummary>();

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
        LoadPageData();
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
            VersionSummaries = BuildVersionSummaries(FormNumber, Versions);
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
            VersionSummaries = BuildVersionSummaries(FormNumber, Versions);
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
            _logger.LogWarning(
                ex,
                "Failed to load data from uploaded file {FileName}. ExceptionChain={ExceptionChain}",
                DataUpload.FileName,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            VersionSummaries = BuildVersionSummaries(FormNumber, Versions);
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Unexpected error while processing uploaded data file {FileName}. ExceptionChain={ExceptionChain}",
                DataUpload.FileName,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            VersionSummaries = BuildVersionSummaries(FormNumber, Versions);
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
            VersionSummaries = BuildVersionSummaries(FormNumber, Versions);
            return Page();
        }

        try
        {
            var result = await _storage.SaveAsync(SchemaUpload, _parser, ct, targetFormNumber: FormNumber);

            if (!result.IsNewVersion)
            {
                TempData["UploadMessage"] = $"No schema changes for {result.FormTitle} (#{result.FormNumber}); current version is v{result.Version}.";
            }
            else if (result.RequiresTypeSetup && result.PendingId is string typePendingIdForMsg)
            {
                TempData["UploadMessage"] = $"Upload staged for {result.FormTitle} (#{result.FormNumber}) v{result.Version}. Please confirm column types to create the version.";
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
                SchemaUpload.FileName,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            VersionSummaries = BuildVersionSummaries(FormNumber, Versions);
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Unexpected error while processing uploaded file {FileName}. ExceptionChain={ExceptionChain}",
                SchemaUpload.FileName,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            VersionSummaries = BuildVersionSummaries(FormNumber, Versions);
            return Page();
        }
    }


    private void LoadPageData()
    {
        Versions = _storage.ListVersions(FormNumber);
        Latest = _storage.TryGetLatestStructure(FormNumber);
        Meta = _storage.TryLoadFormMeta(FormNumber);

        if (Latest is not null)
        {
            LatestUploads = _dataStorage.ListUploads(FormNumber, Latest.Version);
            LoadedFormulasByPath = BuildSchemaFormulasByPath(FormNumber, Latest.Version, Latest);
        }

        VersionSummaries = BuildVersionSummaries(FormNumber, Versions);
    }

    private IReadOnlyDictionary<string, ColumnFormulaInfo> BuildSchemaFormulasByPath(
        string formNumber,
        int version,
        FormStructure structure)
    {
        // Best-effort: extract formulas from the CURRENT SCHEMA TEMPLATE (original.xlsx)
        // so the column description reflects only what the schema defines.
        const int maxRowsToScan = 50;

        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || structure.Columns.Count == 0)
        {
            return new Dictionary<string, ColumnFormulaInfo>(StringComparer.Ordinal);
        }

        var formulasByPath = new Dictionary<string, string>(StringComparer.Ordinal);

        try
        {
            var templatePath = _storage.GetOriginalFilePath(formNumber, version);
            if (!System.IO.File.Exists(templatePath))
            {
                return new Dictionary<string, ColumnFormulaInfo>(StringComparer.Ordinal);
            }

            using var fs = System.IO.File.OpenRead(templatePath);
            var layout = _parser.ParseLayout(fs, sourceFileName: Path.GetFileName(templatePath));

            // Need a fresh stream for reading cells.
            fs.Position = 0;
            using var workbook = new ClosedXML.Excel.XLWorkbook(fs);
            var ws = workbook.Worksheets.FirstOrDefault();
            if (ws is null)
            {
                return new Dictionary<string, ColumnFormulaInfo>(StringComparer.Ordinal);
            }

            var leafIndexByColumnNumber = new Dictionary<int, int>();
            for (var i = 0; i < layout.LeafColumns.Count; i++)
            {
                leafIndexByColumnNumber[layout.LeafColumns[i]] = i;
            }

            int TryAddFormula(int columnIndex, int rowNumber)
            {
                var cell = ws.Cell(rowNumber, columnIndex);
                if (!cell.HasFormula)
                {
                    return 0;
                }

                string? f = null;
                try
                {
                    f = (cell.FormulaA1 ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(f))
                    {
                        f = (cell.FormulaR1C1 ?? string.Empty).Trim();
                    }
                }
                catch
                {
                    // ignore
                }

                if (string.IsNullOrWhiteSpace(f))
                {
                    return 0;
                }

                if (!f.StartsWith("=", StringComparison.Ordinal))
                {
                    f = $"={f}";
                }

                if (!leafIndexByColumnNumber.TryGetValue(columnIndex, out var leafIdx))
                {
                    return 0;
                }

                if (leafIdx < 0 || leafIdx >= structure.Columns.Count)
                {
                    return 0;
                }

                var path = structure.Columns[leafIdx].Path;
                if (formulasByPath.ContainsKey(path))
                {
                    return 0;
                }

                formulasByPath[path] = f;
                return 1;
            }

            var lastRowToScan = Math.Min(layout.UsedLastRow, layout.DataStartRow + maxRowsToScan - 1);
            for (var r = layout.DataStartRow; r <= lastRowToScan; r++)
            {
                foreach (var leafCol in layout.LeafColumns)
                {
                    TryAddFormula(leafCol, r);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to extract schema formulas for {FormNumber} v{Version}", formNumber, version);
            return new Dictionary<string, ColumnFormulaInfo>(StringComparer.Ordinal);
        }

        var result = new Dictionary<string, ColumnFormulaInfo>(StringComparer.Ordinal);
        foreach (var kv in formulasByPath)
        {
            result[kv.Key] = new ColumnFormulaInfo(Formula: kv.Value);
        }

        return result;
    }

    public sealed record ColumnFormulaInfo(string Formula);

    private IReadOnlyList<VersionSummary> BuildVersionSummaries(string formNumber, IReadOnlyList<int> versions)
    {
        var summaries = new List<VersionSummary>(versions.Count);
        foreach (var v in versions)
        {
            var s = _storage.TryLoadStructure(formNumber, v);
            if (s is null)
            {
                continue;
            }

            summaries.Add(new VersionSummary(
                Version: v,
                UploadedAtUtc: s.UploadedAtUtc,
                ColumnsCount: s.Columns.Count,
                FormTitle: s.FormTitle));
        }

        return summaries
            .OrderByDescending(x => x.Version)
            .ToArray();
    }

    public sealed record VersionSummary(int Version, DateTime UploadedAtUtc, int ColumnsCount, string FormTitle);

    public IActionResult OnGetDownloadData(string formNumber, int version, string uploadId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        var structure = _storage.TryLoadStructure(formNumber, version);
        if (structure is null)
        {
            return NotFound();
        }

        var path = _dataStorage.GetOriginalFilePath(formNumber, version, uploadId);
        if (!System.IO.File.Exists(path))
        {
            return NotFound();
        }

        var downloadName = DownloadFileName.ForDataUpload(structure, version, uploadId);
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

        var structure = _storage.TryLoadStructure(formNumber, version);
        if (structure is null)
        {
            return NotFound();
        }

        var path = _storage.GetOriginalFilePath(formNumber, version);
        if (!System.IO.File.Exists(path))
        {
            return NotFound();
        }

        var downloadName = DownloadFileName.ForSchema(structure, version);
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
