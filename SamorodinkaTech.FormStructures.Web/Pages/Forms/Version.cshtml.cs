using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Forms;

public class VersionModel : PageModel
{
    private readonly FormStorage _storage;
    private readonly FormDataStorage _dataStorage;
    private readonly ILogger<VersionModel> _logger;

    public VersionModel(FormStorage storage, FormDataStorage dataStorage, ILogger<VersionModel> logger)
    {
        _storage = storage;
        _dataStorage = dataStorage;
        _logger = logger;
    }

    public string FormNumber { get; private set; } = string.Empty;
    public int Version { get; private set; }
    public FormStructure? Structure { get; private set; }

    public bool EditTypes { get; private set; }
    public int? MapFromVersion { get; private set; }
    public FormStructure? PreviousStructure { get; private set; }

    [BindProperty]
    public List<ColumnTypeEditRow> TypeEdits { get; set; } = new();

    [BindProperty]
    public List<ColumnMapEditRow> MapEdits { get; set; } = new();

    public IActionResult OnGet(string formNumber, int version, bool editTypes = false, int? mapFrom = null)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0)
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Version = version;
        Structure = _storage.TryLoadStructure(FormNumber, Version);

        EditTypes = editTypes;
        MapFromVersion = mapFrom;

        if (Structure is null)
        {
            return NotFound();
        }

        TypeEdits = Structure.Columns
            .OrderBy(c => c.Index)
            .Select(c => new ColumnTypeEditRow
            {
                Path = c.Path,
                Type = c.Type
            })
            .ToList();

        if (MapFromVersion is int prev && prev > 0)
        {
            PreviousStructure = _storage.TryLoadStructure(FormNumber, prev);
            if (PreviousStructure is null)
            {
                ModelState.AddModelError(string.Empty, $"Previous schema v{prev} not found.");
                MapFromVersion = null;
            }
            else
            {
                var prevByPath = PreviousStructure.Columns.ToDictionary(c => c.Path, StringComparer.Ordinal);

                MapEdits = Structure.Columns
                    .OrderBy(c => c.Index)
                    .Select(c => new ColumnMapEditRow
                    {
                        NewPath = c.Path,
                        FromPath = prevByPath.ContainsKey(c.Path) ? c.Path : string.Empty,
                        Type = c.Type
                    })
                    .ToList();
            }
        }

        return Page();
    }

    public async Task<IActionResult> OnPostSaveTypesAsync(string formNumber, int version, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0)
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Version = version;
        Structure = _storage.TryLoadStructure(FormNumber, Version);
        if (Structure is null)
        {
            return NotFound();
        }

        // Keep the editor visible when returning Page() from a POST.
        EditTypes = true;

        if (TypeEdits.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "No column edits provided.");

            // Repopulate defaults so the user can correct and resubmit.
            TypeEdits = Structure.Columns
                .OrderBy(c => c.Index)
                .Select(c => new ColumnTypeEditRow { Path = c.Path, Type = c.Type })
                .ToList();

            return Page();
        }

        var typeByPath = TypeEdits
            .Where(x => !string.IsNullOrWhiteSpace(x.Path))
            .GroupBy(x => x.Path, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.Last().Type, StringComparer.Ordinal);

        var updatedColumns = Structure.Columns
            .Select(c => typeByPath.TryGetValue(c.Path, out var t) ? c with { Type = t } : c)
            .ToArray();

        var updated = Structure with { Columns = updatedColumns };

        try
        {
            await _storage.SaveStructureAsync(FormNumber, Version, updated, ct);
            TempData["SaveMessage"] = "Column types saved.";
            return Redirect($"/forms/{Uri.EscapeDataString(FormNumber)}/v{Version}");
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Failed to save column types for {FormNumber} v{Version}. ExceptionChain={ExceptionChain}",
                FormNumber,
                Version,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, $"Failed to save column types. {ExceptionUtil.FormatExceptionChain(ex)}");
            return Page();
        }
    }

    public async Task<IActionResult> OnPostSaveMappingAsync(string formNumber, int version, int mapFrom, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || mapFrom <= 0)
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Version = version;
        MapFromVersion = mapFrom;

        Structure = _storage.TryLoadStructure(FormNumber, Version);
        if (Structure is null)
        {
            return NotFound();
        }

        PreviousStructure = _storage.TryLoadStructure(FormNumber, mapFrom);
        if (PreviousStructure is null)
        {
            return NotFound();
        }

        var prevByPath = PreviousStructure.Columns.ToDictionary(c => c.Path, StringComparer.Ordinal);

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
            // Rebuild TypeEdits so the page can still render both sections.
            TypeEdits = Structure.Columns
                .OrderBy(c => c.Index)
                .Select(c => new ColumnTypeEditRow { Path = c.Path, Type = c.Type })
                .ToList();
            return Page();
        }

        var mapByNewPath = MapEdits
            .Where(e => !string.IsNullOrWhiteSpace(e.NewPath))
            .GroupBy(e => e.NewPath, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.Last(), StringComparer.Ordinal);

        var updatedColumns = Structure.Columns
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

        var updated = Structure with { Columns = updatedColumns };

        try
        {
            await _storage.SaveStructureAsync(FormNumber, Version, updated, ct);
            TempData["SaveMessage"] = "Column mapping saved.";
            return Redirect($"/forms/{Uri.EscapeDataString(FormNumber)}/v{Version}");
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Failed to save column mapping for {FormNumber} v{Version}. ExceptionChain={ExceptionChain}",
                FormNumber,
                Version,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, $"Failed to save column mapping. {ExceptionUtil.FormatExceptionChain(ex)}");
            return Page();
        }
    }

    public IActionResult OnPostDelete(string formNumber, int version)
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

        // Treat missing folders as already deleted.
        _storage.DeleteVersion(formNumber, version);

        TempData["DeleteMessage"] = $"Deleted schema v{version} (and its uploaded data).";

        var latest = _storage.TryGetLatestStructure(formNumber);
        return latest is null
            ? Redirect("/")
            : RedirectToPage("/Forms/Details", new { formNumber });
    }

    public sealed class ColumnTypeEditRow
    {
        public string Path { get; set; } = string.Empty;
        public SamorodinkaTech.FormStructures.Web.Models.ColumnType Type { get; set; } = SamorodinkaTech.FormStructures.Web.Models.ColumnType.String;
    }

    public sealed class ColumnMapEditRow
    {
        public string NewPath { get; set; } = string.Empty;
        public string FromPath { get; set; } = string.Empty;
        public SamorodinkaTech.FormStructures.Web.Models.ColumnType Type { get; set; } = SamorodinkaTech.FormStructures.Web.Models.ColumnType.String;
    }
}
