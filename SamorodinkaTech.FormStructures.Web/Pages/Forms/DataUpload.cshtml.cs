using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Forms;

public class DataUploadModel : PageModel
{
    private readonly FormStorage _formStorage;
    private readonly FormDataStorage _dataStorage;

    public DataUploadModel(FormStorage formStorage, FormDataStorage dataStorage)
    {
        _formStorage = formStorage;
        _dataStorage = dataStorage;
    }

    public string FormNumber { get; private set; } = string.Empty;
    public int Version { get; private set; }
    public string UploadId { get; private set; } = string.Empty;

    public FormStructure? Structure { get; private set; }
    public FormDataFile? Data { get; private set; }

    public string SortKey { get; private set; } = "row";
    public string SortDir { get; private set; } = "asc";

    public IActionResult OnGet(
        string formNumber,
        int version,
        string uploadId,
        string? sort = null,
        string? dir = null)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Version = version;
        UploadId = uploadId;

        Structure = _formStorage.TryLoadStructure(FormNumber, Version);
        if (Structure is null)
        {
            return NotFound();
        }

        Data = _dataStorage.TryLoadData(FormNumber, Version, UploadId);
        if (Data is null)
        {
            return NotFound();
        }

        SortKey = string.IsNullOrWhiteSpace(sort) ? "row" : sort;
        SortDir = string.Equals(dir, "desc", StringComparison.OrdinalIgnoreCase) ? "desc" : "asc";

        Data = Data with
        {
            Rows = ApplySort(Data.Rows, Structure, SortKey, SortDir)
        };

        return Page();
    }

    public IActionResult OnPostDelete(string formNumber, int version, string uploadId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        var deleted = _dataStorage.DeleteUpload(formNumber, version, uploadId);
        if (!deleted)
        {
            return NotFound();
        }

        return RedirectToPage("/Forms/Data", new { formNumber });
    }

    public IActionResult OnGetDownload(string formNumber, int version, string uploadId)
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

    private static IReadOnlyList<FormDataRow> ApplySort(
        IReadOnlyList<FormDataRow> rows,
        FormStructure structure,
        string sortKey,
        string sortDir)
    {
        var descending = string.Equals(sortDir, "desc", StringComparison.OrdinalIgnoreCase);

        if (string.Equals(sortKey, "row", StringComparison.OrdinalIgnoreCase))
        {
            return descending
                ? rows.OrderByDescending(r => r.RowNumber).ToArray()
                : rows.OrderBy(r => r.RowNumber).ToArray();
        }

        if (sortKey.StartsWith("c", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(sortKey[1..], out var colIndex)
            && colIndex >= 1
            && colIndex <= structure.Columns.Count)
        {
            var colPath = structure.Columns[colIndex - 1].Path;
            var ordered = descending
                ? rows.OrderByDescending(r => GetSortValue(r, colPath), StringComparer.OrdinalIgnoreCase)
                : rows.OrderBy(r => GetSortValue(r, colPath), StringComparer.OrdinalIgnoreCase);

            ordered = descending
                ? ordered.ThenByDescending(r => r.RowNumber)
                : ordered.ThenBy(r => r.RowNumber);

            return ordered.ToArray();
        }

        return rows;
    }

    private static string GetSortValue(FormDataRow row, string colPath)
    {
        if (row.Values.TryGetValue(colPath, out var value) && !string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        return string.Empty;
    }
}
