using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Forms;

public class DataModel : PageModel
{
    private readonly FormStorage _formStorage;
    private readonly FormDataStorage _dataStorage;

    public DataModel(FormStorage formStorage, FormDataStorage dataStorage)
    {
        _formStorage = formStorage;
        _dataStorage = dataStorage;
    }

    public string FormNumber { get; private set; } = string.Empty;
    public FormStructure? LatestStructure { get; private set; }
    public IReadOnlyList<FormDataUpload> Uploads { get; private set; } = Array.Empty<FormDataUpload>();

    public IActionResult OnGet(string formNumber)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        LatestStructure = _formStorage.TryGetLatestStructure(FormNumber);
        if (LatestStructure is null)
        {
            return NotFound();
        }

        Uploads = _dataStorage.ListUploads(FormNumber);

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
}
