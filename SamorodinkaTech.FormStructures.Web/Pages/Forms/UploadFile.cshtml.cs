using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Forms;

public class UploadFileModel : PageModel
{
    private readonly FormStorage _formStorage;
    private readonly FormDataStorage _dataStorage;

    public UploadFileModel(FormStorage formStorage, FormDataStorage dataStorage)
    {
        _formStorage = formStorage;
        _dataStorage = dataStorage;
    }

    public string FormNumber { get; private set; } = string.Empty;
    public int Version { get; private set; }
    public string UploadId { get; private set; } = string.Empty;

    public FormDataUpload? Meta { get; private set; }

    public IActionResult OnGet(string formNumber, int version, string uploadId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Version = version;
        UploadId = uploadId;

        Meta = _dataStorage.TryLoadUploadMeta(FormNumber, Version, UploadId);
        if (Meta is null)
        {
            return NotFound();
        }

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

        var structure = _formStorage.TryLoadStructure(formNumber, version);
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
}
