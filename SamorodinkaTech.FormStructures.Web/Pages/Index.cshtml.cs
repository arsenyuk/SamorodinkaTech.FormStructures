using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages;

public class IndexModel : PageModel
{
    private readonly FormStorage _storage;
    private readonly FormDataStorage _dataStorage;

    public IndexModel(FormStorage storage, FormDataStorage dataStorage)
    {
        _storage = storage;
        _dataStorage = dataStorage;
    }

    public IReadOnlyList<FormRow> Forms { get; private set; } = Array.Empty<FormRow>();

    public void OnGet()
    {
        LoadForms();
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
