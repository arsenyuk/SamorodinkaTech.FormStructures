using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Directories;

public sealed class DetailsModel : PageModel
{
    private readonly FormStorage _storage;

    public DetailsModel(FormStorage storage)
    {
        _storage = storage;
    }

    [FromRoute]
    public string FormNumber { get; set; } = string.Empty;

    [FromRoute]
    public int Version { get; set; }

    [FromRoute]
    public string Id { get; set; } = string.Empty;

    public ReferenceBook? Book { get; private set; }

    public IActionResult OnGet()
    {
        if (string.IsNullOrWhiteSpace(FormNumber) || Version <= 0 || string.IsNullOrWhiteSpace(Id))
        {
            return NotFound();
        }

        var books = _storage.TryLoadReferenceBooks(FormNumber, Version);
        Book = books.FirstOrDefault(b => string.Equals(b.Id, Id, StringComparison.OrdinalIgnoreCase));
        if (Book is null)
        {
            return NotFound();
        }

        return Page();
    }
}
