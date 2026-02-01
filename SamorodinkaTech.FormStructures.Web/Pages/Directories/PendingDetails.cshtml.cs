using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Directories;

public sealed class PendingDetailsModel : PageModel
{
    private readonly FormStorage _storage;

    public PendingDetailsModel(FormStorage storage)
    {
        _storage = storage;
    }

    [FromRoute]
    public string FormNumber { get; set; } = string.Empty;

    [FromRoute]
    public string PendingId { get; set; } = string.Empty;

    [FromRoute]
    public string Id { get; set; } = string.Empty;

    public int PreviousVersion { get; private set; }

    public ReferenceBook? Book { get; private set; }

    public IActionResult OnGet()
    {
        if (string.IsNullOrWhiteSpace(FormNumber) || string.IsNullOrWhiteSpace(PendingId) || string.IsNullOrWhiteSpace(Id))
        {
            return NotFound();
        }

        var pendingMeta = _storage.ListPending(FormNumber)
            .FirstOrDefault(p => string.Equals(p.PendingId, PendingId, StringComparison.OrdinalIgnoreCase));

        PreviousVersion = pendingMeta?.PreviousVersion ?? 0;

        var books = _storage.TryLoadPendingReferenceBooks(FormNumber, PendingId);
        Book = books.FirstOrDefault(b => string.Equals(b.Id, Id, StringComparison.OrdinalIgnoreCase));

        if (Book is null)
        {
            return NotFound();
        }

        return Page();
    }
}
