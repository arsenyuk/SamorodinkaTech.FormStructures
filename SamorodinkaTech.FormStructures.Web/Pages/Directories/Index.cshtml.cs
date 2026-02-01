using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Directories;

public sealed class IndexModel : PageModel
{
    private readonly FormStorage _storage;

    public IndexModel(FormStorage storage)
    {
        _storage = storage;
    }

    public IReadOnlyList<ItemRow> Items { get; private set; } = Array.Empty<ItemRow>();

    public void OnGet()
    {
        var latest = _storage.ListLatestForms();

        var items = new List<ItemRow>();
        foreach (var f in latest)
        {
            var books = _storage.TryLoadReferenceBooks(f.FormNumber, f.Version);
            if (books.Count == 0)
            {
                continue;
            }

            foreach (var b in books)
            {
                items.Add(new ItemRow(
                    FormNumber: f.FormNumber,
                    FormTitle: f.DisplayFormTitle,
                    Version: f.Version,
                    PendingId: null,
                    PreviousVersion: null,
                    ReferenceBookId: b.Id,
                    ReferenceBookTitle: b.Title,
                    Source: FormatSource(b),
                    ValueCount: b.Values.Count));
            }
        }

        // Include pending uploads too (they are not yet committed versions).
        foreach (var f in latest)
        {
            var pending = _storage.ListPending(f.FormNumber);
            if (pending.Count == 0)
            {
                continue;
            }

            foreach (var p in pending)
            {
                var books = _storage.TryLoadPendingReferenceBooks(f.FormNumber, p.PendingId);
                if (books.Count == 0)
                {
                    continue;
                }

                foreach (var b in books)
                {
                    items.Add(new ItemRow(
                        FormNumber: f.FormNumber,
                        FormTitle: f.DisplayFormTitle,
                        Version: p.IntendedVersion,
                        PendingId: p.PendingId,
                        PreviousVersion: p.PreviousVersion,
                        ReferenceBookId: b.Id,
                        ReferenceBookTitle: b.Title,
                        Source: FormatSource(b),
                        ValueCount: b.Values.Count));
                }
            }
        }

        Items = items
            .OrderBy(i => i.FormTitle, StringComparer.OrdinalIgnoreCase)
            .ThenBy(i => i.PendingId is null ? 0 : 1)
            .ThenBy(i => i.ReferenceBookTitle, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string FormatSource(Web.Models.ReferenceBook b)
    {
        if (!string.IsNullOrWhiteSpace(b.SourceSheet) && !string.IsNullOrWhiteSpace(b.SourceRange))
        {
            return $"{b.SourceSheet}!{b.SourceRange}";
        }

        return b.SourceFormula;
    }

    public sealed record ItemRow(
        string FormNumber,
        string FormTitle,
        int Version,
        string? PendingId,
        int? PreviousVersion,
        string ReferenceBookId,
        string ReferenceBookTitle,
        string Source,
        int ValueCount);
}
