using Microsoft.AspNetCore.OData.Query;
using Microsoft.AspNetCore.OData.Results;
using Microsoft.AspNetCore.OData.Routing.Controllers;
using SamorodinkaTech.FormStructures.Web.OData;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Controllers.OData;

public sealed class ReferenceBooksController : ODataController
{
    private readonly FormStorage _storage;

    public ReferenceBooksController(FormStorage storage)
    {
        _storage = storage;
    }

    [EnableQuery(PageSize = 200)]
    public IQueryable<ODataReferenceBook> Get()
    {
        var forms = _storage.ListLatestForms();

        var items = new List<ODataReferenceBook>();

        foreach (var f in forms)
        {
            var meta = _storage.TryLoadFormMeta(f.FormNumber);

            // Committed versions
            foreach (var v in _storage.ListVersions(f.FormNumber))
            {
                var books = _storage.TryLoadReferenceBooks(f.FormNumber, v);
                if (books.Count == 0)
                {
                    continue;
                }

                foreach (var b in books)
                {
                    items.Add(ODataReferenceBook.FromCommitted(f.FormNumber, v, meta, b));
                }
            }

            // Pending uploads
            foreach (var p in _storage.ListPending(f.FormNumber))
            {
                var books = _storage.TryLoadPendingReferenceBooks(f.FormNumber, p.PendingId);
                if (books.Count == 0)
                {
                    continue;
                }

                foreach (var b in books)
                {
                    items.Add(ODataReferenceBook.FromPending(f.FormNumber, p.IntendedVersion, p.PendingId, meta, b));
                }
            }
        }

        return items.AsQueryable();
    }

    [EnableQuery]
    public SingleResult<ODataReferenceBook> Get(string key)
    {
        if (!ODataKeys.TryParseReferenceBookKey(key, out var isPending, out var formNumber, out var version, out var pendingId, out var bookId))
        {
            return SingleResult.Create(Enumerable.Empty<ODataReferenceBook>().AsQueryable());
        }

        var meta = _storage.TryLoadFormMeta(formNumber);

        if (!isPending)
        {
            var books = _storage.TryLoadReferenceBooks(formNumber, version);
            var book = books.FirstOrDefault(b => string.Equals(b.Id, bookId, StringComparison.Ordinal));
            if (book is null)
            {
                return SingleResult.Create(Enumerable.Empty<ODataReferenceBook>().AsQueryable());
            }

            var entity = ODataReferenceBook.FromCommitted(formNumber, version, meta, book);
            return SingleResult.Create(new[] { entity }.AsQueryable());
        }

        var pending = _storage.ListPending(formNumber)
            .FirstOrDefault(p => string.Equals(p.PendingId, pendingId, StringComparison.OrdinalIgnoreCase));
        if (pending is null)
        {
            return SingleResult.Create(Enumerable.Empty<ODataReferenceBook>().AsQueryable());
        }

        var pendingBooks = _storage.TryLoadPendingReferenceBooks(formNumber, pendingId);
        var pendingBook = pendingBooks.FirstOrDefault(b => string.Equals(b.Id, bookId, StringComparison.Ordinal));
        if (pendingBook is null)
        {
            return SingleResult.Create(Enumerable.Empty<ODataReferenceBook>().AsQueryable());
        }

        var pendingEntity = ODataReferenceBook.FromPending(formNumber, pending.IntendedVersion, pendingId, meta, pendingBook);
        return SingleResult.Create(new[] { pendingEntity }.AsQueryable());
    }
}
