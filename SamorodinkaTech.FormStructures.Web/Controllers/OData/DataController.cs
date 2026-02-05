using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.OData.Query;
using Microsoft.AspNetCore.OData.Routing.Controllers;
using Microsoft.AspNetCore.OData.Results;
using SamorodinkaTech.FormStructures.Web.OData;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Controllers.OData;

public sealed class DataController : ODataController
{
    private readonly FormStorage _formStorage;
    private readonly FormDataStorage _dataStorage;

    public DataController(FormStorage formStorage, FormDataStorage dataStorage)
    {
        _formStorage = formStorage;
        _dataStorage = dataStorage;
    }

    [EnableQuery(PageSize = 100)]
    public IQueryable<ODataUpload> Get()
    {
        return _dataStorage
            .ListAllUploads()
            .Select(ODataUpload.From)
            .AsQueryable();
    }

    [EnableQuery]
    public SingleResult<ODataUpload> Get(string key)
    {
        if (!ODataKeys.TryParseUploadKey(key, out var formNumber, out var version, out var uploadId))
        {
            return SingleResult.Create(Enumerable.Empty<ODataUpload>().AsQueryable());
        }

        var meta = _dataStorage.TryLoadUploadMeta(formNumber, version, uploadId);
        if (meta is null)
        {
            return SingleResult.Create(Enumerable.Empty<ODataUpload>().AsQueryable());
        }

        var entity = ODataUpload.From(meta);
        return SingleResult.Create(new[] { entity }.AsQueryable());
    }

    // GET /odata/Data('...')/Rows
    [EnableQuery(PageSize = 200)]
    public IActionResult GetRows(string key)
    {
        if (!ODataKeys.TryParseUploadKey(key, out var formNumber, out var version, out var uploadId))
        {
            return NotFound();
        }

        var structure = _formStorage.TryLoadStructure(formNumber, version);
        if (structure is null)
        {
            return NotFound();
        }

        var data = _dataStorage.TryLoadData(formNumber, version, uploadId);
        if (data is null)
        {
            return NotFound();
        }

        var uploadKey = ODataKeys.UploadKey(formNumber, version, uploadId);

        var propByPath = structure.Columns
            .ToDictionary(c => c.Path, c => $"c{c.Index}", StringComparer.Ordinal);

        var rows = data.Rows
            .Select(r => new ODataRow
            {
                Id = ODataKeys.RowKey(uploadKey, r.RowNumber),
                UploadKey = uploadKey,
                FormNumber = formNumber,
                Version = version,
                UploadId = uploadId,
                RowNumber = r.RowNumber,
                DynamicProperties = ProjectDynamicProperties(r.Values, propByPath),
            })
            .AsQueryable();

        return Ok(rows);
    }

    // GET /odata/Data('...')/Columns
    [EnableQuery]
    public IActionResult GetColumns(string key)
    {
        if (!ODataKeys.TryParseUploadKey(key, out var formNumber, out var version, out var uploadId))
        {
            return NotFound();
        }

        var structure = _formStorage.TryLoadStructure(formNumber, version);
        if (structure is null)
        {
            return NotFound();
        }

        var uploadKey = ODataKeys.UploadKey(formNumber, version, uploadId);

        var cols = structure.Columns
            .OrderBy(c => c.Index)
            .Select(c => new ODataColumn
            {
                Id = ODataKeys.ColumnKey(uploadKey, c.Index),
                UploadKey = uploadKey,
                FormNumber = formNumber,
                Version = version,
                Index = c.Index,
                ODataProperty = $"c{c.Index}",
                Name = c.Name,
                Path = c.Path,
                ColumnNumber = c.ColumnNumber,
                Type = c.Type,
            })
            .AsQueryable();

        return Ok(cols);
    }

    private static IDictionary<string, object?> ProjectDynamicProperties(
        IReadOnlyDictionary<string, string?> values,
        IReadOnlyDictionary<string, string> propByPath)
    {
        var dict = new Dictionary<string, object?>(StringComparer.Ordinal);

        foreach (var (path, value) in values)
        {
            if (!propByPath.TryGetValue(path, out var prop))
            {
                continue;
            }

            dict[prop] = value;
        }

        return dict;
    }
}
