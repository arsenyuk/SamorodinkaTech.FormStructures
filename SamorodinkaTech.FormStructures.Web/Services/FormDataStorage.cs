using SamorodinkaTech.FormStructures.Web.Models;
using System.Collections.Concurrent;
using ClosedXML.Excel;

namespace SamorodinkaTech.FormStructures.Web.Services;

public sealed class FormDataStorage
{
    private static readonly ConcurrentDictionary<string, object> FormLocks = new(StringComparer.OrdinalIgnoreCase);

    private readonly FormStorage _formStorage;
    private readonly ILogger<FormDataStorage> _logger;

    public FormDataStorage(FormStorage formStorage, ILogger<FormDataStorage> logger)
    {
        _formStorage = formStorage;
        _logger = logger;
    }

    public string RootPath => Path.Combine(_formStorage.RootPath, "data");

    private static object GetFormLock(string formNumber)
    {
        var key = (formNumber ?? string.Empty).Trim();
        if (key.Length == 0)
        {
            key = "__empty__";
        }

        return FormLocks.GetOrAdd(key, _ => new object());
    }

    public void EnsureInitialized()
    {
        Directory.CreateDirectory(_formStorage.RootPath);
        Directory.CreateDirectory(RootPath);
    }

    public async Task<SaveDataResult> SaveAsync(
        IFormFile file,
        ExcelFormParser parser,
        CancellationToken ct,
        string? expectedFormNumber = null,
        string? targetFormNumber = null)
    {
        EnsureInitialized();

        if (file.Length == 0)
        {
            throw new FormParseException("Uploaded file is empty.");
        }

        if (!string.Equals(Path.GetExtension(file.FileName), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new FormParseException("Only .xlsx files are supported.");
        }

        _logger.LogInformation("Data upload received: {FileName} ({Size} bytes)", file.FileName, file.Length);

        byte[] bytes;
        await using (var ms = new MemoryStream())
        {
            await file.CopyToAsync(ms, ct);
            bytes = ms.ToArray();
        }

        var fileSha256 = Hashing.Sha256Hex(bytes);

        ExcelFormParser.ExcelFormLayout layout;
        await using (var s = new MemoryStream(bytes, writable: false))
        {
            layout = parser.ParseLayout(s, file.FileName);
        }

        var formKey = string.IsNullOrWhiteSpace(targetFormNumber)
            ? layout.Structure.FormNumber
            : targetFormNumber;

        if (!string.IsNullOrWhiteSpace(expectedFormNumber)
            && !string.Equals(formKey, expectedFormNumber, StringComparison.OrdinalIgnoreCase))
        {
            throw new FormParseException($"Uploaded file is for form #{formKey}, but expected #{expectedFormNumber}.");
        }

        var matchedVersion = _formStorage.TryFindVersionByStructureHash(formKey, layout.Structure.StructureHash);
        if (matchedVersion is null)
        {
            var knownVersions = _formStorage.ListVersions(formKey);
            if (knownVersions.Count == 0)
            {
                throw new FormParseException($"No stored schema found for form #{formKey}. Upload the empty template first.");
            }

            var latest = _formStorage.TryGetLatestStructure(formKey);
            var latestText = latest is null ? "" : $" Latest is v{latest.Version}.";
            throw new FormParseException($"Uploaded file schema does not match any stored version for form #{formKey}.{latestText}");
        }

        IReadOnlyList<FormDataRow> rows;
        await using (var s = new MemoryStream(bytes, writable: false))
        {
            rows = parser.ReadDataRows(s, layout);
        }

        ValidateUniformColumnFormulas(bytes, layout);

        var uploadId = Guid.NewGuid().ToString("N");
        var uploadedAtUtc = DateTime.UtcNow;

        var upload = new FormDataUpload
        {
            UploadId = uploadId,
            FormNumber = formKey,
            FormVersion = matchedVersion.Value,
            StructureHash = layout.Structure.StructureHash,
            OriginalFileName = file.FileName,
            FileSha256 = fileSha256,
            UploadedAtUtc = uploadedAtUtc,
            RowCount = rows.Count
        };

        var dataFile = new FormDataFile
        {
            Upload = upload,
            Rows = rows
        };

        var uploadDir = GetUploadDir(upload.FormNumber, upload.FormVersion, upload.UploadId);
        var dataJson = JsonUtil.ToStableJson(dataFile);
        var metaJson = JsonUtil.ToStableJson(upload);

        lock (GetFormLock(upload.FormNumber))
        {
            if (Directory.Exists(uploadDir))
            {
                throw new InvalidOperationException($"Data upload directory already exists: {uploadDir}");
            }

            Directory.CreateDirectory(uploadDir);

            var originalPath = Path.Combine(uploadDir, "original.xlsx");
            File.WriteAllBytes(originalPath, bytes);

            File.WriteAllText(Path.Combine(uploadDir, "data.json"), dataJson);
            File.WriteAllText(Path.Combine(uploadDir, "meta.json"), metaJson);
        }

        _logger.LogInformation(
            "Stored data upload {FormNumber} v{Version} ({Rows} rows) at {Dir}",
            upload.FormNumber,
            upload.FormVersion,
            upload.RowCount,
            uploadDir);

        return new SaveDataResult(upload.FormNumber, upload.FormVersion, upload.UploadId, upload.RowCount);
    }

    private static void ValidateUniformColumnFormulas(byte[] xlsxBytes, ExcelFormParser.ExcelFormLayout layout)
    {
        // Requirement: if a column uses formulas, the formula must be identical for all non-empty cells.
        // We compare formulas by R1C1 form (stable for copy-down formulas).
        using var ms = new MemoryStream(xlsxBytes, writable: false);
        using var workbook = new XLWorkbook(ms);
        var ws = workbook.Worksheets.FirstOrDefault();
        if (ws is null)
        {
            return;
        }

        static bool IsNonEmpty(IXLCell cell) => !cell.IsEmpty(XLCellsUsedOptions.Contents);

        static string? GetFormulaKeyR1C1(IXLCell cell)
        {
            try
            {
                var f = (cell.FormulaR1C1 ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(f))
                {
                    f = (cell.FormulaA1 ?? string.Empty).Trim();
                }

                if (string.IsNullOrWhiteSpace(f))
                {
                    return null;
                }

                return f.StartsWith("=", StringComparison.Ordinal) ? f : $"={f}";
            }
            catch
            {
                return null;
            }
        }

        static string? GetFormulaA1(IXLCell cell)
        {
            try
            {
                var f = (cell.FormulaA1 ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(f))
                {
                    return null;
                }

                return f.StartsWith("=", StringComparison.Ordinal) ? f : $"={f}";
            }
            catch
            {
                return null;
            }
        }

        var problems = new List<string>();

        for (var i = 0; i < layout.LeafColumns.Count && i < layout.Structure.Columns.Count; i++)
        {
            var excelCol = layout.LeafColumns[i];
            var colDef = layout.Structure.Columns[i];

            string? expectedKey = null;
            int expectedRow = 0;
            string? expectedA1 = null;

            int? firstNonFormulaRow = null;
            string? firstNonFormulaValue = null;

            // Track up to a few distinct formulas for diagnostics.
            var distinct = new Dictionary<string, (int Row, string? A1)>(StringComparer.Ordinal);

            for (var r = layout.DataStartRow; r <= layout.UsedLastRow; r++)
            {
                var cell = ws.Cell(r, excelCol);
                if (!IsNonEmpty(cell))
                {
                    continue;
                }

                if (cell.HasFormula)
                {
                    var key = GetFormulaKeyR1C1(cell);
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        // Treat as formula cell but without readable formula text.
                        key = "=<unreadable formula>";
                    }

                    if (expectedKey is null)
                    {
                        expectedKey = key;
                        expectedRow = r;
                        expectedA1 = GetFormulaA1(cell);
                        distinct[key] = (r, expectedA1);
                        continue;
                    }

                    if (!string.Equals(expectedKey, key, StringComparison.Ordinal))
                    {
                        if (!distinct.ContainsKey(key) && distinct.Count < 5)
                        {
                            distinct[key] = (r, GetFormulaA1(cell));
                        }
                    }

                    continue;
                }

                // Non-formula value in a column that might otherwise be formulas.
                firstNonFormulaRow ??= r;
                if (firstNonFormulaValue is null)
                {
                    var v = (cell.GetFormattedString() ?? string.Empty).Trim();
                    firstNonFormulaValue = v.Length > 60 ? v[..60] + "…" : v;
                }
            }

            if (expectedKey is null)
            {
                continue; // no formulas in this column
            }

            if (firstNonFormulaRow is not null)
            {
                problems.Add(
                    $"Column #{colDef.Index} '{colDef.Name}' contains formulas (first at row {expectedRow}) and non-formula values (first at row {firstNonFormulaRow}: '{firstNonFormulaValue}').");
            }

            if (distinct.Count > 1)
            {
                var variants = distinct
                    .Select(kv =>
                    {
                        var (row, a1) = kv.Value;
                        var a1Text = string.IsNullOrWhiteSpace(a1) ? "" : $"; A1={a1}";
                        return $"row {row}: R1C1={kv.Key}{a1Text}";
                    })
                    .ToArray();

                problems.Add(
                    $"Column #{colDef.Index} '{colDef.Name}' has inconsistent formulas. Expected (row {expectedRow}): R1C1={expectedKey}{(string.IsNullOrWhiteSpace(expectedA1) ? "" : $"; A1={expectedA1}")}. Found: {string.Join(" | ", variants)}");
            }
        }

        if (problems.Count > 0)
        {
            var message = "Upload rejected: formulas must be identical within each column. " +
                          "Fix the Excel file so the column formula is consistent for all filled cells (R1C1 form should match). " +
                          "Details: " + string.Join(" ", problems);
            throw new FormParseException(message);
        }
    }

    public IReadOnlyList<FormDataUpload> ListUploads(string formNumber, int version)
    {
        EnsureInitialized();

        var versionDir = GetVersionDir(formNumber, version);
        if (!Directory.Exists(versionDir))
        {
            return Array.Empty<FormDataUpload>();
        }

        var result = new List<FormDataUpload>();
        foreach (var dir in Directory.EnumerateDirectories(versionDir))
        {
            var metaPath = Path.Combine(dir, "meta.json");
            if (!File.Exists(metaPath))
            {
                continue;
            }

            try
            {
                var json = File.ReadAllText(metaPath);
                var meta = System.Text.Json.JsonSerializer.Deserialize<FormDataUpload>(json, JsonUtil.StableOptions);
                if (meta is not null)
                {
                    result.Add(meta);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to read meta.json in {Dir}", dir);
            }
        }

        return result
            .OrderByDescending(x => x.UploadedAtUtc)
            .ToArray();
    }

    public FormDataUpload? TryGetLatestUpload(string formNumber)
    {
        EnsureInitialized();

        var formDir = GetSafeSubdir(RootPath, formNumber, nameof(formNumber));
        if (!Directory.Exists(formDir))
        {
            return null;
        }

        FormDataUpload? latest = null;

        foreach (var versionDir in Directory.EnumerateDirectories(formDir, "v*"))
        {
            foreach (var uploadDir in Directory.EnumerateDirectories(versionDir))
            {
                var metaPath = Path.Combine(uploadDir, "meta.json");
                if (!File.Exists(metaPath))
                {
                    continue;
                }

                try
                {
                    var json = File.ReadAllText(metaPath);
                    var meta = System.Text.Json.JsonSerializer.Deserialize<FormDataUpload>(json, JsonUtil.StableOptions);
                    if (meta is null)
                    {
                        continue;
                    }

                    if (latest is null || meta.UploadedAtUtc > latest.UploadedAtUtc)
                    {
                        latest = meta;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read meta.json in {Dir}", uploadDir);
                }
            }
        }

        return latest;
    }

    public IReadOnlyList<FormDataUpload> ListUploads(string formNumber)
    {
        EnsureInitialized();

        var formDir = GetSafeSubdir(RootPath, formNumber, nameof(formNumber));
        if (!Directory.Exists(formDir))
        {
            return Array.Empty<FormDataUpload>();
        }

        var result = new List<FormDataUpload>();

        foreach (var versionDir in Directory.EnumerateDirectories(formDir, "v*"))
        {
            foreach (var uploadDir in Directory.EnumerateDirectories(versionDir))
            {
                var metaPath = Path.Combine(uploadDir, "meta.json");
                if (!File.Exists(metaPath))
                {
                    continue;
                }

                try
                {
                    var json = File.ReadAllText(metaPath);
                    var meta = System.Text.Json.JsonSerializer.Deserialize<FormDataUpload>(json, JsonUtil.StableOptions);
                    if (meta is not null)
                    {
                        result.Add(meta);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to read meta.json in {Dir}", uploadDir);
                }
            }
        }

        return result
            .OrderByDescending(x => x.UploadedAtUtc)
            .ToArray();
    }

    public IReadOnlyList<FormDataUpload> ListAllUploads()
    {
        EnsureInitialized();

        if (!Directory.Exists(RootPath))
        {
            return Array.Empty<FormDataUpload>();
        }

        var result = new List<FormDataUpload>();

        foreach (var formDir in Directory.EnumerateDirectories(RootPath))
        {
            foreach (var versionDir in Directory.EnumerateDirectories(formDir, "v*"))
            {
                foreach (var uploadDir in Directory.EnumerateDirectories(versionDir))
                {
                    var metaPath = Path.Combine(uploadDir, "meta.json");
                    if (!File.Exists(metaPath))
                    {
                        continue;
                    }

                    try
                    {
                        var json = File.ReadAllText(metaPath);
                        var meta = System.Text.Json.JsonSerializer.Deserialize<FormDataUpload>(json, JsonUtil.StableOptions);
                        if (meta is not null)
                        {
                            result.Add(meta);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to read meta.json in {Dir}", uploadDir);
                    }
                }
            }
        }

        return result
            .OrderByDescending(x => x.UploadedAtUtc)
            .ToArray();
    }

    public FormDataFile? TryLoadData(string formNumber, int version, string uploadId)
    {
        EnsureInitialized();

        var path = GetDataJsonPath(formNumber, version, uploadId);
        if (!File.Exists(path))
        {
            return null;
        }

        try
        {
            var json = File.ReadAllText(path);
            return System.Text.Json.JsonSerializer.Deserialize<FormDataFile>(json, JsonUtil.StableOptions);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read data.json for {FormNumber} v{Version} {UploadId}", formNumber, version, uploadId);
            return null;
        }
    }

    public FormDataFile? TryLoadData(string formNumber, int version, string uploadId, ExcelFormParser parser)
    {
        var data = TryLoadData(formNumber, version, uploadId);
        if (data is null)
        {
            return null;
        }

        if (!ContainsFormulaLikeValues(data))
        {
            return data;
        }

        try
        {
            var normalized = TryNormalizeLegacyFormulaTextData(formNumber, version, uploadId, data, parser);
            return normalized ?? data;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to normalize legacy data for {FormNumber} v{Version} {UploadId}", formNumber, version, uploadId);
            return data;
        }
    }

    private static bool ContainsFormulaLikeValues(FormDataFile data)
    {
        foreach (var r in data.Rows)
        {
            foreach (var kv in r.Values)
            {
                if (LooksLikeFormulaText(kv.Value))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool LooksLikeFormulaText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var v = value.Trim();
        if (!v.StartsWith("=", StringComparison.Ordinal))
        {
            return false;
        }

        // Heuristic: a typical A1 reference or SUM(...) indicates this is a real formula string,
        // not a user-entered value starting with '='.
        if (v.Contains("SUM(", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        for (var i = 1; i < v.Length; i++)
        {
            if (char.IsLetter(v[i]))
            {
                // expect something like A1 / AB12
                var j = i;
                while (j < v.Length && char.IsLetter(v[j])) j++;
                if (j < v.Length && char.IsDigit(v[j]))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private FormDataFile? TryNormalizeLegacyFormulaTextData(
        string formNumber,
        int version,
        string uploadId,
        FormDataFile existing,
        ExcelFormParser parser)
    {
        var originalPath = GetOriginalFilePath(formNumber, version, uploadId);
        if (!File.Exists(originalPath))
        {
            return null;
        }

        byte[] bytes;
        try
        {
            bytes = File.ReadAllBytes(originalPath);
        }
        catch
        {
            return null;
        }

        ExcelFormParser.ExcelFormLayout layout;
        using (var ms = new MemoryStream(bytes, writable: false))
        {
            layout = parser.ParseLayout(ms, existing.Upload.OriginalFileName);
        }

        if (!string.Equals(layout.Structure.StructureHash, existing.Upload.StructureHash, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        IReadOnlyList<FormDataRow> rows;
        using (var ms = new MemoryStream(bytes, writable: false))
        {
            rows = parser.ReadDataRows(ms, layout);
        }

        var normalized = new FormDataFile
        {
            Upload = existing.Upload,
            Rows = rows
        };

        var dataJsonPath = GetDataJsonPath(formNumber, version, uploadId);
        var dataJson = JsonUtil.ToStableJson(normalized);

        lock (GetFormLock(formNumber))
        {
            try
            {
                if (File.Exists(dataJsonPath))
                {
                    File.WriteAllText(dataJsonPath, dataJson);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to write normalized data.json for {FormNumber} v{Version} {UploadId}", formNumber, version, uploadId);
            }
        }

        return normalized;
    }

    public FormDataUpload? TryLoadUploadMeta(string formNumber, int version, string uploadId)
    {
        EnsureInitialized();

        var path = GetMetaJsonPath(formNumber, version, uploadId);
        if (!File.Exists(path))
        {
            return null;
        }

        try
        {
            var json = File.ReadAllText(path);
            return System.Text.Json.JsonSerializer.Deserialize<FormDataUpload>(json, JsonUtil.StableOptions);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read meta.json for {FormNumber} v{Version} {UploadId}", formNumber, version, uploadId);
            return null;
        }
    }

    public string GetOriginalFilePath(string formNumber, int version, string uploadId)
    {
        return Path.Combine(GetUploadDir(formNumber, version, uploadId), "original.xlsx");
    }

    public string GetDataJsonPath(string formNumber, int version, string uploadId)
    {
        return Path.Combine(GetUploadDir(formNumber, version, uploadId), "data.json");
    }

    public string GetMetaJsonPath(string formNumber, int version, string uploadId)
    {
        return Path.Combine(GetUploadDir(formNumber, version, uploadId), "meta.json");
    }

    public bool DeleteUpload(string formNumber, int version, string uploadId)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return false;
        }

        var uploadDir = GetUploadDir(formNumber, version, uploadId);
        if (!Directory.Exists(uploadDir))
        {
            return false;
        }

        lock (GetFormLock(formNumber))
        {
            if (!Directory.Exists(uploadDir))
            {
                return false;
            }

            try
            {
                Directory.Delete(uploadDir, recursive: true);

                // Best-effort cleanup of empty folders.
                TryDeleteDirectoryIfEmpty(GetVersionDir(formNumber, version));
                try
                {
                    TryDeleteDirectoryIfEmpty(GetSafeSubdir(RootPath, formNumber, nameof(formNumber)));
                }
                catch (ArgumentException)
                {
                    // Ignore invalid names during best-effort cleanup.
                }

                _logger.LogInformation("Deleted data upload {FormNumber} v{Version} {UploadId}", formNumber, version, uploadId);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to delete data upload {FormNumber} v{Version} {UploadId}", formNumber, version, uploadId);
                throw;
            }
        }
    }

    public bool DeleteVersion(string formNumber, int version)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0)
        {
            return false;
        }

        var versionDir = GetVersionDir(formNumber, version);
        if (!Directory.Exists(versionDir))
        {
            return false;
        }

        lock (GetFormLock(formNumber))
        {
            if (!Directory.Exists(versionDir))
            {
                return false;
            }

            try
            {
                Directory.Delete(versionDir, recursive: true);

                // Best-effort cleanup of empty folders.
                try
                {
                    TryDeleteDirectoryIfEmpty(GetSafeSubdir(RootPath, formNumber, nameof(formNumber)));
                }
                catch (ArgumentException)
                {
                    // Ignore invalid names during best-effort cleanup.
                }

                _logger.LogInformation("Deleted data uploads for {FormNumber} v{Version}", formNumber, version);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to delete data uploads for {FormNumber} v{Version}", formNumber, version);
                throw;
            }
        }
    }

    public bool DeleteForm(string formNumber)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return false;
        }

        var formDir = GetSafeSubdir(RootPath, formNumber, nameof(formNumber));
        if (!Directory.Exists(formDir))
        {
            return false;
        }

        lock (GetFormLock(formNumber))
        {
            if (!Directory.Exists(formDir))
            {
                return false;
            }

            try
            {
                Directory.Delete(formDir, recursive: true);
                _logger.LogInformation("Deleted data uploads for form {FormNumber}", formNumber);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to delete data uploads for form {FormNumber}", formNumber);
                throw;
            }
        }
    }

    public bool RenameForm(string oldFormNumber, string newFormNumber)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(oldFormNumber) || string.IsNullOrWhiteSpace(newFormNumber))
        {
            return false;
        }

        if (string.Equals(oldFormNumber, newFormNumber, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var oldDir = GetSafeSubdir(RootPath, oldFormNumber, nameof(oldFormNumber));
        if (!Directory.Exists(oldDir))
        {
            // No data uploaded yet.
            return true;
        }

        // Lock both keys in stable order to avoid deadlocks.
        var first = string.Compare(oldFormNumber, newFormNumber, StringComparison.OrdinalIgnoreCase) <= 0
            ? oldFormNumber
            : newFormNumber;
        var second = string.Equals(first, oldFormNumber, StringComparison.OrdinalIgnoreCase) ? newFormNumber : oldFormNumber;

        lock (GetFormLock(first))
        lock (GetFormLock(second))
        {
            if (!Directory.Exists(oldDir))
            {
                // No data uploaded yet.
                return true;
            }

            var newDir = GetSafeSubdir(RootPath, newFormNumber, nameof(newFormNumber));
            if (Directory.Exists(newDir))
            {
                throw new InvalidOperationException($"Data directory already exists for form '{newFormNumber}'.");
            }

            Directory.Move(oldDir, newDir);
            _logger.LogInformation("Renamed data form {OldFormNumber} -> {NewFormNumber}", oldFormNumber, newFormNumber);
            return true;
        }
    }

    private string GetUploadDir(string formNumber, int version, string uploadId)
    {
        return GetSafeSubdir(GetVersionDir(formNumber, version), uploadId, nameof(uploadId));
    }

    private string GetVersionDir(string formNumber, int version)
    {
        var formDir = GetSafeSubdir(RootPath, formNumber, nameof(formNumber));
        return Path.Combine(formDir, $"v{version}");
    }

    private static string SafeDirName(string name)
    {
        foreach (var ch in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(ch, '_');
        }

        return name.Trim();
    }

    private static string GetSafeSubdir(string baseDir, string segment, string paramName)
    {
        if (string.IsNullOrWhiteSpace(segment))
        {
            throw new ArgumentException("Directory name is required.", paramName);
        }

        var safe = SafeDirName(segment);
        if (string.IsNullOrWhiteSpace(safe) || safe is "." or "..")
        {
            throw new ArgumentException("Invalid directory name.", paramName);
        }

        var baseFull = Path.GetFullPath(baseDir);
        var candidateFull = Path.GetFullPath(Path.Combine(baseFull, safe));

        var basePrefix = baseFull.EndsWith(Path.DirectorySeparatorChar)
            ? baseFull
            : baseFull + Path.DirectorySeparatorChar;

        if (!candidateFull.StartsWith(basePrefix, StringComparison.Ordinal))
        {
            throw new ArgumentException("Invalid directory name (path traversal detected).", paramName);
        }

        return candidateFull;
    }

    private static void TryDeleteDirectoryIfEmpty(string path)
    {
        try
        {
            if (!Directory.Exists(path))
            {
                return;
            }

            if (!Directory.EnumerateFileSystemEntries(path).Any())
            {
                Directory.Delete(path);
            }
        }
        catch
        {
            // Best-effort cleanup; ignore.
        }
    }

    public sealed record SaveDataResult(string FormNumber, int Version, string UploadId, int RowCount);
}
