using SamorodinkaTech.FormStructures.Web.Models;

namespace SamorodinkaTech.FormStructures.Web.Services;

public sealed class FormDataStorage
{
    private readonly FormStorage _formStorage;
    private readonly ILogger<FormDataStorage> _logger;

    public FormDataStorage(FormStorage formStorage, ILogger<FormDataStorage> logger)
    {
        _formStorage = formStorage;
        _logger = logger;
    }

    public string RootPath => Path.Combine(_formStorage.RootPath, "data");

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
        Directory.CreateDirectory(uploadDir);

        var originalPath = Path.Combine(uploadDir, "original.xlsx");
        await File.WriteAllBytesAsync(originalPath, bytes, ct);

        var dataJson = JsonUtil.ToStableJson(dataFile);
        await File.WriteAllTextAsync(Path.Combine(uploadDir, "data.json"), dataJson, ct);

        var metaJson = JsonUtil.ToStableJson(upload);
        await File.WriteAllTextAsync(Path.Combine(uploadDir, "meta.json"), metaJson, ct);

        _logger.LogInformation(
            "Stored data upload {FormNumber} v{Version} ({Rows} rows) at {Dir}",
            upload.FormNumber,
            upload.FormVersion,
            upload.RowCount,
            uploadDir);

        return new SaveDataResult(upload.FormNumber, upload.FormVersion, upload.UploadId, upload.RowCount);
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

        var formDir = Path.Combine(RootPath, SafeDirName(formNumber));
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

        var formDir = Path.Combine(RootPath, SafeDirName(formNumber));
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

        try
        {
            Directory.Delete(uploadDir, recursive: true);

            // Best-effort cleanup of empty folders.
            TryDeleteDirectoryIfEmpty(GetVersionDir(formNumber, version));
            TryDeleteDirectoryIfEmpty(Path.Combine(RootPath, SafeDirName(formNumber)));

            _logger.LogInformation("Deleted data upload {FormNumber} v{Version} {UploadId}", formNumber, version, uploadId);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to delete data upload {FormNumber} v{Version} {UploadId}", formNumber, version, uploadId);
            throw;
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

        try
        {
            Directory.Delete(versionDir, recursive: true);

            // Best-effort cleanup of empty folders.
            TryDeleteDirectoryIfEmpty(Path.Combine(RootPath, SafeDirName(formNumber)));

            _logger.LogInformation("Deleted data uploads for {FormNumber} v{Version}", formNumber, version);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to delete data uploads for {FormNumber} v{Version}", formNumber, version);
            throw;
        }
    }

    public bool DeleteForm(string formNumber)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return false;
        }

        var formDir = Path.Combine(RootPath, SafeDirName(formNumber));
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

        var oldDir = Path.Combine(RootPath, SafeDirName(oldFormNumber));
        if (!Directory.Exists(oldDir))
        {
            // No data uploaded yet.
            return true;
        }

        var newDir = Path.Combine(RootPath, SafeDirName(newFormNumber));
        if (Directory.Exists(newDir))
        {
            throw new InvalidOperationException($"Data directory already exists for form '{newFormNumber}'.");
        }

        Directory.Move(oldDir, newDir);
        _logger.LogInformation("Renamed data form {OldFormNumber} -> {NewFormNumber}", oldFormNumber, newFormNumber);
        return true;
    }

    private string GetUploadDir(string formNumber, int version, string uploadId)
    {
        return Path.Combine(GetVersionDir(formNumber, version), SafeDirName(uploadId));
    }

    private string GetVersionDir(string formNumber, int version)
    {
        return Path.Combine(RootPath, SafeDirName(formNumber), $"v{version}");
    }

    private static string SafeDirName(string name)
    {
        foreach (var ch in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(ch, '_');
        }

        return name.Trim();
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
