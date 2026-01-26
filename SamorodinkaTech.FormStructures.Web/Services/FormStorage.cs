using Microsoft.Extensions.Options;
using SamorodinkaTech.FormStructures.Web.Models;

namespace SamorodinkaTech.FormStructures.Web.Services;

public sealed class FormStorage
{
    private readonly string _root;
    private readonly ILogger<FormStorage> _logger;

    public FormStorage(IOptions<StorageOptions> options, IWebHostEnvironment env, ILogger<FormStorage> logger)
    {
        _logger = logger;
        var storageRoot = options.Value.StorageRoot;
        _root = Path.GetFullPath(Path.Combine(env.ContentRootPath, storageRoot));
    }

    public string RootPath => _root;

    public void EnsureInitialized()
    {
        Directory.CreateDirectory(_root);
        Directory.CreateDirectory(Path.Combine(_root, "forms"));
    }

    public IReadOnlyList<FormLatestInfo> ListLatestForms()
    {
        EnsureInitialized();

        var formsDir = Path.Combine(_root, "forms");
        if (!Directory.Exists(formsDir))
        {
            return Array.Empty<FormLatestInfo>();
        }

        var result = new List<FormLatestInfo>();
        foreach (var dir in Directory.EnumerateDirectories(formsDir))
        {
            var formNumber = Path.GetFileName(dir);
            var latest = TryGetLatestStructure(formNumber);
            if (latest is null)
            {
                continue;
            }

            var meta = TryLoadFormMeta(formNumber);
            var displayNumber = meta?.DisplayFormNumber ?? formNumber;
            var displayTitle = meta?.DisplayFormTitle ?? latest.FormTitle;

            result.Add(new FormLatestInfo(
                FormNumber: formNumber,
                DisplayFormNumber: displayNumber,
                DisplayFormTitle: displayTitle,
                Version: latest.Version,
                UploadedAtUtc: latest.UploadedAtUtc));
        }

        return result
            .OrderBy(f => f.DisplayFormTitle, StringComparer.OrdinalIgnoreCase)
            .ThenBy(f => f.DisplayFormNumber, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public FormStructure? TryGetLatestStructure(string formNumber)
    {
        var versions = ListVersions(formNumber);
        if (versions.Count == 0)
        {
            return null;
        }

        return TryLoadStructure(formNumber, versions.Max());
    }

    public int? TryFindVersionByStructureHash(string formNumber, string structureHash)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(structureHash))
        {
            return null;
        }

        var versions = ListVersions(formNumber);
        foreach (var v in versions)
        {
            var structure = TryLoadStructure(formNumber, v);
            if (structure is null)
            {
                continue;
            }

            if (string.Equals(structure.StructureHash, structureHash, StringComparison.OrdinalIgnoreCase))
            {
                return v;
            }
        }

        return null;
    }

    public IReadOnlyList<int> ListVersions(string formNumber)
    {
        EnsureInitialized();

        var formDir = GetFormDir(formNumber);
        if (!Directory.Exists(formDir))
        {
            return Array.Empty<int>();
        }

        var versions = new List<int>();
        foreach (var dir in Directory.EnumerateDirectories(formDir))
        {
            var name = Path.GetFileName(dir);
            if (!name.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (int.TryParse(name[1..], out var v))
            {
                // Only consider fully committed schema versions.
                // Pending uploads should never show up here.
                var structurePath = Path.Combine(dir, "structure.json");
                if (!File.Exists(structurePath))
                {
                    continue;
                }
                versions.Add(v);
            }
        }

        return versions.OrderByDescending(v => v).ToArray();
    }

    public FormStructure? TryLoadStructure(string formNumber, int version)
    {
        var path = Path.Combine(GetVersionDir(formNumber, version), "structure.json");
        if (!File.Exists(path))
        {
            return null;
        }

        try
        {
            var json = File.ReadAllText(path);
            return System.Text.Json.JsonSerializer.Deserialize<FormStructure>(json, JsonUtil.StableOptions);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read structure.json for {FormNumber} v{Version}", formNumber, version);
            return null;
        }
    }

    public async Task<SaveResult> SaveAsync(IFormFile file, ExcelFormParser parser, CancellationToken ct, string? targetFormNumber = null)
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

        _logger.LogInformation("Upload received: {FileName} ({Size} bytes)", file.FileName, file.Length);

        await using var ms = new MemoryStream();
        await file.CopyToAsync(ms, ct);
        ms.Position = 0;

        var parsed = parser.Parse(ms, file.FileName);
        _logger.LogInformation("Parsed form {FormNumber}: {Title}", parsed.FormNumber, parsed.FormTitle);

        var formKey = string.IsNullOrWhiteSpace(targetFormNumber) ? parsed.FormNumber : targetFormNumber;

        EnsureFormMetaExists(
            formNumber: formKey,
            displayFormNumber: parsed.FormNumber,
            displayFormTitle: parsed.FormTitle);

        var (latestVersion, latestHash) = GetLatestVersionInfo(formKey);
        var isNewVersion = latestVersion == 0 || !string.Equals(latestHash, parsed.StructureHash, StringComparison.OrdinalIgnoreCase);

        var newVersion = isNewVersion ? (latestVersion + 1) : latestVersion;

        if (!isNewVersion)
        {
            _logger.LogInformation("Structure unchanged for {FormNumber}; keeping version v{Version}", parsed.FormNumber, newVersion);
            return new SaveResult(
                formKey,
                parsed.FormTitle,
                newVersion,
                IsNewVersion: false,
                PreviousVersion: latestVersion == 0 ? null : latestVersion,
                RequiresTypeSetup: false,
                RequiresColumnMapping: false,
                UnmatchedNewColumnCount: 0,
                PendingId: null);
        }

        // Carry column types from previous version if possible (match by Path).
        var unmatchedNewColumns = 0;
        IReadOnlyList<ColumnDefinition> columnsWithTypes = parsed.Columns;
        if (newVersion > 1)
        {
            var previous = TryLoadStructure(formKey, newVersion - 1);
            if (previous is not null)
            {
                var prevByPath = previous.Columns.ToDictionary(c => c.Path, StringComparer.Ordinal);
                columnsWithTypes = parsed.Columns
                    .Select(c =>
                    {
                        if (prevByPath.TryGetValue(c.Path, out var prev))
                        {
                            return c with { Type = prev.Type };
                        }

                        unmatchedNewColumns++;
                        return c;
                    })
                    .ToArray();
            }
            else
            {
                unmatchedNewColumns = parsed.Columns.Count;
            }
        }

        var stored = parsed with
        {
            FormNumber = formKey,
            TemplateFormNumber = string.Equals(formKey, parsed.FormNumber, StringComparison.OrdinalIgnoreCase)
                ? null
                : parsed.FormNumber,
            Version = newVersion,
            UploadedAtUtc = DateTime.UtcNow,
            Columns = columnsWithTypes
        };

        // For the first version, require explicit type setup before committing the schema.
        // This keeps the form out of the main list until the user confirms types.
        if (newVersion == 1)
        {
            var pendingId = await SavePendingAsyncInternal(stored, ms, previousVersion: 0, ct);

            _logger.LogInformation(
                "Staged pending upload {FormNumber} v{Version} as {PendingId} (requires type setup)",
                stored.FormNumber,
                stored.Version,
                pendingId);

            return new SaveResult(
                formKey,
                stored.FormTitle,
                stored.Version,
                IsNewVersion: true,
                PreviousVersion: null,
                RequiresTypeSetup: true,
                RequiresColumnMapping: false,
                UnmatchedNewColumnCount: 0,
                PendingId: pendingId);
        }

        // If this upload introduced new columns that can't be auto-matched by Path,
        // do not create a new version yet. Stage it as a pending upload and require
        // the user to confirm column mapping before committing.
        var requiresColumnMapping = newVersion > 1 && unmatchedNewColumns > 0;
        if (requiresColumnMapping)
        {
            var pendingId = await SavePendingAsyncInternal(stored, ms, previousVersion: newVersion - 1, ct);

            _logger.LogInformation(
                "Staged pending upload {FormNumber} v{Version} as {PendingId} (unmatched columns: {Count})",
                stored.FormNumber,
                stored.Version,
                pendingId,
                unmatchedNewColumns);

            return new SaveResult(
                formKey,
                stored.FormTitle,
                stored.Version,
                IsNewVersion: true,
                PreviousVersion: newVersion - 1,
                RequiresTypeSetup: false,
                RequiresColumnMapping: true,
                UnmatchedNewColumnCount: unmatchedNewColumns,
                PendingId: pendingId);
        }

            var formDir = GetFormDir(formKey);
            var versionDir = GetVersionDir(formKey, newVersion);
            Directory.CreateDirectory(formDir);
            Directory.CreateDirectory(versionDir);

            // Save original file
            var originalPath = Path.Combine(versionDir, "original.xlsx");
            ms.Position = 0;
            await using (var fs = File.Create(originalPath))
            {
                await ms.CopyToAsync(fs, ct);
            }

        var structureJson = JsonUtil.ToStableJson(stored);
        await File.WriteAllTextAsync(Path.Combine(versionDir, "structure.json"), structureJson, ct);

        _logger.LogInformation("Stored {FormNumber} v{Version} at {Dir}", formKey, newVersion, versionDir);

        return new SaveResult(
            formKey,
            parsed.FormTitle,
            newVersion,
            IsNewVersion: true,
                PreviousVersion: newVersion > 1 ? newVersion - 1 : null,
                RequiresTypeSetup: false,
            RequiresColumnMapping: false,
            UnmatchedNewColumnCount: unmatchedNewColumns,
            PendingId: null);
    }

    public PendingUpload? TryLoadPending(string formNumber, string pendingId)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(pendingId))
        {
            return null;
        }

        var dir = GetPendingDir(formNumber, pendingId);
        var metaPath = Path.Combine(dir, "meta.json");
        var structurePath = Path.Combine(dir, "structure.json");
        if (!File.Exists(metaPath) || !File.Exists(structurePath))
        {
            return null;
        }

        try
        {
            var metaJson = File.ReadAllText(metaPath);
            var meta = System.Text.Json.JsonSerializer.Deserialize<PendingMeta>(metaJson, JsonUtil.StableOptions);
            if (meta is null)
            {
                return null;
            }

            var structureJson = File.ReadAllText(structurePath);
            var structure = System.Text.Json.JsonSerializer.Deserialize<FormStructure>(structureJson, JsonUtil.StableOptions);
            if (structure is null)
            {
                return null;
            }

            return new PendingUpload(meta, structure);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read pending upload for {FormNumber} ({PendingId})", formNumber, pendingId);
            return null;
        }
    }

    public async Task CommitPendingAsync(string formNumber, string pendingId, FormStructure finalStructure, CancellationToken ct)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(pendingId))
        {
            throw new ArgumentException("Invalid formNumber/pendingId.");
        }

        if (!string.Equals(finalStructure.FormNumber, formNumber, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Structure does not match target form.");
        }

        var pendingDir = GetPendingDir(formNumber, pendingId);
        var metaPath = Path.Combine(pendingDir, "meta.json");
        var originalPath = Path.Combine(pendingDir, "original.xlsx");

        if (!Directory.Exists(pendingDir) || !File.Exists(metaPath) || !File.Exists(originalPath))
        {
            throw new DirectoryNotFoundException("Pending upload not found.");
        }

        var metaJson = await File.ReadAllTextAsync(metaPath, ct);
        var meta = System.Text.Json.JsonSerializer.Deserialize<PendingMeta>(metaJson, JsonUtil.StableOptions)
                   ?? throw new InvalidOperationException("Pending upload metadata is invalid.");

        if (meta.IntendedVersion != finalStructure.Version)
        {
            throw new InvalidOperationException("Pending upload version does not match.");
        }

        var versionDir = GetVersionDir(formNumber, finalStructure.Version);
        if (Directory.Exists(versionDir))
        {
            throw new InvalidOperationException($"Schema version already exists: v{finalStructure.Version}.");
        }

        Directory.CreateDirectory(GetFormDir(formNumber));
        Directory.CreateDirectory(versionDir);

        // Copy original file from pending.
        await using (var src = File.OpenRead(originalPath))
        await using (var dst = File.Create(Path.Combine(versionDir, "original.xlsx")))
        {
            await src.CopyToAsync(dst, ct);
        }

        var stored = finalStructure with { UploadedAtUtc = DateTime.UtcNow };
        var structureJson = JsonUtil.ToStableJson(stored);
        await File.WriteAllTextAsync(Path.Combine(versionDir, "structure.json"), structureJson, ct);

        // Remove pending after successful commit.
        DeletePending(formNumber, pendingId);

        _logger.LogInformation("Committed pending upload {FormNumber} v{Version} ({PendingId})", formNumber, stored.Version, pendingId);
    }

    public bool DeletePending(string formNumber, string pendingId)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(pendingId))
        {
            return false;
        }

        var dir = GetPendingDir(formNumber, pendingId);
        if (!Directory.Exists(dir))
        {
            return false;
        }

        try
        {
            Directory.Delete(dir, recursive: true);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to delete pending upload {FormNumber} ({PendingId})", formNumber, pendingId);
            throw;
        }
    }

    public int CleanupOldPendingUploads(TimeSpan maxAge)
    {
        EnsureInitialized();

        if (maxAge <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(maxAge));
        }

        var cutoffUtc = DateTime.UtcNow - maxAge;
        var formsDir = Path.Combine(_root, "forms");
        if (!Directory.Exists(formsDir))
        {
            return 0;
        }

        var deleted = 0;
        foreach (var formDir in Directory.EnumerateDirectories(formsDir))
        {
            var formNumber = Path.GetFileName(formDir);
            var pendingRoot = Path.Combine(formDir, "_pending");
            if (!Directory.Exists(pendingRoot))
            {
                continue;
            }

            foreach (var pendingDir in Directory.EnumerateDirectories(pendingRoot))
            {
                var pendingId = Path.GetFileName(pendingDir);
                var metaPath = Path.Combine(pendingDir, "meta.json");
                if (!File.Exists(metaPath))
                {
                    continue;
                }

                try
                {
                    var metaJson = File.ReadAllText(metaPath);
                    var meta = System.Text.Json.JsonSerializer.Deserialize<PendingMeta>(metaJson, JsonUtil.StableOptions);
                    if (meta is null)
                    {
                        continue;
                    }

                    if (meta.CreatedAtUtc >= cutoffUtc)
                    {
                        continue;
                    }

                    Directory.Delete(pendingDir, recursive: true);
                    deleted++;
                    _logger.LogInformation(
                        "Deleted old pending upload {FormNumber} ({PendingId}); created at {CreatedAtUtc}",
                        formNumber,
                        pendingId,
                        meta.CreatedAtUtc);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to cleanup pending upload {FormNumber} ({PendingId})", formNumber, pendingId);
                }
            }

            TryDeleteDirectoryIfEmpty(pendingRoot);
        }

        return deleted;
    }

    public async Task SaveStructureAsync(string formNumber, int version, FormStructure structure, CancellationToken ct)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0)
        {
            throw new ArgumentException("Invalid formNumber/version.");
        }

        if (!string.Equals(structure.FormNumber, formNumber, StringComparison.OrdinalIgnoreCase) || structure.Version != version)
        {
            throw new ArgumentException("Structure does not match target form/version.");
        }

        var versionDir = GetVersionDir(formNumber, version);
        if (!Directory.Exists(versionDir))
        {
            throw new DirectoryNotFoundException($"Schema version directory not found: {versionDir}");
        }

        var structureJson = JsonUtil.ToStableJson(structure);
        await File.WriteAllTextAsync(Path.Combine(versionDir, "structure.json"), structureJson, ct);
    }

    public string GetOriginalFilePath(string formNumber, int version)
    {
        return Path.Combine(GetVersionDir(formNumber, version), "original.xlsx");
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
            TryDeleteDirectoryIfEmpty(GetFormDir(formNumber));
            _logger.LogInformation("Deleted schema {FormNumber} v{Version}", formNumber, version);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to delete schema {FormNumber} v{Version}", formNumber, version);
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

        var formDir = GetFormDir(formNumber);
        if (!Directory.Exists(formDir))
        {
            return false;
        }

        try
        {
            Directory.Delete(formDir, recursive: true);
            _logger.LogInformation("Deleted schema form {FormNumber}", formNumber);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to delete schema form {FormNumber}", formNumber);
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

        var oldDir = GetFormDir(oldFormNumber);
        if (!Directory.Exists(oldDir))
        {
            return false;
        }

        var newDir = GetFormDir(newFormNumber);
        if (Directory.Exists(newDir))
        {
            throw new InvalidOperationException($"Schema directory already exists for form '{newFormNumber}'.");
        }

        Directory.Move(oldDir, newDir);
        _logger.LogInformation("Renamed schema form {OldFormNumber} -> {NewFormNumber}", oldFormNumber, newFormNumber);
        return true;
    }

    private (int latestVersion, string? latestHash) GetLatestVersionInfo(string formNumber)
    {
        var versions = ListVersions(formNumber);
        if (versions.Count == 0)
        {
            return (0, null);
        }

        var latestVersion = versions.Max();
        var latestStructure = TryLoadStructure(formNumber, latestVersion);
        return (latestVersion, latestStructure?.StructureHash);
    }

    private string GetFormDir(string formNumber) => Path.Combine(_root, "forms", SafeDirName(formNumber));

    private string GetVersionDir(string formNumber, int version) => Path.Combine(GetFormDir(formNumber), $"v{version}");

    private string GetPendingRootDir(string formNumber) => Path.Combine(GetFormDir(formNumber), "_pending");

    private string GetPendingDir(string formNumber, string pendingId) => Path.Combine(GetPendingRootDir(formNumber), SafeDirName(pendingId));

    private async Task<string> SavePendingAsyncInternal(FormStructure structure, MemoryStream originalXlsx, int previousVersion, CancellationToken ct)
    {
        var pendingId = Guid.NewGuid().ToString("n");
        var pendingDir = GetPendingDir(structure.FormNumber, pendingId);
        Directory.CreateDirectory(GetPendingRootDir(structure.FormNumber));
        Directory.CreateDirectory(pendingDir);

        var meta = new PendingMeta(
            PendingId: pendingId,
            CreatedAtUtc: DateTime.UtcNow,
            PreviousVersion: previousVersion,
            IntendedVersion: structure.Version);

        var metaJson = JsonUtil.ToStableJson(meta);
        await File.WriteAllTextAsync(Path.Combine(pendingDir, "meta.json"), metaJson, ct);

        originalXlsx.Position = 0;
        await using (var fs = File.Create(Path.Combine(pendingDir, "original.xlsx")))
        {
            await originalXlsx.CopyToAsync(fs, ct);
        }

        var structureJson = JsonUtil.ToStableJson(structure);
        await File.WriteAllTextAsync(Path.Combine(pendingDir, "structure.json"), structureJson, ct);

        return pendingId;
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

    public FormMeta? TryLoadFormMeta(string formNumber)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return null;
        }

        var path = GetFormMetaPath(formNumber);
        if (!File.Exists(path))
        {
            return null;
        }

        try
        {
            var json = File.ReadAllText(path);
            return System.Text.Json.JsonSerializer.Deserialize<FormMeta>(json, JsonUtil.StableOptions);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read meta.json for {FormNumber}", formNumber);
            return null;
        }
    }

    public void SaveFormMeta(string formNumber, FormMeta meta)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            throw new ArgumentException("Form number is required.", nameof(formNumber));
        }

        EnsureInitialized();

        var formDir = GetFormDir(formNumber);
        Directory.CreateDirectory(formDir);

        var path = GetFormMetaPath(formNumber);
        var json = JsonUtil.ToStableJson(meta);
        File.WriteAllText(path, json);
    }

    private void EnsureFormMetaExists(string formNumber, string displayFormNumber, string displayFormTitle)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return;
        }

        var path = GetFormMetaPath(formNumber);
        if (File.Exists(path))
        {
            return;
        }

        SaveFormMeta(formNumber, new FormMeta
        {
            DisplayFormNumber = displayFormNumber,
            DisplayFormTitle = displayFormTitle,
            UpdatedAtUtc = DateTime.UtcNow
        });
    }

    private string GetFormMetaPath(string formNumber)
    {
        return Path.Combine(GetFormDir(formNumber), "meta.json");
    }

    public sealed record FormLatestInfo(
        string FormNumber,
        string DisplayFormNumber,
        string DisplayFormTitle,
        int Version,
        DateTime UploadedAtUtc);

    public sealed record PendingMeta(string PendingId, DateTime CreatedAtUtc, int PreviousVersion, int IntendedVersion);

    public sealed record PendingUpload(PendingMeta Meta, FormStructure Structure);

    public sealed record SaveResult(
        string FormNumber,
        string FormTitle,
        int Version,
        bool IsNewVersion,
        int? PreviousVersion,
        bool RequiresTypeSetup,
        bool RequiresColumnMapping,
        int UnmatchedNewColumnCount,
        string? PendingId);
}
