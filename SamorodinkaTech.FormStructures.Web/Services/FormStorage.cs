using Microsoft.Extensions.Options;
using SamorodinkaTech.FormStructures.Web.Models;
using ClosedXML.Excel;
using System.Collections.Concurrent;

namespace SamorodinkaTech.FormStructures.Web.Services;

public sealed class FormStorage
{
    private static readonly ConcurrentDictionary<string, object> FormLocks = new(StringComparer.OrdinalIgnoreCase);

    private readonly string _root;
    private readonly ExcelFormParser _parser;
    private readonly ILogger<FormStorage> _logger;

    public FormStorage(IOptions<StorageOptions> options, IWebHostEnvironment env, ExcelFormParser parser, ILogger<FormStorage> logger)
    {
        _logger = logger;
        _parser = parser;
        var storageRoot = options.Value.StorageRoot;
        _root = Path.GetFullPath(Path.Combine(env.ContentRootPath, storageRoot));
    }

    public string RootPath => _root;

    private string FormsRootPath => Path.Combine(_root, "forms");

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
        Directory.CreateDirectory(_root);
        Directory.CreateDirectory(FormsRootPath);
    }

    public IReadOnlyList<FormLatestInfo> ListLatestForms()
    {
        EnsureInitialized();

        var formsDir = FormsRootPath;
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

    public IReadOnlyList<ReferenceBook> TryLoadReferenceBooks(string formNumber, int version)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0)
        {
            return Array.Empty<ReferenceBook>();
        }

        var path = Path.Combine(GetVersionDir(formNumber, version), "reference-books.json");
        if (!File.Exists(path))
        {
            return Array.Empty<ReferenceBook>();
        }

        try
        {
            var json = File.ReadAllText(path);
            var books = System.Text.Json.JsonSerializer.Deserialize<ReferenceBook[]>(json, JsonUtil.StableOptions)
                       ?? Array.Empty<ReferenceBook>();

            var structure = TryLoadStructure(formNumber, version);
            var originalPath = Path.Combine(GetVersionDir(formNumber, version), "original.xlsx");
            return EnhanceReferenceBookTitles(originalPath, structure, books);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to read reference-books.json for {FormNumber} v{Version}", formNumber, version);
            return Array.Empty<ReferenceBook>();
        }
    }

    private static IReadOnlyList<ReferenceBook> EnhanceReferenceBookTitles(string? originalXlsxPath, FormStructure? structure, IReadOnlyList<ReferenceBook> books)
    {
        if (books.Count == 0)
        {
            return books;
        }

        var withStructureTitles = EnhanceReferenceBookTitlesFromStructure(structure, books);

        if (string.IsNullOrWhiteSpace(originalXlsxPath) || !File.Exists(originalXlsxPath))
        {
            return withStructureTitles;
        }

        try
        {
            using var fs = File.OpenRead(originalXlsxPath);
            using var workbook = new XLWorkbook(fs);

            return withStructureTitles
                .Select(b => TryEnhanceOneFromSourceRange(workbook, b))
                .ToArray();
        }
        catch
        {
            // Best-effort only. If we can't open the template, keep stored titles.
            return withStructureTitles;
        }
    }

    private static IReadOnlyList<ReferenceBook> EnhanceReferenceBookTitlesFromStructure(FormStructure? structure, IReadOnlyList<ReferenceBook> books)
    {
        if (structure is null || structure.Columns.Count == 0)
        {
            return books;
        }

        var columnsByIndex = structure.Columns
            .Where(c => c.Index > 0 && !string.IsNullOrWhiteSpace(c.Name))
            .ToDictionary(c => c.Index, c => c.Name, EqualityComparer<int>.Default);

        if (columnsByIndex.Count == 0)
        {
            return books;
        }

        return books
            .Select(b => TryEnhanceOneFromAppliedTo(columnsByIndex, b))
            .ToArray();
    }

    private static ReferenceBook TryEnhanceOneFromAppliedTo(IReadOnlyDictionary<int, string> columnsByIndex, ReferenceBook b)
    {
        if (columnsByIndex.Count == 0)
        {
            return b;
        }

        if (!TitleLooksTechnical(b))
        {
            return b;
        }

        if (b.AppliedTo is null || b.AppliedTo.Count == 0)
        {
            return b;
        }

        var indices = new SortedSet<int>();
        foreach (var a in b.AppliedTo)
        {
            if (!TryParseColumnIndexFromA1Range(a, out var colIndex))
            {
                continue;
            }

            if (columnsByIndex.ContainsKey(colIndex))
            {
                indices.Add(colIndex);
            }
        }

        if (indices.Count == 0)
        {
            return b;
        }

        var titles = indices
            .Select(i => columnsByIndex.TryGetValue(i, out var name) ? name?.Trim() : null)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (titles.Length == 0)
        {
            return b;
        }

        if (titles.Length == 1)
        {
            return b with { Title = titles[0] };
        }

        return b with { Title = string.Join(", ", titles) };
    }

    private static bool TryParseColumnIndexFromA1Range(string? a1Range, out int columnIndex)
    {
        columnIndex = 0;

        if (string.IsNullOrWhiteSpace(a1Range))
        {
            return false;
        }

        // Examples:
        //   Form!A5:A10
        //   Form!$B$5:$B$5
        var text = a1Range.Trim();
        var excl = text.IndexOf('!');
        if (excl >= 0)
        {
            text = text[(excl + 1)..];
        }

        var colon = text.IndexOf(':');
        if (colon >= 0)
        {
            text = text[..colon];
        }

        // Strip '$' and keep leading letters.
        Span<char> lettersBuf = stackalloc char[Math.Min(text.Length, 8)];
        var n = 0;
        foreach (var ch in text)
        {
            if (ch == '$')
            {
                continue;
            }

            if (ch is >= 'A' and <= 'Z')
            {
                if (n < lettersBuf.Length)
                {
                    lettersBuf[n++] = ch;
                    continue;
                }

                // Too many letters for our buffer; fall back to string parsing.
                var s = new string(text.Where(c => c is >= 'A' and <= 'Z').ToArray());
                return TryConvertColumnLettersToIndex(s, out columnIndex);
            }

            if (ch is >= 'a' and <= 'z')
            {
                var upper = char.ToUpperInvariant(ch);
                if (n < lettersBuf.Length)
                {
                    lettersBuf[n++] = upper;
                    continue;
                }

                var s = new string(text.Where(c => char.IsLetter(c)).Select(char.ToUpperInvariant).ToArray());
                return TryConvertColumnLettersToIndex(s, out columnIndex);
            }

            // Stop at first non-letter once letters started.
            if (n > 0)
            {
                break;
            }
        }

        if (n == 0)
        {
            return false;
        }

        return TryConvertColumnLettersToIndex(lettersBuf[..n].ToString(), out columnIndex);
    }

    private static bool TryConvertColumnLettersToIndex(string letters, out int columnIndex)
    {
        columnIndex = 0;

        if (string.IsNullOrWhiteSpace(letters))
        {
            return false;
        }

        var result = 0;
        foreach (var ch in letters.Trim())
        {
            if (ch is < 'A' or > 'Z')
            {
                return false;
            }

            checked
            {
                result = (result * 26) + (ch - 'A' + 1);
            }
        }

        columnIndex = result;
        return columnIndex > 0;
    }

    private static bool TitleLooksTechnical(ReferenceBook b)
    {
        var technical = (!string.IsNullOrWhiteSpace(b.SourceSheet) && !string.IsNullOrWhiteSpace(b.SourceRange))
            ? $"{b.SourceSheet}!{b.SourceRange}"
            : null;

        if (!string.IsNullOrWhiteSpace(technical)
            && string.Equals(b.Title, technical, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!string.IsNullOrWhiteSpace(b.Title)
            && !string.IsNullOrWhiteSpace(b.SourceFormula)
            && string.Equals(b.Title.Trim(), b.SourceFormula.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        // If title is empty/null, we can enhance it.
        return string.IsNullOrWhiteSpace(b.Title);
    }

    private static ReferenceBook TryEnhanceOneFromSourceRange(XLWorkbook workbook, ReferenceBook b)
    {
        if (string.IsNullOrWhiteSpace(b.SourceSheet) || string.IsNullOrWhiteSpace(b.SourceRange))
        {
            return b;
        }

        if (!TitleLooksTechnical(b))
        {
            return b;
        }

        var ws = workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name, b.SourceSheet, StringComparison.OrdinalIgnoreCase));
        if (ws is null)
        {
            return b;
        }

        try
        {
            var range = ws.Range(b.SourceRange);
            var addr = range.RangeAddress;

            string? header = null;

            if (addr.FirstAddress.ColumnNumber == addr.LastAddress.ColumnNumber && addr.FirstAddress.RowNumber > 1)
            {
                header = ws.Cell(addr.FirstAddress.RowNumber - 1, addr.FirstAddress.ColumnNumber)
                    .GetFormattedString()?
                    .Trim();
            }
            else if (addr.FirstAddress.RowNumber == addr.LastAddress.RowNumber && addr.FirstAddress.ColumnNumber > 1)
            {
                header = ws.Cell(addr.FirstAddress.RowNumber, addr.FirstAddress.ColumnNumber - 1)
                    .GetFormattedString()?
                    .Trim();
            }

            if (string.IsNullOrWhiteSpace(header))
            {
                return b;
            }

            return b with { Title = header };
        }
        catch
        {
            return b;
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

        // Multi-thread safety: serialize user-driven schema mutations per form.
        lock (GetFormLock(formKey))
        {
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
                var pendingId = SavePendingInternal(stored, ms, previousVersion: 0);

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
                var pendingId = SavePendingInternal(stored, ms, previousVersion: newVersion - 1);

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

            var originalPath = Path.Combine(versionDir, "original.xlsx");
            ms.Position = 0;
            using (var fs = File.Create(originalPath))
            {
                ms.CopyTo(fs);
            }

            var structureJson = JsonUtil.ToStableJson(stored);
            File.WriteAllText(Path.Combine(versionDir, "structure.json"), structureJson);

            // Extract and persist reference books (data validation lists) from the template.
            ms.Position = 0;
            var books = parser.ExtractReferenceBooks(ms);
            SaveReferenceBooksInternal(formKey, newVersion, books);

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

    public IReadOnlyList<PendingMeta> ListPending(string formNumber)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return Array.Empty<PendingMeta>();
        }

        var pendingRoot = GetPendingRootDir(formNumber);
        if (!Directory.Exists(pendingRoot))
        {
            return Array.Empty<PendingMeta>();
        }

        var list = new List<PendingMeta>();

        foreach (var dir in Directory.EnumerateDirectories(pendingRoot))
        {
            try
            {
                var metaPath = Path.Combine(dir, "meta.json");
                if (!File.Exists(metaPath))
                {
                    continue;
                }

                var metaJson = File.ReadAllText(metaPath);
                var meta = System.Text.Json.JsonSerializer.Deserialize<PendingMeta>(metaJson, JsonUtil.StableOptions);
                if (meta is null)
                {
                    continue;
                }

                list.Add(meta);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to read pending meta in {Dir}", dir);
            }
        }

        return list
            .OrderByDescending(x => x.CreatedAtUtc)
            .ToArray();
    }

    public IReadOnlyList<ReferenceBook> TryLoadPendingReferenceBooks(string formNumber, string pendingId)
    {
        EnsureInitialized();

        if (string.IsNullOrWhiteSpace(formNumber) || string.IsNullOrWhiteSpace(pendingId))
        {
            return Array.Empty<ReferenceBook>();
        }

        var dir = GetPendingDir(formNumber, pendingId);
        var jsonPath = Path.Combine(dir, "reference-books.json");
        var pendingOriginalPath = Path.Combine(dir, "original.xlsx");

        if (File.Exists(jsonPath))
        {
            try
            {
                var json = File.ReadAllText(jsonPath);
                var books = System.Text.Json.JsonSerializer.Deserialize<ReferenceBook[]>(json, JsonUtil.StableOptions)
                           ?? Array.Empty<ReferenceBook>();

                var structure = TryLoadPending(formNumber, pendingId)?.Structure;
                return EnhanceReferenceBookTitles(pendingOriginalPath, structure, books);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to read pending reference-books.json for {FormNumber} ({PendingId})", formNumber, pendingId);
                return Array.Empty<ReferenceBook>();
            }
        }

        if (!File.Exists(pendingOriginalPath))
        {
            return Array.Empty<ReferenceBook>();
        }

        try
        {
            using var fs = File.OpenRead(pendingOriginalPath);
            var books = _parser.ExtractReferenceBooks(fs);

            var structure = TryLoadPending(formNumber, pendingId)?.Structure;
            books = EnhanceReferenceBookTitles(pendingOriginalPath, structure, books);

            // Best-effort cache for faster pending UI.
            TryWriteReferenceBooksJson(jsonPath, books);

            return books;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to extract pending reference books for {FormNumber} ({PendingId})", formNumber, pendingId);
            return Array.Empty<ReferenceBook>();
        }
    }

    public Task CommitPendingAsync(string formNumber, string pendingId, FormStructure finalStructure, CancellationToken ct)
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

        ct.ThrowIfCancellationRequested();

        lock (GetFormLock(formNumber))
        {
            var pendingDir = GetPendingDir(formNumber, pendingId);
            var metaPath = Path.Combine(pendingDir, "meta.json");
            var originalPath = Path.Combine(pendingDir, "original.xlsx");

            if (!Directory.Exists(pendingDir) || !File.Exists(metaPath) || !File.Exists(originalPath))
            {
                throw new DirectoryNotFoundException("Pending upload not found.");
            }

            var metaJson = File.ReadAllText(metaPath);
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

            using (var src = File.OpenRead(originalPath))
            using (var dst = File.Create(Path.Combine(versionDir, "original.xlsx")))
            {
                src.CopyTo(dst);
            }

            var stored = finalStructure with { UploadedAtUtc = DateTime.UtcNow };
            var structureJson = JsonUtil.ToStableJson(stored);
            File.WriteAllText(Path.Combine(versionDir, "structure.json"), structureJson);

            // Extract and persist reference books (data validation lists) from the committed template.
            using (var fs = File.OpenRead(Path.Combine(versionDir, "original.xlsx")))
            {
                var books = _parser.ExtractReferenceBooks(fs);
                SaveReferenceBooksInternal(formNumber, stored.Version, books);
            }

            // Remove pending after successful commit.
            DeletePending(formNumber, pendingId);

            _logger.LogInformation("Committed pending upload {FormNumber} v{Version} ({PendingId})", formNumber, stored.Version, pendingId);
        }

        return Task.CompletedTask;
    }

    private void SaveReferenceBooksInternal(string formNumber, int version, IReadOnlyList<ReferenceBook> books)
    {
        var path = Path.Combine(GetVersionDir(formNumber, version), "reference-books.json");

        TryWriteReferenceBooksJson(path, books);
    }

    private void TryWriteReferenceBooksJson(string path, IReadOnlyList<ReferenceBook> books)
    {
        try
        {
            if (books.Count == 0)
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                return;
            }

            var json = JsonUtil.ToStableJson(books);
            File.WriteAllText(path, json);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to store reference books at {Path}", path);
        }
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

        lock (GetFormLock(formNumber))
        {
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
    }

    public int CleanupOldPendingUploads(TimeSpan maxAge)
    {
        EnsureInitialized();

        if (maxAge <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(maxAge));
        }

        var cutoffUtc = DateTime.UtcNow - maxAge;
        var formsDir = FormsRootPath;
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

    public Task SaveStructureAsync(string formNumber, int version, FormStructure structure, CancellationToken ct)
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

        ct.ThrowIfCancellationRequested();

        lock (GetFormLock(formNumber))
        {
            var versionDir = GetVersionDir(formNumber, version);
            if (!Directory.Exists(versionDir))
            {
                throw new DirectoryNotFoundException($"Schema version directory not found: {versionDir}");
            }

            var structureJson = JsonUtil.ToStableJson(structure);
            File.WriteAllText(Path.Combine(versionDir, "structure.json"), structureJson);
        }

        return Task.CompletedTask;
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

        lock (GetFormLock(formNumber))
        {
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

        lock (GetFormLock(formNumber))
        {
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

    private string GetFormDir(string formNumber) => GetSafeSubdir(FormsRootPath, formNumber, nameof(formNumber));

    private string GetVersionDir(string formNumber, int version) => Path.Combine(GetFormDir(formNumber), $"v{version}");

    private string GetPendingRootDir(string formNumber) => Path.Combine(GetFormDir(formNumber), "_pending");

    private string GetPendingDir(string formNumber, string pendingId) => GetSafeSubdir(GetPendingRootDir(formNumber), pendingId, nameof(pendingId));

    private string SavePendingInternal(FormStructure structure, MemoryStream originalXlsx, int previousVersion)
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
        File.WriteAllText(Path.Combine(pendingDir, "meta.json"), metaJson);

        originalXlsx.Position = 0;
        using (var fs = File.Create(Path.Combine(pendingDir, "original.xlsx")))
        {
            originalXlsx.CopyTo(fs);
        }

        var structureJson = JsonUtil.ToStableJson(structure);
        File.WriteAllText(Path.Combine(pendingDir, "structure.json"), structureJson);

        // Extract and cache reference books for pending UI.
        try
        {
            originalXlsx.Position = 0;
            var books = _parser.ExtractReferenceBooks(originalXlsx);
            TryWriteReferenceBooksJson(Path.Combine(pendingDir, "reference-books.json"), books);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to extract reference books for pending upload {FormNumber} ({PendingId})", structure.FormNumber, pendingId);
        }

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

        lock (GetFormLock(formNumber))
        {
            var formDir = GetFormDir(formNumber);
            Directory.CreateDirectory(formDir);

            var path = GetFormMetaPath(formNumber);
            var json = JsonUtil.ToStableJson(meta);
            File.WriteAllText(path, json);
        }
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

        lock (GetFormLock(formNumber))
        {
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
