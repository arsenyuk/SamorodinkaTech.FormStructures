using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;
using System.Text.Json;

namespace SamorodinkaTech.FormStructures.Web.Pages.Forms;

public class DataUploadModel : PageModel
{
    private readonly FormStorage _formStorage;
    private readonly FormDataStorage _dataStorage;

    public DataUploadModel(FormStorage formStorage, FormDataStorage dataStorage)
    {
        _formStorage = formStorage;
        _dataStorage = dataStorage;
    }

    public string FormNumber { get; private set; } = string.Empty;
    public int Version { get; private set; }
    public string UploadId { get; private set; } = string.Empty;

    public FormStructure? Structure { get; private set; }
    public FormDataFile? Data { get; private set; }

    public IReadOnlyList<ColumnDefinition> FilterColumns { get; private set; } = Array.Empty<ColumnDefinition>();
    public IReadOnlyList<ColumnDefinition> ReferenceBookColumns { get; private set; } = Array.Empty<ColumnDefinition>();
    public IReadOnlyList<ReferenceBookFilterInput> ReferenceBookFilterInputs { get; private set; } = Array.Empty<ReferenceBookFilterInput>();
    public IReadOnlyList<ReferenceBookFilter> AppliedReferenceBookFilters { get; private set; } = Array.Empty<ReferenceBookFilter>();

    /// <summary>
    /// A JSON dictionary: { "c1": ["A","B"], "c2": ["X","Y"] }.
    /// Used to populate value suggestions on the client side.
    /// </summary>
    public string ReferenceBookValuesJson { get; private set; } = "{}";
    public int TotalRowCount { get; private set; }
    public int FilteredRowCount { get; private set; }

    public string ReferenceBookQuery
        => AppliedReferenceBookFilters.Count == 0
            ? string.Empty
            : string.Concat(AppliedReferenceBookFilters.Select(f =>
                string.Concat(
                    "&rbCol=", Uri.EscapeDataString(f.Column),
                    "&rbValue=", Uri.EscapeDataString(f.Value),
                    "&rbMatch=", Uri.EscapeDataString(ToQueryToken(f.Match)))));

    public string SortKey { get; private set; } = "row";
    public string SortDir { get; private set; } = "asc";

    public IActionResult OnGet(
        string formNumber,
        int version,
        string uploadId,
        string? sort = null,
        string? dir = null,
        string[]? rbCol = null,
        string[]? rbValue = null,
        string[]? rbMatch = null)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        FormNumber = formNumber;
        Version = version;
        UploadId = uploadId;

        Structure = _formStorage.TryLoadStructure(FormNumber, Version);
        if (Structure is null)
        {
            return NotFound();
        }

        FilterColumns = Structure.Columns
            .OrderBy(c => c.Index)
            .ToArray();

        // Keep this list for value suggestions only (datalist), not for deciding filter availability.
        ReferenceBookColumns = Structure.Columns
            .Where(c => c.Type == ColumnType.ReferenceBook)
            .OrderBy(c => c.Index)
            .ToArray();

        Data = _dataStorage.TryLoadData(FormNumber, Version, UploadId);
        if (Data is null)
        {
            return NotFound();
        }

        SortKey = string.IsNullOrWhiteSpace(sort) ? "row" : sort;
        SortDir = string.Equals(dir, "desc", StringComparison.OrdinalIgnoreCase) ? "desc" : "asc";

        TotalRowCount = Data.Rows.Count;

        var referenceBooks = _formStorage.TryLoadReferenceBooks(FormNumber, Version);
        ReferenceBookColumns = ExtractFilterableReferenceBookColumns(Structure, referenceBooks);
        ReferenceBookValuesJson = BuildReferenceBookValuesJson(Structure, ReferenceBookColumns, referenceBooks, Data.Rows);

        ReferenceBookFilterInputs = BuildReferenceBookFilterInputs(rbCol, rbValue, rbMatch);

        var filteredRows = Data.Rows;

        var applied = new List<ReferenceBookFilter>();
        foreach (var input in ReferenceBookFilterInputs)
        {
            if (!TryResolveColumnPath(Structure, input.Column, out var colToken, out var colPath))
            {
                continue;
            }

            var needle = (input.Value ?? string.Empty).Trim();
            if (needle.Length == 0)
            {
                continue;
            }

            var match = NormalizeMatch(input.Match);
            applied.Add(new ReferenceBookFilter(colToken, needle, match));
            filteredRows = filteredRows
                .Where(r => MatchesFilter(r, colPath, needle, match))
                .ToArray();
        }

        AppliedReferenceBookFilters = applied;

        FilteredRowCount = filteredRows.Count;

        Data = Data with { Rows = filteredRows };

        Data = Data with
        {
            Rows = ApplySort(Data.Rows, Structure, SortKey, SortDir)
        };

        return Page();
    }

    private static bool TryResolveColumnPath(
        FormStructure structure,
        string? colTokenInput,
        out string colToken,
        out string colPath)
    {
        colToken = string.Empty;
        colPath = string.Empty;

        var token = (colTokenInput ?? string.Empty).Trim();
        if (token.Length == 0)
        {
            return false;
        }

        if (!token.StartsWith("c", StringComparison.OrdinalIgnoreCase)
            || !int.TryParse(token[1..], out var colIndex)
            || colIndex < 1
            || colIndex > structure.Columns.Count)
        {
            return false;
        }

        var col = structure.Columns[colIndex - 1];
        colToken = $"c{colIndex}";
        colPath = col.Path;
        return true;
    }

    private static IReadOnlyList<ColumnDefinition> ExtractFilterableReferenceBookColumns(
        FormStructure structure,
        IReadOnlyList<ReferenceBook> books)
    {
        var idx = new HashSet<int>();

        // 1) Columns explicitly typed as ReferenceBook.
        foreach (var c in structure.Columns)
        {
            if (c.Type == ColumnType.ReferenceBook)
            {
                idx.Add(c.Index);
            }
        }

        // 2) Columns covered by reference-books AppliedTo ranges (works for older stored schemas).
        var leafCols = ExtractLeafColumns(structure.Header);
        var hasLeafMapping = leafCols.Count == structure.Columns.Count;
        if (hasLeafMapping)
        {
            foreach (var b in books)
            {
                foreach (var appliedTo in b.AppliedTo)
                {
                    if (!TryParseA1ColumnSpan(appliedTo, out var colStart, out var colEnd))
                    {
                        continue;
                    }

                    for (var i = 0; i < leafCols.Count; i++)
                    {
                        var leafCol = leafCols[i];
                        if (leafCol < colStart || leafCol > colEnd)
                        {
                            continue;
                        }

                        idx.Add(i + 1); // 1-based index into structure.Columns
                    }
                }
            }
        }

        return structure.Columns
            .Where(c => idx.Contains(c.Index))
            .OrderBy(c => c.Index)
            .ToArray();
    }

    private static bool MatchesFilter(FormDataRow row, string colPath, string needle, ReferenceBookMatch match)
    {
        if (!row.Values.TryGetValue(colPath, out var value) || string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var hay = value.Trim();
        return match switch
        {
            ReferenceBookMatch.Equals => string.Equals(hay, needle, StringComparison.OrdinalIgnoreCase),
            ReferenceBookMatch.Contains => hay.Contains(needle, StringComparison.OrdinalIgnoreCase),
            _ => string.Equals(hay, needle, StringComparison.OrdinalIgnoreCase)
        };
    }

    public sealed record ReferenceBookFilterInput(string? Column, string? Value, string? Match);
    public sealed record ReferenceBookFilter(string Column, string Value, ReferenceBookMatch Match);

    public enum ReferenceBookMatch
    {
        Equals = 0,
        Contains = 1
    }

    private static ReferenceBookMatch NormalizeMatch(string? match)
        => string.Equals(match?.Trim(), "contains", StringComparison.OrdinalIgnoreCase)
            ? ReferenceBookMatch.Contains
            : ReferenceBookMatch.Equals;

    private static string ToQueryToken(ReferenceBookMatch match)
        => match == ReferenceBookMatch.Contains ? "contains" : "eq";

    private static IReadOnlyList<ReferenceBookFilterInput> BuildReferenceBookFilterInputs(
        string[]? rbCols,
        string[]? rbValues,
        string[]? rbMatches)
    {
        // Show exactly the number of provided filters.
        // If none provided, show a single empty row.
        var count = Math.Max(rbCols?.Length ?? 0, rbValues?.Length ?? 0);
        count = Math.Max(count, rbMatches?.Length ?? 0);
        count = Math.Max(1, count);

        // Prevent unbounded query strings from generating massive forms.
        count = Math.Min(count, 20);

        var result = new List<ReferenceBookFilterInput>(count);
        for (var i = 0; i < count; i++)
        {
            var c = rbCols is not null && i < rbCols.Length ? rbCols[i] : null;
            var v = rbValues is not null && i < rbValues.Length ? rbValues[i] : null;
            var m = rbMatches is not null && i < rbMatches.Length ? rbMatches[i] : null;
            result.Add(new ReferenceBookFilterInput(c, v, m));
        }

        return result;
    }

    private static string BuildReferenceBookValuesJson(
        FormStructure structure,
        IReadOnlyList<ColumnDefinition> referenceBookColumns,
        IReadOnlyList<ReferenceBook> books,
        IReadOnlyList<FormDataRow> dataRows)
    {
        var valuesByToken = BuildReferenceBookValuesByToken(structure, referenceBookColumns, books, dataRows);
        return JsonSerializer.Serialize(valuesByToken);
    }

    private static Dictionary<string, IReadOnlyList<string>> BuildReferenceBookValuesByToken(
        FormStructure structure,
        IReadOnlyList<ColumnDefinition> referenceBookColumns,
        IReadOnlyList<ReferenceBook> books,
        IReadOnlyList<FormDataRow> dataRows)
    {
        var leafCols = ExtractLeafColumns(structure.Header);
        var hasLeafMapping = leafCols.Count == structure.Columns.Count;

        var map = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        // 1) Prefer explicit reference-books.json values, mapped via AppliedTo ranges.
        if (hasLeafMapping)
        {
            foreach (var b in books)
            {
                foreach (var appliedTo in b.AppliedTo)
                {
                    if (!TryParseA1ColumnSpan(appliedTo, out var colStart, out var colEnd))
                    {
                        continue;
                    }

                    for (var i = 0; i < leafCols.Count; i++)
                    {
                        var leafCol = leafCols[i];
                        if (leafCol < colStart || leafCol > colEnd)
                        {
                            continue;
                        }

                        var token = $"c{i + 1}";
                        if (!map.TryGetValue(token, out var set))
                        {
                            set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            map[token] = set;
                        }

                        foreach (var v in b.Values)
                        {
                            if (!string.IsNullOrWhiteSpace(v))
                            {
                                set.Add(v.Trim());
                            }
                        }
                    }
                }
            }
        }

        // 2) Fallback: derive distinct values from the actual stored data.
        foreach (var col in referenceBookColumns)
        {
            var idx = col.Index - 1;
            if (idx < 0 || idx >= structure.Columns.Count)
            {
                continue;
            }

            var token = $"c{col.Index}";
            if (!map.TryGetValue(token, out var set) || set.Count == 0)
            {
                set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                map[token] = set;

                foreach (var r in dataRows)
                {
                    if (r.Values.TryGetValue(col.Path, out var raw) && !string.IsNullOrWhiteSpace(raw))
                    {
                        set.Add(raw.Trim());
                    }
                }
            }
        }

        var result = new Dictionary<string, IReadOnlyList<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var (token, set) in map)
        {
            result[token] = set.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();
        }

        return result;
    }

    private static List<int> ExtractLeafColumns(IReadOnlyList<HeaderNode> roots)
    {
        var cols = new List<int>();

        void Walk(HeaderNode node)
        {
            if (node.Children.Count == 0)
            {
                cols.Add(node.ColStart);
                return;
            }

            foreach (var child in node.Children.OrderBy(c => c.ColStart).ThenBy(c => c.RowStart))
            {
                Walk(child);
            }
        }

        foreach (var root in roots.OrderBy(r => r.ColStart).ThenBy(r => r.RowStart))
        {
            Walk(root);
        }

        return cols;
    }

    private static bool TryParseA1ColumnSpan(string appliedTo, out int colStart, out int colEnd)
    {
        colStart = 0;
        colEnd = 0;

        if (string.IsNullOrWhiteSpace(appliedTo))
        {
            return false;
        }

        var s = appliedTo;
        var bang = s.IndexOf('!');
        if (bang >= 0)
        {
            s = s[(bang + 1)..];
        }

        var parts = s.Split(':', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0)
        {
            return false;
        }

        if (!TryParseA1Col(parts[0], out colStart))
        {
            return false;
        }

        colEnd = colStart;
        if (parts.Length >= 2 && TryParseA1Col(parts[1], out var end))
        {
            colEnd = end;
        }

        if (colEnd < colStart)
        {
            (colStart, colEnd) = (colEnd, colStart);
        }

        return colStart > 0 && colEnd > 0;
    }

    private static bool TryParseA1Col(string a1, out int col)
    {
        col = 0;

        if (string.IsNullOrWhiteSpace(a1))
        {
            return false;
        }

        var s = a1.Trim();
        if (s.StartsWith("$", StringComparison.Ordinal))
        {
            s = s[1..];
        }

        var i = 0;
        while (i < s.Length && char.IsLetter(s[i]))
        {
            i++;
        }

        if (i == 0)
        {
            return false;
        }

        var letters = s[..i].ToUpperInvariant();
        var value = 0;
        foreach (var ch in letters)
        {
            if (ch < 'A' || ch > 'Z')
            {
                return false;
            }

            value = (value * 26) + (ch - 'A' + 1);
        }

        col = value;
        return col > 0;
    }

    public IActionResult OnPostDelete(string formNumber, int version, string uploadId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        var deleted = _dataStorage.DeleteUpload(formNumber, version, uploadId);
        if (!deleted)
        {
            return NotFound();
        }

        return RedirectToPage("/Forms/Data", new { formNumber });
    }

    public IActionResult OnGetDownload(string formNumber, int version, string uploadId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        var structure = _formStorage.TryLoadStructure(formNumber, version);
        if (structure is null)
        {
            return NotFound();
        }

        var path = _dataStorage.GetOriginalFilePath(formNumber, version, uploadId);
        if (!System.IO.File.Exists(path))
        {
            return NotFound();
        }

        var downloadName = DownloadFileName.ForDataUpload(structure, version, uploadId);
        return PhysicalFile(path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", downloadName);
    }

    public IActionResult OnGetDataJson(string formNumber, int version, string uploadId)
    {
        if (string.IsNullOrWhiteSpace(formNumber) || version <= 0 || string.IsNullOrWhiteSpace(uploadId))
        {
            return NotFound();
        }

        var path = _dataStorage.GetDataJsonPath(formNumber, version, uploadId);
        if (!System.IO.File.Exists(path))
        {
            return NotFound();
        }

        var json = System.IO.File.ReadAllText(path);
        return Content(json, "application/json");
    }

    private static IReadOnlyList<FormDataRow> ApplySort(
        IReadOnlyList<FormDataRow> rows,
        FormStructure structure,
        string sortKey,
        string sortDir)
    {
        var descending = string.Equals(sortDir, "desc", StringComparison.OrdinalIgnoreCase);

        if (string.Equals(sortKey, "row", StringComparison.OrdinalIgnoreCase))
        {
            return descending
                ? rows.OrderByDescending(r => r.RowNumber).ToArray()
                : rows.OrderBy(r => r.RowNumber).ToArray();
        }

        if (sortKey.StartsWith("c", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(sortKey[1..], out var colIndex)
            && colIndex >= 1
            && colIndex <= structure.Columns.Count)
        {
            var colPath = structure.Columns[colIndex - 1].Path;
            var ordered = descending
                ? rows.OrderByDescending(r => GetSortValue(r, colPath), StringComparer.OrdinalIgnoreCase)
                : rows.OrderBy(r => GetSortValue(r, colPath), StringComparer.OrdinalIgnoreCase);

            ordered = descending
                ? ordered.ThenByDescending(r => r.RowNumber)
                : ordered.ThenBy(r => r.RowNumber);

            return ordered.ToArray();
        }

        return rows;
    }

    private static string GetSortValue(FormDataRow row, string colPath)
    {
        if (row.Values.TryGetValue(colPath, out var value) && !string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        return string.Empty;
    }
}
