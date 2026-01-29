using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ClosedXML.Excel;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Web.Pages.Forms;

public class AggregatedModel : PageModel
{
    private readonly FormStorage _formStorage;
    private readonly FormDataStorage _dataStorage;
    private readonly ExcelFormParser _parser;
    private readonly ILogger<AggregatedModel> _logger;

    public AggregatedModel(
        FormStorage formStorage,
        FormDataStorage dataStorage,
        ExcelFormParser parser,
        ILogger<AggregatedModel> logger)
    {
        _formStorage = formStorage;
        _dataStorage = dataStorage;
        _parser = parser;
        _logger = logger;
    }

    [BindProperty]
    public IFormFile? DataUpload { get; set; }

    public string FormNumber { get; private set; } = string.Empty;
    public int Version { get; private set; }

    public FormStructure? Structure { get; private set; }

    public string SortKey { get; private set; } = "uploaded";
    public string SortDir { get; private set; } = "desc";

    public bool ShowUploadedColumn { get; private set; }
    public bool ShowFileColumn { get; private set; }
    public bool ShowUploadIdColumn { get; private set; }

    public IReadOnlyList<ColumnDefinition> VisibleColumns { get; private set; } = Array.Empty<ColumnDefinition>();

    /// <summary>
    /// The normalized selected column tokens for query string propagation.
    /// Empty means "default" (all data columns, no tech columns).
    /// </summary>
    public IReadOnlyList<string> SelectedCols { get; private set; } = Array.Empty<string>();

    public string SelectedColsQuery
        => SelectedCols.Count == 0
            ? string.Empty
            : string.Concat(SelectedCols.Select(c => $"&cols={Uri.EscapeDataString(c)}"));

    public int UploadCount { get; private set; }

    public IReadOnlyList<AggregatedRow> Rows { get; private set; } = Array.Empty<AggregatedRow>();

    public IActionResult OnGet(
        string formNumber,
        int? version = null,
        string? sort = null,
        string? dir = null,
        string[]? cols = null)
    {
        var loadResult = TryLoad(formNumber, version, sort, dir, cols);
        if (loadResult is not null)
        {
            return loadResult;
        }

        return Page();
    }

    public async Task<IActionResult> OnPostUploadDataAsync(
        string formNumber,
        int? version = null,
        string? sort = null,
        string? dir = null,
        string[]? cols = null,
        CancellationToken ct = default)
    {
        var loadResult = TryLoad(formNumber, version, sort, dir, cols);
        if (loadResult is not null)
        {
            return loadResult;
        }

        if (DataUpload is null)
        {
            ModelState.AddModelError(nameof(DataUpload), "Please choose a .xlsx file.");
            return Page();
        }

        try
        {
            var result = await _dataStorage.SaveAsync(
                DataUpload,
                _parser,
                ct,
                expectedFormNumber: FormNumber,
                targetFormNumber: FormNumber);

            TempData["UploadMessage"] = $"Uploaded '{DataUpload.FileName}' ({result.RowCount} rows) as v{result.Version}.";

            var redirect = $"/forms/{Uri.EscapeDataString(FormNumber)}/data/aggregated" +
                           $"?version={result.Version}" +
                           $"&sort={Uri.EscapeDataString(SortKey)}" +
                           $"&dir={Uri.EscapeDataString(SortDir)}" +
                           SelectedColsQuery;
            return Redirect(redirect);
        }
        catch (FormParseException ex)
        {
            _logger.LogWarning(
                ex,
                "Failed to parse uploaded data file for form {FormNumber}. ExceptionChain={ExceptionChain}",
                FormNumber,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
            return Page();
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Failed to upload data file for form {FormNumber}. ExceptionChain={ExceptionChain}",
                FormNumber,
                ExceptionUtil.FormatExceptionChain(ex));
            ModelState.AddModelError(string.Empty, ExceptionUtil.FormatExceptionChain(ex));
            return Page();
        }
    }

    public IActionResult OnGetXlsx(
        string formNumber,
        int? version = null,
        string? sort = null,
        string? dir = null,
        string[]? cols = null)
    {
        var loadResult = TryLoad(formNumber, version, sort, dir, cols);
        if (loadResult is not null)
        {
            return loadResult;
        }

        if (Structure is null)
        {
            return NotFound();
        }

        using var workbook = BuildAggregatedWorkbook(
            Structure,
            Rows,
            ShowUploadedColumn,
            ShowFileColumn,
            ShowUploadIdColumn,
            VisibleColumns);
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        var bytes = ms.ToArray();

        var downloadName = DownloadFileName.ForAggregated(Structure, Version);
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", downloadName);
    }

    private IActionResult? TryLoad(
        string formNumber,
        int? version,
        string? sort,
        string? dir,
        string[]? cols)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            return NotFound();
        }

        FormNumber = formNumber;

        var selectedVersion = version;
        if (selectedVersion is null)
        {
            var latest = _formStorage.TryGetLatestStructure(FormNumber);
            if (latest is null)
            {
                return NotFound();
            }

            selectedVersion = latest.Version;
        }

        Version = selectedVersion.Value;

        Structure = _formStorage.TryLoadStructure(FormNumber, Version);
        if (Structure is null)
        {
            return NotFound();
        }

        var uploads = _dataStorage.ListUploads(FormNumber, Version);
        UploadCount = uploads.Count;

        var rows = new List<AggregatedRow>();
        foreach (var u in uploads)
        {
            var data = _dataStorage.TryLoadData(FormNumber, u.FormVersion, u.UploadId);
            if (data is null)
            {
                continue;
            }

            foreach (var r in data.Rows)
            {
                rows.Add(new AggregatedRow(
                    UploadId: u.UploadId,
                    FormVersion: u.FormVersion,
                    OriginalFileName: u.OriginalFileName,
                    UploadedAtUtc: u.UploadedAtUtc,
                    RowNumber: r.RowNumber,
                    Values: r.Values));
            }
        }

        SortKey = string.IsNullOrWhiteSpace(sort) ? "uploaded" : sort;
        SortDir = string.Equals(dir, "asc", StringComparison.OrdinalIgnoreCase) ? "asc" : "desc";

        Rows = ApplySort(rows, Structure, SortKey, SortDir);

        ApplyColumnSelection(Structure, cols);
        return null;
    }

    private static XLWorkbook BuildAggregatedWorkbook(
        FormStructure structure,
        IReadOnlyList<AggregatedRow> rows,
        bool showUploaded,
        bool showFile,
        bool showUploadId,
        IReadOnlyList<ColumnDefinition> visibleColumns)
    {
        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Aggregated");

        var allNodes = new List<HeaderNode>();
        void Walk(HeaderNode node)
        {
            allNodes.Add(node);
            foreach (var child in node.Children)
            {
                Walk(child);
            }
        }

        foreach (var root in structure.Header.OrderBy(n => n.ColStart).ThenBy(n => n.RowStart))
        {
            Walk(root);
        }

        var hasHeader = allNodes.Count > 0;
        var minHeaderRow = hasHeader ? allNodes.Min(n => n.RowStart) : 1;
        var maxHeaderRow = hasHeader ? allNodes.Max(n => n.RowEnd) : 1;
        var minHeaderCol = hasHeader ? allNodes.Min(n => n.ColStart) : 1;
        var maxHeaderCol = hasHeader ? allNodes.Max(n => n.ColEnd) : structure.Columns.Count;

        var headerHeight = maxHeaderRow - minHeaderRow + 1;
        if (headerHeight < 1)
        {
            headerHeight = 1;
        }

        var leaves = hasHeader
            ? allNodes.Where(n => n.Children.Count == 0).OrderBy(n => n.ColStart).ToArray()
            : Array.Empty<HeaderNode>();

        var fixedCols = new List<(string Key, string Label)>();
        if (showUploaded) fixedCols.Add(("uploaded", "Uploaded (UTC)"));
        if (showFile) fixedCols.Add(("file", "File"));
        if (showUploadId) fixedCols.Add(("uploadId", "UploadId"));

        var fixedColCount = fixedCols.Count;

        // Visible data columns (in worksheet order)
        var visibleData = visibleColumns
            .Select(c => (Index: c.Index - 1, Column: c))
            .Where(x => x.Index >= 0 && x.Index < structure.Columns.Count)
            .ToArray();

        var visibleLeafColStarts = hasHeader && leaves.Length == structure.Columns.Count
            ? visibleData
                .Select(v => leaves[v.Index].ColStart)
                .Distinct()
                .OrderBy(x => x)
                .ToArray()
            : Array.Empty<int>();

        var visibleLeafPosByColStart = new Dictionary<int, int>();
        for (var i = 0; i < visibleLeafColStarts.Length; i++)
        {
            visibleLeafPosByColStart[visibleLeafColStarts[i]] = i + 1; // 1-based
        }

        // Fixed columns header (merged vertically)
        for (var c = 0; c < fixedCols.Count; c++)
        {
            var col = c + 1;
            var range = ws.Range(1, col, headerHeight, col);
            range.Merge();
            range.Value = fixedCols[c].Label;
        }

        // Form header nodes (shifted right by fixed columns, normalized to row=1)
        if (hasHeader && leaves.Length == structure.Columns.Count)
        {
            foreach (var n in allNodes.OrderBy(n => n.RowStart).ThenBy(n => n.ColStart))
            {
                var visibleSpan = 0;
                for (var col = n.ColStart; col <= n.ColEnd; col++)
                {
                    if (visibleLeafPosByColStart.ContainsKey(col))
                    {
                        visibleSpan++;
                    }
                }

                if (visibleSpan <= 0)
                {
                    continue;
                }

                var rowStart = (n.RowStart - minHeaderRow) + 1;
                var rowEnd = (n.RowEnd - minHeaderRow) + 1;

                var minVisiblePos = int.MaxValue;
                var maxVisiblePos = int.MinValue;
                for (var col = n.ColStart; col <= n.ColEnd; col++)
                {
                    if (visibleLeafPosByColStart.TryGetValue(col, out var pos))
                    {
                        minVisiblePos = Math.Min(minVisiblePos, pos);
                        maxVisiblePos = Math.Max(maxVisiblePos, pos);
                    }
                }

                var colStart = fixedColCount + minVisiblePos;
                var colEnd = fixedColCount + maxVisiblePos;

                var range = ws.Range(rowStart, colStart, rowEnd, colEnd);
                range.Merge();
                range.Value = n.Label;
            }
        }
        else
        {
            // Fallback: single-row header using visible column names
            for (var i = 0; i < visibleData.Length; i++)
            {
                ws.Cell(1, fixedColCount + i + 1).Value = visibleData[i].Column.Name;
            }
        }

        // Basic header styling
        var totalCols = Math.Max(1, fixedColCount + visibleData.Length);
        var headerRange = ws.Range(1, 1, headerHeight, Math.Max(1, totalCols));
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        headerRange.Style.Alignment.WrapText = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
        headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Data
        var startRow = headerHeight + 1;
        for (var i = 0; i < rows.Count; i++)
        {
            var excelRow = startRow + i;
            var r = rows[i];

            var nextFixedCol = 1;
            if (showUploaded)
            {
                ws.Cell(excelRow, nextFixedCol++).Value = r.UploadedAtUtc.ToString("yyyy-MM-dd HH:mm:ss");
            }

            if (showFile)
            {
                ws.Cell(excelRow, nextFixedCol++).Value = r.OriginalFileName;
            }

            if (showUploadId)
            {
                ws.Cell(excelRow, nextFixedCol++).Value = r.UploadId;
            }

            if (hasHeader && leaves.Length == structure.Columns.Count)
            {
                for (var iCol = 0; iCol < visibleData.Length; iCol++)
                {
                    var originalIndex = visibleData[iCol].Index;
                    var colPath = structure.Columns[originalIndex].Path;
                    r.Values.TryGetValue(colPath, out var value);

                    var leafColStart = leaves[originalIndex].ColStart;
                    if (!visibleLeafPosByColStart.TryGetValue(leafColStart, out var pos))
                    {
                        continue;
                    }

                    var excelCol = fixedColCount + pos;
                    ws.Cell(excelRow, excelCol).Value = value ?? string.Empty;
                }
            }
            else
            {
                for (var iCol = 0; iCol < visibleData.Length; iCol++)
                {
                    var originalIndex = visibleData[iCol].Index;
                    var colPath = structure.Columns[originalIndex].Path;
                    r.Values.TryGetValue(colPath, out var value);
                    ws.Cell(excelRow, fixedColCount + iCol + 1).Value = value ?? string.Empty;
                }
            }
        }

        ws.SheetView.FreezeRows(headerHeight);
        if (fixedColCount > 0)
        {
            ws.SheetView.FreezeColumns(fixedColCount);
        }

        ws.Columns(1, totalCols).AdjustToContents();
        ws.Rows(1, headerHeight).AdjustToContents();

        return workbook;
    }

    private void ApplyColumnSelection(FormStructure structure, string[]? cols)
    {
        var hasExplicit = cols is { Length: > 0 };

        var showUploaded = false;
        var showFile = false;
        var showUploadId = false;

        var visibleDataIndexes = new HashSet<int>();

        if (!hasExplicit)
        {
            // Default: all data columns, no tech columns
            for (var i = 0; i < structure.Columns.Count; i++)
            {
                visibleDataIndexes.Add(i);
            }
        }
        else
        {
            foreach (var raw in cols!)
            {
                var token = (raw ?? string.Empty).Trim();
                if (token.Length == 0)
                {
                    continue;
                }

                if (string.Equals(token, "uploaded", StringComparison.OrdinalIgnoreCase))
                {
                    showUploaded = true;
                    continue;
                }

                if (string.Equals(token, "file", StringComparison.OrdinalIgnoreCase))
                {
                    showFile = true;
                    continue;
                }

                if (string.Equals(token, "uploadId", StringComparison.OrdinalIgnoreCase))
                {
                    showUploadId = true;
                    continue;
                }

                if (token.StartsWith("c", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(token[1..], out var colIndex)
                    && colIndex >= 1
                    && colIndex <= structure.Columns.Count)
                {
                    visibleDataIndexes.Add(colIndex - 1);
                }
            }
        }

        ShowUploadedColumn = showUploaded;
        ShowFileColumn = showFile;
        ShowUploadIdColumn = showUploadId;

        VisibleColumns = structure.Columns
            .Where((_, i) => visibleDataIndexes.Contains(i))
            .ToArray();

        if (!hasExplicit)
        {
            SelectedCols = Array.Empty<string>();
            return;
        }

        var normalized = new List<string>();
        if (showUploaded) normalized.Add("uploaded");
        if (showFile) normalized.Add("file");
        if (showUploadId) normalized.Add("uploadId");
        foreach (var i in visibleDataIndexes.OrderBy(i => i))
        {
            normalized.Add($"c{i + 1}");
        }

        SelectedCols = normalized;
    }

    private static IReadOnlyList<AggregatedRow> ApplySort(
        IReadOnlyList<AggregatedRow> rows,
        FormStructure structure,
        string sortKey,
        string sortDir)
    {
        var descending = string.Equals(sortDir, "desc", StringComparison.OrdinalIgnoreCase);

        if (string.Equals(sortKey, "uploaded", StringComparison.OrdinalIgnoreCase))
        {
            var ordered = descending
                ? rows.OrderByDescending(r => r.UploadedAtUtc)
                : rows.OrderBy(r => r.UploadedAtUtc);

            ordered = descending
                ? ordered.ThenByDescending(r => r.OriginalFileName, StringComparer.OrdinalIgnoreCase)
                : ordered.ThenBy(r => r.OriginalFileName, StringComparer.OrdinalIgnoreCase);

            ordered = descending
                ? ordered.ThenByDescending(r => r.RowNumber)
                : ordered.ThenBy(r => r.RowNumber);

            return ordered.ToArray();
        }

        if (string.Equals(sortKey, "row", StringComparison.OrdinalIgnoreCase))
        {
            var ordered = descending
                ? rows.OrderByDescending(r => r.RowNumber)
                : rows.OrderBy(r => r.RowNumber);

            ordered = descending
                ? ordered.ThenByDescending(r => r.UploadedAtUtc)
                : ordered.ThenBy(r => r.UploadedAtUtc);

            return ordered.ToArray();
        }

        if (string.Equals(sortKey, "file", StringComparison.OrdinalIgnoreCase))
        {
            var ordered = descending
                ? rows.OrderByDescending(r => r.OriginalFileName, StringComparer.OrdinalIgnoreCase)
                : rows.OrderBy(r => r.OriginalFileName, StringComparer.OrdinalIgnoreCase);

            ordered = descending
                ? ordered.ThenByDescending(r => r.UploadedAtUtc)
                : ordered.ThenBy(r => r.UploadedAtUtc);

            ordered = descending
                ? ordered.ThenByDescending(r => r.RowNumber)
                : ordered.ThenBy(r => r.RowNumber);

            return ordered.ToArray();
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
                ? ordered.ThenByDescending(r => r.UploadedAtUtc)
                : ordered.ThenBy(r => r.UploadedAtUtc);

            ordered = descending
                ? ordered.ThenByDescending(r => r.RowNumber)
                : ordered.ThenBy(r => r.RowNumber);

            return ordered.ToArray();
        }

        return rows;
    }

    private static string GetSortValue(AggregatedRow row, string colPath)
    {
        if (row.Values.TryGetValue(colPath, out var value) && !string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        return string.Empty;
    }

    public sealed record AggregatedRow(
        string UploadId,
        int FormVersion,
        string OriginalFileName,
        DateTime UploadedAtUtc,
        int RowNumber,
        IReadOnlyDictionary<string, string?> Values);
}
