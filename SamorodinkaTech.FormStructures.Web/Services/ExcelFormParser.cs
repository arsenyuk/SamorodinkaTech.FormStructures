using System.Reflection;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using SamorodinkaTech.FormStructures.Web.Models;

namespace SamorodinkaTech.FormStructures.Web.Services;

public sealed class ExcelFormParser
{
    private static readonly Regex FirstNumberRegex = new(@"\d+", RegexOptions.Compiled);

    public sealed record ExcelFormLayout(
        FormStructure Structure,
        int HeaderRowStart,
        int LastHeaderRow,
        int DataStartRow,
        int UsedLastRow,
        IReadOnlyList<int> LeafColumns);

    public FormStructure Parse(Stream xlsxStream, string? sourceFileName)
    {
        try
        {
            return ParseLayout(xlsxStream, sourceFileName).Structure;
        }
        catch (FormParseException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new FormParseException("Failed to parse Excel file.", ex);
        }
    }

    public ExcelFormLayout ParseLayout(Stream xlsxStream, string? sourceFileName)
    {
        try
        {
            using var workbook = new XLWorkbook(xlsxStream);
            var ws = workbook.Worksheets.FirstOrDefault();
            if (ws is null)
            {
                throw new FormParseException("Excel file contains no worksheets.");
            }

            // Important: Excel templates may contain formatting (number/date types, styles) far below the header
            // without any actual values. RangeUsed() can include those, inflating usedLastRow/Col.
            // We want the last row/col based on contents (values/formulas), not formatting.
            var lastRow = ws.LastRowUsed(XLCellsUsedOptions.Contents);
            var lastCol = ws.LastColumnUsed(XLCellsUsedOptions.Contents);
            if (lastRow is null || lastCol is null)
            {
                throw new FormParseException("Excel file is empty: no cells with content found.");
            }

            var usedLastRow = lastRow.RowNumber();
            var usedLastCol = lastCol.ColumnNumber();

            var formNumber = ExtractFirstNonEmptyCellValue(ws, rowNumber: 1, usedLastCol);
            if (string.IsNullOrWhiteSpace(formNumber))
            {
                throw new FormParseException("Form number is missing: row 1 has no non-empty cells.");
            }

            formNumber = NormalizeFormNumber(formNumber);

            var formTitle = ExtractFirstNonEmptyCellValue(ws, rowNumber: 2, usedLastCol);
            if (string.IsNullOrWhiteSpace(formTitle))
            {
                throw new FormParseException("Form title is missing: row 2 has no non-empty cells.");
            }

            const int headerRowStart = 3;
            var headerResult = ParseHeader(ws, headerRowStart, usedLastRow, usedLastCol);
            var columns = BuildColumns(headerResult.Nodes);
            var leafColumns = BuildLeafColumns(headerResult.Nodes);

            // Infer column types from Excel Data Validation (List).
            // If any leaf column has a list validation applied to it, treat it as a reference book column.
            var refBookLeafCols = DetectReferenceBookLeafColumns(ws, leafColumns);
            if (refBookLeafCols.Count > 0)
            {
                columns = columns
                    .Select((c, i) => i < leafColumns.Count && refBookLeafCols.Contains(leafColumns[i])
                        ? c with { Type = ColumnType.ReferenceBook }
                        : c)
                    .ToArray();
            }

            // Some forms include an extra header row with column indices (1..N).
            // We keep the header tree unchanged (so the index row is not shown as a header label),
            // but extend the header boundary so data rows start after the index row.
            var lastHeaderRow = headerResult.LastHeaderRow;
            if (TryReadSequentialColumnIndexRow(ws, lastHeaderRow + 1, leafColumns, out var columnNumbers))
            {
                lastHeaderRow += 1;
                columns = columns
                    .Select((c, i) => i < columnNumbers.Length ? c with { ColumnNumber = columnNumbers[i] } : c)
                    .ToArray();
            }

            if (columns.Count != leafColumns.Count)
            {
                throw new FormParseException($"Header leaf count mismatch: columns={columns.Count}, leafColumns={leafColumns.Count}.");
            }

            var signature = new
            {
                Header = headerResult.Nodes,
                Columns = columns.Select(c => new { c.Index, c.Name, c.Path })
            };

            var signatureJson = JsonUtil.ToStableJson(signature);
            var structureHash = Hashing.Sha256Hex(signatureJson);

            var structure = new FormStructure
            {
                FormNumber = formNumber,
                FormTitle = formTitle,
                Version = 0,
                UploadedAtUtc = DateTime.UtcNow,
                Header = headerResult.Nodes,
                Columns = columns,
                StructureHash = structureHash,
                SourceFileName = sourceFileName
            };

            var dataStartRow = lastHeaderRow + 1;
            return new ExcelFormLayout(
                Structure: structure,
                HeaderRowStart: headerRowStart,
                LastHeaderRow: lastHeaderRow,
                DataStartRow: dataStartRow,
                UsedLastRow: usedLastRow,
                LeafColumns: leafColumns);
        }
        catch (FormParseException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new FormParseException("Failed to parse Excel file.", ex);
        }
    }

    public IReadOnlyList<FormDataRow> ReadDataRows(Stream xlsxStream, ExcelFormLayout layout)
    {
        try
        {
            using var workbook = new XLWorkbook(xlsxStream);
            var ws = workbook.Worksheets.FirstOrDefault();
            if (ws is null)
            {
                throw new FormParseException("Excel file contains no worksheets.");
            }

            var rows = new List<FormDataRow>();
            for (var r = layout.DataStartRow; r <= layout.UsedLastRow; r++)
            {
                var hasAny = false;
                foreach (var col in layout.LeafColumns)
                {
                    var cell = ws.Cell(r, col);
                    if (!cell.IsEmpty(XLCellsUsedOptions.Contents))
                    {
                        hasAny = true;
                        break;
                    }
                }

                if (!hasAny)
                {
                    continue;
                }

                var values = new Dictionary<string, string?>(StringComparer.Ordinal);
                for (var i = 0; i < layout.LeafColumns.Count; i++)
                {
                    var col = layout.LeafColumns[i];
                    var key = layout.Structure.Columns[i].Path;
                    var cell = ws.Cell(r, col);
                    if (cell.IsEmpty(XLCellsUsedOptions.Contents))
                    {
                        values[key] = null;
                        continue;
                    }

                    var raw = cell.GetFormattedString()?.Trim();
                    values[key] = string.IsNullOrWhiteSpace(raw) ? null : raw;
                }

                rows.Add(new FormDataRow
                {
                    RowNumber = r,
                    Values = values
                });
            }

            return rows;
        }
        catch (FormParseException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new FormParseException("Failed to read data rows from Excel file.", ex);
        }
    }

    public IReadOnlyList<ReferenceBook> ExtractReferenceBooks(Stream xlsxStream)
    {
        try
        {
            using var workbook = new XLWorkbook(xlsxStream);

            var bySource = new Dictionary<string, MutableReferenceBook>(StringComparer.OrdinalIgnoreCase);

            foreach (var ws in workbook.Worksheets)
            {
                foreach (var dv in ws.DataValidations)
                {
                    if (dv.AllowedValues != XLAllowedValues.List)
                    {
                        continue;
                    }

                    var formula = (GetValidationFormula1(dv) ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(formula))
                    {
                        continue;
                    }

                    if (LooksLikeExplicitList(formula))
                    {
                        continue;
                    }

                    if (!TryResolveListSource(workbook, ws, formula, out var sourceSheet, out var sourceRange, out var values))
                    {
                        continue;
                    }

                    var sourceKey = $"{sourceSheet}!{sourceRange}";
                    var title = TryGetRangeHeaderTitle(workbook, sourceSheet, sourceRange) ?? sourceKey;
                    if (!bySource.TryGetValue(sourceKey, out var book))
                    {
                        book = new MutableReferenceBook
                        {
                            SourceFormula = formula,
                            SourceSheet = sourceSheet,
                            SourceRange = sourceRange,
                            Title = title,
                            Values = values,
                        };
                        bySource[sourceKey] = book;
                    }
                    else
                    {
                        // Prefer longer value list if multiple validations point to the same source.
                        if (values.Count > book.Values.Count)
                        {
                            book.Values = values;
                        }
                    }

                    foreach (var r in dv.Ranges)
                    {
                        book.AppliedTo.Add($"{ws.Name}!{r.RangeAddress}");
                    }
                }
            }

            return bySource.Values
                .Select(b => new ReferenceBook
                {
                    Id = Hashing.Sha256Hex($"{b.SourceSheet}!{b.SourceRange}"),
                    Title = b.Title,
                    SourceFormula = b.SourceFormula,
                    SourceSheet = b.SourceSheet,
                    SourceRange = b.SourceRange,
                    AppliedTo = b.AppliedTo
                        .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                        .ToArray(),
                    Values = b.Values
                })
                .OrderBy(b => b.Title, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
        catch (Exception ex)
        {
            throw new FormParseException("Failed to read reference books (data validation lists) from Excel file.", ex);
        }
    }

    private static bool LooksLikeExplicitList(string formula)
    {
        var f = formula.Trim();
        if (f.StartsWith("=", StringComparison.Ordinal))
        {
            f = f[1..].Trim();
        }

        // Explicit list is usually quoted or comma-separated constants, not a sheet/range reference.
        // We only want lists that point to a value interval.
        return f.Contains(',', StringComparison.Ordinal) && !f.Contains('!', StringComparison.Ordinal);
    }

    private static string? GetValidationFormula1(IXLDataValidation dv)
    {
        // ClosedXML doesn't expose formula properties on IXLDataValidation in all versions.
        // Use reflection on the concrete type.
        var t = dv.GetType();

        // Seen in various ClosedXML builds:
        // - Formula1
        // - DataValidationForumla1 (misspelled)
        // - DataValidationFormula1
        var prop =
            t.GetProperty("Formula1", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?? t.GetProperty("DataValidationForumla1", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?? t.GetProperty("DataValidationFormula1", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

        var value = prop?.GetValue(dv);
        if (value is null)
        {
            // In ClosedXML 0.105, list validations often store the list source in Value/MinValue
            // and don't populate Formula1 at all.
            value =
                t.GetProperty("Value", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)?.GetValue(dv)
                ?? t.GetProperty("MinValue", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)?.GetValue(dv)
                ?? t.GetProperty("MaxValue", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)?.GetValue(dv);

            if (value is null)
            {
                return null;
            }
        }

        if (value is string s)
        {
            return s;
        }

        return value.ToString();
    }

    private static HashSet<int> DetectReferenceBookLeafColumns(IXLWorksheet ws, IReadOnlyList<int> leafColumns)
    {
        var leafSet = new HashSet<int>(leafColumns);
        var result = new HashSet<int>();

        foreach (var dv in ws.DataValidations)
        {
            if (dv.AllowedValues != XLAllowedValues.List)
            {
                continue;
            }

            // If this is a list validation (explicit list or range-based), treat it as a reference book.
            foreach (var r in dv.Ranges)
            {
                var addr = r.RangeAddress;
                for (var c = addr.FirstAddress.ColumnNumber; c <= addr.LastAddress.ColumnNumber; c++)
                {
                    if (leafSet.Contains(c))
                    {
                        result.Add(c);
                    }
                }
            }
        }

        return result;
    }

    private static bool TryResolveListSource(
        XLWorkbook workbook,
        IXLWorksheet currentSheet,
        string formula,
        out string sourceSheet,
        out string sourceRange,
        out IReadOnlyList<string> values)
    {
        sourceSheet = string.Empty;
        sourceRange = string.Empty;
        values = Array.Empty<string>();

        var f = formula.Trim();
        if (f.StartsWith("=", StringComparison.Ordinal))
        {
            f = f[1..].Trim();
        }

        // Typical formats:
        //  - Sheet1!$A$1:$A$10
        //  - 'Sheet 1'!$A$1:$A$10
        //  - $A$1:$A$10 (same sheet)
        IXLWorksheet sourceWs;
        string rangeRef;

        if (TryParseSheetRange(f, out var sheetName, out rangeRef))
        {
            var ws = workbook.Worksheets
                .FirstOrDefault(w => string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (ws is null)
            {
                return false;
            }

            sourceWs = ws;
            sourceSheet = ws.Name ?? sheetName;
        }
        else
        {
            // If there is no explicit sheet, assume the current one.
            sourceWs = currentSheet;
            sourceSheet = currentSheet.Name ?? string.Empty;
            rangeRef = f;
        }

        try
        {
            var range = sourceWs.Range(rangeRef);
            sourceRange = range.RangeAddress.ToString() ?? rangeRef;
            values = ExtractRangeValues(range);
            return values.Count > 0;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryParseSheetRange(string f, out string sheetName, out string rangeRef)
    {
        sheetName = string.Empty;
        rangeRef = string.Empty;

        var bang = f.IndexOf('!');
        if (bang <= 0 || bang >= f.Length - 1)
        {
            return false;
        }

        sheetName = f[..bang].Trim();
        rangeRef = f[(bang + 1)..].Trim();

        if (sheetName.Length >= 2 && sheetName.StartsWith("'", StringComparison.Ordinal) && sheetName.EndsWith("'", StringComparison.Ordinal))
        {
            sheetName = sheetName[1..^1].Replace("''", "'");
        }

        return !string.IsNullOrWhiteSpace(sheetName) && !string.IsNullOrWhiteSpace(rangeRef);
    }

    private static IReadOnlyList<string> ExtractRangeValues(IXLRange range)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var list = new List<string>();

        foreach (var cell in range.Cells()
            .OrderBy(c => c.Address.RowNumber)
            .ThenBy(c => c.Address.ColumnNumber))
        {
            if (cell.IsEmpty(XLCellsUsedOptions.Contents))
            {
                continue;
            }

            var raw = cell.GetFormattedString()?.Trim();
            if (string.IsNullOrWhiteSpace(raw))
            {
                continue;
            }

            if (seen.Add(raw))
            {
                list.Add(raw);
            }
        }

        return list;
    }

    private static string? TryGetRangeHeaderTitle(XLWorkbook workbook, string sourceSheet, string sourceRange)
    {
        if (string.IsNullOrWhiteSpace(sourceSheet) || string.IsNullOrWhiteSpace(sourceRange))
        {
            return null;
        }

        var ws = workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name, sourceSheet, StringComparison.OrdinalIgnoreCase));
        if (ws is null)
        {
            return null;
        }

        try
        {
            var range = ws.Range(sourceRange);
            var addr = range.RangeAddress;

            // Most validation lists are 1-dimensional.
            // Prefer the adjacent "header" cell of that list (above for vertical, left for horizontal).
            if (addr.FirstAddress.ColumnNumber == addr.LastAddress.ColumnNumber && addr.FirstAddress.RowNumber > 1)
            {
                var headerCell = ws.Cell(addr.FirstAddress.RowNumber - 1, addr.FirstAddress.ColumnNumber);
                var header = headerCell.GetFormattedString()?.Trim();
                return string.IsNullOrWhiteSpace(header) ? null : header;
            }

            if (addr.FirstAddress.RowNumber == addr.LastAddress.RowNumber && addr.FirstAddress.ColumnNumber > 1)
            {
                var headerCell = ws.Cell(addr.FirstAddress.RowNumber, addr.FirstAddress.ColumnNumber - 1);
                var header = headerCell.GetFormattedString()?.Trim();
                return string.IsNullOrWhiteSpace(header) ? null : header;
            }

            return null;
        }
        catch
        {
            return null;
        }
    }

    private static string NormalizeFormNumber(string raw)
    {
        var trimmed = raw.Trim();
        var m = FirstNumberRegex.Match(trimmed);
        return m.Success ? m.Value : trimmed;
    }

    private static string? ExtractFirstNonEmptyCellValue(IXLWorksheet ws, int rowNumber, int lastCol)
    {
        for (var col = 1; col <= lastCol; col++)
        {
            var value = ws.Cell(rowNumber, col).GetString()?.Trim();
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    private static bool TryReadSequentialColumnIndexRow(
        IXLWorksheet ws,
        int rowNumber,
        IReadOnlyList<int> leafColumns,
        out string[] columnNumbers)
    {
        columnNumbers = Array.Empty<string>();

        if (rowNumber <= 0 || leafColumns.Count == 0)
        {
            return false;
        }

        var numbers = new string[leafColumns.Count];

        for (var i = 0; i < leafColumns.Count; i++)
        {
            var c = leafColumns[i];
            var raw = ws.Cell(rowNumber, c).GetString()?.Trim();
            if (string.IsNullOrWhiteSpace(raw))
            {
                return false;
            }

            // Allow formats like "1" or "1." by extracting the first number.
            var m = FirstNumberRegex.Match(raw);
            if (!m.Success || !int.TryParse(m.Value, out var n))
            {
                return false;
            }

            var expected = i + 1;
            if (n != expected)
            {
                return false;
            }

            numbers[i] = m.Value;
        }

        columnNumbers = numbers;
        return true;
    }

    private sealed record HeaderParseResult(IReadOnlyList<HeaderNode> Nodes, int LastHeaderRow, int LastHeaderCol);
    private static void EnsureNoMergedCellsInBottomHeaderRow(IXLRangeAddress[] merged, int bottomHeaderRow, int lastHeaderCol)
    {
        // We only support templates where each column has a clearly named header cell.
        // Therefore the bottom header row must not contain any merged cells.
        var offender = merged.FirstOrDefault(a =>
            a.FirstAddress.RowNumber <= bottomHeaderRow
            && a.LastAddress.RowNumber >= bottomHeaderRow
            && a.FirstAddress.ColumnNumber <= lastHeaderCol
            && a.LastAddress.ColumnNumber >= 1);

        if (offender is not null)
        {
            throw new FormParseException(
                $"Bottom header row (row {bottomHeaderRow}) contains a merged range (r{offender.FirstAddress.RowNumber}-{offender.LastAddress.RowNumber}, c{offender.FirstAddress.ColumnNumber}-{offender.LastAddress.ColumnNumber}). "
                + "Unmerge the bottom header row so each column has a unique header cell.");
        }
    }

    private static HeaderParseResult ParseHeader(IXLWorksheet ws, int headerRowStart, int usedLastRow, int usedLastCol)
    {
        if (usedLastRow < headerRowStart)
        {
            throw new FormParseException("Header is missing: no rows after title.");
        }

        var merged = ws.MergedRanges
            .Select(r => r.RangeAddress)
            .ToArray();

        // Important: data rows may contain values too, so we can't detect header end by "last row with any value".
        // Instead, find the smallest header end row that yields a valid header tree.

        var lastMergedBottom = merged
            .Where(a => a.FirstAddress.RowNumber >= headerRowStart)
            .Select(a => a.LastAddress.RowNumber)
            .DefaultIfEmpty(0)
            .Max();

        var maxProbeRow = Math.Min(usedLastRow, headerRowStart + 50);
        var probeStartRow = Math.Max(headerRowStart, lastMergedBottom == 0 ? headerRowStart : lastMergedBottom);

        FormParseException? lastError = null;

        for (var lastHeaderRow = probeStartRow; lastHeaderRow <= maxProbeRow; lastHeaderRow++)
        {
            var lastHeaderCol = GetLastHeaderCol(ws, headerRowStart, lastHeaderRow, usedLastCol, merged);
            if (lastHeaderCol < 1)
            {
                continue;
            }

            var regions = new List<HeaderRegion>();
            try
            {
                EnsureNoMergedCellsInBottomHeaderRow(merged, lastHeaderRow, lastHeaderCol);

                for (var r = headerRowStart; r <= lastHeaderRow; r++)
                {
                    for (var c = 1; c <= lastHeaderCol; c++)
                    {
                        var cell = ws.Cell(r, c);
                        var regionAddress = FindMergedAddress(merged, r, c);
                        if (regionAddress is not null)
                        {
                            if (regionAddress.FirstAddress.RowNumber != r || regionAddress.FirstAddress.ColumnNumber != c)
                            {
                                continue; // not top-left
                            }

                            // If a merged header extends below the current probe boundary, this boundary is invalid.
                            if (regionAddress.LastAddress.RowNumber > lastHeaderRow)
                            {
                                throw new FormParseException("Header probe ended inside a merged range.");
                            }

                            var label = cell.GetString()?.Trim();
                            if (string.IsNullOrWhiteSpace(label))
                            {
                                throw new FormParseException($"Header cell is empty at row {r}, col {c} (merged range).");
                            }

                            regions.Add(new HeaderRegion(
                                RowStart: regionAddress.FirstAddress.RowNumber,
                                RowEnd: regionAddress.LastAddress.RowNumber,
                                ColStart: regionAddress.FirstAddress.ColumnNumber,
                                ColEnd: regionAddress.LastAddress.ColumnNumber,
                                Label: label));
                        }
                        else
                        {
                            var label = cell.GetString()?.Trim();
                            if (string.IsNullOrWhiteSpace(label))
                            {
                                continue;
                            }

                            regions.Add(new HeaderRegion(
                                RowStart: r,
                                RowEnd: r,
                                ColStart: c,
                                ColEnd: c,
                                Label: label));
                        }
                    }
                }

                if (regions.Count == 0)
                {
                    continue;
                }

                // Validate that there are no gaps in the bottom-level columns.
                var leafCoverage = new bool[lastHeaderCol + 1];
                foreach (var region in regions)
                {
                    if (region.RowEnd == lastHeaderRow)
                    {
                        for (var c = region.ColStart; c <= region.ColEnd; c++)
                        {
                            leafCoverage[c] = true;
                        }
                    }
                }

                // If bottom row has no labels (structured header where leaves are above), accept.
                // But if it has some, ensure no gaps between 1..lastHeaderCol.
                if (leafCoverage.Skip(1).Any(v => v))
                {
                    for (var c = 1; c <= lastHeaderCol; c++)
                    {
                        if (!leafCoverage[c])
                        {
                            throw new FormParseException($"Header has a gap in bottom row at col {c}.");
                        }
                    }
                }

                // Build parent/child relationships.
                var nodes = regions
                    .Distinct()
                    .OrderBy(r => r.RowStart)
                    .ThenBy(r => r.ColStart)
                    .Select(r => new MutableNode(r))
                    .ToList();

                foreach (var child in nodes)
                {
                    MutableNode? parent = null;
                    foreach (var candidate in nodes)
                    {
                        if (candidate.Region.RowStart >= child.Region.RowStart)
                        {
                            continue;
                        }

                        if (candidate.Region.ColStart <= child.Region.ColStart
                            && candidate.Region.ColEnd >= child.Region.ColEnd)
                        {
                            if (parent is null || candidate.Region.RowStart > parent.Region.RowStart)
                            {
                                parent = candidate;
                            }
                        }
                    }

                    if (parent is not null)
                    {
                        parent.Children.Add(child);
                        child.Parent = parent;
                    }
                }

                var roots = nodes
                    .Where(n => n.Parent is null)
                    .OrderBy(n => n.Region.ColStart)
                    .ThenBy(n => n.Region.RowStart)
                    .Select(n => n.ToHeaderNode())
                    .ToArray();

                // Ensure this boundary yields a valid leaf set.
                _ = BuildLeafColumns(roots);

                return new HeaderParseResult(roots, lastHeaderRow, lastHeaderCol);
            }
            catch (FormParseException ex)
            {
                lastError = ex;
                // Try a larger boundary.
                continue;
            }
        }

        // If we saw a more specific validation error while probing, surface it instead of a generic
        // "no valid boundary" message. This helps users fix malformed templates.
        if (lastError is not null)
        {
            throw lastError;
        }

        throw new FormParseException("Header is missing: no valid header boundary found starting from row 3.");
    }

    private static IReadOnlyList<int> BuildLeafColumns(IReadOnlyList<HeaderNode> roots)
    {
        var cols = new List<int>();

        void Walk(HeaderNode node)
        {
            if (node.Children.Count == 0)
            {
                if (node.ColStart != node.ColEnd)
                {
                    throw new FormParseException($"Leaf header '{node.Label}' spans multiple columns (c{node.ColStart}-c{node.ColEnd}).");
                }

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

        var distinct = cols.Distinct().ToArray();
        if (distinct.Length != cols.Count)
        {
            throw new FormParseException("Header leaf columns are not unique.");
        }

        return cols;
    }

    private static int GetLastHeaderRow(IXLWorksheet ws, int headerRowStart, int usedLastRow, int usedLastCol, IXLRangeAddress[] merged)
    {
        var lastRowWithValue = 0;
        for (var r = headerRowStart; r <= usedLastRow; r++)
        {
            if (RowHasAnyValue(ws, r, usedLastCol))
            {
                lastRowWithValue = r;
            }
        }

        var lastMergedBottom = merged
            .Where(a => a.FirstAddress.RowNumber >= headerRowStart)
            .Select(a => a.LastAddress.RowNumber)
            .DefaultIfEmpty(0)
            .Max();

        return Math.Max(lastRowWithValue, lastMergedBottom);
    }

    private static int GetLastHeaderCol(IXLWorksheet ws, int headerRowStart, int headerRowEnd, int usedLastCol, IXLRangeAddress[] merged)
    {
        var lastColWithValue = 0;
        for (var r = headerRowStart; r <= headerRowEnd; r++)
        {
            for (var c = 1; c <= usedLastCol; c++)
            {
                if (!string.IsNullOrWhiteSpace(ws.Cell(r, c).GetString()))
                {
                    lastColWithValue = Math.Max(lastColWithValue, c);
                }
            }
        }

        var lastMergedRight = merged
            .Where(a => a.FirstAddress.RowNumber >= headerRowStart && a.FirstAddress.RowNumber <= headerRowEnd)
            .Select(a => a.LastAddress.ColumnNumber)
            .DefaultIfEmpty(0)
            .Max();

        return Math.Max(lastColWithValue, lastMergedRight);
    }

    private static bool RowHasAnyValue(IXLWorksheet ws, int rowNumber, int usedLastCol)
    {
        for (var c = 1; c <= usedLastCol; c++)
        {
            if (!string.IsNullOrWhiteSpace(ws.Cell(rowNumber, c).GetString()))
            {
                return true;
            }
        }

        return false;
    }

    private static IXLRangeAddress? FindMergedAddress(IXLRangeAddress[] merged, int row, int col)
    {
        foreach (var addr in merged)
        {
            if (row >= addr.FirstAddress.RowNumber
                && row <= addr.LastAddress.RowNumber
                && col >= addr.FirstAddress.ColumnNumber
                && col <= addr.LastAddress.ColumnNumber)
            {
                return addr;
            }
        }

        return null;
    }

    private static IReadOnlyList<ColumnDefinition> BuildColumns(IReadOnlyList<HeaderNode> roots)
    {
        var leaves = new List<(string name, string path)>();

        void Walk(HeaderNode node, List<string> stack)
        {
            stack.Add(node.Label);

            if (node.Children.Count == 0)
            {
                var path = string.Join(" / ", stack);
                leaves.Add((node.Label, path));
            }
            else
            {
                foreach (var child in node.Children.OrderBy(c => c.ColStart).ThenBy(c => c.RowStart))
                {
                    Walk(child, stack);
                }
            }

            stack.RemoveAt(stack.Count - 1);
        }

        foreach (var root in roots.OrderBy(r => r.ColStart).ThenBy(r => r.RowStart))
        {
            Walk(root, new List<string>());
        }

        return leaves
            .Select((leaf, i) => new ColumnDefinition
            {
                Index = i + 1,
                Name = leaf.name,
                Path = leaf.path
            })
            .ToArray();
    }

    private sealed record HeaderRegion(int RowStart, int RowEnd, int ColStart, int ColEnd, string Label);

    private sealed class MutableNode
    {
        public HeaderRegion Region { get; }
        public MutableNode? Parent { get; set; }
        public List<MutableNode> Children { get; } = new();

        public MutableNode(HeaderRegion region)
        {
            Region = region;
        }

        public HeaderNode ToHeaderNode()
        {
            return new HeaderNode
            {
                Label = Region.Label,
                RowStart = Region.RowStart,
                RowEnd = Region.RowEnd,
                ColStart = Region.ColStart,
                ColEnd = Region.ColEnd,
                Children = Children
                    .OrderBy(c => c.Region.ColStart)
                    .ThenBy(c => c.Region.RowStart)
                    .Select(c => c.ToHeaderNode())
                    .ToArray()
            };
        }
    }

    private sealed class MutableReferenceBook
    {
        public string Title { get; set; } = string.Empty;
        public string SourceFormula { get; set; } = string.Empty;
        public string SourceSheet { get; set; } = string.Empty;
        public string SourceRange { get; set; } = string.Empty;
        public HashSet<string> AppliedTo { get; } = new(StringComparer.OrdinalIgnoreCase);
        public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();
    }
}
