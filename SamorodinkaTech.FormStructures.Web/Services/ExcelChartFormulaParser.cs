using System.Globalization;
using System.Text.RegularExpressions;

namespace SamorodinkaTech.FormStructures.Web.Services;

public enum ChartAggregationKind
{
    Sum = 0,
    Average = 1,
    Difference = 2,
}

public sealed record ChartFormulaInfo(ChartAggregationKind Kind, IReadOnlyList<int> Columns, int? TitleColumn = null);

public static class ExcelChartFormulaParser
{
    private static readonly Regex SumFnRegex = new(
        @"\b(SUM|СУММ)\s*\(",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex AverageFnRegex = new(
        @"\b(AVERAGE|AVERAGEA|СРЗНАЧ|СРЗНАЧА)\s*\(",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex CountFnRegex = new(
        @"\b(COUNT|COUNTA)\s*\(",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    // Excel A1 references can be absolute ($A$1) and ranges (A1:C1).
    // We only care about columns to map them to the rendered table.
    private static readonly Regex RangeRefRegex = new(
        @"(?<![A-Z0-9_])(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?\$?([A-Z]{1,3})\$?\d+\s*:\s*(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?\$?([A-Z]{1,3})\$?\d+",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex CellRefRegex = new(
        @"(?<![A-Z0-9_])(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?\$?([A-Z]{1,3})\$?\d+",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex ConstantDivisorRegex = new(
        @"/\s*(\d+)(?![A-Z0-9_])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex A1CellRefWithRowColRegex = new(
        @"(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?\$?(?<col>[A-Z]{1,3})\$?(?<row>\d+)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex A1DifferenceRegex = new(
        @"^\s*=\s*(?<y>(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?\$?[A-Z]{1,3}\$?\d+)\s*-\s*(?<z>(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?\$?[A-Z]{1,3}\$?\d+)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex R1C1DifferenceRegex = new(
        @"^\s*=\s*(?<y>(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?(?:R(?:\[?-?\d*\]?))?(?:C(?:\[?-?\d+\]?)))\s*-\s*(?<z>(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?(?:R(?:\[?-?\d*\]?))?(?:C(?:\[?-?\d+\]?)))\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    // R1C1 refs: R4C2, R[0]C[-1], RC[-1], R[-1]C, etc.
    // We only support refs that resolve to the SAME ROW as the origin (so parts are in the same table row).
    private static readonly Regex R1C1RangeRegex = new(
        @"(?<![A-Z0-9_])(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?(R(?:\[?-?\d*\]?))?(C(?:\[?-?\d+\]?))\s*:\s*(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?(R(?:\[?-?\d*\]?))?(C(?:\[?-?\d+\]?))",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex R1C1CellRegex = new(
        @"(?<![A-Z0-9_])(?:(?:'[^']+'|[A-Z0-9_\.]+)!)?(R(?:\[?-?\d*\]?))?(C(?:\[?-?\d+\]?))",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    public static bool TryParse(string? formula, out ChartFormulaInfo info)
    {
        info = new ChartFormulaInfo(ChartAggregationKind.Sum, Array.Empty<int>());

        if (string.IsNullOrWhiteSpace(formula))
        {
            return false;
        }

        var f = formula.Trim();
        if (!f.StartsWith("=", StringComparison.Ordinal))
        {
            return false;
        }

        var columns = ExtractReferencedColumns(f);
        if (columns.Count < 2)
        {
            return false;
        }

        var kind = DetectKind(f, columns.Count);
        if (kind is null)
        {
            return false;
        }

        info = new ChartFormulaInfo(kind.Value, columns);
        return true;
    }

    public static bool TryParse(string? formula, int originColumn, int originRow, out ChartFormulaInfo info)
    {
        // Prefer A1 parsing (no context required). If that fails, try R1C1 using the origin cell.
        if (TryParse(formula, out info))
        {
            return true;
        }

        info = new ChartFormulaInfo(ChartAggregationKind.Sum, Array.Empty<int>());

        if (string.IsNullOrWhiteSpace(formula))
        {
            return false;
        }

        if (originColumn <= 0 || originRow <= 0)
        {
            return false;
        }

        var f = formula.Trim();
        if (!f.StartsWith("=", StringComparison.Ordinal))
        {
            return false;
        }

        if (TryParseDifferenceA1(f, originColumn, originRow, out info))
        {
            return true;
        }

        if (TryParseDifferenceR1C1(f, originColumn, originRow, out info))
        {
            return true;
        }

        var columns = ExtractReferencedColumnsR1C1(f, originColumn, originRow);
        if (columns.Count < 2)
        {
            return false;
        }

        var kind = DetectKind(f, columns.Count);
        if (kind is null)
        {
            return false;
        }

        info = new ChartFormulaInfo(kind.Value, columns);
        return true;
    }

    private static bool TryParseDifferenceA1(string formula, int originColumn, int originRow, out ChartFormulaInfo info)
    {
        info = new ChartFormulaInfo(ChartAggregationKind.Difference, Array.Empty<int>());

        var m = A1DifferenceRegex.Match(formula);
        if (!m.Success)
        {
            return false;
        }

        if (!TryParseA1CellRef(m.Groups["y"].Value, out var yCol, out var yRow))
        {
            return false;
        }
        if (!TryParseA1CellRef(m.Groups["z"].Value, out var zCol, out var zRow))
        {
            return false;
        }

        if (yRow != originRow || zRow != originRow)
        {
            return false;
        }

        if (originColumn == zCol)
        {
            return false;
        }

        info = new ChartFormulaInfo(
            ChartAggregationKind.Difference,
            new[] { originColumn, zCol },
            TitleColumn: yCol);
        return true;
    }

    private static bool TryParseDifferenceR1C1(string formula, int originColumn, int originRow, out ChartFormulaInfo info)
    {
        info = new ChartFormulaInfo(ChartAggregationKind.Difference, Array.Empty<int>());

        var m = R1C1DifferenceRegex.Match(formula);
        if (!m.Success)
        {
            return false;
        }

        var y = m.Groups["y"].Value;
        var z = m.Groups["z"].Value;

        if (!TryParseR1C1CellRef(y, originColumn, originRow, out var yResolved))
        {
            return false;
        }
        if (!TryParseR1C1CellRef(z, originColumn, originRow, out var zResolved))
        {
            return false;
        }

        if (yResolved.Row != originRow || zResolved.Row != originRow)
        {
            return false;
        }

        if (originColumn == zResolved.Col)
        {
            return false;
        }

        info = new ChartFormulaInfo(
            ChartAggregationKind.Difference,
            new[] { originColumn, zResolved.Col },
            TitleColumn: yResolved.Col);
        return true;
    }

    private static bool TryParseA1CellRef(string token, out int col, out int row)
    {
        col = 0;
        row = 0;

        var m = A1CellRefWithRowColRegex.Match(token ?? string.Empty);
        if (!m.Success)
        {
            return false;
        }

        var colLetters = m.Groups["col"].Value;
        var rowText = m.Groups["row"].Value;
        if (!int.TryParse(rowText, NumberStyles.Integer, CultureInfo.InvariantCulture, out row) || row <= 0)
        {
            return false;
        }

        var c = ColumnLettersToNumber(colLetters);
        if (c is null || c.Value <= 0)
        {
            return false;
        }

        col = c.Value;
        return true;
    }

    private static bool TryParseR1C1CellRef(string token, int originColumn, int originRow, out (int Row, int Col) resolved)
    {
        resolved = default;

        var m = R1C1CellRegex.Match(token ?? string.Empty);
        if (!m.Success)
        {
            return false;
        }

        var r = m.Groups[1].Value;
        var c = m.Groups[2].Value;
        var rc = TryResolveR1C1(r, c, originColumn, originRow);
        if (rc is null)
        {
            return false;
        }

        resolved = rc.Value;
        return true;
    }

    private static ChartAggregationKind? DetectKind(string formula, int distinctColumnCount)
    {
        // Direct average functions.
        if (AverageFnRegex.IsMatch(formula))
        {
            return ChartAggregationKind.Average;
        }

        // Common average patterns:
        //   SUM(...)/COUNT(...)
        //   SUM(...)/COUNTA(...)
        //   (A+B+...)/N
        //   SUM(...)/N
        var hasDivision = formula.Contains('/', StringComparison.Ordinal);
        if (hasDivision)
        {
            var hasSum = SumFnRegex.IsMatch(formula);
            var hasCount = CountFnRegex.IsMatch(formula);

            if (hasSum && hasCount)
            {
                return ChartAggregationKind.Average;
            }

            var m = ConstantDivisorRegex.Match(formula);
            if (m.Success
                && int.TryParse(m.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var n)
                && n >= 2
                && n == distinctColumnCount)
            {
                return ChartAggregationKind.Average;
            }
        }

        // Summation formulas (used for pie decomposition).
        if (SumFnRegex.IsMatch(formula)
            || formula.Contains('+', StringComparison.Ordinal))
        {
            return ChartAggregationKind.Sum;
        }

        return null;
    }

    private static IReadOnlyList<int> ExtractReferencedColumns(string formula)
    {
        var cols = new SortedSet<int>();

        // Expand ranges first (A1:C1 => A,B,C).
        foreach (Match m in RangeRefRegex.Matches(formula))
        {
            var a = m.Groups[1].Value;
            var b = m.Groups[2].Value;
            if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b))
            {
                continue;
            }

            var start = ColumnLettersToNumber(a);
            var end = ColumnLettersToNumber(b);
            if (start is null || end is null)
            {
                continue;
            }

            var lo = Math.Min(start.Value, end.Value);
            var hi = Math.Max(start.Value, end.Value);
            for (var c = lo; c <= hi; c++)
            {
                cols.Add(c);
            }
        }

        // Add single-cell references.
        foreach (Match m in CellRefRegex.Matches(formula))
        {
            var letters = m.Groups[1].Value;
            if (string.IsNullOrWhiteSpace(letters))
            {
                continue;
            }

            var n = ColumnLettersToNumber(letters);
            if (n is null)
            {
                continue;
            }

            cols.Add(n.Value);
        }

        return cols.ToArray();
    }

    private static IReadOnlyList<int> ExtractReferencedColumnsR1C1(string formula, int originColumn, int originRow)
    {
        var cols = new SortedSet<int>();

        foreach (Match m in R1C1RangeRegex.Matches(formula))
        {
            var r1 = m.Groups[1].Value;
            var c1 = m.Groups[2].Value;
            var r2 = m.Groups[3].Value;
            var c2 = m.Groups[4].Value;

            var start = TryResolveR1C1(r1, c1, originColumn, originRow);
            var end = TryResolveR1C1(r2, c2, originColumn, originRow);
            if (start is null || end is null)
            {
                continue;
            }

            // Only same-row ranges are meaningful for pie decomposition.
            if (start.Value.Row != originRow || end.Value.Row != originRow)
            {
                continue;
            }

            var lo = Math.Min(start.Value.Col, end.Value.Col);
            var hi = Math.Max(start.Value.Col, end.Value.Col);
            for (var c = lo; c <= hi; c++)
            {
                cols.Add(c);
            }
        }

        foreach (Match m in R1C1CellRegex.Matches(formula))
        {
            var r = m.Groups[1].Value;
            var c = m.Groups[2].Value;
            var resolved = TryResolveR1C1(r, c, originColumn, originRow);
            if (resolved is null)
            {
                continue;
            }

            if (resolved.Value.Row != originRow)
            {
                continue;
            }

            cols.Add(resolved.Value.Col);
        }

        return cols.ToArray();
    }

    private static (int Row, int Col)? TryResolveR1C1(string? rToken, string? cToken, int originColumn, int originRow)
    {
        // cToken is required in our regex.
        if (string.IsNullOrWhiteSpace(cToken))
        {
            return null;
        }

        var row = originRow;
        var col = originColumn;

        if (!string.IsNullOrWhiteSpace(rToken))
        {
            var r = rToken.Trim();
            if (!r.StartsWith('R') && !r.StartsWith('r'))
            {
                return null;
            }

            var payload = r[1..];
            if (payload.Length == 0)
            {
                // "R" means current row.
            }
            else if (payload.StartsWith('[') && payload.EndsWith(']'))
            {
                var inner = payload[1..^1];
                if (inner.Length == 0)
                {
                    // "R[]" is treated as current row.
                }
                else if (int.TryParse(inner, NumberStyles.Integer, CultureInfo.InvariantCulture, out var delta))
                {
                    row = originRow + delta;
                }
                else
                {
                    return null;
                }
            }
            else if (int.TryParse(payload, NumberStyles.Integer, CultureInfo.InvariantCulture, out var absRow))
            {
                row = absRow;
            }
            else
            {
                return null;
            }
        }

        var cText = cToken.Trim();
        if (!cText.StartsWith('C') && !cText.StartsWith('c'))
        {
            return null;
        }

        var cPayload = cText[1..];
        if (cPayload.Length == 0)
        {
            // "C" means current column.
        }
        else if (cPayload.StartsWith('[') && cPayload.EndsWith(']'))
        {
            var inner = cPayload[1..^1];
            if (inner.Length == 0)
            {
                // "C[]" treated as current.
            }
            else if (int.TryParse(inner, NumberStyles.Integer, CultureInfo.InvariantCulture, out var delta))
            {
                col = originColumn + delta;
            }
            else
            {
                return null;
            }
        }
        else if (int.TryParse(cPayload, NumberStyles.Integer, CultureInfo.InvariantCulture, out var absCol))
        {
            col = absCol;
        }
        else
        {
            return null;
        }

        return col > 0 && row > 0 ? (row, col) : null;
    }

    private static int? ColumnLettersToNumber(string letters)
    {
        var s = letters.Trim().ToUpperInvariant();
        if (s.Length == 0 || s.Length > 3)
        {
            return null;
        }

        var n = 0;
        foreach (var ch in s)
        {
            if (ch < 'A' || ch > 'Z')
            {
                return null;
            }

            n = (n * 26) + (ch - 'A' + 1);
        }

        return n;
    }
}
