using System;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using SamorodinkaTech.FormStructures.Web.Services;
using Xunit;

namespace SamorodinkaTech.FormStructures.Tests;

public sealed class ExcelChartFormulaParserTests
{
    public static TheoryData<string, ChartAggregationKind, int[]> ChartableFormulas
        => new()
        {
            // SUM variants
            { "=A4+B4+C4", ChartAggregationKind.Sum, new[] { 1, 2, 3 } },
            { "=SUM(A4:C4)", ChartAggregationKind.Sum, new[] { 1, 2, 3 } },
            { "=SUM($A$4:$C$4)", ChartAggregationKind.Sum, new[] { 1, 2, 3 } },
            { "=SUM ( A4 : C4 )", ChartAggregationKind.Sum, new[] { 1, 2, 3 } },
            { "=СУММ(A4:C4)", ChartAggregationKind.Sum, new[] { 1, 2, 3 } },
            { "=ROUND(SUM(A4:C4),2)", ChartAggregationKind.Sum, new[] { 1, 2, 3 } },
            { "=SUM(A4,C4,E4)", ChartAggregationKind.Sum, new[] { 1, 3, 5 } },
            { "=Sheet1!A4+Sheet1!B4", ChartAggregationKind.Sum, new[] { 1, 2 } },
            { "=SUM(Sheet1!A4:Sheet1!C4)", ChartAggregationKind.Sum, new[] { 1, 2, 3 } },

            // AVERAGE variants
            { "=AVERAGE(A4:C4)", ChartAggregationKind.Average, new[] { 1, 2, 3 } },
            { "=AVERAGE($A$4:$C$4)", ChartAggregationKind.Average, new[] { 1, 2, 3 } },
            { "=SUM(A4:C4)/COUNT(A4:C4)", ChartAggregationKind.Average, new[] { 1, 2, 3 } },
            { "=SUM(A4:C4)/COUNTA(A4:C4)", ChartAggregationKind.Average, new[] { 1, 2, 3 } },
            { "=(A4+B4+C4)/3", ChartAggregationKind.Average, new[] { 1, 2, 3 } },
            { "=SUM(A4:C4)/3", ChartAggregationKind.Average, new[] { 1, 2, 3 } },
        };

    public static TheoryData<string, int, int, ChartAggregationKind, int[]> ChartableFormulasR1C1
        => new()
        {
            // origin is D4 (col 4, row 4). Refer to A4:C4 via relative RC[-3]:RC[-1].
            { "=SUM(RC[-3]:RC[-1])", 4, 4, ChartAggregationKind.Sum, new[] { 1, 2, 3 } },
            { "=AVERAGE(RC[-3]:RC[-1])", 4, 4, ChartAggregationKind.Average, new[] { 1, 2, 3 } },

            // Non-range: SUM of explicit cells.
            { "=RC[-3]+RC[-2]+RC[-1]", 4, 4, ChartAggregationKind.Sum, new[] { 1, 2, 3 } },

            // Same-row, reversed range.
            { "=SUM(RC[-1]:RC[-3])", 4, 4, ChartAggregationKind.Sum, new[] { 1, 2, 3 } },
        };

    public static TheoryData<string, int, int, int, int[], string> DifferenceFormulas
        => new()
        {
            // X (origin) = Y - Z; title comes from Y, pie parts are X and Z.
            { "=A4-B4", 3, 4, 1, new[] { 3, 2 }, "A1" },
            { "=RC[-2]-RC[-1]", 3, 4, 1, new[] { 3, 2 }, "R1C1" },
        };

    [Theory]
    [MemberData(nameof(ChartableFormulas))]
    public void TryParse_RecognizesChartableFormulas_AndExtractsColumns(
        string formula,
        ChartAggregationKind expectedKind,
        int[] expectedColumns)
    {
        Assert.True(ExcelChartFormulaParser.TryParse(formula, out var info));
        Assert.Equal(expectedKind, info.Kind);
        Assert.Equal(expectedColumns, info.Columns);
    }

    [Theory]
    [MemberData(nameof(ChartableFormulasR1C1))]
    public void TryParse_R1C1_RecognizesChartableFormulas_AndExtractsColumns(
        string formula,
        int originCol,
        int originRow,
        ChartAggregationKind expectedKind,
        int[] expectedColumns)
    {
        Assert.True(ExcelChartFormulaParser.TryParse(formula, originCol, originRow, out var info));
        Assert.Equal(expectedKind, info.Kind);
        Assert.Equal(expectedColumns, info.Columns);
    }

    [Theory]
    [MemberData(nameof(DifferenceFormulas))]
    public void TryParse_Difference_XEqualsYMinusZ_UsesYAsTitle_AndXAndZAsParts(
        string formula,
        int originCol,
        int originRow,
        int expectedTitleCol,
        int[] expectedPartCols,
        string caseName)
    {
        Assert.True(ExcelChartFormulaParser.TryParse(formula, originCol, originRow, out var info), caseName);
        Assert.Equal(ChartAggregationKind.Difference, info.Kind);
        Assert.Equal(expectedTitleCol, info.TitleColumn);
        Assert.Equal(expectedPartCols, info.Columns);
    }

    [Theory]
    [InlineData("AVG-001-AVERAGE", ChartAggregationKind.Average)]
    [InlineData("AVG-002-SUM-COUNT", ChartAggregationKind.Average)]
    [InlineData("AVG-003-SUM-DIV", ChartAggregationKind.Average)]
    [InlineData("AVG-004-PLUS-DIV", ChartAggregationKind.Average)]
    [InlineData("AVG-005-RU", ChartAggregationKind.Average)]
    public void FixtureExample_WithDataAndAverageFormula_IsRecognized(
        string fixtureBaseName,
        ChartAggregationKind expectedKind)
    {
        using var stream = LoadXlsxFromBase64Fixture(fixtureBaseName);
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();

        Assert.Equal(10, ws.Cell(4, 1).GetValue<int>());
        Assert.Equal(20, ws.Cell(4, 2).GetValue<int>());
        Assert.Equal(30, ws.Cell(4, 3).GetValue<int>());

        var cell = ws.Cell(4, 4);
        Assert.True(cell.HasFormula);

        var raw = (cell.FormulaA1 ?? string.Empty).Trim();
        var normalized = raw.StartsWith('=') ? raw : "=" + raw;

        Assert.True(ExcelChartFormulaParser.TryParse(normalized, out var info));
        Assert.Equal(expectedKind, info.Kind);
        Assert.Equal(new[] { 1, 2, 3 }, info.Columns);
    }

    [Fact]
    public void FixtureExample_WithDataAndDifferenceFormula_IsRecognized_WithTitleAndParts()
    {
        using var stream = LoadXlsxFromBase64Fixture("DIFF-001-Y-MINUS-Z");
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();

        // Headers are row 3; data is row 4.
        Assert.Equal(100, ws.Cell(4, 1).GetValue<int>()); // Y
        Assert.Equal(30, ws.Cell(4, 2).GetValue<int>());  // Z

        var cellX = ws.Cell(4, 3);
        Assert.True(cellX.HasFormula);

        var raw = (cellX.FormulaA1 ?? string.Empty).Trim();
        var normalized = raw.StartsWith('=') ? raw : "=" + raw;

        Assert.True(ExcelChartFormulaParser.TryParse(normalized, originColumn: 3, originRow: 4, out var info));
        Assert.Equal(ChartAggregationKind.Difference, info.Kind);
        Assert.Equal(1, info.TitleColumn);
        Assert.Equal(new[] { 3, 2 }, info.Columns);
    }

    [Theory]
    [InlineData("=A4")]
    [InlineData("=AVERAGE(A4)")]
    [InlineData("=IF(A4>0,1,0)")]
    [InlineData("not a formula")]
    public void TryParse_ReturnsFalse_WhenUnsupportedOrTooFewParts(string formula)
    {
        Assert.False(ExcelChartFormulaParser.TryParse(formula, out _));
    }

    private static MemoryStream LoadXlsxFromBase64Fixture(string fixtureBaseName)
    {
        var fileName = fixtureBaseName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
            ? fixtureBaseName
            : fixtureBaseName + ".xlsx";

        var xlsxPath = Path.Combine(AppContext.BaseDirectory, "TestData", fileName);
        if (File.Exists(xlsxPath))
        {
            return new MemoryStream(File.ReadAllBytes(xlsxPath));
        }

        var base64Path = Path.Combine(AppContext.BaseDirectory, "TestData", $"{fileName}.base64");
        var base64 = File.ReadAllText(base64Path, Encoding.UTF8);
        var bytes = Convert.FromBase64String(base64);
        return new MemoryStream(bytes);
    }
}
