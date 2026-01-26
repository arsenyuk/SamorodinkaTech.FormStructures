using ClosedXML.Excel;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;
using Xunit;

namespace SamorodinkaTech.FormStructures.Tests;

public class ExcelFormParserTests
{
    [Fact]
    public void ParseLayout_NormalizesFormNumber_AndDetectsHeaderBoundary_WhenDataRowsExist()
    {
        using var stream = LoadXlsxFromBase64Fixture("TEST-001-filled.xlsx");

        var parser = new ExcelFormParser();
        var layout = parser.ParseLayout(stream, sourceFileName: "TEST-001-filled.xlsx");

        Assert.Equal("001", layout.Structure.FormNumber);
        Assert.Equal("Demo form (5 columns)", layout.Structure.FormTitle);

        Assert.Equal(3, layout.HeaderRowStart);
        Assert.Equal(4, layout.LastHeaderRow);
        Assert.Equal(5, layout.DataStartRow);

        Assert.Equal(5, layout.Structure.Columns.Count);
        Assert.Equal(5, layout.LeafColumns.Count);
        Assert.All(layout.LeafColumns, c => Assert.InRange(c, 1, 5));
    }

    [Fact]
    public void ReadDataRows_SkipsEmptyRows_AndReadsValuesByLeafColumns()
    {
        using var stream = BuildWorkbook(includeDataRows: true, includeEmptyRow: true);

        var parser = new ExcelFormParser();
        var layout = parser.ParseLayout(stream, sourceFileName: "TEST-001-filled.xlsx");

        // Need a new stream for reading rows.
        stream.Position = 0;
        var rows = parser.ReadDataRows(stream, layout);

        // We generate 3 data rows and 1 fully empty row -> should return 3.
        Assert.Equal(3, rows.Count);

        var first = rows[0];
        Assert.Equal(5, first.RowNumber);

        Assert.Equal("alpha", first.Values[layout.Structure.Columns[0].Path]);
        Assert.Equal("10", first.Values[layout.Structure.Columns[1].Path]);
        Assert.Equal("x", first.Values[layout.Structure.Columns[2].Path]);
    }

    [Fact]
    public void ParseLayout_DetectsColumnIndexRow_AndDoesNotTreatItAsData()
    {
        using var stream = LoadXlsxFromBase64Fixture("TEST-002.xlsx");

        var parser = new ExcelFormParser();
        var layout = parser.ParseLayout(stream, sourceFileName: "TEST-002.xlsx");

        Assert.Equal("002", layout.Structure.FormNumber);
        Assert.Equal("Demo form with column indices (5 columns)", layout.Structure.FormTitle);

        Assert.Equal(3, layout.HeaderRowStart);
        Assert.Equal(5, layout.LastHeaderRow);
        Assert.Equal(6, layout.DataStartRow);

        Assert.Equal(5, layout.Structure.Columns.Count);
        Assert.All(layout.Structure.Columns, c => Assert.False(string.IsNullOrWhiteSpace(c.ColumnNumber)));
        Assert.Equal(new[] { "1", "2", "3", "4", "5" }, layout.Structure.Columns.Select(c => c.ColumnNumber).ToArray());

        // Ensure the index row was not interpreted as data.
        stream.Position = 0;
        var rows = parser.ReadDataRows(stream, layout);
        Assert.Empty(rows);
    }

    [Fact]
    public void ParseLayout_IgnoresEmptyButFormattedRowAfterHeader()
    {
        using var stream = LoadXlsxFromBase64Fixture("TEST-003-types-empty-row.xlsx");

        var parser = new ExcelFormParser();
        var layout = parser.ParseLayout(stream, sourceFileName: "TEST-003-types-empty-row.xlsx");

        Assert.Equal("003", layout.Structure.FormNumber);
        Assert.Equal("Demo form: typed empty row", layout.Structure.FormTitle);

        // Simple 1-row header.
        Assert.Equal(3, layout.HeaderRowStart);
        Assert.Equal(3, layout.LastHeaderRow);
        Assert.Equal(4, layout.DataStartRow);

        // The sheet contains an extra empty row (row 4) with formats applied, but no actual content.
        // It must NOT extend the used range.
        Assert.Equal(3, layout.UsedLastRow);

        Assert.Equal(5, layout.Structure.Columns.Count);
        Assert.Equal(new[] { "String", "Date", "DateTime", "Int", "Decimal" }, layout.Structure.Columns.Select(c => c.Name).ToArray());

        stream.Position = 0;
        var rows = parser.ReadDataRows(stream, layout);
        Assert.Empty(rows);
    }

    [Fact]
    public void StructureHash_IsStableBetweenTemplateAndFilledFile()
    {
        var parser = new ExcelFormParser();

        using var template = LoadXlsxFromBase64Fixture("TEST-001.xlsx");
        var templateLayout = parser.ParseLayout(template, sourceFileName: "TEST-001.xlsx");

        using var filled = LoadXlsxFromBase64Fixture("TEST-001-filled.xlsx");
        var filledLayout = parser.ParseLayout(filled, sourceFileName: "TEST-001-filled.xlsx");

        Assert.Equal(templateLayout.Structure.StructureHash, filledLayout.Structure.StructureHash);
    }

    [Fact]
    public void ParseLayout_Throws_WhenStreamIsNotXlsx()
    {
        var parser = new ExcelFormParser();

        using var ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("not an xlsx"));

        var ex = Assert.Throws<FormParseException>(() => parser.ParseLayout(ms, sourceFileName: "bad.txt"));
        Assert.Contains("Failed to parse Excel file", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseLayout_Throws_WhenWorkbookHasNoUsedCells()
    {
        var parser = new ExcelFormParser();
        using var stream = CreateWorkbook(ws =>
        {
            // Intentionally empty worksheet (no used cells).
        });

        var ex = Assert.Throws<FormParseException>(() => parser.ParseLayout(stream, sourceFileName: "empty.xlsx"));
        Assert.Contains("Excel file is empty", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseLayout_Throws_WhenFormNumberMissing()
    {
        var parser = new ExcelFormParser();
        using var stream = CreateWorkbook(ws =>
        {
            // Row 1 intentionally blank.
            ws.Cell(2, 1).Value = "Title";
            ws.Cell(3, 1).Value = "H";
        });

        var ex = Assert.Throws<FormParseException>(() => parser.ParseLayout(stream, sourceFileName: "no-form.xlsx"));
        Assert.Contains("Form number is missing", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseLayout_Throws_WhenFormTitleMissing()
    {
        var parser = new ExcelFormParser();
        using var stream = CreateWorkbook(ws =>
        {
            ws.Cell(1, 1).Value = "TEST-001";
            // Row 2 intentionally blank.
            ws.Cell(3, 1).Value = "H";
        });

        var ex = Assert.Throws<FormParseException>(() => parser.ParseLayout(stream, sourceFileName: "no-title.xlsx"));
        Assert.Contains("Form title is missing", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseLayout_Throws_WhenHeaderMissing_NoRowsAfterTitle()
    {
        var parser = new ExcelFormParser();
        using var stream = CreateWorkbook(ws =>
        {
            ws.Cell(1, 1).Value = "TEST-001";
            ws.Cell(2, 1).Value = "Title";
            // No header rows.
        });

        var ex = Assert.Throws<FormParseException>(() => parser.ParseLayout(stream, sourceFileName: "no-header.xlsx"));
        Assert.Contains("Header is missing", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseLayout_Throws_WhenHeaderHasBottomRowGap()
    {
        var parser = new ExcelFormParser();
        using var stream = CreateWorkbook(ws =>
        {
            ws.Cell(1, 1).Value = "TEST-001";
            ws.Cell(2, 1).Value = "Title";

            // Header row 3 and 4, but row 4 has a gap (col 2 missing).
            ws.Cell(3, 1).Value = "Group";
            ws.Range(3, 1, 3, 3).Merge();

            ws.Cell(4, 1).Value = "A1";
            // ws.Cell(4, 2) intentionally blank.
            ws.Cell(4, 3).Value = "A3";
        });

        var ex = Assert.Throws<FormParseException>(() => parser.ParseLayout(stream, sourceFileName: "gap.xlsx"));
        Assert.Contains("gap", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseLayout_Throws_WhenMergedHeaderTopLeftIsEmpty()
    {
        var parser = new ExcelFormParser();
        using var stream = CreateWorkbook(ws =>
        {
            ws.Cell(1, 1).Value = "TEST-001";
            ws.Cell(2, 1).Value = "Title";

            // Merged header with empty label should fail.
            ws.Range(3, 1, 3, 2).Merge();
            ws.Cell(3, 1).Value = "";

            ws.Cell(4, 1).Value = "A1";
            ws.Cell(4, 2).Value = "A2";
        });

        var ex = Assert.Throws<FormParseException>(() => parser.ParseLayout(stream, sourceFileName: "merged-empty.xlsx"));
        Assert.Contains("Header cell is empty", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseLayout_Throws_WhenLeafHeaderSpansMultipleColumns()
    {
        var parser = new ExcelFormParser();
        using var stream = CreateWorkbook(ws =>
        {
            ws.Cell(1, 1).Value = "TEST-001";
            ws.Cell(2, 1).Value = "Title";

            // Single-row header (row 3) where a leaf spans multiple columns.
            ws.Cell(3, 1).Value = "A";
            ws.Range(3, 1, 3, 2).Merge();
            ws.Cell(3, 3).Value = "B";
        });

        var ex = Assert.Throws<FormParseException>(() => parser.ParseLayout(stream, sourceFileName: "leaf-span.xlsx"));
        Assert.Contains("Leaf header", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static MemoryStream BuildWorkbook(bool includeDataRows, bool includeEmptyRow = false)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Form");

        ws.Cell(1, 1).Value = "TEST-001";
        ws.Cell(2, 1).Value = "Demo form (5 columns)";

        ws.Cell(3, 1).Value = "Group A";
        ws.Range(3, 1, 3, 2).Merge();

        ws.Cell(3, 3).Value = "Group B";
        ws.Range(3, 3, 3, 5).Merge();

        ws.Cell(4, 1).Value = "A1";
        ws.Cell(4, 2).Value = "A2";
        ws.Cell(4, 3).Value = "B1";
        ws.Cell(4, 4).Value = "B2";
        ws.Cell(4, 5).Value = "B3";

        if (includeDataRows)
        {
            // Data starts at row 5.
            ws.Cell(5, 1).Value = "alpha";
            ws.Cell(5, 2).Value = "10";
            ws.Cell(5, 3).Value = "x";
            ws.Cell(5, 4).Value = "2026-01-22";
            ws.Cell(5, 5).Value = "true";

            if (includeEmptyRow)
            {
                // Row 6 intentionally left empty.
            }

            ws.Cell(includeEmptyRow ? 7 : 6, 1).Value = "beta";
            ws.Cell(includeEmptyRow ? 7 : 6, 2).Value = "20";
            ws.Cell(includeEmptyRow ? 7 : 6, 3).Value = "y";

            ws.Cell(includeEmptyRow ? 8 : 7, 1).Value = "gamma";
            ws.Cell(includeEmptyRow ? 8 : 7, 2).Value = "30";
            ws.Cell(includeEmptyRow ? 8 : 7, 3).Value = "z";
        }

        var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        return ms;
    }

    private static MemoryStream LoadXlsxFromBase64Fixture(string xlsxFileName)
    {
        var xlsxPath = Path.Combine(AppContext.BaseDirectory, "TestData", xlsxFileName);
        if (File.Exists(xlsxPath))
        {
            return new MemoryStream(File.ReadAllBytes(xlsxPath));
        }

        var base64Path = Path.Combine(AppContext.BaseDirectory, "TestData", $"{xlsxFileName}.base64");
        var base64 = File.ReadAllText(base64Path);
        var bytes = Convert.FromBase64String(base64);
        return new MemoryStream(bytes);
    }

    private static MemoryStream CreateWorkbook(Action<IXLWorksheet> configure)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Form");
        configure(ws);

        var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        return ms;
    }
}
