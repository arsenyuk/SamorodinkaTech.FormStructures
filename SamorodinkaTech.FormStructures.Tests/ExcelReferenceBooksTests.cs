using SamorodinkaTech.FormStructures.Web.Services;
using Xunit;
using System.Text;
using SamorodinkaTech.FormStructures.Web.Models;

namespace SamorodinkaTech.FormStructures.Tests;

public sealed class ExcelReferenceBooksTests
{
    [Fact]
    public void ExtractReferenceBooks_FindsListValidationWithRange_AndReadsValues()
    {
        using var stream = LoadXlsxFromBase64Fixture("REFBOOK-001.xlsx");

        var parser = new ExcelFormParser();
        var books = parser.ExtractReferenceBooks(stream);

        Assert.Single(books);
        var book = books[0];

        Assert.Equal(new[] { "Alpha", "Beta", "Gamma" }, book.Values);
        Assert.Equal("Lists", book.SourceSheet);
        Assert.Equal("$A$1:$A$3", book.SourceRange);

        Assert.Contains(book.AppliedTo, a => a.Contains("Form!", StringComparison.OrdinalIgnoreCase) && a.Contains("A5", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ExtractReferenceBooks_SupportsTwoDifferentColumns_WithDifferentLists()
    {
        using var stream = LoadXlsxFromBase64Fixture("REFBOOK-002.xlsx");

        var parser = new ExcelFormParser();

        // 1) Ensure we extract two different reference books.
        var books = parser.ExtractReferenceBooks(stream);
        Assert.Equal(2, books.Count);

        Assert.Contains(books, b => b.SourceSheet == "Lists" && b.SourceRange == "$A$1:$A$3" && b.Values.SequenceEqual(new[] { "Red", "Green", "Blue" }));
        Assert.Contains(books, b => b.SourceSheet == "Lists" && b.SourceRange == "$B$1:$B$3" && b.Values.SequenceEqual(new[] { "S", "M", "L" }));

        // 2) Ensure layout auto-detects ReferenceBook type for both columns.
        stream.Position = 0;
        var layout = parser.ParseLayout(stream, sourceFileName: "REFBOOK-002.xlsx");

        Assert.Equal(2, layout.Structure.Columns.Count);
        Assert.All(layout.Structure.Columns, c => Assert.Equal(ColumnType.ReferenceBook, c.Type));

        // 3) Ensure we can read data rows and values match the filled cells.
        stream.Position = 0;
        var rows = parser.ReadDataRows(stream, layout);
        Assert.Equal(2, rows.Count);

        var colorPath = layout.Structure.Columns[0].Path;
        var sizePath = layout.Structure.Columns[1].Path;

        Assert.Equal("Red", rows[0].Values[colorPath]);
        Assert.Equal("M", rows[0].Values[sizePath]);
        Assert.Equal("Blue", rows[1].Values[colorPath]);
        Assert.Equal("S", rows[1].Values[sizePath]);
    }

    private static MemoryStream LoadXlsxFromBase64Fixture(string xlsxFileName)
    {
        var xlsxPath = Path.Combine(AppContext.BaseDirectory, "TestData", xlsxFileName);
        if (File.Exists(xlsxPath))
        {
            return new MemoryStream(File.ReadAllBytes(xlsxPath));
        }

        var base64Path = Path.Combine(AppContext.BaseDirectory, "TestData", $"{xlsxFileName}.base64");
        var base64 = File.ReadAllText(base64Path, Encoding.UTF8);
        var bytes = Convert.FromBase64String(base64);
        return new MemoryStream(bytes);
    }
}
