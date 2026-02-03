using SamorodinkaTech.FormStructures.Web.Services;
using Xunit;
using System.Text;
using SamorodinkaTech.FormStructures.Web.Models;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.FileProviders;

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

    [Fact]
    public void ExtractReferenceBooks_UsesHeaderCellAsTitle_WhenRangeHasHeaderAbove()
    {
        using var wb = new XLWorkbook();

        var form = wb.AddWorksheet("Form");
        form.Cell(1, 1).Value = "REFBOOK-003";
        form.Cell(2, 1).Value = "Form with reference book header";
        form.Cell(3, 1).Value = "Column";

        var lists = wb.AddWorksheet("Lists");
        lists.Cell(1, 1).Value = "Colors";
        lists.Cell(2, 1).Value = "Red";
        lists.Cell(3, 1).Value = "Green";
        lists.Cell(4, 1).Value = "Blue";

        var target = form.Cell(5, 1);
        var dv = target.CreateDataValidation();
        dv.AllowedValues = XLAllowedValues.List;
        dv.InCellDropdown = true;
        dv.List("=Lists!$A$2:$A$4");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        var parser = new ExcelFormParser();
        var books = parser.ExtractReferenceBooks(ms);

        Assert.Single(books);
        var book = books[0];

        Assert.Equal("Colors", book.Title);
        Assert.Equal("Lists", book.SourceSheet);
        Assert.Equal("$A$2:$A$4", book.SourceRange);
        Assert.Equal(new[] { "Red", "Green", "Blue" }, book.Values);
    }

    [Fact]
    public void FormStorage_EnhancesReferenceBookTitles_FromAppliedToColumnHeaders()
    {
        using var stream = LoadXlsxFromBase64Fixture("REFBOOK-002.xlsx");

        var parser = new ExcelFormParser();
        var structure = parser.Parse(stream, "REFBOOK-002.xlsx");

        stream.Position = 0;
        var books = parser.ExtractReferenceBooks(stream);

        // Store the same files FormStorage expects.
        var tempRoot = Path.Combine(Path.GetTempPath(), "FormStructuresTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);

        try
        {
            var storageRoot = Path.Combine(tempRoot, "storage");
            var versionDir = Path.Combine(storageRoot, "forms", structure.FormNumber, "v1");
            Directory.CreateDirectory(versionDir);

            structure = structure with { Version = 1 };

            File.WriteAllText(
                Path.Combine(versionDir, "structure.json"),
                System.Text.Json.JsonSerializer.Serialize(structure, JsonUtil.StableOptions));

            File.WriteAllText(
                Path.Combine(versionDir, "reference-books.json"),
                System.Text.Json.JsonSerializer.Serialize(books, JsonUtil.StableOptions));

            stream.Position = 0;
            File.WriteAllBytes(Path.Combine(versionDir, "original.xlsx"), stream.ToArray());

            var env = new TestHostEnvironment { ContentRootPath = tempRoot };
            var storage = new FormStorage(
                Options.Create(new StorageOptions { StorageRoot = "storage" }),
                env,
                parser,
                NullLogger<FormStorage>.Instance);

            var loaded = storage.TryLoadReferenceBooks(structure.FormNumber, version: 1);

            // Titles should come from the form's column headers (Color / Size), not technical sheet/range strings.
            Assert.Contains(loaded, b => string.Equals(b.Title, "Color", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(loaded, b => string.Equals(b.Title, "Size", StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            try { Directory.Delete(tempRoot, recursive: true); } catch { /* best-effort */ }
        }
    }

    private sealed class TestHostEnvironment : IWebHostEnvironment
    {
        public string EnvironmentName { get; set; } = "Development";
        public string ApplicationName { get; set; } = "Tests";
        public string WebRootPath { get; set; } = string.Empty;
        public IFileProvider WebRootFileProvider { get; set; } = new NullFileProvider();
        public string ContentRootPath { get; set; } = string.Empty;
        public IFileProvider ContentRootFileProvider { get; set; } = new NullFileProvider();
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
