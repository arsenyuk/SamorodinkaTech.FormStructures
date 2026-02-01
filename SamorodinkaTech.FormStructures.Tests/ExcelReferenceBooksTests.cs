using SamorodinkaTech.FormStructures.Web.Services;
using Xunit;
using System.Text;

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
