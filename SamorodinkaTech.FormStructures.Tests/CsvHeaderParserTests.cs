using System.Text;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

namespace SamorodinkaTech.FormStructures.Tests;

public sealed class CsvHeaderParserTests
{
    private static MemoryStream LoadCsvFixture(string fileName)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", fileName);
        return new MemoryStream(File.ReadAllBytes(path));
    }

    [Fact]
    public void ParseHeaderRow_CommaSeparated_Works()
    {
        var csv = "A,B,C\n1,2,3\n";
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        var result = CsvHeaderParser.ParseHeaderRow(ms);

        Assert.Equal(',', result.Separator);
        Assert.Equal(["A", "B", "C"], result.Headers);
    }

    [Fact]
    public void ParseHeaderRow_SemicolonSeparated_Works()
    {
        var csv = "A;B;C\n";
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        var result = CsvHeaderParser.ParseHeaderRow(ms);

        Assert.Equal(';', result.Separator);
        Assert.Equal(["A", "B", "C"], result.Headers);
    }

    [Fact]
    public void ParseHeaderRow_TabSeparated_Works()
    {
        var csv = "A\tB\tC\n";
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        var result = CsvHeaderParser.ParseHeaderRow(ms);

        Assert.Equal('\t', result.Separator);
        Assert.Equal(["A", "B", "C"], result.Headers);
    }

    [Fact]
    public void ParseHeaderRow_QuotedHeadersWithComma_Works()
    {
        var csv = "\"A,1\",B\n";
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        var result = CsvHeaderParser.ParseHeaderRow(ms);

        Assert.Equal(',', result.Separator);
        Assert.Equal(["A,1", "B"], result.Headers);
    }

    [Fact]
    public void ParseHeaderRow_EscapedQuotes_Works()
    {
        var csv = "\"A\"\"B\",C\n";
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        var result = CsvHeaderParser.ParseHeaderRow(ms);

        Assert.Equal(["A\"B", "C"], result.Headers);
    }

    [Fact]
    public void ParseHeaderRow_EmptyHeader_Throws()
    {
        var csv = ",B\n";
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        var ex = Assert.Throws<FormParseException>(() => CsvHeaderParser.ParseHeaderRow(ms));
        Assert.Contains("empty", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseHeaderRow_DuplicateHeaders_Throws()
    {
        var csv = "A,B,a\n";
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        var ex = Assert.Throws<FormParseException>(() => CsvHeaderParser.ParseHeaderRow(ms));
        Assert.Contains("duplicate", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseHeaderRow_Utf8Bom_DoesNotLeakIntoFirstHeader()
    {
        var utf8Bom = Encoding.UTF8.GetPreamble();
        var payload = Encoding.UTF8.GetBytes("A,B\n");
        var bytes = utf8Bom.Concat(payload).ToArray();

        using var ms = new MemoryStream(bytes);

        var result = CsvHeaderParser.ParseHeaderRow(ms);

        Assert.Equal(["A", "B"], result.Headers);
    }

    [Fact]
    public void ParseHeaderRow_FixtureExample_Positive_Works()
    {
        using var ms = LoadCsvFixture("CSV-SCHEMA-NEWFORM-OK.csv");

        var result = CsvHeaderParser.ParseHeaderRow(ms);

        Assert.Equal(';', result.Separator);
        Assert.Equal(["Имя", "Возраст", "Город;регион", "A\"B"], result.Headers);
    }

    [Fact]
    public void ParseHeaderRow_FixtureExample_NegativeDuplicate_Throws()
    {
        using var ms = LoadCsvFixture("CSV-SCHEMA-NEWFORM-NEGATIVE-DUPLICATE.csv");

        var ex = Assert.Throws<FormParseException>(() => CsvHeaderParser.ParseHeaderRow(ms));
        Assert.Contains("duplicate", ex.Message, StringComparison.OrdinalIgnoreCase);
    }
}
