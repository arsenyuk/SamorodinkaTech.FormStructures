using SamorodinkaTech.FormStructures.Web.Services;
using Xunit;

namespace SamorodinkaTech.FormStructures.Tests;

public class CsvUtilTests
{
    [Fact]
    public void Escape_ReturnsEmpty_ForNullOrEmpty()
    {
        Assert.Equal(string.Empty, CsvUtil.Escape(null));
        Assert.Equal(string.Empty, CsvUtil.Escape(""));
    }

    [Fact]
    public void Escape_Quotes_WhenContainsSeparatorOrNewlineOrQuote()
    {
        Assert.Equal("\"a,b\"", CsvUtil.Escape("a,b"));
        Assert.Equal("\"a\r\nb\"", CsvUtil.Escape("a\r\nb"));
        Assert.Equal("\"a\"\"b\"", CsvUtil.Escape("a\"b"));
    }

    [Fact]
    public void BuildUtf8CsvWithBom_WritesHeadersAndBom_AndKeepsUnicode()
    {
        var bytes = CsvUtil.BuildUtf8CsvWithBom(
            headers: new[] { "Имя", "City" },
            rows: new[] { (IReadOnlyList<string?>)new[] { "Алиса", "Zürich" } });

        Assert.True(bytes.Length > 3);
        Assert.Equal(0xEF, bytes[0]);
        Assert.Equal(0xBB, bytes[1]);
        Assert.Equal(0xBF, bytes[2]);

        var text = System.Text.Encoding.UTF8.GetString(bytes);
        Assert.Contains("Имя,City", text);
        Assert.Contains("Алиса,Zürich", text);
    }
}
