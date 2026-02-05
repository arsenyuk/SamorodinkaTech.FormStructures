using ClosedXML.Excel;
using SamorodinkaTech.FormStructures.Web.Models;

namespace SamorodinkaTech.FormStructures.Web.Services;

public static class CsvSchemaTemplateBuilder
{
    public static byte[] BuildXlsxTemplateBytes(string formNumber, string formTitle, IReadOnlyList<string> headers)
    {
        if (string.IsNullOrWhiteSpace(formNumber))
        {
            throw new ArgumentException("Form number is required.", nameof(formNumber));
        }

        if (string.IsNullOrWhiteSpace(formTitle))
        {
            throw new ArgumentException("Form title is required.", nameof(formTitle));
        }

        if (headers is null || headers.Count == 0)
        {
            throw new FormParseException("CSV header row has no columns.");
        }

        if (headers.Count > CsvHeaderParser.MaxColumns)
        {
            throw new FormParseException($"CSV header row has too many columns ({headers.Count}). Max is {CsvHeaderParser.MaxColumns}.");
        }

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell(1, 1).Value = formNumber.Trim();
        ws.Cell(2, 1).Value = formTitle.Trim();

        // Row 3 is the header row (ExcelFormParser expects header starting from row 3).
        for (var i = 0; i < headers.Count; i++)
        {
            ws.Cell(3, i + 1).Value = headers[i];
        }

        // Make it a bit more readable if someone downloads the generated template.
        ws.Row(1).Style.Font.Bold = true;
        ws.Row(2).Style.Font.Bold = true;
        ws.Row(3).Style.Font.Bold = true;
        ws.Columns(1, headers.Count).AdjustToContents();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }
}
