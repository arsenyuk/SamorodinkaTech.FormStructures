using ClosedXML.Excel;

static XLWorkbook BuildInconsistentFormulaWorkbook()
{
    var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("Form");

    // Form number: will normalize to 010.
    ws.Cell(1, 1).Value = "FORMULA-010";
    ws.Cell(2, 1).Value = "Formula validation: inconsistent formulas";

    // Single-row header (row 3). Data starts at row 4.
    ws.Cell(3, 1).Value = "A";
    ws.Cell(3, 2).Value = "B";
    ws.Cell(3, 3).Value = "Sum";

    // Row 4
    ws.Cell(4, 1).Value = 10;
    ws.Cell(4, 2).Value = 20;
    ws.Cell(4, 3).FormulaA1 = "A4+B4";

    // Row 5 (different formula -> should fail)
    ws.Cell(5, 1).Value = 5;
    ws.Cell(5, 2).Value = 7;
    ws.Cell(5, 3).FormulaA1 = "A5+B5+1";

    // Row 6
    ws.Cell(6, 1).Value = 100;
    ws.Cell(6, 2).Value = 200;
    ws.Cell(6, 3).FormulaA1 = "A6+B6";

    return wb;
}

var outPath = args.Length > 0 ? args[0] : throw new ArgumentException("Expected output .base64 path as arg0");
Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outPath))!);

using var wb = BuildInconsistentFormulaWorkbook();
using var ms = new MemoryStream();
wb.SaveAs(ms);
var bytes = ms.ToArray();
var base64 = Convert.ToBase64String(bytes);

File.WriteAllText(outPath, base64);
Console.WriteLine($"Wrote {outPath} ({bytes.Length} bytes, base64 length {base64.Length})");
