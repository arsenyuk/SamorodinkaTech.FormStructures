using ClosedXML.Excel;

using var wb = new XLWorkbook();

var form = wb.AddWorksheet("Form");
form.Cell(1, 1).Value = "REFBOOK-001";
form.Cell(2, 1).Value = "Form with reference book";
form.Cell(3, 1).Value = "Column";

var lists = wb.AddWorksheet("Lists");
lists.Cell(1, 1).Value = "Alpha";
lists.Cell(2, 1).Value = "Beta";
lists.Cell(3, 1).Value = "Gamma";

var target = form.Cell(5, 1);
var dv = target.CreateDataValidation();
dv.AllowedValues = XLAllowedValues.List;
dv.InCellDropdown = true;

// Cross-sheet list source as an explicit formula.
dv.List("=Lists!$A$1:$A$3");

using var ms = new MemoryStream();
wb.SaveAs(ms);
var bytes = ms.ToArray();

Console.Write(Convert.ToBase64String(bytes));
