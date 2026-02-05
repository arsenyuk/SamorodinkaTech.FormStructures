using ClosedXML.Excel;

static void PrintUsage()
{
    Console.Error.WriteLine("Usage:");
    Console.Error.WriteLine("  dotnet run --project tools/ExcelFixtureGen/ExcelFixtureGen.csproj -- --list");
    Console.Error.WriteLine("  dotnet run --project tools/ExcelFixtureGen/ExcelFixtureGen.csproj -- --fixture <NAME> [--out-xlsx <PATH>] [--out-base64 <PATH>]");
    Console.Error.WriteLine();
    Console.Error.WriteLine("If no output path is provided, base64 is printed to stdout.");
}

static IReadOnlyList<string> ListFixtures() => new[]
{
    "REFBOOK-001",
    "REFBOOK-002",
    "BOTTOM-MERGED-001",
    "FORMULA-001",
    "AVG-001-AVERAGE",
    "AVG-002-SUM-COUNT",
    "AVG-003-SUM-DIV",
    "AVG-004-PLUS-DIV",
    "AVG-005-RU",
    "DIFF-001-Y-MINUS-Z",
    "TYPES-FORMAT-001",
    "TYPES-NOFORMAT-001",
};

static byte[] BuildFixture(string name)
{
    return name.ToUpperInvariant() switch
    {
        "REFBOOK-001" => BuildRefbook001(),
        "REFBOOK-002" => BuildRefbook002(),
        "BOTTOM-MERGED-001" => BuildBottomMerged001(),
        "FORMULA-001" => BuildFormula001(),
        "AVG-001-AVERAGE" => BuildAvg001Average(),
        "AVG-002-SUM-COUNT" => BuildAvg002SumCount(),
        "AVG-003-SUM-DIV" => BuildAvg003SumDiv(),
        "AVG-004-PLUS-DIV" => BuildAvg004PlusDiv(),
        "AVG-005-RU" => BuildAvg005Ru(),
        "DIFF-001-Y-MINUS-Z" => BuildDiff001YMinusZ(),
        "TYPES-FORMAT-001" => BuildTypesFormat001(),
        "TYPES-NOFORMAT-001" => BuildTypesNoFormat001(),
        _ => throw new ArgumentException($"Unknown fixture: {name}")
    };
}

static byte[] BuildDiff001YMinusZ()
{
    using var wb = new XLWorkbook();

    var ws = wb.AddWorksheet("Form");
    ws.Cell(1, 1).Value = "DIFF-001";
    ws.Cell(2, 1).Value = "Difference: X = Y - Z (title from Y; parts X and Z)";

    // 1-row header (row 3), 3 leaf columns.
    ws.Cell(3, 1).Value = "Y";
    ws.Cell(3, 2).Value = "Z";
    ws.Cell(3, 3).Value = "X";

    // Data starts at row 4.
    ws.Cell(4, 1).Value = 100;
    ws.Cell(4, 2).Value = 30;
    ws.Cell(4, 3).FormulaA1 = "A4-B4";

    using var ms = new MemoryStream();
    wb.SaveAs(ms);
    return ms.ToArray();
}

static byte[] BuildAvgBase(string formNumber, string title, string formulaA1)
{
    using var wb = new XLWorkbook();

    var ws = wb.AddWorksheet("Form");
    ws.Cell(1, 1).Value = formNumber;
    ws.Cell(2, 1).Value = title;

    // Simple 1-row header (row 3), 4 leaf columns.
    ws.Cell(3, 1).Value = "A";
    ws.Cell(3, 2).Value = "B";
    ws.Cell(3, 3).Value = "C";
    ws.Cell(3, 4).Value = "Avg";

    // Data starts at row 4.
    ws.Cell(4, 1).Value = 10;
    ws.Cell(4, 2).Value = 20;
    ws.Cell(4, 3).Value = 30;
    ws.Cell(4, 4).FormulaA1 = formulaA1;

    using var ms = new MemoryStream();
    wb.SaveAs(ms);
    return ms.ToArray();
}

static byte[] BuildAvg001Average()
    => BuildAvgBase(
        formNumber: "AVG-001",
        title: "Average: AVERAGE(A4:C4)",
        formulaA1: "AVERAGE(A4:C4)");

static byte[] BuildAvg002SumCount()
    => BuildAvgBase(
        formNumber: "AVG-002",
        title: "Average: SUM(A4:C4)/COUNT(A4:C4)",
        formulaA1: "SUM(A4:C4)/COUNT(A4:C4)");

static byte[] BuildAvg003SumDiv()
    => BuildAvgBase(
        formNumber: "AVG-003",
        title: "Average: SUM(A4:C4)/3",
        formulaA1: "SUM(A4:C4)/3");

static byte[] BuildAvg004PlusDiv()
    => BuildAvgBase(
        formNumber: "AVG-004",
        title: "Average: (A4+B4+C4)/3",
        formulaA1: "(A4+B4+C4)/3");

static byte[] BuildAvg005Ru()
    => BuildAvgBase(
        formNumber: "AVG-005",
        title: "Average (RU): СРЗНАЧ(A4:C4)",
        formulaA1: "СРЗНАЧ(A4:C4)");

static byte[] BuildFormula001()
{
    using var wb = new XLWorkbook();

    var ws = wb.AddWorksheet("Form");
    ws.Cell(1, 1).Value = "FORMULA-001";
    ws.Cell(2, 1).Value = "Form with formulas (c3 = c1 + c2)";

    // Simple 1-row header (row 3), 3 leaf columns.
    ws.Cell(3, 1).Value = "c1";
    ws.Cell(3, 2).Value = "c2";
    ws.Cell(3, 3).Value = "c3";

    // Data starts at row 4.
    ws.Cell(4, 1).Value = 10;
    ws.Cell(4, 2).Value = 20;
    ws.Cell(4, 3).FormulaA1 = "A4+B4";

    ws.Cell(5, 1).Value = 5;
    ws.Cell(5, 2).Value = 7;
    ws.Cell(5, 3).FormulaA1 = "A5+B5";

    ws.Cell(6, 1).Value = 100;
    ws.Cell(6, 2).Value = 200;
    ws.Cell(6, 3).FormulaA1 = "A6+B6";

    using var ms = new MemoryStream();
    wb.SaveAs(ms);
    return ms.ToArray();
}

static byte[] BuildTypesFormat001()
{
    using var wb = new XLWorkbook();

    var ws = wb.AddWorksheet("Form");
    ws.Cell(1, 1).Value = "TYPES-FORMAT-001";
    ws.Cell(2, 1).Value = "Form with explicit type hints via number formats";

    // Simple 1-row header (row 3), 5 leaf columns.
    ws.Cell(3, 1).Value = "String";
    ws.Cell(3, 2).Value = "Date";
    ws.Cell(3, 3).Value = "DateTime";
    ws.Cell(3, 4).Value = "Int";
    ws.Cell(3, 5).Value = "Decimal";

    // Row 4 is the first data row. Leave values empty but set explicit number formats.
    // These formats are the only source of type inference.
    ws.Cell(4, 1).Style.NumberFormat.Format = "@";
    ws.Cell(4, 2).Style.NumberFormat.Format = "yyyy-mm-dd";
    ws.Cell(4, 3).Style.NumberFormat.Format = "yyyy-mm-dd hh:mm:ss";
    ws.Cell(4, 4).Style.NumberFormat.Format = "0";
    ws.Cell(4, 5).Style.NumberFormat.Format = "0.00";

    using var ms = new MemoryStream();
    wb.SaveAs(ms);
    return ms.ToArray();
}

static byte[] BuildTypesNoFormat001()
{
    using var wb = new XLWorkbook();

    var ws = wb.AddWorksheet("Form");
    ws.Cell(1, 1).Value = "TYPES-NOFORMAT-001";
    ws.Cell(2, 1).Value = "Form without explicit type hints";

    // Simple 1-row header (row 3), 3 leaf columns.
    ws.Cell(3, 1).Value = "A";
    ws.Cell(3, 2).Value = "B";
    ws.Cell(3, 3).Value = "C";

    // Data starts at row 4. Put numeric/date-ish values but do not set number formats.
    // The parser should NOT infer types from values alone.
    ws.Cell(4, 1).Value = 10;
    ws.Cell(4, 2).Value = 20.5;
    ws.Cell(4, 3).Value = "2026-02-03";

    using var ms = new MemoryStream();
    wb.SaveAs(ms);
    return ms.ToArray();
}

static byte[] BuildRefbook001()
{
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
    return ms.ToArray();
}

static byte[] BuildRefbook002()
{
    using var wb = new XLWorkbook();

    var form = wb.AddWorksheet("Form");
    form.Cell(1, 1).Value = "REFBOOK-002";
    form.Cell(2, 1).Value = "Form with two reference books";

    // Simple 1-row header (row 3), 2 leaf columns.
    form.Cell(3, 1).Value = "Color";
    form.Cell(3, 2).Value = "Size";

    var lists = wb.AddWorksheet("Lists");
    // Two different reference books (distinct ranges).
    lists.Cell(1, 1).Value = "Red";
    lists.Cell(2, 1).Value = "Green";
    lists.Cell(3, 1).Value = "Blue";

    lists.Cell(1, 2).Value = "S";
    lists.Cell(2, 2).Value = "M";
    lists.Cell(3, 2).Value = "L";

    // Apply list validations to two different columns.
    var colorCell = form.Cell(5, 1);
    var colorDv = colorCell.CreateDataValidation();
    colorDv.AllowedValues = XLAllowedValues.List;
    colorDv.InCellDropdown = true;
    colorDv.List("=Lists!$A$1:$A$3");

    var sizeCell = form.Cell(5, 2);
    var sizeDv = sizeCell.CreateDataValidation();
    sizeDv.AllowedValues = XLAllowedValues.List;
    sizeDv.InCellDropdown = true;
    sizeDv.List("=Lists!$B$1:$B$3");

    // Put a couple of data rows.
    form.Cell(5, 1).Value = "Red";
    form.Cell(5, 2).Value = "M";
    form.Cell(6, 1).Value = "Blue";
    form.Cell(6, 2).Value = "S";

    using var ms = new MemoryStream();
    wb.SaveAs(ms);
    return ms.ToArray();
}

static byte[] BuildBottomMerged001()
{
    using var wb = new XLWorkbook();

    var ws = wb.AddWorksheet("Form");
    ws.Cell(1, 1).Value = "TEST-001";
    ws.Cell(2, 1).Value = "Title";

    // Single-row header (row 3) that contains a merged cell.
    ws.Cell(3, 1).Value = "A";
    ws.Range(3, 1, 3, 2).Merge();
    ws.Cell(3, 3).Value = "B";

    using var ms = new MemoryStream();
    wb.SaveAs(ms);
    return ms.ToArray();
}

static string? GetArgValue(string[] args, string key)
{
    for (var i = 0; i < args.Length - 1; i++)
    {
        if (string.Equals(args[i], key, StringComparison.OrdinalIgnoreCase))
        {
            return args[i + 1];
        }
    }

    return null;
}

static bool HasFlag(string[] args, string flag) => args.Any(a => string.Equals(a, flag, StringComparison.OrdinalIgnoreCase));

var cliArgs = Environment.GetCommandLineArgs().Skip(1).ToArray();

if (HasFlag(cliArgs, "--list"))
{
    foreach (var f in ListFixtures())
    {
        Console.WriteLine(f);
    }

    return;
}

var fixtureName = GetArgValue(cliArgs, "--fixture");
if (string.IsNullOrWhiteSpace(fixtureName))
{
    PrintUsage();
    Console.Error.WriteLine();
    Console.Error.WriteLine("Available fixtures:");
    foreach (var f in ListFixtures())
    {
        Console.Error.WriteLine($"  - {f}");
    }

    Environment.ExitCode = 2;
    return;
}

var bytes = BuildFixture(fixtureName);
var base64 = Convert.ToBase64String(bytes);

var outXlsx = GetArgValue(cliArgs, "--out-xlsx");
var outBase64 = GetArgValue(cliArgs, "--out-base64");

if (!string.IsNullOrWhiteSpace(outXlsx))
{
    File.WriteAllBytes(outXlsx, bytes);
}

if (!string.IsNullOrWhiteSpace(outBase64))
{
    File.WriteAllText(outBase64, base64);
}

if (string.IsNullOrWhiteSpace(outXlsx) && string.IsNullOrWhiteSpace(outBase64))
{
    Console.Write(base64);
}
