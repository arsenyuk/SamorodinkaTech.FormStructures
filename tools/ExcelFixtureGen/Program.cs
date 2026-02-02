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
    "BOTTOM-MERGED-001",
};

static byte[] BuildFixture(string name)
{
    return name.ToUpperInvariant() switch
    {
        "REFBOOK-001" => BuildRefbook001(),
        "BOTTOM-MERGED-001" => BuildBottomMerged001(),
        _ => throw new ArgumentException($"Unknown fixture: {name}")
    };
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
