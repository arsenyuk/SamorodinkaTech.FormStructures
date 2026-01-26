using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;

var repoRoot = Directory.GetCurrentDirectory();
var webRoot = Path.Combine(repoRoot, "SamorodinkaTech.FormStructures.Web");
if (!Directory.Exists(webRoot))
{
    Console.Error.WriteLine($"Expected web project folder at: {webRoot}");
    return 1;
}

var outDir = Path.Combine(repoRoot, "tools", "SeedForm", "out");
Directory.CreateDirectory(outDir);

var mode = args.FirstOrDefault()?.Trim().ToLowerInvariant();

static void BuildTemplate(string path, string rawFormNumber, string formTitle, bool includeColumnIndexRow)
{
    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("Form");

    // Row 1: form number (any non-empty cell works)
    ws.Cell(1, 1).Value = rawFormNumber;

    // Row 2: form title
    ws.Cell(2, 1).Value = formTitle;

    // Header starts at row 3.
    // Row 3: grouped headers
    ws.Cell(3, 1).Value = "Group A";
    ws.Range(3, 1, 3, 2).Merge();

    ws.Cell(3, 3).Value = "Group B";
    ws.Range(3, 3, 3, 5).Merge();

    // Row 4: leaf headers (5 columns)
    ws.Cell(4, 1).Value = "A1";
    ws.Cell(4, 2).Value = "A2";
    ws.Cell(4, 3).Value = "B1";
    ws.Cell(4, 4).Value = "B2";
    ws.Cell(4, 5).Value = "B3";

    if (includeColumnIndexRow)
    {
        // Row 5: column indices (1..N)
        for (var i = 1; i <= 5; i++)
        {
            ws.Cell(5, i).Value = i;
        }
    }

    ws.Columns(1, 5).AdjustToContents();
    wb.SaveAs(path);
}

static async Task SeedSchemaAsync(FormStorage storage, ExcelFormParser parser, string templatePath, CancellationToken ct)
{
    await using var templateFs = File.OpenRead(templatePath);
    var templateFile = new FormFile(templateFs, 0, templateFs.Length, "Upload", Path.GetFileName(templatePath))
    {
        Headers = new HeaderDictionary(),
        ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    };

    var schemaResult = await storage.SaveAsync(templateFile, parser, ct);
    Console.WriteLine($"Seeded form schema: #{schemaResult.FormNumber} v{schemaResult.Version} (new={schemaResult.IsNewVersion})");

    // New-form uploads may be staged as pending until types are confirmed.
    if (schemaResult.PendingId is not null)
    {
        var pending = storage.TryLoadPending(schemaResult.FormNumber, schemaResult.PendingId);
        if (pending is null)
        {
            throw new Exception($"Pending upload not found for {schemaResult.FormNumber} ({schemaResult.PendingId})");
        }

        await storage.CommitPendingAsync(schemaResult.FormNumber, schemaResult.PendingId, pending.Structure, ct);
        Console.WriteLine($"Committed pending schema: #{schemaResult.FormNumber} v{pending.Structure.Version} ({schemaResult.PendingId})");
    }
}

var demo1Raw = "TEST-001";
var demo1Title = "Demo form (5 columns)";
var demo1TemplatePath = Path.Combine(outDir, $"{demo1Raw}.xlsx");
BuildTemplate(demo1TemplatePath, demo1Raw, demo1Title, includeColumnIndexRow: false);

var demo2Raw = "TEST-002";
var demo2Title = "Demo form with column indices (5 columns)";
var demo2TemplatePath = Path.Combine(outDir, $"{demo2Raw}.xlsx");
BuildTemplate(demo2TemplatePath, demo2Raw, demo2Title, includeColumnIndexRow: true);

var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
var storageLogger = loggerFactory.CreateLogger<FormStorage>();
var dataLogger = loggerFactory.CreateLogger<FormDataStorage>();

var env = new SimpleHostEnvironment(webRoot);
var options = Options.Create(new StorageOptions { StorageRoot = "storage" });

var parser = new ExcelFormParser();
var storage = new FormStorage(options, env, storageLogger);
var dataStorage = new FormDataStorage(storage, dataLogger);

// Seed schemas (idempotent)
await SeedSchemaAsync(storage, parser, demo1TemplatePath, CancellationToken.None);
await SeedSchemaAsync(storage, parser, demo2TemplatePath, CancellationToken.None);

Console.WriteLine($"Template file: {demo1TemplatePath}");
Console.WriteLine($"Template file: {demo2TemplatePath}");

if (mode == "data")
{
    var expectedFormNumber = "001"; // ExcelFormParser normalizes to first number
    var rawFormNumber = demo1Raw;
    var formTitle = demo1Title;
    var filledPath = Path.Combine(outDir, $"{rawFormNumber}-filled.xlsx");

    // Create a filled file with same header and a few data rows.
    using (var wb = new XLWorkbook())
    {
        var ws = wb.AddWorksheet("Form");

        ws.Cell(1, 1).Value = rawFormNumber;
        ws.Cell(2, 1).Value = formTitle;

        ws.Cell(3, 1).Value = "Group A";
        ws.Range(3, 1, 3, 2).Merge();

        ws.Cell(3, 3).Value = "Group B";
        ws.Range(3, 3, 3, 5).Merge();

        ws.Cell(4, 1).Value = "A1";
        ws.Cell(4, 2).Value = "A2";
        ws.Cell(4, 3).Value = "B1";
        ws.Cell(4, 4).Value = "B2";
        ws.Cell(4, 5).Value = "B3";

        // Data rows start at row 5
        ws.Cell(5, 1).Value = "alpha";
        ws.Cell(5, 2).Value = "10";
        ws.Cell(5, 3).Value = "x";
        ws.Cell(5, 4).Value = "2026-01-22";
        ws.Cell(5, 5).Value = "true";

        ws.Cell(6, 1).Value = "beta";
        ws.Cell(6, 2).Value = "20";
        ws.Cell(6, 3).Value = "y";
        ws.Cell(6, 4).Value = "2026-01-23";
        ws.Cell(6, 5).Value = "false";

        ws.Cell(7, 1).Value = "gamma";
        ws.Cell(7, 2).Value = "30";
        ws.Cell(7, 3).Value = "z";
        ws.Cell(7, 4).Value = "2026-01-24";
        ws.Cell(7, 5).Value = "true";

        ws.Columns(1, 5).AdjustToContents();

        wb.SaveAs(filledPath);
    }

    Console.WriteLine($"Filled file: {filledPath}");

    // Store as a data upload
    await using (var dataFs = File.OpenRead(filledPath))
    {
        var dataFile = new FormFile(dataFs, 0, dataFs.Length, "Upload", Path.GetFileName(filledPath))
        {
            Headers = new HeaderDictionary(),
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        };

        var saved = await dataStorage.SaveAsync(dataFile, parser, CancellationToken.None, expectedFormNumber: expectedFormNumber);
        Console.WriteLine($"Seeded data upload: #{saved.FormNumber} v{saved.Version} uploadId={saved.UploadId} rows={saved.RowCount}");
        Console.WriteLine($"UI: http://localhost:5148/forms/{saved.FormNumber}/data");
        Console.WriteLine($"UI: http://localhost:5148/forms/{saved.FormNumber}/data/v{saved.Version}/{saved.UploadId}");
    }
}

Console.WriteLine($"Stored schemas under: {Path.Combine(webRoot, "storage", "forms")}");
Console.WriteLine($"Stored data under: {Path.Combine(webRoot, "storage", "data")}");

return 0;

sealed class SimpleHostEnvironment : IWebHostEnvironment
{
    public SimpleHostEnvironment(string contentRootPath)
    {
        ContentRootPath = contentRootPath;
        ContentRootFileProvider = new Microsoft.Extensions.FileProviders.PhysicalFileProvider(contentRootPath);

        WebRootPath = Path.Combine(contentRootPath, "wwwroot");
        WebRootFileProvider = new Microsoft.Extensions.FileProviders.PhysicalFileProvider(WebRootPath);

        EnvironmentName = Environments.Development;
        ApplicationName = "SeedForm";
    }

    public string ApplicationName { get; set; }
    public Microsoft.Extensions.FileProviders.IFileProvider ContentRootFileProvider { get; set; }
    public string ContentRootPath { get; set; }
    public string EnvironmentName { get; set; }
    public Microsoft.Extensions.FileProviders.IFileProvider WebRootFileProvider { get; set; }
    public string WebRootPath { get; set; }
}
