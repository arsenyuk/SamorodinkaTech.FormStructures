using System.Text;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using SamorodinkaTech.FormStructures.Web.Models;
using SamorodinkaTech.FormStructures.Web.Services;
using Xunit;

namespace SamorodinkaTech.FormStructures.Tests;

public sealed class FormDataFormulaValidationTests
{
    [Fact]
    public async Task SaveAsync_RejectsUpload_WhenFormulasDifferWithinSameColumn()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), "FormStructuresTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);

        try
        {
            var parser = new ExcelFormParser();

            // 1) Create and store a matching schema (v1) so FormDataStorage can validate structure hash.
            var schemaBytes = BuildSchemaTemplateBytes();
            var schemaStructure = parser.Parse(new MemoryStream(schemaBytes, writable: false), sourceFileName: "schema.xlsx") with
            {
                Version = 1,
                UploadedAtUtc = DateTime.UtcNow
            };

            var storageRoot = Path.Combine(tempRoot, "storage");
            var versionDir = Path.Combine(storageRoot, "forms", schemaStructure.FormNumber, "v1");
            Directory.CreateDirectory(versionDir);

            File.WriteAllText(
                Path.Combine(versionDir, "structure.json"),
                System.Text.Json.JsonSerializer.Serialize(schemaStructure, JsonUtil.StableOptions),
                Encoding.UTF8);

            File.WriteAllBytes(Path.Combine(versionDir, "original.xlsx"), schemaBytes);

            var env = new TestHostEnvironment { ContentRootPath = tempRoot };
            var formStorage = new FormStorage(
                Options.Create(new StorageOptions { StorageRoot = "storage" }),
                env,
                parser,
                NullLogger<FormStorage>.Instance);

            var dataStorage = new FormDataStorage(formStorage, NullLogger<FormDataStorage>.Instance);

            // 2) Upload a data file with inconsistent formulas.
            await using var uploadStream = LoadXlsxFromBase64Fixture("FORMULA-INCONSISTENT-001.xlsx");
            var file = ToFormFile(uploadStream, fileName: "FORMULA-INCONSISTENT-001.xlsx");

            var ex = await Assert.ThrowsAsync<FormParseException>(() =>
                dataStorage.SaveAsync(file, parser, CancellationToken.None));

            Assert.Contains("Upload rejected", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("inconsistent", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Sum", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            try { Directory.Delete(tempRoot, recursive: true); } catch { /* best-effort */ }
        }
    }

    private static byte[] BuildSchemaTemplateBytes()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Form");

        ws.Cell(1, 1).Value = "FORMULA-010";
        ws.Cell(2, 1).Value = "Formula validation: inconsistent formulas";

        ws.Cell(3, 1).Value = "A";
        ws.Cell(3, 2).Value = "B";
        ws.Cell(3, 3).Value = "Sum";

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private static IFormFile ToFormFile(MemoryStream ms, string fileName)
    {
        ms.Position = 0;
        return new FormFile(ms, 0, ms.Length, name: "file", fileName: fileName)
        {
            Headers = new HeaderDictionary(),
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        };
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
