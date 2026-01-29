using System.Text;

using SamorodinkaTech.FormStructures.Web.Models;

namespace SamorodinkaTech.FormStructures.Web.Services;

public static class DownloadFileName
{
    private const int MaxPartLength = 80;

    public static string ForSchema(FormStructure structure, int version)
    {
        var baseName = ForForm(structure);
        return $"{baseName}-v{version}.xlsx";
    }

    public static string ForDataUpload(FormStructure structure, int version, string uploadId)
    {
        var baseName = ForForm(structure);
        var safeUploadId = ToSafePart(uploadId, fallback: "upload");
        return $"{baseName}-v{version}-{safeUploadId}.xlsx";
    }

    public static string ForAggregated(FormStructure structure, int version)
    {
        var baseName = ForForm(structure);
        return $"{baseName}-v{version}-aggregated.xlsx";
    }

    public static string ForForm(FormStructure structure)
    {
        var safeNumber = ToSafePart(structure.FormNumber, fallback: "form");
        var safeTitle = ToSafePart(structure.FormTitle, fallback: string.Empty);

        return string.IsNullOrWhiteSpace(safeTitle)
            ? safeNumber
            : $"{safeNumber}-{safeTitle}";
    }

    public static string ToSafePart(string? value, string fallback)
    {
        var s = (value ?? string.Empty).Trim();
        if (s.Length == 0)
        {
            return fallback;
        }

        var sb = new StringBuilder(s.Length);
        var invalid = Path.GetInvalidFileNameChars();

        foreach (var ch in s)
        {
            sb.Append(invalid.Contains(ch) ? '_' : ch);
        }

        var result = sb.ToString().Trim();
        if (result.Length > MaxPartLength)
        {
            result = result[..MaxPartLength].Trim();
        }

        return result.Length == 0 ? fallback : result;
    }
}
