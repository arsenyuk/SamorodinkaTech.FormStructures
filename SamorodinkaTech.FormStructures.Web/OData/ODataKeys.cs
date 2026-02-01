using System.Text;

namespace SamorodinkaTech.FormStructures.Web.OData;

public static class ODataKeys
{
    public static string UploadKey(string formNumber, int version, string uploadId)
    {
        return $"u:{B64(formNumber)}:{version}:{B64(uploadId)}";
    }

    public static bool TryParseUploadKey(string key, out string formNumber, out int version, out string uploadId)
    {
        formNumber = string.Empty;
        uploadId = string.Empty;
        version = 0;

        if (string.IsNullOrWhiteSpace(key))
        {
            return false;
        }

        var parts = key.Split(':', StringSplitOptions.None);
        if (parts.Length != 4)
        {
            return false;
        }

        if (!string.Equals(parts[0], "u", StringComparison.Ordinal) || parts[1].Length == 0)
        {
            return false;
        }

        if (!TryUnb64(parts[1], out formNumber))
        {
            return false;
        }

        if (!int.TryParse(parts[2], out version) || version <= 0)
        {
            return false;
        }

        if (!TryUnb64(parts[3], out uploadId))
        {
            return false;
        }

        return !string.IsNullOrWhiteSpace(formNumber) && !string.IsNullOrWhiteSpace(uploadId);
    }

    public static string RowKey(string uploadKey, int rowNumber)
    {
        return $"r:{uploadKey}:{rowNumber}";
    }

    public static bool TryParseRowKey(string key, out string uploadKey, out int rowNumber)
    {
        uploadKey = string.Empty;
        rowNumber = 0;

        if (string.IsNullOrWhiteSpace(key))
        {
            return false;
        }

        if (!key.StartsWith("r:u:", StringComparison.Ordinal))
        {
            return false;
        }

        var lastColon = key.LastIndexOf(':');
        if (lastColon <= 0 || lastColon >= key.Length - 1)
        {
            return false;
        }

        uploadKey = key.Substring(2, lastColon - 2);

        var rowPart = key[(lastColon + 1)..];
        return int.TryParse(rowPart, out rowNumber) && rowNumber > 0;
    }

    public static string ColumnKey(string uploadKey, int index)
    {
        return $"c:{uploadKey}:{index}";
    }

    private static string B64(string value)
    {
        var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
        var s = Convert.ToBase64String(bytes);
        return s.TrimEnd('=').Replace('+', '-').Replace('/', '_');
    }

    private static bool TryUnb64(string value, out string decoded)
    {
        decoded = string.Empty;

        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var s = value.Replace('-', '+').Replace('_', '/');
        switch (s.Length % 4)
        {
            case 2: s += "=="; break;
            case 3: s += "="; break;
        }

        try
        {
            var bytes = Convert.FromBase64String(s);
            decoded = Encoding.UTF8.GetString(bytes);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
