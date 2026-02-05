using System.Text;

namespace SamorodinkaTech.FormStructures.Web.Services;

public static class CsvUtil
{
    public static byte[] BuildUtf8CsvWithBom(
        IReadOnlyList<string> headers,
        IReadOnlyList<IReadOnlyList<string?>> rows,
        char separator = ',')
    {
        var encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);

        using var ms = new MemoryStream();
        using (var writer = new StreamWriter(ms, encoding, bufferSize: 16 * 1024, leaveOpen: true))
        {
            writer.NewLine = "\r\n";

            writer.WriteLine(string.Join(separator, headers.Select(h => Escape(h, separator))));

            foreach (var row in rows)
            {
                var cells = row.Select(v => Escape(v, separator));
                writer.WriteLine(string.Join(separator, cells));
            }
        }

        return ms.ToArray();
    }

    public static string Escape(string? value, char separator = ',')
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        var mustQuote = false;
        for (var i = 0; i < value.Length; i++)
        {
            var ch = value[i];
            if (ch == '"' || ch == '\r' || ch == '\n' || ch == separator)
            {
                mustQuote = true;
                break;
            }
        }

        if (!mustQuote)
        {
            return value;
        }

        var sb = new StringBuilder(value.Length + 2);
        sb.Append('"');
        foreach (var ch in value)
        {
            if (ch == '"')
            {
                sb.Append("\"\"");
            }
            else
            {
                sb.Append(ch);
            }
        }
        sb.Append('"');
        return sb.ToString();
    }
}
