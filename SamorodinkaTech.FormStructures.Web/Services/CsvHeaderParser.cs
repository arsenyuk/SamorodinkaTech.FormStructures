using System.Text;
using SamorodinkaTech.FormStructures.Web.Models;

namespace SamorodinkaTech.FormStructures.Web.Services;

public static class CsvHeaderParser
{
    private static readonly char[] CandidateSeparators = [',', ';', '\t'];

    public const int MaxColumns = 2000;

    public sealed record ParseResult(IReadOnlyList<string> Headers, char Separator);

    public static ParseResult ParseHeaderRow(Stream csvStream)
    {
        if (csvStream is null)
        {
            throw new ArgumentNullException(nameof(csvStream));
        }

        if (!csvStream.CanRead)
        {
            throw new FormParseException("CSV stream is not readable.");
        }

        // StreamReader will auto-detect UTF BOMs (UTF-8/UTF-16/UTF-32).
        using var reader = new StreamReader(csvStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: true);

        string? line;
        do
        {
            line = reader.ReadLine();
            if (line is null)
            {
                throw new FormParseException("CSV file is empty: no header row found.");
            }

            line = line.TrimEnd();
        }
        while (line.Length == 0);

        // Choose the separator that yields the most fields when parsed.
        // This is a pragmatic approach for real-world CSVs (comma vs semicolon vs tab).
        (char sep, List<string> fields)? best = null;
        foreach (var sep in CandidateSeparators)
        {
            var fields = ParseFields(line, sep);
            if (best is null || fields.Count > best.Value.fields.Count)
            {
                best = (sep, fields);
            }
        }

        if (best is null)
        {
            throw new FormParseException("Failed to parse CSV header row.");
        }

        var headers = best.Value.fields
            .Select((h, i) =>
            {
                var x = (h ?? string.Empty).Trim();
                if (i == 0)
                {
                    x = x.TrimStart('\uFEFF');
                }

                return x;
            })
            .ToArray();

        ValidateHeaders(headers);

        return new ParseResult(headers, best.Value.sep);
    }

    private static void ValidateHeaders(string[] headers)
    {
        if (headers.Length == 0)
        {
            throw new FormParseException("CSV header row has no columns.");
        }

        if (headers.Length > MaxColumns)
        {
            throw new FormParseException($"CSV header row has too many columns ({headers.Length}). Max is {MaxColumns}.");
        }

        for (var i = 0; i < headers.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(headers[i]))
            {
                throw new FormParseException($"CSV header column #{i + 1} is empty.");
            }
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var duplicates = new List<string>();
        foreach (var h in headers)
        {
            if (!seen.Add(h))
            {
                duplicates.Add(h);
            }
        }

        if (duplicates.Count > 0)
        {
            var uniq = duplicates.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            throw new FormParseException($"CSV header row contains duplicate column names: {string.Join(", ", uniq)}.");
        }
    }

    private static List<string> ParseFields(string line, char separator)
    {
        var result = new List<string>();

        var sb = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var c = line[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    // Escaped quote: ""
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                        continue;
                    }

                    inQuotes = false;
                    continue;
                }

                sb.Append(c);
                continue;
            }

            if (c == '"')
            {
                inQuotes = true;
                continue;
            }

            if (c == separator)
            {
                result.Add(sb.ToString());
                sb.Clear();
                continue;
            }

            sb.Append(c);
        }

        result.Add(sb.ToString());

        return result;
    }
}
