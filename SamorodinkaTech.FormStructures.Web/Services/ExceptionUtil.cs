using System.Collections.Generic;
using System.Text;

namespace SamorodinkaTech.FormStructures.Web.Services;

public static class ExceptionUtil
{
    public static string FormatExceptionChain(Exception ex)
    {
        if (ex is null)
        {
            throw new ArgumentNullException(nameof(ex));
        }

        var sb = new StringBuilder();
        var seen = new HashSet<Exception>(ReferenceEqualityComparer.Instance);
        AppendException(sb, ex, seen, indentLevel: 0, prefix: null);
        return sb.ToString();
    }

    private static void AppendException(
        StringBuilder sb,
        Exception ex,
        HashSet<Exception> seen,
        int indentLevel,
        string? prefix)
    {
        var indent = new string(' ', indentLevel * 2);

        if (!seen.Add(ex))
        {
            sb.Append(indent);
            if (!string.IsNullOrWhiteSpace(prefix))
            {
                sb.Append(prefix);
            }

            sb.AppendLine("(cycle detected)");
            return;
        }

        sb.Append(indent);
        if (!string.IsNullOrWhiteSpace(prefix))
        {
            sb.Append(prefix);
        }

        var typeName = ex.GetType().Name;
        var message = ex.Message?.Trim();
        sb.AppendLine(string.IsNullOrWhiteSpace(message) ? typeName : $"{typeName}: {message}");

        if (ex is AggregateException aggregateException)
        {
            var flattened = aggregateException.Flatten();
            for (var i = 0; i < flattened.InnerExceptions.Count; i++)
            {
                var inner = flattened.InnerExceptions[i];
                AppendException(sb, inner, seen, indentLevel + 1, prefix: $"[{i}] ");
            }

            return;
        }

        if (ex.InnerException is not null)
        {
            AppendException(sb, ex.InnerException, seen, indentLevel + 1, prefix: "Inner: ");
        }
    }
}
