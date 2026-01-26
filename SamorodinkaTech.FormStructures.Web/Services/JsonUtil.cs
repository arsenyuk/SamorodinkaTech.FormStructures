using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SamorodinkaTech.FormStructures.Web.Services;

public static class JsonUtil
{
    public static readonly JsonSerializerOptions StableOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        Converters =
        {
            new JsonStringEnumConverter(JsonNamingPolicy.CamelCase)
        }
    };

    public static string ToStableJson<T>(T value)
    {
        return JsonSerializer.Serialize(value, StableOptions);
    }
}
