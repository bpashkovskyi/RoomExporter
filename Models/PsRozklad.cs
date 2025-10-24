using System.Text.Json.Serialization;

namespace RoomExporter.Models;

public sealed class PsRozklad
{
    [JsonPropertyName("roz_items")] public List<RozItem> RozItems { get; set; } = new();
}