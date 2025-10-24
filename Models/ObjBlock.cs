using System.Text.Json.Serialization;

namespace RoomExporter.Models;

public sealed class ObjBlock
{
    [JsonPropertyName("name")] public string? Name { get; set; }

    [JsonPropertyName("objects")] public List<ObjRoom>? Objects { get; set; }
}