using System.Text.Json.Serialization;

namespace RoomExporter.Models;

public sealed class ObjRoom
{
    [JsonPropertyName("name")] public string? Name { get; set; }

    [JsonPropertyName("ID")] public string? ID { get; set; }
}