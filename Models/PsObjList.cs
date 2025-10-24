using System.Text.Json.Serialization;

namespace RoomExporter.Models;

public sealed class PsObjList
{
    [JsonPropertyName("blocks")] public List<ObjBlock>? Blocks { get; set; }
}