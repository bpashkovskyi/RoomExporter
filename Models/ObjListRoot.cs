using System.Text.Json.Serialization;

namespace RoomExporter.Models;

public sealed class ObjListRoot
{
    [JsonPropertyName("psrozklad_export")] public PsObjList? PsExport { get; set; }
}