using System.Text.Json.Serialization;

namespace RoomExporter.Models;

public sealed class RozkladRoot
{
    [JsonPropertyName("psrozklad_export")] public PsRozklad? PsExport { get; set; }
}