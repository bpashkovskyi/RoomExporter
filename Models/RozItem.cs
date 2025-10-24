using System.Text.Json.Serialization;

namespace RoomExporter.Models;

public sealed class RozItem
{
    [JsonPropertyName("object")] public string? Object { get; set; }
    [JsonPropertyName("date")] public string? Date { get; set; }
    [JsonPropertyName("comment")] public string? Comment { get; set; }
    [JsonPropertyName("lesson_number")] public string? LessonNumber { get; set; }
    [JsonPropertyName("lesson_name")] public string? LessonName { get; set; }
    [JsonPropertyName("lesson_time")] public string? LessonTime { get; set; }

    [JsonPropertyName("lesson_description")]
    public string? LessonDescription { get; set; }

    // computed:
    public DateTime? ParsedDate { get; set; }
    public int? Lesson { get; set; }
    public bool IsBusy { get; set; }
}