namespace RoomExporter.Models;

public sealed record RowResult(string Id, string Name, Dictionary<int, double> Percentages);