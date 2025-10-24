using System.Text;

namespace RoomExporter;

public static class Program
{
    public sealed class Settings
    {
        // ===== Hardcoded settings =====
        public string BeginDateStr { get; init; } = "01.09.2025"; // dd.MM.yyyy
        public string EndDateStr { get; init; } = "31.12.2025"; // dd.MM.yyyy
        public string OutputPath { get; init; } = "rooms-load.xlsx";
        public int? MaxRooms { get; init; } = null; // e.g., 50 for testing; null = all rooms

        public string BaseUrl { get; init; } = "https://dekanat.nung.edu.ua/cgi-bin/timetable_export.cgi";
        public int MaxParallel { get; init; } = 8; // be polite to the server

        public static Settings Default => new();
    }
    
    public static async Task Main()
    {
        // Enable Windows-1251
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Initialize
        var settings = Settings.Default;
        var (begin, end) = DateUtils.ParseRange(settings.BeginDateStr, settings.EndDateStr);
        var workdays = DateUtils.Workdays(begin, end).ToArray();
        if (workdays.Length == 0)
        {
            Console.WriteLine("No workdays in the given range. Exiting.");
            return;
        }

        Console.WriteLine($"Date range: {begin:yyyy-MM-dd} .. {end:yyyy-MM-dd} | Workdays: {workdays.Length}");

        using var http = HttpClientFactory.Create();
        var api = new NungApiClient(http, settings.BaseUrl);

        // 1) Fetch all rooms
        var rooms = await api.FetchRoomsAsync();
        if (settings.MaxRooms.HasValue)
        {
            rooms = rooms.Take(settings.MaxRooms.Value).ToList();
        }

        Console.WriteLine($"Rooms to process: {rooms.Count}");

        // 2) For each room, fetch schedule + compute percentages
        var results = await Processor.BuildResultsAsync(
            api,
            rooms,
            settings.BeginDateStr,
            settings.EndDateStr,
            workdays,
            maxParallel: settings.MaxParallel
        );

        // 3) Write Excel
        await ExcelWriter.WriteAsync(settings.OutputPath, results);

        Console.WriteLine($"✅ Done. Excel saved to: {settings.OutputPath}");
    }
}