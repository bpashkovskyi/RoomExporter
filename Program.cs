using OfficeOpenXml;

using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

public class Program()
{
    // ===== Hardcoded settings =====
    const string BeginDateStr = "01.09.2025";      // dd.MM.yyyy
    const string EndDateStr = "31.12.2025";      // dd.MM.yyyy
    const string OutputPath = "rooms-load.xlsx"; // where to save the Excel
    static int? MaxRooms = null;              // e.g., 50 for testing; null = all rooms

    const string BaseUrl = "https://dekanat.nung.edu.ua/cgi-bin/timetable_export.cgi";
    const int MaxParallel = 8; // be polite to the server

    public static async Task Main()
    {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); // for Windows-1251

        // Workdays in range
        var begin = DateTime.ParseExact(BeginDateStr, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        var end = DateTime.ParseExact(EndDateStr, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        var workdays = EnumerateDays(begin, end).Where(IsWorkday).ToArray();
        int totalWorkdays = workdays.Length;
        if (totalWorkdays == 0)
        {
            Console.WriteLine("No workdays in the given range. Exiting.");
            return;
        }

        Console.WriteLine($"Date range: {begin:yyyy-MM-dd} .. {end:yyyy-MM-dd} | Workdays: {totalWorkdays}");

        var http = new HttpClient(new HttpClientHandler
        {
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        });
        http.DefaultRequestHeaders.UserAgent.ParseAdd("RoomLoadExporter/1.0");

        // 1) Fetch all rooms
        var rooms = await FetchRoomsAsync(http);
        if (MaxRooms.HasValue) rooms = rooms.Take(MaxRooms.Value).ToList();
        Console.WriteLine($"Rooms to process: {rooms.Count}");

        // 2) For each room, fetch schedule + compute percentages
        var throttler = new SemaphoreSlim(MaxParallel);
        var results = new List<RowResult>(rooms.Count);
        var tasks = rooms.Select(async room =>
        {
            await throttler.WaitAsync();
            try
            {
                var schedule = await FetchRoomScheduleAsync(http, room.Id, BeginDateStr, EndDateStr);
                var percentages = ComputeLessonPercentages(schedule, workdays, totalWorkdays);
                lock (results) results.Add(new RowResult(room.Id, room.Name, percentages));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[WARN] Room {room.Id} '{room.Name}': {ex.Message}");
                lock (results) results.Add(new RowResult(room.Id, room.Name, new Dictionary<int, double>())); // zeros
            }
            finally
            {
                throttler.Release();
            }
        });
        await Task.WhenAll(tasks);

        // 3) Write Excel
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        using (var pkg = new ExcelPackage())
        {
            var ws = pkg.Workbook.Worksheets.Add("Load");
            int col = 1;
            ws.Cells[1, col++].Value = "Room ID";
            ws.Cells[1, col++].Value = "Room Name";
            for (int l = 1; l <= 8; l++) ws.Cells[1, col++].Value = l.ToString();

            int row = 2;
            foreach (var r in results.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase))
            {
                col = 1;
                ws.Cells[row, col++].Value = r.Id;
                ws.Cells[row, col++].Value = r.Name;

                for (int l = 1; l <= 8; l++)
                {
                    r.Percentages.TryGetValue(l, out var pct);
                    ws.Cells[row, col++].Value = pct;
                    ws.Cells[row, col - 1].Style.Numberformat.Format = "0.0%";
                }
                row++;
            }

            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            await pkg.SaveAsAsync(new FileInfo(OutputPath));
        }

        Console.WriteLine($"✅ Done. Excel saved to: {OutputPath}");

        // ----------------- Helpers -----------------

        static IEnumerable<DateTime> EnumerateDays(DateTime start, DateTime endInclusive)
        {
            for (var d = start.Date; d <= endInclusive.Date; d = d.AddDays(1)) yield return d;
        }

        static bool IsWorkday(DateTime d) => d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday;

        static async Task<List<Room>> FetchRoomsAsync(HttpClient http)
        {
            // req_type=obj_list&req_mode=room&req_format=json&coding_mode=WINDOWS-1251&bs=ok
            var uri = $"{BaseUrl}?req_type=obj_list&req_mode=room&show_ID=yes&req_format=json&coding_mode=WINDOWS-1251&bs=ok";

            var bytes = await http.GetByteArrayAsync(uri);
            var json = Encoding.GetEncoding(1251).GetString(bytes);

            var root = JsonSerializer.Deserialize<ObjListRoot>(json, JsonOptions());
            if (root?.PsExport?.Blocks == null) return new List<Room>();

            var rooms = new List<Room>();
            foreach (var block in root.PsExport.Blocks)
            {
                if (block.Objects == null) continue;
                foreach (var o in block.Objects)
                {
                    if (string.IsNullOrWhiteSpace(o.ID) || string.IsNullOrWhiteSpace(o.Name)) continue;
                    rooms.Add(new Room(o.ID.Trim(), o.Name.Trim()));
                }
            }
            return rooms;
        }

        static async Task<List<RozItem>> FetchRoomScheduleAsync(HttpClient http, string roomId, string beginDate, string endDate)
        {
            // req_type=rozklad&req_mode=room&OBJ_ID={id}&begin_date=dd.MM.yyyy&end_date=dd.MM.yyyy&req_format=json&coding_mode=WINDOWS-1251&bs=ok
            var qp = $"req_type=rozklad&req_mode=room&OBJ_ID={WebUtility.UrlEncode(roomId)}&OBJ_name=&dep_name=&ros_text=united&begin_date={beginDate}&end_date={endDate}&req_format=json&coding_mode=WINDOWS-1251&bs=ok";
            var uri = $"{BaseUrl}?{qp}";

            var bytes = await http.GetByteArrayAsync(uri);
            var json = Encoding.GetEncoding(1251).GetString(bytes);

            var root = JsonSerializer.Deserialize<RozkladRoot>(json, JsonOptions());
            var items = root?.PsExport?.RozItems ?? new();
            foreach (var i in items)
            {
                i.ParsedDate = DateTime.TryParseExact(i.Date, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
                    ? d.Date : (DateTime?)null;
                i.Lesson = int.TryParse(i.LessonNumber, out var ln) ? ln : (int?)null;
                i.IsBusy = !string.IsNullOrWhiteSpace(i.LessonDescription);
            }
            return items;
        }

        static Dictionary<int, double> ComputeLessonPercentages(List<RozItem> items, DateTime[] workdays, int totalWorkdays)
        {
            var result = new Dictionary<int, double>();
            if (totalWorkdays == 0) { for (int l = 1; l <= 8; l++) result[l] = 0; return result; }

            for (int lesson = 1; lesson <= 8; lesson++)
            {
                var busyDates = new HashSet<DateTime>();
                foreach (var i in items)
                {
                    if (i.ParsedDate is null || i.Lesson is null) continue;
                    if (i.Lesson == lesson && i.IsBusy && IsWorkday(i.ParsedDate.Value))
                        busyDates.Add(i.ParsedDate.Value.Date);
                }

                var count = busyDates.Count(d => workdays.Contains(d));
                var pct = (double)count / totalWorkdays;
                result[lesson] = pct;
            }
            return result;
        }

        static JsonSerializerOptions JsonOptions() => new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        };

        // ------------- Models for obj_list -------------

    }


    public sealed class ObjListRoot
    {
        [JsonPropertyName("psrozklad_export")]
        public PsObjList? PsExport { get; set; }
    }
    public sealed class PsObjList
    {
        [JsonPropertyName("blocks")]
        public List<ObjBlock>? Blocks { get; set; }
    }
    public sealed class ObjBlock
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("objects")]
        public List<ObjRoom>? Objects { get; set; }
    }
    public sealed class ObjRoom
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("ID")]
        public string? ID { get; set; }
    }

    public sealed record Room(string Id, string Name);

    // ------------- Models for rozklad -------------
    public sealed class RozkladRoot
    {
        [JsonPropertyName("psrozklad_export")]
        public PsRozklad? PsExport { get; set; }
    }
    public sealed class PsRozklad
    {
        [JsonPropertyName("roz_items")]
        public List<RozItem> RozItems { get; set; } = new();
    }
    public sealed class RozItem
    {
        [JsonPropertyName("object")] public string? Object { get; set; }
        [JsonPropertyName("date")] public string? Date { get; set; }
        [JsonPropertyName("comment")] public string? Comment { get; set; }
        [JsonPropertyName("lesson_number")] public string? LessonNumber { get; set; }
        [JsonPropertyName("lesson_name")] public string? LessonName { get; set; }
        [JsonPropertyName("lesson_time")] public string? LessonTime { get; set; }
        [JsonPropertyName("lesson_description")] public string? LessonDescription { get; set; }

        // computed:
        public DateTime? ParsedDate { get; set; }
        public int? Lesson { get; set; }
        public bool IsBusy { get; set; }
    }

    public sealed record RowResult(string Id, string Name, Dictionary<int, double> Percentages);

}