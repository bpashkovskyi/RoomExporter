using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using RoomExporter.Models;

namespace RoomExporter;

public sealed class NungApiClient
{
    private readonly HttpClient _http;
    private readonly string _baseUrl;
    private static readonly Encoding Cp1251 = Encoding.GetEncoding(1251);

    public NungApiClient(HttpClient http, string baseUrl)
    {
        _http = http;
        _baseUrl = baseUrl.TrimEnd('/');
    }

    public async Task<List<Room>> FetchRoomsAsync()
    {
        // req_type=obj_list&req_mode=room&req_format=json&coding_mode=WINDOWS-1251&bs=ok
        var uri =
            $"{_baseUrl}?req_type=obj_list&req_mode=room&show_ID=yes&req_format=json&coding_mode=WINDOWS-1251&bs=ok";

        var json = await GetJsonAsync(uri);
        var root = JsonSerializer.Deserialize<ObjListRoot>(json, Json.Options);
        if (root?.PsExport?.Blocks is null)
        {
            return new();
        }

        var rooms = new List<Room>();
        foreach (var block in root.PsExport.Blocks)
        {
            if (block.Objects is null)
            {
                continue;
            }

            foreach (var o in block.Objects)
            {
                if (string.IsNullOrWhiteSpace(o.ID) || string.IsNullOrWhiteSpace(o.Name))
                {
                    continue;
                }

                rooms.Add(new Room(o.ID.Trim(), o.Name.Trim()));
            }
        }

        return rooms;
    }

    public async Task<List<RozItem>> FetchRoomScheduleAsync(string roomId, string beginDate, string endDate)
    {
        // req_type=rozklad&req_mode=room&OBJ_ID={id}&begin_date=dd.MM.yyyy&end_date=dd.MM.yyyy&req_format=json&coding_mode=WINDOWS-1251&bs=ok
        var qp =
            $"req_type=rozklad&req_mode=room&OBJ_ID={WebUtility.UrlEncode(roomId)}&OBJ_name=&dep_name=&ros_text=united&begin_date={beginDate}&end_date={endDate}&req_format=json&coding_mode=WINDOWS-1251&bs=ok";
        var uri = $"{_baseUrl}?{qp}";

        var json = await GetJsonAsync(uri);
        var root = JsonSerializer.Deserialize<RozkladRoot>(json, Json.Options);

        var items = root?.PsExport?.RozItems ?? new();
        foreach (var i in items)
        {
            i.ParsedDate = DateTime.TryParseExact(i.Date, "dd.MM.yyyy", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var d)
                ? d.Date
                : null;

            i.Lesson = int.TryParse(i.LessonNumber, out var ln) ? ln : null;
            i.IsBusy = !string.IsNullOrWhiteSpace(i.LessonDescription);
        }

        return items;
    }

    private async Task<string> GetJsonAsync(string uri)
    {
        var bytes = await _http.GetByteArrayAsync(uri).ConfigureAwait(false);
        return Cp1251.GetString(bytes);
    }
}