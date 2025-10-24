using RoomExporter.Models;

namespace RoomExporter;

public static class Processor
{
    public static async Task<List<RowResult>> BuildResultsAsync(
        NungApiClient api,
        List<Room> rooms,
        string beginDateStr,
        string endDateStr,
        IReadOnlyCollection<DateTime> workdays,
        int maxParallel)
    {
        var results = new List<RowResult>(rooms.Count);
        using var throttler = new SemaphoreSlim(maxParallel);

        var tasks = rooms.Select(async room =>
        {
            await throttler.WaitAsync().ConfigureAwait(false);
            try
            {
                var schedule = await api.FetchRoomScheduleAsync(room.Id, beginDateStr, endDateStr).ConfigureAwait(false);
                var percentages = ComputeLessonPercentages(schedule, workdays);
                lock (results)
                    results.Add(new RowResult(room.Id, room.Name, percentages));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[WARN] Room {room.Id} '{room.Name}': {ex.Message}");
                lock (results)
                    results.Add(new RowResult(room.Id, room.Name, new Dictionary<int, double>()));
            }
            finally
            {
                throttler.Release();
            }
        });

        await Task.WhenAll(tasks).ConfigureAwait(false);
        return results;
    }

    private static Dictionary<int, double> ComputeLessonPercentages(
        List<RozItem> items,
        IReadOnlyCollection<DateTime> workdays)
    {
        var totalWorkdays = workdays.Count;
        var result = new Dictionary<int, double>(capacity: 8);

        if (totalWorkdays == 0)
        {
            for (var l = 1; l <= 8; l++)
            {
                result[l] = 0d;
            }

            return result;
        }

        // Pre-index workdays for O(1) lookups
        var workdaySet = new HashSet<DateTime>(workdays);

        for (var lesson = 1; lesson <= 8; lesson++)
        {
            var busyDates = new HashSet<DateTime>();
            foreach (var i in items)
            {
                if (i.ParsedDate is null || i.Lesson is null)
                {
                    continue;
                }

                if (i.Lesson == lesson && i.IsBusy && DateUtils.IsWorkday(i.ParsedDate.Value))
                {
                    busyDates.Add(i.ParsedDate.Value.Date);
                }
            }

            var count = busyDates.Count(d => workdaySet.Contains(d));
            result[lesson] = (double)count / totalWorkdays;
        }

        return result;
    }
}