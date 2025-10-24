using System.Globalization;

namespace RoomExporter;

public static class DateUtils
{
    public static (DateTime Begin, DateTime End) ParseRange(string beginStr, string endStr)
    {
        var begin = DateTime.ParseExact(beginStr, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        var end = DateTime.ParseExact(endStr, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        return (begin, end);
    }

    public static IEnumerable<DateTime> Workdays(DateTime begin, DateTime endInclusive) =>
        EnumerateDays(begin, endInclusive).Where(IsWorkday);

    private static IEnumerable<DateTime> EnumerateDays(DateTime start, DateTime endInclusive)
    {
        for (var d = start.Date; d <= endInclusive.Date; d = d.AddDays(1))
        {
            yield return d;
        }
    }

    public static bool IsWorkday(DateTime d) =>
        d.DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday);
}