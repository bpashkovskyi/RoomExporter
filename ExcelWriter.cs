using OfficeOpenXml;
using RoomExporter.Models;

namespace RoomExporter;

public static class ExcelWriter
{
    public static async Task WriteAsync(string path, IEnumerable<RowResult> rows)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Load");

        var col = 1;
        ws.Cells[1, col++].Value = "Room ID";
        ws.Cells[1, col++].Value = "Room Name";
        for (var l = 1; l <= 8; l++)
        {
            ws.Cells[1, col++].Value = l.ToString();
        }

        var row = 2;
        foreach (var r in rows.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase))
        {
            col = 1;
            ws.Cells[row, col++].Value = r.Id;
            ws.Cells[row, col++].Value = r.Name;

            for (var l = 1; l <= 8; l++)
            {
                r.Percentages.TryGetValue(l, out var pct);
                ws.Cells[row, col++].Value = pct;
                ws.Cells[row, col - 1].Style.Numberformat.Format = "0.0%";
            }

            row++;
        }

        ws.Cells[ws.Dimension.Address].AutoFitColumns();
        await pkg.SaveAsAsync(new FileInfo(path)).ConfigureAwait(false);
    }
}