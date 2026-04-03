using EmailParser.Helpers;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmailParser.Services;

/// <summary>
/// Writes an Excel report listing file paths that exceed a configurable
/// character-length threshold, with each directory segment in its own column.
/// </summary>
public class LongPathReportService
{
    private static readonly ILogger Log = Serilog.Log.ForContext<LongPathReportService>();

    /// <summary>
    /// Writes an Excel workbook to <paramref name="outputPath"/> containing one
    /// row per path, with columns for the full path, its length, each directory
    /// segment, and the file name.
    /// </summary>
    public void WriteLongPathReport(IReadOnlyList<string> longPaths, string outputPath)
    {
        Log.Information("Writing long-path report with {Count} entries to {Path}",
            longPaths.Count, outputPath);

        var rows = longPaths.Select(p => new
        {
            FullPath    = p,
            Length      = p.Length,
            DirSegments = (Path.GetDirectoryName(p) ?? string.Empty)
                              .Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
                                     StringSplitOptions.RemoveEmptyEntries),
            FileName    = Path.GetFileName(p),
        }).ToArray();

        int maxDirSegments = rows.Max(r => r.DirSegments.Length);
        int totalCols      = 2 + maxDirSegments + 1;
        int totalRows      = rows.Length + 1;

        var data = new object[totalRows, totalCols];

        data[0, 0] = "Full Path";
        data[0, 1] = "Path Length";
        for (int i = 0; i < maxDirSegments; i++)
            data[0, 2 + i] = $"Folder Level {i + 1}";
        data[0, 2 + maxDirSegments] = "File Name";

        for (int i = 0; i < rows.Length; i++)
        {
            var r = rows[i];
            data[i + 1, 0] = r.FullPath;
            data[i + 1, 1] = r.Length;
            for (int s = 0; s < r.DirSegments.Length; s++)
                data[i + 1, 2 + s] = r.DirSegments[s];
            data[i + 1, 2 + maxDirSegments] = r.FileName;
        }

        Excel.Application? app       = null;
        Excel.Workbooks?   workbooks = null;
        Excel.Workbook?    workbook  = null;
        Excel.Sheets?      sheets    = null;
        Excel.Worksheet?   worksheet = null;
        Excel.Range?       dataRange = null;
        Excel.Range?       headerRow = null;
        Excel.Range?       topLeft   = null;
        Excel.Range?       botRight  = null;

        try
        {
            app       = new Excel.Application { Visible = false, DisplayAlerts = false };
            workbooks = app.Workbooks;
            workbook  = workbooks.Add();
            sheets    = workbook.Worksheets;
            worksheet = (Excel.Worksheet)sheets[1];
            worksheet.Name = "Long Paths";

            topLeft  = (Excel.Range)worksheet.Cells[1, 1];
            botRight = (Excel.Range)worksheet.Cells[totalRows, totalCols];
            dataRange = worksheet.Range[topLeft, botRight];
            dataRange.Value2 = data;

            headerRow = (Excel.Range)worksheet.Rows[1];
            headerRow.Font.Bold = true;
            dataRange.Columns.AutoFit();

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
            workbook.SaveAs(outputPath, Excel.XlFileFormat.xlOpenXMLWorkbook);

            Log.Information("Long-path report saved to {Path}", outputPath);
        }
        finally
        {
            workbook?.Close(SaveChanges: false);
            app?.Quit();

            ComHelper.ReleaseComObject(headerRow);
            ComHelper.ReleaseComObject(dataRange);
            ComHelper.ReleaseComObject(topLeft);
            ComHelper.ReleaseComObject(botRight);
            ComHelper.ReleaseComObject(worksheet);
            ComHelper.ReleaseComObject(sheets);
            ComHelper.ReleaseComObject(workbook);
            ComHelper.ReleaseComObject(workbooks);
            ComHelper.ReleaseComObject(app);
        }
    }
}
