using System.Runtime.InteropServices;
using EmailParser.Helpers;
using EmailParser.Models;
using Serilog;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmailParser.Services;

/// <summary>
/// Loads term patterns from an Excel data-dictionary file. The patterns are
/// used to strip unwanted terms from output file and folder names.
/// </summary>
public class DataDictionaryService
{
    private static readonly ILogger Log = Serilog.Log.ForContext<DataDictionaryService>();

    /// <summary>
    /// Finds the most-recently modified Excel file in <paramref name="dictionaryDir"/>
    /// and loads all unique cell values as patterns, ordered longest-first.
    /// </summary>
    public DataDictionary LoadLatestDataDictionary(string dictionaryDir)
    {
        Directory.CreateDirectory(dictionaryDir);

        string? latestExcelFile = Directory
            .EnumerateFiles(dictionaryDir)
            .Where(path =>
            {
                string ext = Path.GetExtension(path);
                return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                    || ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
                    || ext.Equals(".xls", StringComparison.OrdinalIgnoreCase);
            })
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .FirstOrDefault();

        if (latestExcelFile is null)
        {
            Log.Information("No Excel data-dictionary file found in {Directory}", dictionaryDir);
            return new DataDictionary(dictionaryDir, null, Array.Empty<string>());
        }

        Log.Information("Loading data dictionary from {FilePath}", latestExcelFile);
        var patterns = LoadPatternsFromExcel(latestExcelFile);
        Log.Information("Loaded {Count} dictionary patterns", patterns.Count);

        return new DataDictionary(dictionaryDir, latestExcelFile, patterns);
    }

    // Private helpers

    private static IReadOnlyList<string> LoadPatternsFromExcel(string excelPath)
    {
        Excel.Application? app = null;
        Excel.Workbooks? workbooks = null;
        Excel.Workbook? workbook = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? worksheet = null;
        Excel.Range? usedRange = null;

        try
        {
            app = new Excel.Application { Visible = false, DisplayAlerts = false };
            workbooks = app.Workbooks;
            workbook = workbooks.Open(excelPath, ReadOnly: true);

            sheets = workbook.Worksheets;
            worksheet = (Excel.Worksheet)sheets[1];
            usedRange = worksheet.UsedRange;

            var values = usedRange.Value2;
            var patterns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (values is object[,] matrix)
            {
                foreach (object? value in matrix)
                {
                    string candidate = value?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(candidate))
                        patterns.Add(candidate);
                }
            }
            else
            {
                string candidate = values?.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(candidate))
                    patterns.Add(candidate);
            }

            return patterns
                .OrderByDescending(static p => p.Length)
                .ToArray();
        }
        finally
        {
            if (workbook is not null)
                workbook.Close(SaveChanges: false);
            if (app is not null)
                app.Quit();

            ComHelper.ReleaseComObject(usedRange);
            ComHelper.ReleaseComObject(worksheet);
            ComHelper.ReleaseComObject(sheets);
            ComHelper.ReleaseComObject(workbook);
            ComHelper.ReleaseComObject(workbooks);
            ComHelper.ReleaseComObject(app);
        }
    }
}
