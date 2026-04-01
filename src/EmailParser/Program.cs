using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using EmailParser.Models;
using EmailParser.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmailParser;

class Program
{
    static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        Console.WriteLine("Email Parser — Save Outlook Emails to PDF");
        Console.WriteLine("==========================================");
        Console.WriteLine();

        // =======================================================================
        // PATH CONFIGURATION
        // Change these three lines to move files to a different location.
        // =======================================================================
        string outputBaseDir     = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "EmailParser");
        string reportsDir        = @"C:\Users\JessicaAnyanwu\OneDrive - Suir Engineering Ltd\Documents\EmailParser\Email Parcer Directory";
        string dataDictionaryDir = reportsDir;  // data dictionary lives in the same folder as reports
        // =======================================================================

        // -----------------------------------------------------------------------
        // 1. Resolve the source: an Outlook folder name or a local directory of
        //    .msg files.
        // -----------------------------------------------------------------------
        string folderPath;

        if (args.Length > 0)
        {
            folderPath = args[0].Trim();
            Console.WriteLine($"Source (from argument): {folderPath}");
        }
        else
        {
            Console.Write(
                "Enter an Outlook folder name (e.g. 'Inbox' or 'Inbox/Projects'),\n" +
                "or a path to a local directory containing .msg files: ");
            folderPath = Console.ReadLine()?.Trim() ?? string.Empty;
        }

        if (string.IsNullOrWhiteSpace(folderPath))
        {
            Console.Error.WriteLine("Error: Source cannot be empty.");
            Environment.Exit(1);
        }

        // -----------------------------------------------------------------------
        // 2. Determine whether the input is a local directory of .msg files
        //    (Office-free mode) or an Outlook folder name (requires Outlook).
        // -----------------------------------------------------------------------
        bool isMsgDirectory = Directory.Exists(folderPath);

        // -----------------------------------------------------------------------
        // 3. Load the latest Excel data dictionary for terms to strip.
        // -----------------------------------------------------------------------
        IReadOnlyList<string> dictionaryPatterns;
        try
        {
            var dictionary = LoadLatestDataDictionary(dataDictionaryDir);
            dictionaryPatterns = dictionary.Patterns;

            if (dictionary.SourcePath is null)
            {
                Console.WriteLine($"Data dictionary: No Excel file found in '{dictionary.DirectoryPath}'.");
                Console.WriteLine("                 No terms will be stripped from names.");
            }
            else
            {
                Console.WriteLine($"Data dictionary: {dictionary.SourcePath}");
                Console.WriteLine($"Terms loaded   : {dictionaryPatterns.Count}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error loading Excel data dictionary: {ex.Message}");
            Environment.Exit(1);
            return;
        }

        // -----------------------------------------------------------------------
        // 4. Prepare output directory:  My Documents\EmailParser\<name>
        // -----------------------------------------------------------------------
        string outputSubDir = isMsgDirectory
            ? SanitizePath(
                Path.GetFileName(Path.TrimEndingDirectorySeparator(folderPath))
                ?? "Messages")
            : SanitizePath(folderPath);

        outputSubDir = StripDictionaryTerms(outputSubDir, dictionaryPatterns);
        if (string.IsNullOrWhiteSpace(outputSubDir))
            outputSubDir = "Messages";

        string outputDir = Path.Combine(outputBaseDir, outputSubDir);
        Directory.CreateDirectory(outputDir);

        if (isMsgDirectory)
            Console.WriteLine($"Mode:            Reading .msg files from '{folderPath}'");
        else
            Console.WriteLine($"Mode:            Reading from Outlook folder '{folderPath}'");
        Console.WriteLine($"Output directory : {outputDir}");
        Console.WriteLine();

        // -----------------------------------------------------------------------
        // 5. Fetch emails and convert each one to PDF.
        // -----------------------------------------------------------------------
        try
        {
            MsgFileService? msgService = null;
            IEnumerable<EmailData> emails = isMsgDirectory
                ? (msgService = new MsgFileService()).GetEmailsFromDirectory(folderPath)
                : new OutlookService().GetEmailsFromFolder(folderPath);

            var pdfService = new PdfService();

            int processed = 0;
            int failed = 0;

            // Track used file names per output directory to handle duplicate subjects.
            var usedNames = new Dictionary<string, HashSet<string>>(
                StringComparer.OrdinalIgnoreCase);

            foreach (var email in emails)
            {
                string safeSubject = SanitizeFileName(email.Subject);
                safeSubject = StripDictionaryTerms(safeSubject, dictionaryPatterns);
                if (string.IsNullOrWhiteSpace(safeSubject))
                    safeSubject = "No Subject";

                // Replicate the source subdirectory structure inside outputDir.
                string emailOutputDir = outputDir;
                if (isMsgDirectory && !string.IsNullOrEmpty(email.SourceFilePath))
                {
                    string sourceFileDir = Path.GetDirectoryName(email.SourceFilePath)
                                          ?? folderPath;
                    string relativeDir = Path.GetRelativePath(folderPath, sourceFileDir);
                    if (!string.IsNullOrEmpty(relativeDir) && relativeDir != ".")
                    {
                        relativeDir = StripDictionaryTerms(relativeDir, dictionaryPatterns);
                        emailOutputDir = Path.Combine(outputDir, relativeDir);
                    }
                }

                Directory.CreateDirectory(emailOutputDir);

                // Append a counter if the same subject already appeared in this directory.
                if (!usedNames.TryGetValue(emailOutputDir, out var dirUsedNames))
                {
                    dirUsedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    usedNames[emailOutputDir] = dirUsedNames;
                }

                string fileName = safeSubject;
                int counter = 2;
                while (!dirUsedNames.Add(fileName))
                    fileName = $"{safeSubject} ({counter++})";

                string outputPath = Path.Combine(emailOutputDir, fileName + ".pdf");

                // Save original attachments before the PDF service consumes (and then
                // deletes) the temp files.  Attachments go into a folder named
                // "<subject> attachments" (e.g. "re: window enquiry attachments") so
                // each email thread gets its own clearly-labelled attachment folder.
                if (email.Attachments.Count > 0)
                {
                    string attachmentsDir = Path.Combine(emailOutputDir, fileName + " attachments");
                    try
                    {
                        SaveAttachmentsToFolder(email, attachmentsDir);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(
                            $"  Warning: Could not save attachments for '{email.Subject}': {ex.Message}");
                    }
                }

                //Console.Write($"  Processing: {email.Subject} ... ");

                try
                {
                    pdfService.SaveEmailAsPdf(email, outputPath);
                    //Console.WriteLine($"saved → {outputPath}");
                    processed++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("FAILED");
                    Console.Error.WriteLine($"    Error: {ex.Message}");
                    failed++;
                }
            }

            if (msgService?.LongPaths.Count > 0)
            {
                string reportPath = Path.Combine(reportsDir,
                    $"FilePathReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                Console.WriteLine(
                    $"Writing file path report ({msgService.LongPaths.Count} paths over 250 chars)...");
                try
                {
                    WriteLongPathReport(msgService.LongPaths, reportPath);
                    Console.WriteLine($"File path report : {reportPath}");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Warning: Could not write long path report: {ex.Message}");
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Done.  Processed: {processed}  |  Failed: {failed}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(OfficeAvailability.IsOfficeUnavailableException(ex)
                ? $"Fatal error: {ex.Message}\n\n" +
                  "This tool requires Microsoft Office to be installed and configured.\n" +
                  "Please install Microsoft Office (including Outlook) and try again.\n\n" +
                  "Tip: If Outlook is unavailable, export your emails as .msg files and\n" +
                  "run the tool with the path to that directory instead, e.g.:\n" +
                  "  EmailParser.exe \"C:\\Users\\you\\exports\\inbox\""
                : $"Fatal error: {ex.Message}");
            Environment.Exit(1);
        }
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    // Build once; Path.GetInvalidFileNameChars() returns the same values every call.
    private static readonly HashSet<char> InvalidFileNameChars =
        new(Path.GetInvalidFileNameChars());

    private sealed record DataDictionary(string DirectoryPath, string? SourcePath, IReadOnlyList<string> Patterns);

    private static DataDictionary LoadLatestDataDictionary(string dictionaryDir)
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
            return new DataDictionary(dictionaryDir, null, Array.Empty<string>());

        var patterns = LoadPatternsFromExcel(latestExcelFile);
        return new DataDictionary(dictionaryDir, latestExcelFile, patterns);
    }

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

            ReleaseComObject(usedRange);
            ReleaseComObject(worksheet);
            ReleaseComObject(sheets);
            ReleaseComObject(workbook);
            ReleaseComObject(workbooks);
            ReleaseComObject(app);
        }
    }

    private static void ReleaseComObject(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
            Marshal.ReleaseComObject(comObject);
    }

    /// <summary>
    /// Writes an Excel report listing every file path that exceeded 250 characters.
    /// Each row shows the full path, its length, and each folder segment in its own
    /// column so the depth at which paths grow long is immediately visible.
    /// </summary>
    private static void WriteLongPathReport(IReadOnlyList<string> longPaths, string outputPath)
    {
        // Split each path into directory segments + filename.
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
        int totalCols      = 2 + maxDirSegments + 1;  // FullPath + Length + dir levels + FileName
        int totalRows      = rows.Length + 1;          // header + data

        // Build a 2-D object array for a single bulk write — far faster than
        // setting cells one at a time through the COM boundary.
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

            // Bold the header row and auto-fit all columns.
            headerRow = (Excel.Range)worksheet.Rows[1];
            headerRow.Font.Bold = true;
            dataRange.Columns.AutoFit();

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
            workbook.SaveAs(outputPath, Excel.XlFileFormat.xlOpenXMLWorkbook);
        }
        finally
        {
            workbook?.Close(SaveChanges: false);
            app?.Quit();

            ReleaseComObject(headerRow);
            ReleaseComObject(dataRange);
            ReleaseComObject(topLeft);
            ReleaseComObject(botRight);
            ReleaseComObject(worksheet);
            ReleaseComObject(sheets);
            ReleaseComObject(workbook);
            ReleaseComObject(workbooks);
            ReleaseComObject(app);
        }
    }

    /// <summary>
    /// Removes every dictionary term from anywhere within a file or folder name,
    /// case-insensitively. After removal, consecutive spaces are collapsed and
    /// any leading separators are trimmed.
    /// Note: this method is intentionally NOT called for attachment file names.
    /// </summary>
    private static string StripDictionaryTerms(string? text, IReadOnlyList<string> patterns)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        string result = text.Trim();

        foreach (string pattern in patterns)
        {
            int idx;
            while ((idx = result.IndexOf(pattern, StringComparison.OrdinalIgnoreCase)) >= 0)
                result = result.Remove(idx, pattern.Length);
        }

        // Collapse any runs of whitespace left behind by the removals.
        result = string.Join(" ", result.Split(' ', StringSplitOptions.RemoveEmptyEntries));

        // Remove any stray leading separators or spaces.
        result = result.TrimStart(' ', '-', '_').Trim();

        return result;
    }

    /// <summary>
    /// Removes characters that are invalid in file names.
    /// </summary>
    private static string SanitizeFileName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "No Subject";

        string result = string.Concat(name.Select(c => InvalidFileNameChars.Contains(c) ? '_' : c));
        return result.Trim().TrimEnd('.');
    }

    /// <summary>
    /// Converts a folder path (e.g. "Inbox/Projects" or an absolute file-system
    /// path) into a safe relative path for use as a directory name.
    /// For absolute paths the generic OS prefix is stripped first:
    ///   - Paths under the current user profile lose the C:\Users\&lt;username&gt; prefix.
    ///   - All other rooted paths lose the drive root (e.g. "C:\").
    /// </summary>
    private static string SanitizePath(string folderPath)
    {
        // Strip generic OS prefix from absolute paths so that segments like
        // "C:", "Users" and the username never appear in the output folder name.
        if (Path.IsPathRooted(folderPath))
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string fullPath = Path.GetFullPath(folderPath);

            if (fullPath.StartsWith(userProfile, StringComparison.OrdinalIgnoreCase))
            {
                // e.g. C:\Users\Jessica\Documents\emails → Documents\emails
                folderPath = Path.GetRelativePath(userProfile, fullPath);
            }
            else
            {
                // e.g. D:\ProjectFiles\emails → ProjectFiles\emails
                string? root = Path.GetPathRoot(fullPath);
                if (!string.IsNullOrEmpty(root))
                    folderPath = fullPath.Substring(root.Length);
            }
        }

        // Replace forward and back slashes with the OS directory separator,
        // then sanitize each segment individually.
        string[] segments = folderPath.Split(new[] { '/', '\\' },
            StringSplitOptions.RemoveEmptyEntries);

        string[] safeSegments = segments.Select(seg =>
            string.Concat(seg.Select(c => InvalidFileNameChars.Contains(c) ? '_' : c)).Trim()
        ).ToArray();

        return Path.Combine(safeSegments);
    }

    // -------------------------------------------------------------------------
    // Attachment saving helpers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Copies every attachment in <paramref name="email"/> to
    /// <paramref name="attachmentsDir"/> in its original format.
    /// ZIP attachments are extracted directly into <paramref name="attachmentsDir"/>
    /// with no folder structure preserved.
    /// </summary>
    private static void SaveAttachmentsToFolder(EmailData email, string attachmentsDir)
    {
        Directory.CreateDirectory(attachmentsDir);

        foreach (AttachmentData attachment in email.Attachments)
        {
            if (!File.Exists(attachment.TempFilePath))
                continue;

            string ext = Path.GetExtension(attachment.FileName);

            if (string.Equals(ext, ".zip", StringComparison.OrdinalIgnoreCase))
            {
                // Extract ZIP contents directly into the attachments folder.
                ExtractZipToFolder(attachment.TempFilePath, attachmentsDir);
            }
            else
            {
                string safeFileName = SanitizeFileName(attachment.FileName);
                if (string.IsNullOrWhiteSpace(safeFileName))
                    safeFileName = "attachment" + ext;

                string destPath = GetUniqueFilePath(attachmentsDir, safeFileName);
                File.Copy(attachment.TempFilePath, destPath, overwrite: false);
            }
        }
    }

    /// <summary>
    /// Extracts a ZIP archive to <paramref name="extractDir"/> using a
    /// zip-slip safe strategy and flattens all files into a single folder.
    /// </summary>
    private static void ExtractZipToFolder(string zipPath, string extractDir)
    {
        Directory.CreateDirectory(extractDir);

        try
        {
            using var archive = ZipFile.OpenRead(zipPath);
            foreach (var entry in archive.Entries)
            {
                // Skip directory-only entries.
                if (string.IsNullOrEmpty(entry.Name))
                    continue;

                // Flatten: ignore any subfolder path in the ZIP entry.
                string safeFileName = SanitizeFileName(entry.Name);
                if (string.IsNullOrWhiteSpace(safeFileName))
                    safeFileName = "file";

                string destPath = GetUniqueFilePath(extractDir, safeFileName);
                entry.ExtractToFile(destPath, overwrite: false);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(
                $"  Warning: Could not extract ZIP '{zipPath}': {ex.Message}");
        }
    }

    /// <summary>
    /// Returns a path for a new file inside <paramref name="directory"/>.
    /// If <paramref name="fileName"/> already exists, appends a numeric
    /// counter (e.g. "report (2).docx") until a free name is found.
    /// </summary>
    private static string GetUniqueFilePath(string directory, string fileName)
    {
        string dest = Path.Combine(directory, fileName);
        if (!File.Exists(dest))
            return dest;

        string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
        string ext = Path.GetExtension(fileName);
        int counter = 2;
        do
        {
            dest = Path.Combine(directory, $"{nameWithoutExt} ({counter++}){ext}");
        }
        while (File.Exists(dest));

        return dest;
    }

    /// <summary>
    /// Returns a path for a new subdirectory inside <paramref name="parent"/>.
    /// If the name already exists, appends a numeric counter until free.
    /// </summary>
    private static string GetUniqueDirectoryPath(string parent, string name)
    {
        string dest = Path.Combine(parent, name);
        if (!Directory.Exists(dest))
            return dest;

        int counter = 2;
        string candidate;
        do
        {
            candidate = Path.Combine(parent, $"{name} ({counter++})");
        }
        while (Directory.Exists(candidate));

        return candidate;
    }
}

