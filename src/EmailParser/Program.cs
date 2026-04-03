using System.Text;
using EmailParser.Helpers;
using EmailParser.Models;
using EmailParser.Services;
using Serilog;

namespace EmailParser;

class Program
{
    static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // PATH CONFIGURATION — change these lines to move files to a different location.
        string outputBaseDir     = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "EmailParser");
        string reportsDir        = @"C:\Users\JessicaAnyanwu\OneDrive - Suir Engineering Ltd\Documents\EmailParser\Email Parcer Directory";
        string dataDictionaryDir = reportsDir;  // data dictionary lives in the same folder as reports

        // Configure Serilog: write to both console and a rolling log file.
        string logDir = Path.Combine(outputBaseDir, "Logs");
        Directory.CreateDirectory(logDir);

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.Console(
                outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
            .WriteTo.File(
                Path.Combine(logDir, "EmailParser-.log"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 30,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {SourceContext}{NewLine}  {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        try
        {
            Run(args, outputBaseDir, reportsDir, dataDictionaryDir);
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Unhandled exception — application terminating");
            Environment.Exit(1);
        }
        finally
        {
            Log.CloseAndFlush();
        }

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    private static void Run(
        string[] args,
        string outputBaseDir,
        string reportsDir,
        string dataDictionaryDir)
    {
        Log.Information("Email Parser — Save Outlook Emails to PDF");
        Log.Information("==========================================");

        // 1. Resolve the source: an Outlook folder name or a local directory of .msg files.
        string folderPath = ResolveSource(args);

        // 2. True = local .msg files (Office-free); false = Outlook folder (requires Outlook).
        bool isMsgDirectory = Directory.Exists(folderPath);

        // 3. Load the latest Excel data dictionary for terms to strip from file names.
        IReadOnlyList<string> dictionaryPatterns;
        try
        {
            var dictionaryService = new DataDictionaryService();
            var dictionary = dictionaryService.LoadLatestDataDictionary(dataDictionaryDir);
            dictionaryPatterns = dictionary.Patterns;

            if (dictionary.SourcePath is null)
                Log.Information("Data dictionary: no file found — no terms will be stripped");
            else
                Log.Information("Data dictionary: {Path} ({Count} terms)",
                    dictionary.SourcePath, dictionaryPatterns.Count);
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Failed to load Excel data dictionary");
            Environment.Exit(1);
            return;
        }

        // 4. Prepare output directory: My Documents\EmailParser\<name>
        string outputSubDir = isMsgDirectory
            ? FileNameHelper.SanitizePath(
                Path.GetFileName(Path.TrimEndingDirectorySeparator(folderPath))
                ?? "Messages")
            : FileNameHelper.SanitizePath(folderPath);

        outputSubDir = FileNameHelper.StripDictionaryTerms(outputSubDir, dictionaryPatterns);
        if (string.IsNullOrWhiteSpace(outputSubDir))
            outputSubDir = "Messages";

        string outputDir = Path.Combine(outputBaseDir, outputSubDir);
        Directory.CreateDirectory(outputDir);

        if (isMsgDirectory)
            Log.Information("Mode: reading .msg files from {Source}", folderPath);
        else
            Log.Information("Mode: reading from Outlook folder {Source}", folderPath);
        Log.Information("Output directory: {OutputDir}", outputDir);

        // 5. Fetch emails and convert each one to PDF.
        try
        {
            ProcessEmails(
                folderPath, isMsgDirectory, outputDir, reportsDir,
                dictionaryPatterns);
        }
        catch (Exception ex) when (OfficeAvailability.IsOfficeUnavailableException(ex))
        {
            Log.Fatal(ex,
                "Microsoft Office is not installed or configured. " +
                "Install Outlook and try again, or supply a directory of .msg files instead");
            Environment.Exit(1);
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Fatal error during email processing");
            Environment.Exit(1);
        }
    }

    /// <summary>
    /// Resolves the email source from command-line arguments or an interactive prompt.
    /// </summary>
    private static string ResolveSource(string[] args)
    {
        string folderPath;

        if (args.Length > 0)
        {
            folderPath = args[0].Trim();
            Log.Information("Source (from argument): {Source}", folderPath);
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
            Log.Fatal("Source cannot be empty");
            Environment.Exit(1);
        }

        return folderPath;
    }

    /// <summary>
    /// Core processing loop: fetches emails, saves attachments, and converts to PDF.
    /// </summary>
    private static void ProcessEmails(
        string folderPath,
        bool isMsgDirectory,
        string outputDir,
        string reportsDir,
        IReadOnlyList<string> dictionaryPatterns)
    {
        MsgFileService? msgService = null;
        IEnumerable<EmailData> emails = isMsgDirectory
            ? (msgService = new MsgFileService()).GetEmailsFromDirectory(folderPath)
            : new OutlookService().GetEmailsFromFolder(folderPath);

        var pdfService      = new PdfService();
        var attachmentSaver = new AttachmentSaver();

        int processed = 0;
        int failed    = 0;

        // Track used file names per output directory to handle duplicate subjects.
        var usedNames = new Dictionary<string, HashSet<string>>(
            StringComparer.OrdinalIgnoreCase);

        foreach (var email in emails)
        {
            string safeSubject = FileNameHelper.SanitizeFileName(email.Subject);
            safeSubject = FileNameHelper.StripDictionaryTerms(safeSubject, dictionaryPatterns);
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
                    relativeDir = FileNameHelper.StripDictionaryTerms(
                        relativeDir, dictionaryPatterns);
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
            // deletes) the temp files.
            if (email.Attachments.Count > 0)
            {
                string attachmentsDir = Path.Combine(emailOutputDir, fileName + " attachments");
                try
                {
                    attachmentSaver.SaveAttachmentsToFolder(email, attachmentsDir);
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "Could not save attachments for '{Subject}'", email.Subject);
                }
            }

            Log.Information("Processing: {Subject}", email.Subject);

            try
            {
                pdfService.SaveEmailAsPdf(email, outputPath);
                Log.Information("Saved -> {OutputPath}", outputPath);
                processed++;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to convert email '{Subject}' to PDF", email.Subject);
                failed++;
            }
        }

        // Write a report of any overly-long file paths encountered.
        if (msgService?.LongPaths.Count > 0)
        {
            var reportService = new LongPathReportService();
            string reportPath = Path.Combine(reportsDir,
                $"FilePathReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            Log.Information(
                "Writing file-path report ({Count} paths over 250 chars)",
                msgService.LongPaths.Count);

            try
            {
                reportService.WriteLongPathReport(msgService.LongPaths, reportPath);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Could not write long-path report");
            }
        }

        Log.Information("Done. Processed: {Processed} | Failed: {Failed}", processed, failed);
    }
}
