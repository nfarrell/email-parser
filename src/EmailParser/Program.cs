using EmailParser.Models;
using EmailParser.Services;

namespace EmailParser;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Email Parser — Save Outlook Emails to PDF");
        Console.WriteLine("==========================================");
        Console.WriteLine();

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
        // 3. Prepare output directory:  My Documents\EmailParser\<name>
        // -----------------------------------------------------------------------
        string outputSubDir = isMsgDirectory
            ? SanitizePath(
                Path.GetFileName(Path.TrimEndingDirectorySeparator(folderPath))
                ?? "Messages")
            : SanitizePath(folderPath);

        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string outputDir = Path.Combine(documentsPath, "EmailParser", outputSubDir);
        Directory.CreateDirectory(outputDir);

        if (isMsgDirectory)
            Console.WriteLine($"Mode:            Reading .msg files from '{folderPath}'");
        else
            Console.WriteLine($"Mode:            Reading from Outlook folder '{folderPath}'");
        Console.WriteLine($"Output directory : {outputDir}");
        Console.WriteLine();

        // -----------------------------------------------------------------------
        // 4. Fetch emails and convert each one to PDF.
        // -----------------------------------------------------------------------
        try
        {
            IEnumerable<EmailData> emails = isMsgDirectory
                ? new MsgFileService().GetEmailsFromDirectory(folderPath)
                : new OutlookService().GetEmailsFromFolder(folderPath);

            var pdfService = new PdfService();

            int processed = 0;
            int failed = 0;

            // Track used file names to handle duplicate subjects.
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var email in emails)
            {
                string safeSubject = SanitizeFileName(email.Subject);
                if (string.IsNullOrWhiteSpace(safeSubject))
                    safeSubject = "No Subject";

                // Append a counter if the same subject already appeared.
                string fileName = safeSubject;
                int counter = 1;
                while (!usedNames.Add(fileName))
                    fileName = $"{safeSubject} ({counter++})";

                string outputPath = Path.Combine(outputDir, fileName + ".pdf");

                Console.Write($"  Processing: {email.Subject} ... ");

                try
                {
                    pdfService.SaveEmailAsPdf(email, outputPath);
                    Console.WriteLine($"saved → {outputPath}");
                    processed++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("FAILED");
                    Console.Error.WriteLine($"    Error: {ex.Message}");
                    failed++;
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
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    // Build once; Path.GetInvalidFileNameChars() returns the same values every call.
    private static readonly HashSet<char> InvalidFileNameChars =
        new(Path.GetInvalidFileNameChars());

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
    /// Converts an Outlook folder path (e.g. "Inbox/Projects") into a safe
    /// relative path for use as a directory name, replacing path separators
    /// with OS-appropriate directory separators.
    /// </summary>
    private static string SanitizePath(string folderPath)
    {
        // Replace forward and back slashes with the OS directory separator,
        // then sanitize each segment individually.
        string[] segments = folderPath.Split(new[] { '/', '\\' },
            StringSplitOptions.RemoveEmptyEntries);

        string[] safeSegments = segments.Select(seg =>
            string.Concat(seg.Select(c => InvalidFileNameChars.Contains(c) ? '_' : c)).Trim()
        ).ToArray();

        return Path.Combine(safeSegments);
    }
}

