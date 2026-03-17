using System.IO.Compression;
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

            // Track used file names per output directory to handle duplicate subjects.
            var usedNames = new Dictionary<string, HashSet<string>>(
                StringComparer.OrdinalIgnoreCase);

            foreach (var email in emails)
            {
                string safeSubject = SanitizeFileName(email.Subject);
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
                        emailOutputDir = Path.Combine(outputDir, relativeDir);
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

    // -------------------------------------------------------------------------
    // Attachment saving helpers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Copies every attachment in <paramref name="email"/> to
    /// <paramref name="attachmentsDir"/> in its original format.
    /// ZIP attachments are extracted into a subfolder named after the archive.
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
                // Extract the ZIP into a subfolder named after the archive.
                string zipFolderName = SanitizeFileName(
                    Path.GetFileNameWithoutExtension(attachment.FileName));
                if (string.IsNullOrWhiteSpace(zipFolderName))
                    zipFolderName = "archive";

                string zipExtractDir = GetUniqueDirectoryPath(attachmentsDir, zipFolderName);
                ExtractZipToFolder(attachment.TempFilePath, zipExtractDir);
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
    /// zip-slip safe extraction strategy.
    /// </summary>
    private static void ExtractZipToFolder(string zipPath, string extractDir)
    {
        Directory.CreateDirectory(extractDir);
        string canonicalExtractDir = Path.GetFullPath(extractDir)
            + Path.DirectorySeparatorChar;

        try
        {
            using var archive = ZipFile.OpenRead(zipPath);
            foreach (var entry in archive.Entries)
            {
                // Skip directory-only entries.
                if (string.IsNullOrEmpty(entry.Name))
                    continue;

                string destPath = Path.GetFullPath(
                    Path.Combine(extractDir, entry.FullName));

                // Zip-slip guard: skip entries that escape the target directory.
                if (!destPath.StartsWith(canonicalExtractDir,
                        StringComparison.OrdinalIgnoreCase))
                {
                    Console.Error.WriteLine(
                        $"  Warning: Skipping ZIP entry with unsafe path: '{entry.FullName}'");
                    continue;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(destPath) ?? extractDir);
                entry.ExtractToFile(destPath, overwrite: true);
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

