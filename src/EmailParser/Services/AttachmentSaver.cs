using System.IO.Compression;
using EmailParser.Helpers;
using EmailParser.Models;
using Serilog;

namespace EmailParser.Services;

/// <summary>
/// Saves email attachments to an output folder in their original format.
/// ZIP attachments are extracted into a named subfolder.
/// </summary>
public class AttachmentSaver
{
    private static readonly ILogger Log = Serilog.Log.ForContext<AttachmentSaver>();

    /// <summary>
    /// Copies every attachment in <paramref name="email"/> to
    /// <paramref name="attachmentsDir"/> in its original format.
    /// ZIP attachments are extracted and flattened directly into the folder.
    /// </summary>
    public void SaveAttachmentsToFolder(EmailData email, string attachmentsDir)
    {
        Directory.CreateDirectory(attachmentsDir);

        foreach (AttachmentData attachment in email.Attachments)
        {
            if (!File.Exists(attachment.TempFilePath))
            {
                Log.Warning("Temp file missing for attachment {FileName}, skipping",
                    attachment.FileName);
                continue;
            }

            string ext = Path.GetExtension(attachment.FileName);

            if (string.Equals(ext, ".zip", StringComparison.OrdinalIgnoreCase))
            {
                Log.Debug("Extracting ZIP attachment {FileName} to {Directory}",
                    attachment.FileName, attachmentsDir);
                ExtractZipToFolder(attachment.TempFilePath, attachmentsDir);
            }
            else
            {
                string safeFileName = FileNameHelper.SanitizeFileName(attachment.FileName);
                if (string.IsNullOrWhiteSpace(safeFileName))
                    safeFileName = "attachment" + ext;

                string destPath = FileNameHelper.GetUniqueFilePath(attachmentsDir, safeFileName);
                File.Copy(attachment.TempFilePath, destPath, overwrite: false);
                Log.Debug("Saved attachment {FileName} to {Path}", attachment.FileName, destPath);
            }
        }
    }

    // Private helpers

    /// <summary>
    /// Extracts a ZIP archive into <paramref name="extractDir"/> using a
    /// zip-slip-safe strategy, flattening all entries into a single folder.
    /// </summary>
    private static void ExtractZipToFolder(string zipPath, string extractDir)
    {
        Directory.CreateDirectory(extractDir);

        try
        {
            using var archive = ZipFile.OpenRead(zipPath);
            foreach (var entry in archive.Entries)
            {
                if (string.IsNullOrEmpty(entry.Name))
                    continue;

                string safeFileName = FileNameHelper.SanitizeFileName(entry.Name);
                if (string.IsNullOrWhiteSpace(safeFileName))
                    safeFileName = "file";

                string destPath = FileNameHelper.GetUniqueFilePath(extractDir, safeFileName);
                entry.ExtractToFile(destPath, overwrite: false);
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Could not extract ZIP {ZipPath}", zipPath);
        }
    }
}
