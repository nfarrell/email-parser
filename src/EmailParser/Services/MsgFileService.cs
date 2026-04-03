using EmailParser.Models;
using MsgReader.Outlook;

namespace EmailParser.Services;

/// <summary>
/// Reads emails from Outlook .msg files stored in a local directory, without
/// requiring Microsoft Outlook to be installed.
/// </summary>
public class MsgFileService
{
    private readonly List<string> _longPaths = [];

    /// <summary>
    /// File paths longer than 250 characters discovered during
    /// <see cref="GetEmailsFromDirectory"/>. Written to an Excel report by the caller.
    /// </summary>
    public IReadOnlyList<string> LongPaths => _longPaths;

    /// <summary>
    /// Returns all emails parsed from .msg files found inside
    /// <paramref name="directoryPath"/> and its subdirectories. Files that
    /// cannot be parsed are skipped with a warning.
    /// </summary>
    /// <param name="directoryPath">
    /// Path to an existing directory that contains one or more .msg files.
    /// </param>
    public IEnumerable<EmailData> GetEmailsFromDirectory(string directoryPath)
    {
        string[] msgFiles = Directory.GetFiles(
            directoryPath, "*.msg", SearchOption.AllDirectories);

        if (msgFiles.Length == 0)
        {
            Console.Error.WriteLine(
                $"Warning: No .msg files found in '{directoryPath}'.");
        }

        foreach (string msgFile in msgFiles)
        {
            if (msgFile.Length > 250)
                _longPaths.Add(msgFile);

            EmailData email;
            try
            {
                email = ReadMsgFile(msgFile);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(
                    $"Warning: Could not read '{msgFile}': [{ex.GetType().Name}] {ex.Message}");
                if (ex.InnerException is not null)
                    Console.Error.WriteLine(
                        $"         Inner: [{ex.InnerException.GetType().Name}] {ex.InnerException.Message}");
                continue;
            }

            email.SourceFilePath = msgFile;
            yield return email;
        }
    }

    // Private helpers

    private static EmailData ReadMsgFile(string msgPath)
    {
        using var msg = new Storage.Message(msgPath);

        var email = new EmailData
        {
            Subject      = msg.Subject ?? "(No Subject)",
            HtmlBody     = msg.BodyHtml ?? string.Empty,
            TextBody     = msg.BodyText ?? string.Empty,
            ReceivedTime = (msg.ReceivedOn ?? msg.SentOn)?.LocalDateTime ?? DateTime.Now,
            From         = FormatSender(msg.Sender),
            To           = msg.GetEmailRecipients(RecipientType.To, false, false)
                           ?? string.Empty,
        };

        foreach (object att in msg.Attachments)
        {
            // Embedded messages (Storage.Message) and other non-file objects are
            // not Storage.Attachment instances and are skipped here.
            if (att is not Storage.Attachment attachment)
                continue;
            if (attachment.Data is not { Length: > 0 })
                continue;
            if (attachment.Hidden || attachment.IsInline)
                continue;

            string ext      = Path.GetExtension(attachment.FileName ?? string.Empty);
            string tempPath = Path.Combine(
                Path.GetTempPath(),
                $"ep_{Path.GetRandomFileName()}{ext}");

            File.WriteAllBytes(tempPath, attachment.Data);

            email.Attachments.Add(new AttachmentData
            {
                FileName     = attachment.FileName ?? "attachment",
                TempFilePath = tempPath,
            });
        }

        return email;
    }

    /// <summary>
    /// Formats a <see cref="Storage.Sender"/> as a human-readable string,
    /// e.g. "Alice Smith &lt;alice@example.com&gt;" or just the display name /
    /// e-mail address when only one is available.
    /// </summary>
    private static string FormatSender(Storage.Sender? sender)
    {
        if (sender == null)
            return string.Empty;

        string displayName = sender.DisplayName ?? string.Empty;
        string email       = sender.Email       ?? string.Empty;

        if (!string.IsNullOrWhiteSpace(displayName) && !string.IsNullOrWhiteSpace(email))
            return $"{displayName} <{email}>";

        return !string.IsNullOrWhiteSpace(displayName) ? displayName : email;
    }
}
