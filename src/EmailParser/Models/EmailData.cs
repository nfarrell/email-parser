namespace EmailParser.Models;

/// <summary>
/// Represents an email fetched from Outlook.
/// </summary>
public class EmailData
{
    public string Subject { get; set; } = string.Empty;
    public string HtmlBody { get; set; } = string.Empty;
    public string TextBody { get; set; } = string.Empty;
    public DateTime ReceivedTime { get; set; }
    public string From { get; set; } = string.Empty;
    public string To { get; set; } = string.Empty;
    public List<AttachmentData> Attachments { get; set; } = new();
}

/// <summary>
/// Represents a single email attachment saved to a temporary file.
/// </summary>
public class AttachmentData
{
    /// <summary>Original file name of the attachment (e.g. "report.docx").</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>
    /// Full path to the temporary file where the attachment has been saved.
    /// The caller is responsible for deleting this file when finished.
    /// </summary>
    public string TempFilePath { get; set; } = string.Empty;
}
