using EmailParser.Models;
using iText.Html2pdf;
using iText.Kernel.Pdf;

namespace EmailParser.Services;

/// <summary>
/// Converts an <see cref="EmailData"/> (body + attachments) into a single PDF file.
/// </summary>
public class PdfService
{
    private readonly AttachmentProcessor _attachmentProcessor = new();

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    /// <summary>
    /// Renders <paramref name="email"/> and all of its attachments to a single PDF
    /// at <paramref name="outputPath"/>.
    /// </summary>
    public void SaveEmailAsPdf(EmailData email, string outputPath)
    {
        var tempFilesToDelete = new List<string>();

        try
        {
            // 1. Convert the email body to a temporary PDF.
            string bodyPdf = Path.Combine(
                Path.GetTempPath(),
                $"ep_{Path.GetRandomFileName()}.pdf");
            tempFilesToDelete.Add(bodyPdf);

            ConvertEmailBodyToPdf(email, bodyPdf);

            // 2. Convert each attachment to one or more temporary PDFs.
            var attachmentPdfs = new List<string>();
            foreach (AttachmentData attachment in email.Attachments)
            {
                IReadOnlyList<string> pdfs = _attachmentProcessor.ProcessAttachment(attachment);
                attachmentPdfs.AddRange(pdfs);
                tempFilesToDelete.AddRange(pdfs);
            }

            // 3. Merge body + attachment PDFs into the final output file.
            var allPdfs = new List<string>(attachmentPdfs.Count + 1) { bodyPdf };
            allPdfs.AddRange(attachmentPdfs);
            MergePdfs(allPdfs, outputPath);
        }
        finally
        {
            // Clean up all temporary PDFs created during conversion.
            foreach (string tmp in tempFilesToDelete)
                TryDeleteFile(tmp);

            // Clean up the Outlook attachment temp files saved by OutlookService.
            foreach (AttachmentData att in email.Attachments)
                TryDeleteFile(att.TempFilePath);
        }
    }

    // -------------------------------------------------------------------------
    // Email body → PDF
    // -------------------------------------------------------------------------

    private static void ConvertEmailBodyToPdf(EmailData email, string outputPath)
    {
        string html = BuildEmailHtml(email);

        using var writer = new PdfWriter(outputPath);
        using var pdfDoc = new PdfDocument(writer);
        HtmlConverter.ConvertToPdf(html, pdfDoc, new ConverterProperties());
    }

    /// <summary>
    /// Wraps the email's HTML (or plain-text) body in a full HTML page that
    /// prepends a formatted header (From / To / Subject / Received).
    /// </summary>
    private static string BuildEmailHtml(EmailData email)
    {
        string body = string.IsNullOrWhiteSpace(email.HtmlBody)
            ? $"<pre>{HtmlEncode(email.TextBody)}</pre>"
            : email.HtmlBody;

        // Use $$""" so that CSS braces are literal and {{expr}} is interpolation.
        return $$"""
            <!DOCTYPE html>
            <html>
            <head>
              <meta charset="UTF-8">
              <style>
                body { font-family: Arial, sans-serif; margin: 20px; color: #222; }
                .ep-header { border-bottom: 2px solid #ccc; padding-bottom: 10px; margin-bottom: 20px; }
                .ep-row { margin: 4px 0; }
                .ep-label { font-weight: bold; display: inline-block; min-width: 80px; }
              </style>
            </head>
            <body>
              <div class="ep-header">
                <div class="ep-row"><span class="ep-label">From:</span> {{HtmlEncode(email.From)}}</div>
                <div class="ep-row"><span class="ep-label">To:</span> {{HtmlEncode(email.To)}}</div>
                <div class="ep-row"><span class="ep-label">Subject:</span> {{HtmlEncode(email.Subject)}}</div>
                <div class="ep-row"><span class="ep-label">Received:</span> {{email.ReceivedTime:yyyy-MM-dd HH:mm:ss}}</div>
              </div>
              {{body}}
            </body>
            </html>
            """;
    }

    // -------------------------------------------------------------------------
    // PDF merging
    // -------------------------------------------------------------------------

    private static void MergePdfs(IList<string> inputPaths, string outputPath)
    {
        using var writer = new PdfWriter(outputPath);
        using var outputPdf = new PdfDocument(writer);

        foreach (string inputPath in inputPaths)
        {
            if (!File.Exists(inputPath))
                continue;

            try
            {
                using var reader = new PdfReader(inputPath);
                using var inputPdf = new PdfDocument(reader);
                inputPdf.CopyPagesTo(1, inputPdf.GetNumberOfPages(), outputPdf);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(
                    $"  Warning: Could not merge '{inputPath}': {ex.Message}");
            }
        }
    }

    // -------------------------------------------------------------------------
    // Utility helpers
    // -------------------------------------------------------------------------

    private static string HtmlEncode(string value) =>
        System.Net.WebUtility.HtmlEncode(value ?? string.Empty);

    private static void TryDeleteFile(string path)
    {
        try { File.Delete(path); } catch { /* best-effort */ }
    }
}
