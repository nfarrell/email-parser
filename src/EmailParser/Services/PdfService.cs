using EmailParser.Models;
using iText.Html2pdf;
using iText.Kernel.Pdf;
using Serilog;

namespace EmailParser.Services;

/// <summary>
/// Converts an <see cref="EmailData"/> (body + attachments) into a single PDF file.
/// </summary>
public class PdfService
{
    private static readonly ILogger Log = Serilog.Log.ForContext<PdfService>();

    // Public API

    /// <summary>
    /// Renders <paramref name="email"/> body to a PDF at <paramref name="outputPath"/>.
    /// Attachments are saved separately in their original format.
    /// </summary>
    public void SaveEmailAsPdf(EmailData email, string outputPath)
    {
        Log.Debug("Converting email '{Subject}' to PDF at {OutputPath}", email.Subject, outputPath);

        var tempFilesToDelete = new List<string>();

        try
        {
            // Convert the email body to a PDF.
            string bodyPdf = Path.Combine(
                Path.GetTempPath(),
                $"ep_{Path.GetRandomFileName()}.pdf");
            tempFilesToDelete.Add(bodyPdf);

            ConvertEmailBodyToPdf(email, bodyPdf);

            // Copy the body PDF to the output path.
            File.Copy(bodyPdf, outputPath, overwrite: true);

            Log.Debug("PDF created successfully at {OutputPath}", outputPath);
        }
        finally
        {
            // Clean up all temporary PDFs.
            foreach (string tmp in tempFilesToDelete)
                TryDeleteFile(tmp);

            // Clean up the Outlook attachment temp files saved by OutlookService.
            foreach (AttachmentData att in email.Attachments)
                TryDeleteFile(att.TempFilePath);
        }
    }

    // Email body → PDF

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

    // Utility helpers

    private static string HtmlEncode(string value) =>
        System.Net.WebUtility.HtmlEncode(value ?? string.Empty);

    private static void TryDeleteFile(string path)
    {
        try { File.Delete(path); } catch { /* best-effort */ }
    }
}
