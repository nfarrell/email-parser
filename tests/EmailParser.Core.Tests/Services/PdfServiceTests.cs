using EmailParser.Core.Models;
using EmailParser.Core.Services;

namespace EmailParser.Core.Tests.Services;

public class PdfServiceTests : IDisposable
{
    private readonly string _tempDir;
    private readonly PdfService _pdfService;

    public PdfServiceTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"ep_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _pdfService = new PdfService();
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { }
    }

    [Fact]
    public void SaveEmailAsPdf_HtmlBody_CreatesPdfFile()
    {
        var email = new EmailData
        {
            Subject = "Test Email",
            HtmlBody = "<html><body><p>Hello World</p></body></html>",
            TextBody = "Hello World",
            ReceivedTime = new DateTime(2024, 6, 15, 14, 30, 0),
            From = "sender@example.com",
            To = "recipient@example.com",
        };

        string outputPath = Path.Combine(_tempDir, "test_email.pdf");
        _pdfService.SaveEmailAsPdf(email, outputPath);

        Assert.True(File.Exists(outputPath));
        // PDF files start with %PDF
        byte[] bytes = File.ReadAllBytes(outputPath);
        Assert.True(bytes.Length > 0);
        Assert.Equal((byte)'%', bytes[0]);
        Assert.Equal((byte)'P', bytes[1]);
        Assert.Equal((byte)'D', bytes[2]);
        Assert.Equal((byte)'F', bytes[3]);
    }

    [Fact]
    public void SaveEmailAsPdf_TextBodyOnly_CreatesPdfFile()
    {
        var email = new EmailData
        {
            Subject = "Plain Text Email",
            HtmlBody = "", // empty HTML body
            TextBody = "This is a plain text email body.",
            ReceivedTime = new DateTime(2024, 1, 1, 8, 0, 0),
            From = "alice@example.com",
            To = "bob@example.com",
        };

        string outputPath = Path.Combine(_tempDir, "plain_text.pdf");
        _pdfService.SaveEmailAsPdf(email, outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void SaveEmailAsPdf_SpecialCharsInFields_DoesNotThrow()
    {
        var email = new EmailData
        {
            Subject = "RE: <Important> & \"Urgent\" Email",
            HtmlBody = "<html><body><p>Content with &amp; special chars</p></body></html>",
            TextBody = "Content with & special chars",
            ReceivedTime = DateTime.Now,
            From = "user <admin@example.com>",
            To = "\"Team\" <team@example.com>",
        };

        string outputPath = Path.Combine(_tempDir, "special_chars.pdf");
        _pdfService.SaveEmailAsPdf(email, outputPath);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SaveEmailAsPdf_EmptyEmailFields_DoesNotThrow()
    {
        var email = new EmailData
        {
            Subject = "",
            HtmlBody = "",
            TextBody = "",
            ReceivedTime = default,
            From = "",
            To = "",
        };

        string outputPath = Path.Combine(_tempDir, "empty_email.pdf");
        _pdfService.SaveEmailAsPdf(email, outputPath);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SaveEmailAsPdf_WithAttachments_CleansUpTempFiles()
    {
        // Create a temp attachment file
        string tempAttachment = Path.Combine(_tempDir, "temp_attachment.txt");
        File.WriteAllText(tempAttachment, "attachment content");

        var email = new EmailData
        {
            Subject = "Email With Attachment",
            HtmlBody = "<p>See attached</p>",
            ReceivedTime = DateTime.Now,
            From = "sender@example.com",
            To = "recipient@example.com",
            Attachments = new List<AttachmentData>
            {
                new() { FileName = "attachment.txt", TempFilePath = tempAttachment }
            }
        };

        string outputPath = Path.Combine(_tempDir, "with_attachment.pdf");
        _pdfService.SaveEmailAsPdf(email, outputPath);

        Assert.True(File.Exists(outputPath));
        // The PdfService should clean up the temp attachment file
        Assert.False(File.Exists(tempAttachment));
    }

    [Fact]
    public void SaveEmailAsPdf_LargeHtmlBody_HandlesProperly()
    {
        string largeBody = "<html><body>" +
            string.Join("", Enumerable.Range(0, 100).Select(i => $"<p>Paragraph {i} with some content.</p>")) +
            "</body></html>";

        var email = new EmailData
        {
            Subject = "Large Email",
            HtmlBody = largeBody,
            ReceivedTime = DateTime.Now,
            From = "sender@example.com",
            To = "recipient@example.com",
        };

        string outputPath = Path.Combine(_tempDir, "large_email.pdf");
        _pdfService.SaveEmailAsPdf(email, outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }
}
