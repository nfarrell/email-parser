using EmailParser.Core.Models;

namespace EmailParser.Core.Tests.Models;

public class EmailDataTests
{
    // ───────────────────────── EmailData defaults ─────────────────────────

    [Fact]
    public void EmailData_DefaultProperties_AreInitialized()
    {
        var email = new EmailData();

        Assert.Equal(string.Empty, email.Subject);
        Assert.Equal(string.Empty, email.HtmlBody);
        Assert.Equal(string.Empty, email.TextBody);
        Assert.Equal(string.Empty, email.From);
        Assert.Equal(string.Empty, email.To);
        Assert.Equal(string.Empty, email.SourceFilePath);
        Assert.NotNull(email.Attachments);
        Assert.Empty(email.Attachments);
        Assert.Equal(default(DateTime), email.ReceivedTime);
    }

    [Fact]
    public void EmailData_CanSetAllProperties()
    {
        var receivedTime = new DateTime(2024, 1, 15, 10, 30, 0);
        var email = new EmailData
        {
            Subject = "Test Subject",
            HtmlBody = "<html><body>Test</body></html>",
            TextBody = "Test",
            ReceivedTime = receivedTime,
            From = "sender@example.com",
            To = "recipient@example.com",
            SourceFilePath = @"C:\emails\test.msg",
        };

        Assert.Equal("Test Subject", email.Subject);
        Assert.Equal("<html><body>Test</body></html>", email.HtmlBody);
        Assert.Equal("Test", email.TextBody);
        Assert.Equal(receivedTime, email.ReceivedTime);
        Assert.Equal("sender@example.com", email.From);
        Assert.Equal("recipient@example.com", email.To);
        Assert.Equal(@"C:\emails\test.msg", email.SourceFilePath);
    }

    [Fact]
    public void EmailData_Attachments_CanBeModified()
    {
        var email = new EmailData();
        email.Attachments.Add(new AttachmentData
        {
            FileName = "report.pdf",
            TempFilePath = @"C:\temp\ep_abc123.pdf",
        });

        Assert.Single(email.Attachments);
        Assert.Equal("report.pdf", email.Attachments[0].FileName);
    }

    [Fact]
    public void EmailData_Attachments_CanHoldMultiple()
    {
        var email = new EmailData();
        email.Attachments.Add(new AttachmentData { FileName = "file1.pdf" });
        email.Attachments.Add(new AttachmentData { FileName = "file2.docx" });
        email.Attachments.Add(new AttachmentData { FileName = "file3.xlsx" });

        Assert.Equal(3, email.Attachments.Count);
    }

    // ───────────────────────── AttachmentData defaults ─────────────────────────

    [Fact]
    public void AttachmentData_DefaultProperties_AreInitialized()
    {
        var attachment = new AttachmentData();

        Assert.Equal(string.Empty, attachment.FileName);
        Assert.Equal(string.Empty, attachment.TempFilePath);
    }

    [Fact]
    public void AttachmentData_CanSetAllProperties()
    {
        var attachment = new AttachmentData
        {
            FileName = "document.pdf",
            TempFilePath = @"C:\temp\ep_xyz.pdf",
        };

        Assert.Equal("document.pdf", attachment.FileName);
        Assert.Equal(@"C:\temp\ep_xyz.pdf", attachment.TempFilePath);
    }
}
