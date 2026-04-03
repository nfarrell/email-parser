using System.IO.Compression;
using EmailParser.Core.Helpers;
using EmailParser.Core.Models;
using EmailParser.Core.Services;

namespace EmailParser.Core.Tests.Services;

public class AttachmentSaverTests : IDisposable
{
    private readonly string _tempDir;
    private readonly AttachmentSaver _saver;

    public AttachmentSaverTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"ep_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _saver = new AttachmentSaver();
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { }
    }

    // ───────────────────────── Regular attachments ─────────────────────────

    [Fact]
    public void SaveAttachmentsToFolder_CopiesFileToOutputDir()
    {
        string tempFile = CreateTempFile("test content", ".txt");
        var email = new EmailData
        {
            Attachments = new List<AttachmentData>
            {
                new() { FileName = "document.txt", TempFilePath = tempFile }
            }
        };

        string outputDir = Path.Combine(_tempDir, "output");
        _saver.SaveAttachmentsToFolder(email, outputDir);

        Assert.True(Directory.Exists(outputDir));
        string[] files = Directory.GetFiles(outputDir);
        Assert.Single(files);
        Assert.Equal("document.txt", Path.GetFileName(files[0]));
        Assert.Equal("test content", File.ReadAllText(files[0]));
    }

    [Fact]
    public void SaveAttachmentsToFolder_MultipleAttachments_AllCopied()
    {
        string tempFile1 = CreateTempFile("content1", ".txt");
        string tempFile2 = CreateTempFile("content2", ".pdf");
        var email = new EmailData
        {
            Attachments = new List<AttachmentData>
            {
                new() { FileName = "file1.txt", TempFilePath = tempFile1 },
                new() { FileName = "file2.pdf", TempFilePath = tempFile2 },
            }
        };

        string outputDir = Path.Combine(_tempDir, "output");
        _saver.SaveAttachmentsToFolder(email, outputDir);

        Assert.Equal(2, Directory.GetFiles(outputDir).Length);
    }

    [Fact]
    public void SaveAttachmentsToFolder_MissingTempFile_SkipsGracefully()
    {
        var email = new EmailData
        {
            Attachments = new List<AttachmentData>
            {
                new() { FileName = "ghost.txt", TempFilePath = @"C:\nonexistent\file.txt" }
            }
        };

        string outputDir = Path.Combine(_tempDir, "output");
        _saver.SaveAttachmentsToFolder(email, outputDir);

        // Directory is created but no files copied
        Assert.True(Directory.Exists(outputDir));
        Assert.Empty(Directory.GetFiles(outputDir));
    }

    [Fact]
    public void SaveAttachmentsToFolder_NoAttachments_OnlyCreatesDir()
    {
        var email = new EmailData();

        string outputDir = Path.Combine(_tempDir, "output");
        _saver.SaveAttachmentsToFolder(email, outputDir);

        Assert.True(Directory.Exists(outputDir));
        Assert.Empty(Directory.GetFiles(outputDir));
    }

    [Fact]
    public void SaveAttachmentsToFolder_SanitizesFileName()
    {
        string tempFile = CreateTempFile("test content", ".txt");
        var email = new EmailData
        {
            Attachments = new List<AttachmentData>
            {
                new() { FileName = "file<name>:test.txt", TempFilePath = tempFile }
            }
        };

        string outputDir = Path.Combine(_tempDir, "output");
        _saver.SaveAttachmentsToFolder(email, outputDir);

        string[] files = Directory.GetFiles(outputDir);
        Assert.Single(files);
        string savedFileName = Path.GetFileName(files[0]);
        Assert.DoesNotContain("<", savedFileName);
        Assert.DoesNotContain(">", savedFileName);
        Assert.DoesNotContain(":", savedFileName);
    }

    // ───────────────────────── ZIP attachments ─────────────────────────

    [Fact]
    public void SaveAttachmentsToFolder_ZipFile_ExtractsContents()
    {
        string zipPath = CreateTestZip(new Dictionary<string, string>
        {
            ["inner_file.txt"] = "inner content",
            ["readme.md"] = "# Read Me"
        });

        var email = new EmailData
        {
            Attachments = new List<AttachmentData>
            {
                new() { FileName = "archive.zip", TempFilePath = zipPath }
            }
        };

        string outputDir = Path.Combine(_tempDir, "output");
        _saver.SaveAttachmentsToFolder(email, outputDir);

        // The ZIP should be extracted (flattened) into the output dir
        string[] files = Directory.GetFiles(outputDir);
        Assert.Equal(2, files.Length);

        var fileNames = files.Select(Path.GetFileName).OrderBy(n => n).ToArray();
        Assert.Contains("inner_file.txt", fileNames);
        Assert.Contains("readme.md", fileNames);
    }

    [Fact]
    public void SaveAttachmentsToFolder_MixOfZipAndRegular_AllProcessed()
    {
        string regularFile = CreateTempFile("regular content", ".docx");
        string zipPath = CreateTestZip(new Dictionary<string, string>
        {
            ["zipped.txt"] = "zipped content"
        });

        var email = new EmailData
        {
            Attachments = new List<AttachmentData>
            {
                new() { FileName = "report.docx", TempFilePath = regularFile },
                new() { FileName = "data.zip", TempFilePath = zipPath },
            }
        };

        string outputDir = Path.Combine(_tempDir, "output");
        _saver.SaveAttachmentsToFolder(email, outputDir);

        string[] files = Directory.GetFiles(outputDir);
        Assert.Equal(2, files.Length);
    }

    // ───────────────────────── Helpers ─────────────────────────

    private string CreateTempFile(string content, string extension)
    {
        string path = Path.Combine(_tempDir, $"temp_{Guid.NewGuid():N}{extension}");
        File.WriteAllText(path, content);
        return path;
    }

    private string CreateTestZip(Dictionary<string, string> entries)
    {
        string zipPath = Path.Combine(_tempDir, $"test_{Guid.NewGuid():N}.zip");
        using var stream = File.Create(zipPath);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);
        foreach (var (name, content) in entries)
        {
            var entry = archive.CreateEntry(name);
            using var writer = new StreamWriter(entry.Open());
            writer.Write(content);
        }
        return zipPath;
    }
}
