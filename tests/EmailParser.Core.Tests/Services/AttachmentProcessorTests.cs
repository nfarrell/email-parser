using System.IO.Compression;
using EmailParser.Core.Models;
using EmailParser.Core.Services;

namespace EmailParser.Core.Tests.Services;

public class AttachmentProcessorTests : IDisposable
{
    private readonly string _tempDir;
    private readonly AttachmentProcessor _processor;

    public AttachmentProcessorTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"ep_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _processor = new AttachmentProcessor();
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { }
    }

    // ───────── PDF pass-through ─────────

    [Fact]
    public void ProcessAttachment_PdfFile_ReturnsCopyPath()
    {
        string pdfPath = CreateMinimalPdf();
        var attachment = new AttachmentData
        {
            FileName = "document.pdf",
            TempFilePath = pdfPath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Single(result);
        Assert.True(File.Exists(result[0]));
        Assert.NotEqual(pdfPath, result[0]);
        File.Delete(result[0]);
    }

    [Fact]
    public void ProcessAttachment_PdfFile_CaseInsensitive()
    {
        string pdfPath = CreateMinimalPdf();
        var attachment = new AttachmentData
        {
            FileName = "DOCUMENT.PDF",
            TempFilePath = pdfPath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Single(result);
        File.Delete(result[0]);
    }

    // ───────── Unsupported types ─────────

    [Fact]
    public void ProcessAttachment_UnsupportedExtension_ReturnsEmpty()
    {
        string filePath = Path.Combine(_tempDir, "data.json");
        File.WriteAllText(filePath, "{}");
        var attachment = new AttachmentData
        {
            FileName = "data.json",
            TempFilePath = filePath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Empty(result);
    }

    [Fact]
    public void ProcessAttachment_TxtFile_ReturnsEmpty()
    {
        string filePath = Path.Combine(_tempDir, "readme.txt");
        File.WriteAllText(filePath, "Hello");
        var attachment = new AttachmentData
        {
            FileName = "readme.txt",
            TempFilePath = filePath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Empty(result);
    }

    // ───────── Missing file ─────────

    [Fact]
    public void ProcessAttachment_FileDoesNotExist_ReturnsEmpty()
    {
        var attachment = new AttachmentData
        {
            FileName = "missing.pdf",
            TempFilePath = Path.Combine(_tempDir, "nonexistent.pdf"),
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Empty(result);
    }

    // ───────── Image conversion ─────────

    [Fact]
    public void ProcessAttachment_PngImage_ReturnsPdf()
    {
        string imagePath = CreateMinimalPng();
        var attachment = new AttachmentData
        {
            FileName = "screenshot.png",
            TempFilePath = imagePath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Single(result);
        Assert.True(File.Exists(result[0]));
        File.Delete(result[0]);
    }

    [Fact]
    public void ProcessAttachment_BmpImage_ReturnsPdf()
    {
        string imagePath = CreateMinimalBmp();
        var attachment = new AttachmentData
        {
            FileName = "image.bmp",
            TempFilePath = imagePath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Single(result);
        Assert.True(File.Exists(result[0]));
        File.Delete(result[0]);
    }

    // ───────── Word/Excel (graceful on CI without Office) ─────────

    [Fact]
    public void ProcessAttachment_WordFile_DoesNotThrow()
    {
        string filePath = Path.Combine(_tempDir, "document.docx");
        File.WriteAllBytes(filePath, new byte[] { 0x50, 0x4B, 0x03, 0x04 });
        var attachment = new AttachmentData
        {
            FileName = "document.docx",
            TempFilePath = filePath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        // Without Office, gracefully returns empty; with Office returns PDF
        Assert.NotNull(result);
    }

    [Fact]
    public void ProcessAttachment_ExcelFile_DoesNotThrow()
    {
        string filePath = Path.Combine(_tempDir, "spreadsheet.xlsx");
        File.WriteAllBytes(filePath, new byte[] { 0x50, 0x4B, 0x03, 0x04 });
        var attachment = new AttachmentData
        {
            FileName = "spreadsheet.xlsx",
            TempFilePath = filePath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.NotNull(result);
    }

    // ───────── ZIP handling ─────────

    [Fact]
    public void ProcessAttachment_ZipWithPdf_ReturnsPdfs()
    {
        string pdfContent = CreateMinimalPdf();
        byte[] pdfBytes = File.ReadAllBytes(pdfContent);

        string zipPath = Path.Combine(_tempDir, "archive.zip");
        using (var stream = File.Create(zipPath))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            var entry = archive.CreateEntry("inner.pdf");
            using var entryStream = entry.Open();
            entryStream.Write(pdfBytes, 0, pdfBytes.Length);
        }

        var attachment = new AttachmentData
        {
            FileName = "archive.zip",
            TempFilePath = zipPath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Single(result);
        Assert.True(File.Exists(result[0]));
        foreach (string path in result)
            File.Delete(path);
    }

    [Fact]
    public void ProcessAttachment_ZipWithUnsupportedFiles_ReturnsEmpty()
    {
        string zipPath = Path.Combine(_tempDir, "text_archive.zip");
        using (var stream = File.Create(zipPath))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            var entry = archive.CreateEntry("readme.txt");
            using var writer = new StreamWriter(entry.Open());
            writer.Write("plain text");
        }

        var attachment = new AttachmentData
        {
            FileName = "text_archive.zip",
            TempFilePath = zipPath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Empty(result);
    }

    [Fact]
    public void ProcessAttachment_EmptyZip_ReturnsEmpty()
    {
        string zipPath = Path.Combine(_tempDir, "empty.zip");
        using (var stream = File.Create(zipPath))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            // empty archive
        }

        var attachment = new AttachmentData
        {
            FileName = "empty.zip",
            TempFilePath = zipPath,
        };

        IReadOnlyList<string> result = _processor.ProcessAttachment(attachment);

        Assert.Empty(result);
    }

    // ───────── Helpers ─────────

    private string CreateMinimalPdf()
    {
        string path = Path.Combine(_tempDir, $"test_{Guid.NewGuid():N}.pdf");
        using var writer = new iText.Kernel.Pdf.PdfWriter(path);
        using var doc = new iText.Kernel.Pdf.PdfDocument(writer);
        doc.AddNewPage();
        return path;
    }

    // Creates a minimal valid 1x1 PNG image.
    private string CreateMinimalPng()
    {
        string path = Path.Combine(_tempDir, $"test_{Guid.NewGuid():N}.png");
        // Minimal 1x1 white PNG
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==");
        File.WriteAllBytes(path, png);
        return path;
    }

    // Creates a minimal valid 1x1 BMP image.
    private string CreateMinimalBmp()
    {
        string path = Path.Combine(_tempDir, $"test_{Guid.NewGuid():N}.bmp");
        // Minimal 1x1 24-bit BMP: 14-byte header + 40-byte DIB header + 4 bytes pixel data (padded)
        byte[] bmp = new byte[58];
        // BM signature
        bmp[0] = 0x42; bmp[1] = 0x4D;
        // File size (58 bytes)
        bmp[2] = 58; bmp[3] = 0; bmp[4] = 0; bmp[5] = 0;
        // Pixel data offset (54 bytes)
        bmp[10] = 54; bmp[11] = 0; bmp[12] = 0; bmp[13] = 0;
        // DIB header size (40 bytes)
        bmp[14] = 40; bmp[15] = 0; bmp[16] = 0; bmp[17] = 0;
        // Width (1)
        bmp[18] = 1; bmp[19] = 0; bmp[20] = 0; bmp[21] = 0;
        // Height (1)
        bmp[22] = 1; bmp[23] = 0; bmp[24] = 0; bmp[25] = 0;
        // Planes (1)
        bmp[26] = 1; bmp[27] = 0;
        // Bits per pixel (24)
        bmp[28] = 24; bmp[29] = 0;
        // Pixel data: BGR (white) + 1 byte padding
        bmp[54] = 255; bmp[55] = 255; bmp[56] = 255; bmp[57] = 0;
        File.WriteAllBytes(path, bmp);
        return path;
    }
}
