using System.IO.Compression;
using System.Runtime.InteropServices;
using EmailParser.Models;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Serilog;

namespace EmailParser.Services;

/// <summary>
/// Converts email attachments (images, Word, Excel, PDF, ZIP) into one or more
/// temporary PDF files that can later be merged into the final output.
/// </summary>
public class AttachmentProcessor
{
    private static readonly ILogger Log = Serilog.Log.ForContext<AttachmentProcessor>();

    private static readonly HashSet<string> ImageExtensions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif",
        };

    private static readonly HashSet<string> WordExtensions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ".doc", ".docx", ".rtf",
        };

    private static readonly HashSet<string> ExcelExtensions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ".xls", ".xlsx", ".csv",
        };

    // Public API

    /// <summary>
    /// Converts a single attachment into one or more temporary PDF files.
    /// ZIP archives are extracted and each contained file is processed in turn.
    /// </summary>
    /// <returns>
    /// Paths to temporary PDF files. The caller is responsible for deleting
    /// every path in the returned list once it has been consumed.
    /// </returns>
    public IReadOnlyList<string> ProcessAttachment(AttachmentData attachment)
    {
        string ext = System.IO.Path.GetExtension(attachment.FileName);
        Log.Debug("Processing attachment {FileName} (extension: {Extension})",
            attachment.FileName, ext);

        if (string.Equals(ext, ".zip", StringComparison.OrdinalIgnoreCase))
            return ProcessZipFile(attachment.TempFilePath);

        string? pdf = ConvertFileToPdf(attachment.TempFilePath, attachment.FileName);
        return pdf is null ? Array.Empty<string>() : new[] { pdf };
    }

    // ZIP handling

    private IReadOnlyList<string> ProcessZipFile(string zipPath)
    {
        var results = new List<string>();
        string extractDir = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"ep_zip_{System.IO.Path.GetRandomFileName()}");

        try
        {
            // Do NOT use ZipFile.ExtractToDirectory because it is susceptible to
            // zip-slip path traversal when entries contain relative paths such as
            // "../../evil.exe".  Instead, extract each entry individually after
            // verifying that its resolved destination stays within extractDir.
            Directory.CreateDirectory(extractDir);
            string canonicalExtractDir = System.IO.Path.GetFullPath(extractDir)
                + System.IO.Path.DirectorySeparatorChar;

            using var archive = System.IO.Compression.ZipFile.OpenRead(zipPath);
            foreach (var entry in archive.Entries)
            {
                // Skip directory entries.
                if (string.IsNullOrEmpty(entry.Name))
                    continue;

                string destPath = System.IO.Path.GetFullPath(
                    System.IO.Path.Combine(extractDir, entry.FullName));

                // Zip-slip guard: reject entries that escape the extraction directory.
                if (!destPath.StartsWith(canonicalExtractDir, StringComparison.OrdinalIgnoreCase))
                {
                    Log.Warning("Skipping ZIP entry with unsafe path: {EntryName}", entry.FullName);
                    continue;
                }

                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destPath)!);
                entry.ExtractToFile(destPath, overwrite: true);
            }

            foreach (string file in Directory.GetFiles(extractDir, "*", SearchOption.AllDirectories))
            {
                // Recursively handle nested ZIPs.
                if (string.Equals(System.IO.Path.GetExtension(file), ".zip",
                        StringComparison.OrdinalIgnoreCase))
                {
                    var nested = ProcessZipFile(file);
                    results.AddRange(nested);
                }
                else
                {
                    string? pdf = ConvertFileToPdf(file, System.IO.Path.GetFileName(file));
                    if (pdf is not null)
                        results.Add(pdf);
                }
            }
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Could not process ZIP {ZipPath}", zipPath);
        }
        finally
        {
            try { Directory.Delete(extractDir, recursive: true); } catch { /* best-effort */ }
        }

        return results;
    }

    // Generic file → PDF dispatcher

    /// <summary>
    /// Converts a file at <paramref name="filePath"/> to a temporary PDF,
    /// using <paramref name="originalFileName"/> to determine the file type.
    /// Returns <c>null</c> when the file type is not supported.
    /// </summary>
    private string? ConvertFileToPdf(string filePath, string originalFileName)
    {
        if (!File.Exists(filePath))
            return null;

        string ext = System.IO.Path.GetExtension(originalFileName);

        if (string.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase))
            return CopyToNewTempPdf(filePath);

        if (ImageExtensions.Contains(ext))
            return ConvertImageToPdf(filePath);

        if (WordExtensions.Contains(ext))
        {
            try
            {
                return ConvertWordToPdf(filePath);
            }
            catch (Exception ex) when (OfficeAvailability.IsOfficeUnavailableException(ex))
            {
                Log.Warning("Skipping Word attachment {FileName} — Microsoft Word is not " +
                    "installed or accessible", originalFileName);
                return null;
            }
        }

        if (ExcelExtensions.Contains(ext))
        {
            try
            {
                return ConvertExcelToPdf(filePath);
            }
            catch (Exception ex) when (OfficeAvailability.IsOfficeUnavailableException(ex))
            {
                Log.Warning("Skipping Excel attachment {FileName} — Microsoft Excel is not " +
                    "installed or accessible", originalFileName);
                return null;
            }
        }

        Log.Warning("Skipping unsupported attachment type: {FileName}", originalFileName);
        return null;
    }

    // Conversion methods

    /// <summary>Copies an existing PDF to a fresh temp file owned by the caller.</summary>
    private static string CopyToNewTempPdf(string sourcePdf)
    {
        string dest = TempPdfPath();
        File.Copy(sourcePdf, dest, overwrite: true);
        return dest;
    }

    private static string ConvertImageToPdf(string imagePath)
    {
        string outputPath = TempPdfPath();

        ImageData imageData = ImageDataFactory.Create(imagePath);

        // Scale to fit within A4 (595 × 842 pt) while preserving aspect ratio.
        float imgW = imageData.GetWidth();
        float imgH = imageData.GetHeight();
        float maxW = PageSize.A4.GetWidth();
        float maxH = PageSize.A4.GetHeight();
        float scale = Math.Min(maxW / imgW, maxH / imgH);
        if (scale > 1f) scale = 1f;

        var pageSize = new PageSize(imgW * scale, imgH * scale);

        using var writer = new PdfWriter(outputPath);
        using var pdfDoc = new PdfDocument(writer);
        pdfDoc.SetDefaultPageSize(pageSize);

        using var document = new iText.Layout.Document(pdfDoc);
        document.SetMargins(0, 0, 0, 0);

        var image = new iText.Layout.Element.Image(imageData).SetAutoScale(true);
        document.Add(image);

        return outputPath;
    }

    private static string? ConvertWordToPdf(string wordPath)
    {
        string outputPath = TempPdfPath();

        Microsoft.Office.Interop.Word.Application? wordApp = null;
        Microsoft.Office.Interop.Word.Document? doc = null;

        try
        {
            wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
            doc = wordApp.Documents.Open(
                wordPath,
                ReadOnly: true,
                AddToRecentFiles: false);

            doc.ExportAsFixedFormat(outputPath, WdExportFormat.wdExportFormatPDF);
            return outputPath;
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to convert Word document {WordPath}", wordPath);
            TryDeleteFile(outputPath);
            return null;
        }
        finally
        {
            if (doc is not null)
            {
                doc.Close(SaveChanges: false);
                Marshal.ReleaseComObject(doc);
            }

            if (wordApp is not null)
            {
                wordApp.Quit(SaveChanges: false);
                Marshal.ReleaseComObject(wordApp);
            }
        }
    }

    private static string? ConvertExcelToPdf(string excelPath)
    {
        string outputPath = TempPdfPath();

        Microsoft.Office.Interop.Excel.Application? excelApp = null;
        Workbook? workbook = null;

        try
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
            };

            workbook = excelApp.Workbooks.Open(
                excelPath,
                ReadOnly: true,
                AddToMru: false);

            workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputPath);
            return outputPath;
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to convert Excel document {ExcelPath}", excelPath);
            TryDeleteFile(outputPath);
            return null;
        }
        finally
        {
            if (workbook is not null)
            {
                workbook.Close(SaveChanges: false);
                Marshal.ReleaseComObject(workbook);
            }

            if (excelApp is not null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }

    // Utility helpers

    private static string TempPdfPath() =>
        System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"ep_{System.IO.Path.GetRandomFileName()}.pdf");

    private static void TryDeleteFile(string path)
    {
        try { File.Delete(path); } catch { /* best-effort */ }
    }
}
