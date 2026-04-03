using EmailParser.Core.Helpers;

namespace EmailParser.Core.Tests.Helpers;

public class FileNameHelperTests
{
    // ───────────────────────── SanitizeFileName ─────────────────────────

    [Fact]
    public void SanitizeFileName_NullInput_ReturnsNoSubject()
    {
        Assert.Equal("No Subject", FileNameHelper.SanitizeFileName(null));
    }

    [Fact]
    public void SanitizeFileName_EmptyString_ReturnsNoSubject()
    {
        Assert.Equal("No Subject", FileNameHelper.SanitizeFileName(string.Empty));
    }

    [Fact]
    public void SanitizeFileName_WhitespaceOnly_ReturnsNoSubject()
    {
        Assert.Equal("No Subject", FileNameHelper.SanitizeFileName("   "));
    }

    [Fact]
    public void SanitizeFileName_ValidName_ReturnsSameName()
    {
        Assert.Equal("report.docx", FileNameHelper.SanitizeFileName("report.docx"));
    }

    [Fact]
    public void SanitizeFileName_ContainsInvalidChars_ReplacesWithUnderscore()
    {
        // Characters like < > : " / \ | ? * are invalid in file names on Windows.
        string result = FileNameHelper.SanitizeFileName("file<name>test:doc");
        Assert.DoesNotContain("<", result);
        Assert.DoesNotContain(">", result);
        Assert.DoesNotContain(":", result);
        Assert.Contains("_", result);
    }

    [Fact]
    public void SanitizeFileName_TrailingDots_AreTrimmed()
    {
        string result = FileNameHelper.SanitizeFileName("filename...");
        Assert.False(result.EndsWith('.'));
    }

    [Fact]
    public void SanitizeFileName_TrailingSpaces_AreTrimmed()
    {
        string result = FileNameHelper.SanitizeFileName("filename   ");
        Assert.False(result.EndsWith(' '));
    }

    [Fact]
    public void SanitizeFileName_LeadingSpaces_AreTrimmed()
    {
        string result = FileNameHelper.SanitizeFileName("   filename");
        Assert.False(result.StartsWith(' '));
    }

    [Fact]
    public void SanitizeFileName_AllInvalidChars_ReturnsUnderscores()
    {
        string result = FileNameHelper.SanitizeFileName("<>:\"/\\|?*");
        Assert.All(result, c => Assert.Equal('_', c));
    }

    [Fact]
    public void SanitizeFileName_MixedValidAndInvalid_PreservesValidChars()
    {
        string result = FileNameHelper.SanitizeFileName("RE: Important Email");
        Assert.Equal("RE_ Important Email", result);
    }

    // ───────────────────────── SanitizePath ─────────────────────────

    [Fact]
    public void SanitizePath_RelativePath_PreservesSegments()
    {
        string result = FileNameHelper.SanitizePath("folder1/folder2/folder3");
        Assert.Contains("folder1", result);
        Assert.Contains("folder2", result);
        Assert.Contains("folder3", result);
    }

    [Fact]
    public void SanitizePath_RemovesEmptySegments()
    {
        string result = FileNameHelper.SanitizePath("folder1//folder2");
        // Should not have empty segments
        string[] segments = result.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
        Assert.All(segments, s => Assert.False(string.IsNullOrWhiteSpace(s)));
    }

    [Fact]
    public void SanitizePath_InvalidCharsInSegments_AreReplaced()
    {
        string result = FileNameHelper.SanitizePath("folder<1>/folder:2");
        Assert.DoesNotContain("<", result);
        Assert.DoesNotContain(">", result);
    }

    [Fact]
    public void SanitizePath_BackslashSeparators_AreSplit()
    {
        string result = FileNameHelper.SanitizePath(@"folder1\folder2\folder3");
        string[] segments = result.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(3, segments.Length);
    }

    [Fact]
    public void SanitizePath_SingleSegment_ReturnsSameSegment()
    {
        string result = FileNameHelper.SanitizePath("Inbox");
        Assert.Equal("Inbox", result);
    }

    // ───────────────────────── StripDictionaryTerms ─────────────────────────

    [Fact]
    public void StripDictionaryTerms_NullInput_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, FileNameHelper.StripDictionaryTerms(null, Array.Empty<string>()));
    }

    [Fact]
    public void StripDictionaryTerms_EmptyInput_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, FileNameHelper.StripDictionaryTerms("", Array.Empty<string>()));
    }

    [Fact]
    public void StripDictionaryTerms_WhitespaceInput_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, FileNameHelper.StripDictionaryTerms("   ", Array.Empty<string>()));
    }

    [Fact]
    public void StripDictionaryTerms_NoPatterns_ReturnsTrimmedInput()
    {
        Assert.Equal("hello world", FileNameHelper.StripDictionaryTerms("  hello world  ", Array.Empty<string>()));
    }

    [Fact]
    public void StripDictionaryTerms_SinglePattern_IsRemoved()
    {
        var patterns = new[] { "PROJECT" };
        string result = FileNameHelper.StripDictionaryTerms("PROJECT Report", patterns);
        Assert.Equal("Report", result);
    }

    [Fact]
    public void StripDictionaryTerms_CaseInsensitive_RemovesPattern()
    {
        var patterns = new[] { "PROJECT" };
        string result = FileNameHelper.StripDictionaryTerms("project Report", patterns);
        Assert.Equal("Report", result);
    }

    [Fact]
    public void StripDictionaryTerms_MultiplePatterns_AllRemoved()
    {
        var patterns = new[] { "PROJECT", "ACME" };
        string result = FileNameHelper.StripDictionaryTerms("PROJECT ACME Report", patterns);
        Assert.Equal("Report", result);
    }

    [Fact]
    public void StripDictionaryTerms_RepeatedPattern_AllOccurrencesRemoved()
    {
        var patterns = new[] { "test" };
        string result = FileNameHelper.StripDictionaryTerms("test data test results", patterns);
        Assert.Equal("data results", result);
    }

    [Fact]
    public void StripDictionaryTerms_LeadingSeparators_AreTrimmed()
    {
        var patterns = new[] { "PREFIX" };
        string result = FileNameHelper.StripDictionaryTerms("PREFIX - Report", patterns);
        Assert.Equal("Report", result);
    }

    [Fact]
    public void StripDictionaryTerms_LeadingUnderscores_AreTrimmed()
    {
        var patterns = new[] { "PREFIX" };
        string result = FileNameHelper.StripDictionaryTerms("PREFIX _ Report", patterns);
        Assert.Equal("Report", result);
    }

    [Fact]
    public void StripDictionaryTerms_AllContentStripped_ReturnsEmpty()
    {
        var patterns = new[] { "everything" };
        string result = FileNameHelper.StripDictionaryTerms("everything", patterns);
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void StripDictionaryTerms_CollapsesMultipleSpaces()
    {
        var patterns = new[] { "MIDDLE" };
        string result = FileNameHelper.StripDictionaryTerms("Start MIDDLE End", patterns);
        Assert.Equal("Start End", result);
    }

    // ───────────────────────── GetUniqueFilePath ─────────────────────────

    [Fact]
    public void GetUniqueFilePath_NoConflict_ReturnsOriginalPath()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), $"ep_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        try
        {
            string result = FileNameHelper.GetUniqueFilePath(tempDir, "test.txt");
            Assert.Equal(Path.Combine(tempDir, "test.txt"), result);
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void GetUniqueFilePath_FileExists_AppendsCounter()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), $"ep_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        try
        {
            // Create conflicting file
            File.WriteAllText(Path.Combine(tempDir, "test.txt"), "existing");

            string result = FileNameHelper.GetUniqueFilePath(tempDir, "test.txt");
            Assert.Equal(Path.Combine(tempDir, "test (2).txt"), result);
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void GetUniqueFilePath_MultipleConflicts_IncrementsCounter()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), $"ep_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        try
        {
            File.WriteAllText(Path.Combine(tempDir, "test.txt"), "existing");
            File.WriteAllText(Path.Combine(tempDir, "test (2).txt"), "existing2");

            string result = FileNameHelper.GetUniqueFilePath(tempDir, "test.txt");
            Assert.Equal(Path.Combine(tempDir, "test (3).txt"), result);
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void GetUniqueFilePath_FileWithoutExtension_StillWorks()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), $"ep_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        try
        {
            File.WriteAllText(Path.Combine(tempDir, "README"), "existing");

            string result = FileNameHelper.GetUniqueFilePath(tempDir, "README");
            Assert.Equal(Path.Combine(tempDir, "README (2)"), result);
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }
}
