using EmailParser.Core.Models;

namespace EmailParser.Core.Tests.Models;

public class DataDictionaryTests
{
    [Fact]
    public void Constructor_SetsAllProperties()
    {
        var patterns = new[] { "pattern1", "pattern2" };
        var dict = new DataDictionary("C:\\dict", "C:\\dict\\file.xlsx", patterns);

        Assert.Equal("C:\\dict", dict.DirectoryPath);
        Assert.Equal("C:\\dict\\file.xlsx", dict.SourcePath);
        Assert.Same(patterns, dict.Patterns);
    }

    [Fact]
    public void Constructor_NullSourcePath_IsAllowed()
    {
        var dict = new DataDictionary("C:\\dict", null, Array.Empty<string>());

        Assert.Null(dict.SourcePath);
    }

    [Fact]
    public void Constructor_EmptyPatterns_IsAllowed()
    {
        var dict = new DataDictionary("C:\\dict", null, Array.Empty<string>());

        Assert.Empty(dict.Patterns);
    }

    [Fact]
    public void RecordEquality_SameValues_AreEqual()
    {
        var patterns = new[] { "a", "b" };
        var dict1 = new DataDictionary("dir", "source", patterns);
        var dict2 = new DataDictionary("dir", "source", patterns);

        Assert.Equal(dict1, dict2);
    }

    [Fact]
    public void RecordEquality_DifferentValues_AreNotEqual()
    {
        var dict1 = new DataDictionary("dir1", "source1", new[] { "a" });
        var dict2 = new DataDictionary("dir2", "source2", new[] { "b" });

        Assert.NotEqual(dict1, dict2);
    }

    [Fact]
    public void With_CreatesModifiedCopy()
    {
        var original = new DataDictionary("dir", "source", new[] { "a" });
        var modified = original with { DirectoryPath = "newDir" };

        Assert.Equal("newDir", modified.DirectoryPath);
        Assert.Equal(original.SourcePath, modified.SourcePath);
        Assert.Same(original.Patterns, modified.Patterns);
    }
}
