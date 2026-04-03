namespace EmailParser.Models;

/// <summary>
/// Represents a loaded data dictionary: the directory it was found in,
/// the path to the source Excel file (if any), and the term patterns extracted.
/// </summary>
public sealed record DataDictionary(
    string DirectoryPath,
    string? SourcePath,
    IReadOnlyList<string> Patterns);
