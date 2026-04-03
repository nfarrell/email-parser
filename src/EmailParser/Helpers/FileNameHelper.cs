namespace EmailParser.Helpers;

/// <summary>
/// Utility methods for sanitising file and folder names and stripping
/// data-dictionary terms from paths.
/// </summary>
internal static class FileNameHelper
{
    // Build once; Path.GetInvalidFileNameChars() returns the same values every call.
    private static readonly HashSet<char> InvalidFileNameChars =
        new(Path.GetInvalidFileNameChars());

    /// <summary>
    /// Replaces characters that are invalid in file names with underscores.
    /// </summary>
    internal static string SanitizeFileName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "No Subject";

        string result = string.Concat(name.Select(c => InvalidFileNameChars.Contains(c) ? '_' : c));
        return result.Trim().TrimEnd('.');
    }

    /// <summary>
    /// Converts a folder path into a safe relative path for use as a directory name.
    /// For absolute paths, strips the drive root and user-profile prefix so that
    /// "C:", "Users", and the username never appear in the output folder name.
    /// </summary>
    internal static string SanitizePath(string folderPath)
    {
        if (Path.IsPathRooted(folderPath))
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string fullPath = Path.GetFullPath(folderPath);

            if (fullPath.StartsWith(userProfile, StringComparison.OrdinalIgnoreCase))
            {
                folderPath = Path.GetRelativePath(userProfile, fullPath);
            }
            else
            {
                string? root = Path.GetPathRoot(fullPath);
                if (!string.IsNullOrEmpty(root))
                    folderPath = fullPath.Substring(root.Length);
            }
        }

        string[] segments = folderPath.Split(new[] { '/', '\\' },
            StringSplitOptions.RemoveEmptyEntries);

        string[] safeSegments = segments.Select(seg =>
            string.Concat(seg.Select(c => InvalidFileNameChars.Contains(c) ? '_' : c)).Trim()
        ).ToArray();

        return Path.Combine(safeSegments);
    }

    /// <summary>
    /// Strips every dictionary term from a file/folder name (case-insensitive),
    /// then collapses runs of whitespace and trims leading separators.
    /// </summary>
    internal static string StripDictionaryTerms(string? text, IReadOnlyList<string> patterns)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        string result = text.Trim();

        foreach (string pattern in patterns)
        {
            int idx;
            while ((idx = result.IndexOf(pattern, StringComparison.OrdinalIgnoreCase)) >= 0)
                result = result.Remove(idx, pattern.Length);
        }

        result = string.Join(" ", result.Split(' ', StringSplitOptions.RemoveEmptyEntries));
        result = result.TrimStart(' ', '-', '_').Trim();

        return result;
    }

    /// <summary>
    /// Returns a unique file path inside <paramref name="directory"/>; appends a
    /// counter if the name is already taken.
    /// </summary>
    internal static string GetUniqueFilePath(string directory, string fileName)
    {
        string dest = Path.Combine(directory, fileName);
        if (!File.Exists(dest))
            return dest;

        string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
        string ext = Path.GetExtension(fileName);
        int counter = 2;
        do
        {
            dest = Path.Combine(directory, $"{nameWithoutExt} ({counter++}){ext}");
        }
        while (File.Exists(dest));

        return dest;
    }
}
