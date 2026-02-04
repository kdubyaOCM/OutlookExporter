using System.Text;

namespace OutlookExportTool.Core.Utilities;

public static class FileNameSanitizer
{
    public static string Sanitize(string? name, string fallback = "file")
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return fallback;
        }

        var invalid = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder(name.Length);
        foreach (var ch in name)
        {
            builder.Append(invalid.Contains(ch) ? '_' : ch);
        }

        var cleaned = builder.ToString().Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? fallback : cleaned;
    }

    public static string EnsureUnique(string directory, string baseName)
    {
        var candidate = baseName;
        var counter = 1;
        while (File.Exists(Path.Combine(directory, candidate)))
        {
            candidate = AppendSuffix(baseName, counter++);
        }

        return candidate;
    }

    private static string AppendSuffix(string name, int counter)
    {
        var ext = Path.GetExtension(name);
        var stem = Path.GetFileNameWithoutExtension(name);
        return string.IsNullOrEmpty(ext)
            ? $"{stem}_{counter}"
            : $"{stem}_{counter}{ext}";
    }
}
