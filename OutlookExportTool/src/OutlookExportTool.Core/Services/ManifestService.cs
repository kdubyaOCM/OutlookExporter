using System.Text.Json;
using OutlookExportTool.Core.Models;

namespace OutlookExportTool.Core.Services;

public sealed class ManifestService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = false,
        PropertyNamingPolicy = null
    };

    public Dictionary<string, ManifestEntry> LoadIndex(string manifestPath)
    {
        var index = new Dictionary<string, ManifestEntry>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(manifestPath))
        {
            return index;
        }

        foreach (var line in File.ReadLines(manifestPath))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            try
            {
                var entry = JsonSerializer.Deserialize<ManifestEntry>(line, JsonOptions);
                if (entry != null && !string.IsNullOrWhiteSpace(entry.EntryId))
                {
                    index[entry.EntryId] = entry;
                }
            }
            catch
            {
                // Ignore malformed lines for resilience.
            }
        }

        return index;
    }

    public void Append(string manifestPath, ManifestEntry entry)
    {
        var json = JsonSerializer.Serialize(entry, JsonOptions);
        Directory.CreateDirectory(Path.GetDirectoryName(manifestPath)!);
        File.AppendAllText(manifestPath, json + Environment.NewLine);
    }
}
