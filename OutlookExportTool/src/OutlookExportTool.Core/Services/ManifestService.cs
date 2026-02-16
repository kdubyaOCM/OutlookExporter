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

    public async Task<Dictionary<string, ManifestEntry>> LoadIndexAsync(string manifestPath, CancellationToken cancellationToken = default)
    {
        var index = new Dictionary<string, ManifestEntry>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(manifestPath))
        {
            return index;
        }

        await foreach (var line in File.ReadLinesAsync(manifestPath, cancellationToken))
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

    public async Task AppendAsync(string manifestPath, ManifestEntry entry, CancellationToken cancellationToken = default)
    {
        var json = JsonSerializer.Serialize(entry, JsonOptions);
        Directory.CreateDirectory(Path.GetDirectoryName(manifestPath)!);
        await File.AppendAllTextAsync(manifestPath, json + Environment.NewLine, cancellationToken);
    }
}
