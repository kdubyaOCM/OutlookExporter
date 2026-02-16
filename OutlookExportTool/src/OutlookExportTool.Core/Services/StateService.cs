using System.Text.Json;
using OutlookExportTool.Core.Models;

namespace OutlookExportTool.Core.Services;

public sealed class StateService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = null
    };

    public ExportState? Load(string statePath)
    {
        if (!File.Exists(statePath))
        {
            return null;
        }

        var json = File.ReadAllText(statePath);
        return JsonSerializer.Deserialize<ExportState>(json, JsonOptions);
    }

    public async Task<ExportState?> LoadAsync(string statePath, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(statePath))
        {
            return null;
        }

        var json = await File.ReadAllTextAsync(statePath, cancellationToken);
        return JsonSerializer.Deserialize<ExportState>(json, JsonOptions);
    }

    public void Save(string statePath, ExportState state)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(statePath)!);
        var json = JsonSerializer.Serialize(state, JsonOptions);
        File.WriteAllText(statePath, json);
    }

    public async Task SaveAsync(string statePath, ExportState state, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(statePath)!);
        var json = JsonSerializer.Serialize(state, JsonOptions);
        await File.WriteAllTextAsync(statePath, json, cancellationToken);
    }

    public void Delete(string statePath)
    {
        if (File.Exists(statePath))
        {
            File.Delete(statePath);
        }
    }

    public Task DeleteAsync(string statePath, CancellationToken cancellationToken = default)
    {
        return Task.Run(() =>
        {
            if (File.Exists(statePath))
            {
                File.Delete(statePath);
            }
        }, cancellationToken);
    }
}
