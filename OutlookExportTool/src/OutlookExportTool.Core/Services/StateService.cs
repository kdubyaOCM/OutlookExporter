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

    public void Save(string statePath, ExportState state)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(statePath)!);
        var json = JsonSerializer.Serialize(state, JsonOptions);
        File.WriteAllText(statePath, json);
    }

    public void Delete(string statePath)
    {
        if (File.Exists(statePath))
        {
            File.Delete(statePath);
        }
    }
}
