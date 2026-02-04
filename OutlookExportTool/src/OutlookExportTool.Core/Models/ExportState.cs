namespace OutlookExportTool.Core.Models;

public sealed class ExportState
{
    public string RunId { get; set; } = string.Empty;
    public string ExportRoot { get; set; } = string.Empty;
    public string FolderEntryId { get; set; } = string.Empty;
    public string FolderStoreId { get; set; } = string.Empty;
    public string FolderPath { get; set; } = string.Empty;
    public int CurrentIndex { get; set; }
    public List<string> PendingEntryIds { get; set; } = new();
}
