namespace OutlookExportTool.Core.Models;

public sealed class ExportOptions
{
    public string OutlookFolderEntryId { get; init; } = string.Empty;
    public string OutlookFolderStoreId { get; init; } = string.Empty;
    public string OutlookFolderPath { get; init; } = string.Empty;
    public string OutputRoot { get; init; } = string.Empty;
    public bool ResumeEnabled { get; init; } = true;
    public DateTime? StartDate { get; init; }
    public DateTime? EndDate { get; init; }
}
