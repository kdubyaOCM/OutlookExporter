namespace OutlookExportTool.Core.Models;

public sealed class ProgressInfo
{
    public int Processed { get; set; }
    public int Total { get; set; }
    public string Status { get; set; } = string.Empty;
}
