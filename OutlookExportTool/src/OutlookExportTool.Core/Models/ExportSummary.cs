namespace OutlookExportTool.Core.Models;

public sealed class ExportSummary
{
    public int Total { get; set; }
    public int Processed { get; set; }
    public int Skipped { get; set; }
    public int Failed { get; set; }
    public int Partial { get; set; }
    public int Collisions { get; set; }
    public string ExportRoot { get; set; } = string.Empty;
    public bool Canceled { get; set; }
}
