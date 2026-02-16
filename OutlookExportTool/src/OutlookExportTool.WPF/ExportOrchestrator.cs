using OutlookExportTool.Core.Models;

namespace OutlookExportTool.WPF;

public sealed class ExportOrchestrator
{
    public Task<ExportSummary> StartExportAsync(ExportOptions options, Action<ProgressInfo> progress, CancellationToken token)
    {
        var tcs = new TaskCompletionSource<ExportSummary>();

        var thread = new Thread(() =>
        {
            try
            {
                var worker = new OutlookExportWorker();
                var summary = worker.Run(options, progress, token);
                tcs.TrySetResult(summary);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Export worker failed: {ex.GetType().Name} - {ex.Message}. Stack trace: {ex.StackTrace}");
                var summary = new ExportSummary
                {
                    Canceled = token.IsCancellationRequested,
                    Failed = 1,
                    Total = 0,
                    Processed = 0
                };
                tcs.TrySetResult(summary);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();

        return tcs.Task;
    }
}
