using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using Microsoft.Office.Interop.Outlook;
using OutlookExportTool.Core.Models;
using OutlookExportTool.Core.Services;
using OutlookExportTool.Outlook;

namespace OutlookExportTool.WPF;

public sealed class MainWindowViewModel : INotifyPropertyChanged
{
    private readonly ExportOrchestrator _orchestrator = new();
    private CancellationTokenSource? _cts;
    private string? _folderEntryId;
    private string? _folderStoreId;
    private string _outlookFolderPath = string.Empty;
    private string _outputFolder = string.Empty;
    private string _statusText = "Ready.";
    private double _progressPercent;
    private bool _isRunning;
    private DateTime? _startDate;
    private DateTime? _endDate;

    public event PropertyChangedEventHandler? PropertyChanged;

    public RelayCommand PickOutlookFolderCommand { get; }
    public RelayCommand PickOutputFolderCommand { get; }
    public RelayCommand StartCommand { get; }
    public RelayCommand CancelCommand { get; }

    public MainWindowViewModel()
    {
        PickOutlookFolderCommand = new RelayCommand(PickOutlookFolder);
        PickOutputFolderCommand = new RelayCommand(PickOutputFolder);
        StartCommand = new RelayCommand(StartExport, () => CanStart);
        CancelCommand = new RelayCommand(CancelExport, () => CanCancel);
    }

    public string OutlookFolderPath
    {
        get => _outlookFolderPath;
        private set => SetField(ref _outlookFolderPath, value);
    }

    public string OutputFolder
    {
        get => _outputFolder;
        private set => SetField(ref _outputFolder, value);
    }

    public string StatusText
    {
        get => _statusText;
        private set => SetField(ref _statusText, value);
    }

    public double ProgressPercent
    {
        get => _progressPercent;
        private set => SetField(ref _progressPercent, value);
    }

    public DateTime? StartDate
    {
        get => _startDate;
        set => SetField(ref _startDate, value);
    }

    public DateTime? EndDate
    {
        get => _endDate;
        set => SetField(ref _endDate, value);
    }

    public bool CanStart => !_isRunning && !string.IsNullOrWhiteSpace(_folderEntryId) && !string.IsNullOrWhiteSpace(OutputFolder);
    public bool CanCancel => _isRunning;

    private void PickOutlookFolder()
    {
        Application? outlookApp = null;
        NameSpace? ns = null;
        MAPIFolder? folder = null;

        try
        {
            outlookApp = new Application();
            ns = outlookApp.GetNamespace("MAPI");
            folder = ns.PickFolder();

            if (folder != null)
            {
                _folderEntryId = folder.EntryID;
                _folderStoreId = folder.StoreID;
                OutlookFolderPath = folder.FolderPath;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to pick Outlook folder: {ex.Message}", "Outlook Export Tool", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ComRelease.Release(folder);
            ComRelease.Release(ns);
            ComRelease.Release(outlookApp);
        }

        RaiseCommandState();
    }

    private void PickOutputFolder()
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select export output folder",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            OutputFolder = dialog.SelectedPath;
        }

        RaiseCommandState();
    }

    private async void StartExport()
    {
        if (string.IsNullOrWhiteSpace(_folderEntryId) || string.IsNullOrWhiteSpace(_folderStoreId))
        {
            return;
        }

        _cts = new CancellationTokenSource();
        _isRunning = true;
        StatusText = "Starting export...";
        ProgressPercent = 0;
        RaiseCommandState();

        var options = new ExportOptions
        {
            OutlookFolderEntryId = _folderEntryId,
            OutlookFolderStoreId = _folderStoreId,
            OutlookFolderPath = OutlookFolderPath,
            OutputRoot = OutputFolder,
            ResumeEnabled = true,
            StartDate = StartDate,
            EndDate = EndDate
        };

        var summary = await _orchestrator.StartExportAsync(options, UpdateProgress, _cts.Token);

        _isRunning = false;
        RaiseCommandState();

        if (summary.Canceled)
        {
            StatusText = $"Canceled. Processed {summary.Processed}/{summary.Total}.";
        }
        else
        {
            StatusText = $"Completed. Total {summary.Total}, processed {summary.Processed}, skipped {summary.Skipped}, failed {summary.Failed}, partial {summary.Partial}.";
        }
    }

    private void CancelExport()
    {
        _cts?.Cancel();
        StatusText = "Cancel requested. Finishing current email...";
    }

    private void UpdateProgress(ProgressInfo progress)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            ProgressPercent = progress.Total == 0 ? 0 : (double)progress.Processed / progress.Total * 100d;
            StatusText = progress.Status;
        });
    }

    private void RaiseCommandState()
    {
        StartCommand.RaiseCanExecuteChanged();
        CancelCommand.RaiseCanExecuteChanged();
    }

    private void SetField<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
