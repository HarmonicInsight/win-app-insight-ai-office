using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using InsightCommon.License;
using InsightCommon.UI;
using InsightAiOffice.App.Helpers;
using Microsoft.Win32;

namespace InsightAiOffice.App.ViewModels;

public partial class MainViewModel : ObservableObject, IDisposable
{
    private readonly InsightLicenseManager _licenseManager;

    [ObservableProperty] private string _statusText = "";
    [ObservableProperty] private bool _isFileLoaded;
    [ObservableProperty] private string? _currentFilePath;
    [ObservableProperty] private string _documentTitle = "";

    /// <summary>Called by MainWindow after file is loaded into the correct editor.</summary>
    public Action<string>? OnFileOpenRequested { get; set; }

    public MainViewModel(InsightLicenseManager licenseManager)
    {
        _licenseManager = licenseManager;
        StatusText = LanguageManager.Get("Status_Ready");
    }

    [RelayCommand]
    private void NewDocument()
    {
        DocumentTitle = LanguageManager.Get("File_New");
        IsFileLoaded = true;
        StatusText = LanguageManager.Get("Status_Ready");
    }

    [RelayCommand]
    private void OpenDocument()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "All Supported|*.iaof;*.docx;*.doc;*.xlsx;*.xls;*.csv;*.pptx;*.ppt;*.pdf|" +
                     "IAOF Project|*.iaof|" +
                     "Word|*.docx;*.doc|Excel|*.xlsx;*.xls;*.csv|PowerPoint|*.pptx;*.ppt|" +
                     "PDF|*.pdf|All Files|*.*",
        };

        if (dialog.ShowDialog() != true) return;
        OnFileOpenRequested?.Invoke(dialog.FileName);
    }

    public async Task OpenFileByPathAsync(string filePath)
    {
        OnFileOpenRequested?.Invoke(filePath);
        await Task.CompletedTask;
    }

    [RelayCommand]
    private void SaveDocument()
    {
        if (string.IsNullOrEmpty(CurrentFilePath))
        {
            SaveAsDocument();
            return;
        }
        StatusText = LanguageManager.Get("File_Save");
    }

    [RelayCommand]
    private void SaveAsDocument()
    {
        var dialog = new SaveFileDialog
        {
            Filter = "IAOF Project|*.iaof|Word|*.docx|Excel|*.xlsx|PowerPoint|*.pptx",
        };

        if (dialog.ShowDialog() != true) return;
        CurrentFilePath = dialog.FileName;
        DocumentTitle = Path.GetFileName(dialog.FileName);
        StatusText = LanguageManager.Get("File_Save");
    }

    [RelayCommand]
    private void ShowHelp()
    {
        Views.HelpWindow.ShowSection(Application.Current.MainWindow, "overview");
        StatusText = LanguageManager.Get("Menu_Help");
    }

    [RelayCommand]
    private void ShowHelpSection(string? sectionId)
    {
        Views.HelpWindow.ShowSection(Application.Current.MainWindow, sectionId ?? "overview");
    }

    [RelayCommand]
    private void ZoomIn() => InsightScaleManager.Instance.ZoomIn();

    [RelayCommand]
    private void ZoomOut() => InsightScaleManager.Instance.ZoomOut();

    [RelayCommand]
    private void ResetZoom() => InsightScaleManager.Instance.Reset();

    public void Dispose()
    {
        GC.SuppressFinalize(this);
    }
}
