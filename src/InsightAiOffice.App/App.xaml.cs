using System.IO;
using System.Windows;
using System.Windows.Threading;
using InsightCommon.AI;
using InsightCommon.License;
using InsightAiOffice.App.ViewModels;
using InsightAiOffice.App.Helpers;
using Microsoft.Extensions.DependencyInjection;

namespace InsightAiOffice.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Syncfusion ライセンス登録（Binary License Key — 各コンポーネント用）
        // IAOF は Word + Excel + PPTX + Ribbon を統合するため、全コンポーネントのキーを登録
        // 1. RichTextBoxAdv / DocIO (Word) — IOSD と同じキー
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
            "Ngo9BigBOggjHTQxAR8/V1JGaF5cXGpCf1FpRmJGdld5fUVHYVZUTXxaS00DNHVRdkdlWX1ccXVVRGFfV0JwV0VWYEs=");
        // 2. SfSpreadsheet UI (Excel) — IOSH と同じキー
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
            "IAk8BicRIAEqCzQhAR8kAxMHIgRJXmFXf013TGhYfUFzdUpPaVVYVHdeSFhqQ3taZiUeUn1ecnJVRGJdUEZzXEFaZ0h4Un1GYQ==");
        // 3. XlsIO (Excel library) — IOSH と同じキー
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
            "Ngo9BigBOggjHTQxAR8/V1JGaF5cXGpCf1FpRmJGdld5fUVHYVZUTXxaS00DNHVRdkdlWX1ccXVXQ2ZYVUF2XkBWYEs=");
        // 4. Presentation (PowerPoint) — INSS と同じキー
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
            "Ngo9BigBOggjHTQxAR8/V1JGaF5cXGpCf1FpRmJGdld5fUVHYVZUTXxaS00DNHVRdkdlWX1ccXVXQmBfWEx0V0JWYEs=");

        // カスタム WindowChrome タイトルバーが Syncfusion テーマで上書きされるのを防止
        Syncfusion.SfSkinManager.SfSkinManager.ApplyStylesOnApplication = false;

        LanguageManager.SetLanguage("ja");

        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;
        DispatcherUnhandledException += OnDispatcherUnhandledException;

        try
        {
            var services = ServiceConfiguration.ConfigureServices();
            var licenseManager = services.GetRequiredService<InsightLicenseManager>();
            var mainViewModel = services.GetRequiredService<MainViewModel>();
            var presetService = services.GetRequiredService<PromptPresetService>();
            var recentFiles = services.GetRequiredService<Helpers.RecentFilesService>();

            var mainWindow = new MainWindow(licenseManager, presetService, recentFiles)
                { DataContext = mainViewModel };

            // Wire up ViewModel → Window file open
            mainViewModel.OnFileOpenRequested = filePath => mainWindow.OpenFileByPath(filePath);

            // Handle command-line file open
            if (e.Args.Length > 0 && File.Exists(e.Args[0]))
            {
                mainWindow.Loaded += (_, _) => mainWindow.OpenFileByPath(e.Args[0]);
            }

            mainWindow.Show();
        }
        catch (Exception ex)
        {
            MessageBox.Show(LanguageManager.Format("Error_StartupFailed", ex.Message),
                LanguageManager.Get("Error_StartupError"), MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
        }
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs args)
    {
        try
        {
            MessageBox.Show(
                LanguageManager.Format("Error_Unexpected", args.Exception.Message),
                LanguageManager.Get("Error_Title"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch
        {
            MessageBox.Show(args.Exception.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        args.Handled = true;
    }

    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs args)
    {
        System.Diagnostics.Trace.WriteLine($"[UnobservedTask] {args.Exception?.InnerException?.Message}");
        args.SetObserved();
    }

    private static void OnAppDomainUnhandledException(object sender, UnhandledExceptionEventArgs args)
    {
        if (args.ExceptionObject is Exception ex)
        {
            System.Diagnostics.Trace.WriteLine($"[AppDomain] {ex.Message}");
        }
    }
}
