using System.IO;
using System.Windows;
using System.Windows.Threading;
using InsightCommon.AI;
using InsightCommon.Addon;
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

        // Syncfusion 初期化 (ライセンス登録 + テーマ設定)
        InsightCommon.Theme.SyncfusionInitializer.Initialize();

        LanguageManager.SetLanguage("ja");

        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        AppDomain.CurrentDomain.UnhandledException += OnAppDomainUnhandledException;
        DispatcherUnhandledException += OnDispatcherUnhandledException;

        try
        {
            var services = ServiceConfiguration.ConfigureServices();
            var licenseManager = services.GetRequiredService<InsightLicenseManager>();
            var mainViewModel = services.GetRequiredService<MainViewModel>();

            var aiService = services.GetRequiredService<AiService>();
            var presetService = services.GetRequiredService<PromptPresetService>();
            var referenceService = services.GetRequiredService<ReferenceMaterialsService>();

            var mainWindow = new MainWindow(licenseManager, aiService, presetService, referenceService)
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
