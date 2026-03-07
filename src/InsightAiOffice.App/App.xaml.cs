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

        // Syncfusion ライセンス登録
        // FindConfigPath は最大4階層上まで探索するが、src/ 配下のプロジェクトは
        // bin/Debug/net8.0-windows/ から5階層上がリポジトリルートのため明示指定
        var configPath = Path.GetFullPath(Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "..",
            "insight-common", "config", "third-party-licenses.json"));
        InsightCommon.License.ThirdPartyLicenseProvider.RegisterSyncfusion("uiEdition",
            File.Exists(configPath) ? configPath : null);

        // Syncfusion テーマ初期化（ライセンスは上で登録済みだが Initialize 内の
        // SfSkinManager.ApplyStylesOnApplication = false の設定が必要）
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
