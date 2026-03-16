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

        // --generate-samples: サンプル出力ファイル再生成モード（開発用）
        if (e.Args.Length > 0 && e.Args[0] == "--generate-samples")
        {
            // Syncfusion ライセンスを先に登録してからサンプル生成
            RegisterSyncfusionLicenses();
            var assetsDir = e.Args.Length > 1
                ? e.Args[1]
                : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets");
            Console.WriteLine("サンプル出力ファイルを再生成しています...");
            Tools.SampleOutputGenerator.GenerateToSource(assetsDir);
            Console.WriteLine("完了しました。");
            Shutdown(0);
            return;
        }

        RegisterSyncfusionLicenses();

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

    /// <summary>Syncfusion ライセンス登録（33.x — 各コンポーネント）</summary>
    private static void RegisterSyncfusionLicenses()
    {
        // 1. UI Edition（Ribbon / SfSkinManager / 共通UI）
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
            "Ngo9BigBOggjGyl/VkV+XU9AclRDX3xKf0x/TGpQb19xflBPallYVBYiSV9jS3hTdURlWXpfeXZVRmVfVk91XA==");
        // 2. DOCX Editor（SfRichTextBoxAdv / DocIO）
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
            "NxYtGyMROh0gHDMgDk1jX09FaFtGVmJLYVB3WmpQdldgdVRMZVVbQX9PIiBoS35RcEVgWHleeXRTRWBeUEJzVEFe");
        // 3. Spreadsheet Editor（SfSpreadsheet / XlsIO）
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
            "IAk8BicRIAEqCzQhAR8kAxMHIgRJXmBXf01yQWhYfUFzdUpPaVVYVHdeSFhqQ3taZiUeUn1ecnJVRWFaUU12WUdbYEx8Un1GYA==");
        // 4. PDF Viewer
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
            "Ix0oFS8QJAw9HSQvXkVjQlBacltJXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxVdkZiX39XcHNUR2JeUE19XEE=");
        // 5. Document SDK（DocToPDF / Presentation / Pdf）
        Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
            "NxYtFisQPR08Cit/VkV+XU9AclRDX3xKf0x/TGpQb19xflBPallYVBYiSV9jS3hTdURlWXZed3dcQmBVWU91XA==");
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
