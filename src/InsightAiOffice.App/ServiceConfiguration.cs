using InsightCommon.AI;
using InsightCommon.License;
using InsightAiOffice.App.Helpers;
using InsightAiOffice.App.ViewModels;
using Microsoft.Extensions.DependencyInjection;

namespace InsightAiOffice.App;

public static class ServiceConfiguration
{
    public static IServiceProvider ConfigureServices()
    {
        var services = new ServiceCollection();

        // License
        services.AddSingleton(_ =>
            new InsightLicenseManager("IAOF", "Insight AI Office"));

        // Prompt Preset Service (PromptEditorDialog 用 — ユーザープリセットの CRUD)
        services.AddSingleton(_ => new PromptPresetService("IAOF", BuiltInPresets.GetAll));

        // Recent Files
        services.AddSingleton(_ => new RecentFilesService("IAOF"));

        // ViewModels
        services.AddTransient(sp => new MainViewModel(
            sp.GetRequiredService<InsightLicenseManager>()));

        return services.BuildServiceProvider();
    }
}
