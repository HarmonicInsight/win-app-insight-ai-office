using InsightCommon.AI;
using InsightCommon.Addon;
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

        // AI Services
        services.AddSingleton(_ => new AiService("IAOF"));
        services.AddSingleton(_ => new PromptPresetService("IAOF", BuiltInPresets.GetAll));
        services.AddSingleton(_ => new ReferenceMaterialsService("InsightAiOffice"));

        // ViewModels
        services.AddTransient(sp => new MainViewModel(
            sp.GetRequiredService<InsightLicenseManager>()));

        return services.BuildServiceProvider();
    }
}
