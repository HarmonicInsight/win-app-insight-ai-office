using System.Windows;
using System.Windows.Controls;
using InsightAiOffice.App.Helpers;
using InsightCommon.AI;

namespace InsightAiOffice.App.Views;

public partial class SettingsWindow : Window
{
    private readonly string _initialLanguage;
    private readonly AiService? _aiService;

    public SettingsWindow(AiService? aiService = null)
    {
        InitializeComponent();
        _initialLanguage = LanguageManager.CurrentLanguage;
        _aiService = aiService;

        // Select current language
        foreach (ComboBoxItem item in LanguageCombo.Items)
        {
            if ((string)item.Tag == _initialLanguage)
            {
                LanguageCombo.SelectedItem = item;
                break;
            }
        }

        ApplyLocalization();
    }

    private void ApplyLocalization()
    {
        var L = LanguageManager.Get;
        HeaderText.Text = L("File_Settings");
        AiSettingsLabel.Text = L("Settings_AiSettings");
        AiSettingsDesc.Text = L("Settings_AiSettingsDesc");
        OpenAiSettingsBtn.Content = L("Settings_OpenAiSettings");
    }

    private void LanguageCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LanguageCombo.SelectedItem is ComboBoxItem selected)
        {
            var lang = (string)selected.Tag;
            LanguageManager.SetLanguage(lang);

            RestartHint.Text = lang != _initialLanguage
                ? (lang == "ja"
                    ? "一部の変更はアプリ再起動後に反映されます"
                    : "Some changes will take effect after restarting the app")
                : "";

            ApplyLocalization();
        }
    }

    private void OpenAiSettings_Click(object sender, RoutedEventArgs e)
    {
        if (_aiService == null) return;

        var theme = InsightCommon.Theme.InsightTheme.Create();
        _aiService.ShowSettingsDialog(this, theme, "Insight AI Office",
            LanguageManager.CurrentLanguage);
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
