using System.Windows;
using System.Windows.Controls;
using InsightCommon.AI;

namespace InsightAiOffice.App;

public partial class MainWindow
{
    // ── AI Settings ───────────────────────────────────────────────

    private void AiSettings_Click(object sender, RoutedEventArgs e)
    {
        ShowAiSettingsDialog();
    }

    private void ShowAiSettingsDialog()
    {
        var theme = InsightCommon.Theme.InsightTheme.Create();
        _chatVm.AiService.ShowSettingsDialog(this, theme, "Insight AI Office",
            Helpers.LanguageManager.CurrentLanguage);
        StatusText.Text = _chatVm.HasApiKey
            ? Helpers.LanguageManager.Format("AI_ConfigDone", _chatVm.AiService.CurrentModel)
            : Helpers.LanguageManager.Get("AI_ConfigNotSet");
    }

    // ── AI Write-Back ─────────────────────────────────────────────

    private void InsertAiResponseText(string text)
    {
        var L = Helpers.LanguageManager.Get;
        if (_activeEditorType == "word")
        {
            try
            {
                var editor = RichTextEditor;
                if (editor.Selection != null)
                {
                    editor.Selection.InsertText(text);
                    StatusText.Text = L("AI_InsertedToDoc");
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = $"{L("Error_Title")}: {ex.Message}";
            }
        }
        else
        {
            System.Windows.Clipboard.SetText(text);
            StatusText.Text = L("AI_CopiedFallback");
        }
    }

    private void CopyAiResponseText(string text)
    {
        System.Windows.Clipboard.SetText(text);
        StatusText.Text = Helpers.LanguageManager.Get("AI_CopiedToClipboard");
    }
}
