using System.Windows;
using System.Windows.Controls;
using InsightCommon.AI;

namespace InsightAiOffice.App;

public partial class MainWindow
{
    // ── AI One-Click Actions ──────────────────────────────────────

    private void AiSummarize_Click(object sender, RoutedEventArgs e) =>
        ExecuteAiPresetAction("preset_summarize",
            "以下のドキュメントの内容を要約してください。重要なポイントを箇条書きで示してください。");

    private void AiAnalyze_Click(object sender, RoutedEventArgs e) =>
        ExecuteAiPresetAction("preset_analyze",
            "以下のドキュメントの構造と内容を分析してください。");

    private void AiProofread_Click(object sender, RoutedEventArgs e) =>
        ExecuteAiPresetAction("preset_proofread",
            "以下のドキュメントを校正してください。");

    private void ExecuteAiPresetAction(string presetId, string fallbackPrompt)
    {
        var L = Helpers.LanguageManager.Get;
        if (!_chatVm.HasApiKey)
        {
            StatusText.Text = L("AI_ConfigShort");
            return;
        }

        var docContent = ExtractDocumentContent();
        if (string.IsNullOrEmpty(docContent))
        {
            StatusText.Text = L("Doc_OpenFirst");
            return;
        }

        if (!_isRightPanelOpen)
            ToggleRightPanel();

        // プリセットツリーからプロンプトを検索して入力にセット
        var lang = Helpers.LanguageManager.CurrentLanguage == "ja" ? "JA" : "EN";
        var preset = Helpers.BuiltInPresets.GetPresetPrompts()
            .FirstOrDefault(p => p.Id == presetId);
        var prompt = preset?.GetPrompt(lang) ?? fallbackPrompt;

        _chatVm.AiInput = prompt;
        StatusText.Text = L("AI_Responding");
        _chatVm.ExecuteFromInputCommand.Execute(null);
    }

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

    // ── Prompt Preset Management ─────────────────────────────────

    private void TogglePromptPane_Click(object sender, RoutedEventArgs e)
    {
        var theme = InsightCommon.Theme.InsightTheme.Create();
        var dialog = new PromptPresetManagerDialog(new PromptPresetManagerOptions
        {
            Theme = theme,
            Locale = Helpers.LanguageManager.CurrentLanguage,
            PresetService = _presetService,
        });
        dialog.Owner = this;
        dialog.ShowDialog();

        if (dialog.Changed)
        {
            _chatVm.RefreshUserPromptGroups();
            StatusText.Text = Helpers.LanguageManager.Get("AI_PresetUpdated");
        }
    }

    // ── AI Excel Formula Action ───────────────────────────────────

    private void AiExcelFormula_Click(object sender, RoutedEventArgs e)
    {
        if (!_chatVm.HasApiKey)
        {
            StatusText.Text = Helpers.LanguageManager.Get("AI_ConfigShort");
            return;
        }

        if (!_isRightPanelOpen) ToggleRightPanel();

        var lang = Helpers.LanguageManager.CurrentLanguage == "ja" ? "JA" : "EN";
        var preset = Helpers.BuiltInPresets.GetPresetPrompts()
            .FirstOrDefault(p => p.Id == "preset_excel_formula");
        var prompt = preset?.GetPrompt(lang)
            ?? "ユーザーが求める計算・集計を実現する Excel 数式を提案してください。";

        _chatVm.AiInput = "現在開いているスプレッドシートのデータに基づいて、便利な数式を提案してください。\n\n" + prompt;
        _chatVm.ExecuteFromInputCommand.Execute(null);
    }

    // ── AI Data Analysis Action ──────────────────────────────────

    private void AiDataAnalysis_Click(object sender, RoutedEventArgs e) =>
        ExecuteAiPresetAction("preset_data_insight",
            "以下のスプレッドシートのデータを分析し、重要なインサイトを提示してください。");

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
