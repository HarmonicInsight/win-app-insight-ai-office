using System.Windows;
using System.Windows.Controls;
using InsightAiOffice.App.Services;
using InsightCommon.AI;

namespace InsightAiOffice.App;

public partial class MainWindow
{
    private DocumentToolExecutor? _toolExecutor;
    private bool _pendingDocGeneration;

    private DocumentToolExecutor GetToolExecutor()
    {
        return _toolExecutor ??= new DocumentToolExecutor(
            insertText: text =>
            {
                if (_activeEditorType == "word" && RichTextEditor.Selection != null)
                    RichTextEditor.Selection.InsertText(text);
            },
            getSelectedText: () =>
            {
                if (_activeEditorType == "word" && RichTextEditor.Selection != null)
                    return RichTextEditor.Selection.Text ?? "";
                return "";
            },
            setStatus: msg => StatusText.Text = msg
        );
    }

    /// <summary>Intercepts chat messages for pending document generation.</summary>
    private async Task<bool> OnBeforeChatSend(string input)
    {
        if (!_pendingDocGeneration) return false;
        await GenerateDocumentFromPrompt(input);
        return true;
    }

    // ── AI One-Click Actions ──────────────────────────────────────

    private async void AiSummarize_Click(object sender, RoutedEventArgs e)
    {
        var preset = _presetService.LoadAll().Find(p => p.Id == "builtin_summarize");
        var prompt = preset?.SystemPrompt ?? "以下のドキュメントの内容を要約してください。重要なポイントを箇条書きで示してください。";
        if (preset != null) _presetService.IncrementUsage(preset.Id);
        await ExecuteAiAction(prompt);
    }

    private async void AiAnalyze_Click(object sender, RoutedEventArgs e)
    {
        var preset = _presetService.LoadAll().Find(p => p.Id == "builtin_analyze");
        var prompt = preset?.SystemPrompt ?? "以下のドキュメントの構造と内容を分析してください。";
        if (preset != null) _presetService.IncrementUsage(preset.Id);
        await ExecuteAiAction(prompt);
    }

    private async void AiProofread_Click(object sender, RoutedEventArgs e)
    {
        var preset = _presetService.LoadAll().Find(p => p.Id == "builtin_proofread");
        var prompt = preset?.SystemPrompt ?? "以下のドキュメントを校正してください。";
        if (preset != null) _presetService.IncrementUsage(preset.Id);
        await ExecuteAiAction(prompt);
    }

    private async Task ExecuteAiAction(string prompt)
    {
        var L = Helpers.LanguageManager.Get;
        if (!_aiService.IsConfigured)
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

        await _chatVm.ExecuteActionAsync(prompt);
    }

    // ── AI Settings ───────────────────────────────────────────────

    private void AiSettings_Click(object sender, RoutedEventArgs e)
    {
        ShowAiSettingsDialog();
    }

    private void ShowAiSettingsDialog()
    {
        var theme = InsightCommon.Theme.InsightTheme.Create();
        _aiService.ShowSettingsDialog(this, theme, "Insight AI Office",
            Helpers.LanguageManager.CurrentLanguage);
        StatusText.Text = _aiService.IsConfigured
            ? Helpers.LanguageManager.Format("AI_ConfigDone", _aiService.CurrentModel)
            : Helpers.LanguageManager.Get("AI_ConfigNotSet");
        _chatVm.NotifyApiKeyChanged();
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
            _chatVm.LoadPresetGroups();
            StatusText.Text = Helpers.LanguageManager.Get("AI_PresetUpdated");
        }
    }

    // ── AI Excel Formula Action ───────────────────────────────────

    private async void AiExcelFormula_Click(object sender, RoutedEventArgs e)
    {
        var preset = _presetService.LoadAll().Find(p => p.Id == "builtin_excel_formula");
        var prompt = preset?.SystemPrompt ?? "ユーザーが求める計算・集計を実現する Excel 数式を提案してください。";
        if (preset != null) _presetService.IncrementUsage(preset.Id);

        if (!_isRightPanelOpen) ToggleRightPanel();

        var docContent = ExtractDocumentContent();
        var systemPrompt = prompt + (string.IsNullOrEmpty(docContent) ? "" : $"\n\n--- データ ---\n{docContent}");
        await _chatVm.ExecuteWithSystemPromptAsync(
            "現在開いているスプレッドシートのデータに基づいて、便利な数式を提案してください。",
            systemPrompt);
    }

    // ── AI Document Generation ────────────────────────────────────

    private async void AiGenerateDoc_Click(object sender, RoutedEventArgs e) => await AiGenerateDocument();

    private async Task AiGenerateDocument()
    {
        var L = Helpers.LanguageManager.Get;
        if (!_aiService.IsConfigured)
        {
            StatusText.Text = L("AI_ConfigShort");
            return;
        }

        if (!_isRightPanelOpen) ToggleRightPanel();

        _chatVm.AddAssistantMessage("ドキュメントを自動生成します。チャットにテーマを入力してください。\n例: 「来月の営業報告書を作成して」「プロジェクト計画書のテンプレート」");
        _pendingDocGeneration = true;
        StatusText.Text = L("AI_InputTheme");
    }

    private async Task GenerateDocumentFromPrompt(string userPrompt)
    {
        var L = Helpers.LanguageManager.Get;
        _pendingDocGeneration = false;

        var refContext = _referenceService.Count > 0
            ? _referenceService.BuildAutoContext(userPrompt, 5)
            : "";

        var systemPrompt = "あなたはビジネス文書作成の専門家です。" +
            "ユーザーのリクエストに基づいて、Word ドキュメント用の本文テキストを生成してください。" +
            "見出しは行頭に ## を付けて示し、箇条書きは - を使ってください。" +
            "回答はドキュメント本文のみを返してください（前置きや説明は不要）。";

        if (!string.IsNullOrEmpty(refContext))
            systemPrompt += $"\n\n{refContext}";

        await _chatVm.ExecuteWithSystemPromptAsync(userPrompt, systemPrompt);

        var lastResponse = _chatVm.GetLastAssistantMessage();
        if (!string.IsNullOrEmpty(lastResponse))
        {
            var displayName = $"AI生成_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            LoadMarkdownIntoWordEditor(lastResponse, displayName);
            StatusText.Text = L("AI_DocGenerated");
        }
    }

    // ── AI Report Generation (Excel → Word) ─────────────────────

    private async void AiGenerateReport_Click(object sender, RoutedEventArgs e) => await AiGenerateReport();

    private async Task AiGenerateReport()
    {
        var L = Helpers.LanguageManager.Get;
        if (!_aiService.IsConfigured)
        {
            StatusText.Text = L("AI_ConfigShort");
            return;
        }

        var excelData = ExtractExcelContent();
        if (string.IsNullOrEmpty(excelData))
        {
            StatusText.Text = L("Doc_OpenExcelFirst");
            return;
        }

        if (!_isRightPanelOpen) ToggleRightPanel();

        var refContext = _referenceService.Count > 0
            ? _referenceService.BuildAutoContext("レポート生成", 5)
            : "";

        var systemPrompt = "あなたはデータ分析レポートの専門家です。" +
            "提供されたスプレッドシートのデータを分析し、ビジネスレポートを作成してください。" +
            "レポートには以下を含めてください:\n" +
            "# レポートタイトル\n" +
            "## 概要（データの全体像）\n" +
            "## 主要な指標・数値\n" +
            "## 分析結果・傾向\n" +
            "## 提言・次のアクション\n\n" +
            "見出しは ## で、箇条書きは - で示してください。";

        if (!string.IsNullOrEmpty(refContext))
            systemPrompt += $"\n\n{refContext}";

        systemPrompt += $"\n\n--- スプレッドシートデータ ---\n{excelData}\n--- ここまで ---";

        await _chatVm.ExecuteWithSystemPromptAsync(
            "このスプレッドシートデータからビジネスレポートを作成してください。", systemPrompt);

        var lastResponse = _chatVm.GetLastAssistantMessage();
        if (!string.IsNullOrEmpty(lastResponse))
        {
            var displayName = $"AI_Report_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            LoadMarkdownIntoWordEditor(lastResponse, displayName);
            StatusText.Text = L("AI_ReportGenerated");
        }
    }

    /// <summary>
    /// Renders markdown-formatted text into a new Word document in the RichTextEditor.
    /// </summary>
    private void LoadMarkdownIntoWordEditor(string markdownText, string displayName)
    {
        CloseAllEditors();
        WordEditorPanel.Visibility = Visibility.Visible;
        _activeEditorType = "word";
        SwitchRibbon("word");

        RichTextEditor.Document?.Sections.Clear();
        var doc = RichTextEditor.Document;
        if (doc != null)
        {
            var section = new Syncfusion.Windows.Controls.RichTextBoxAdv.SectionAdv();
            doc.Sections.Add(section);

            foreach (var line in markdownText.Split('\n'))
            {
                var trimmed = line.TrimEnd();
                if (string.IsNullOrWhiteSpace(trimmed)) continue;

                var para = new Syncfusion.Windows.Controls.RichTextBoxAdv.ParagraphAdv();
                section.Blocks.Add(para);

                var text = trimmed;
                var isBold = false;
                var fontSize = 11.0;

                if (trimmed.StartsWith("## "))
                {
                    text = trimmed[3..];
                    isBold = true;
                    fontSize = 16.0;
                }
                else if (trimmed.StartsWith("# "))
                {
                    text = trimmed[2..];
                    isBold = true;
                    fontSize = 20.0;
                }
                else if (trimmed.StartsWith("- ") || trimmed.StartsWith("* "))
                {
                    text = "  \u2022 " + trimmed[2..];
                }

                var span = new Syncfusion.Windows.Controls.RichTextBoxAdv.SpanAdv();
                span.Text = text;
                span.CharacterFormat.FontSize = fontSize;
                span.CharacterFormat.Bold = isBold;
                para.Inlines.Add(span);
            }
        }

        WordFileName.Text = displayName;
        FileNameLabel.Text = displayName;
        FileTypeLabel.Text = "DOCX";
        _currentDocPath = "";

        if (DataContext is ViewModels.MainViewModel vm)
        {
            vm.DocumentTitle = displayName;
            vm.IsFileLoaded = true;
        }
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

    // ── Cross-Format AI Analysis ──────────────────────────────────

    private async void CrossFormatAnalysis_Click(object sender, RoutedEventArgs e) => await CrossFormatAnalysis();

    private async Task CrossFormatAnalysis()
    {
        var L = Helpers.LanguageManager.Get;
        if (!_aiService.IsConfigured)
        {
            StatusText.Text = L("AI_ConfigShort");
            return;
        }

        var docContent = ExtractDocumentContent();
        var refContext = _referenceService.Count > 0
            ? _referenceService.BuildAutoContext("クロスフォーマット分析", 10)
            : "";

        if (string.IsNullOrEmpty(docContent) && string.IsNullOrEmpty(refContext))
        {
            StatusText.Text = L("Doc_NeedBoth");
            return;
        }

        if (!_isRightPanelOpen) ToggleRightPanel();

        var systemPrompt = "あなたは Insight AI Office のアシスタントです。" +
            "複数のドキュメント（Word/Excel/PowerPoint）を横断的に分析し、" +
            "データの関連性、矛盾点、改善提案を提示してください。回答は簡潔に、日本語で行ってください。";

        if (!string.IsNullOrEmpty(docContent))
            systemPrompt += $"\n\n--- 現在のドキュメント ---\n{docContent}\n--- ここまで ---";
        if (!string.IsNullOrEmpty(refContext))
            systemPrompt += $"\n\n{refContext}";

        await _chatVm.ExecuteWithSystemPromptAsync(
            "現在開いているドキュメントと参考資料を横断的に分析してください。データの関連性、矛盾点、改善提案を示してください。",
            systemPrompt);
    }
}
