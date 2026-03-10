using System.Windows;
using InsightCommon.AI;
using InsightCommon.License;

namespace InsightAiOffice.App;

public partial class MainWindow : Window
{
    private readonly InsightLicenseManager _licenseManager;
    private readonly PromptPresetService _presetService;
    private readonly Helpers.RecentFilesService _recentFiles;
    private readonly AiChatViewModel _chatVm;
    private bool _isRightPanelOpen;
    private string _activeEditorType = ""; // "word", "excel", "pptx"
    private string _currentDocPath = "";

    // チャットメッセージ管理用（AiChatViewModel の IsSending をフックして管理）
    private string? _pendingUserInput;

    public MainWindow(
        InsightLicenseManager licenseManager,
        PromptPresetService presetService,
        Helpers.RecentFilesService recentFiles)
    {
        InitializeComponent();
        _licenseManager = licenseManager;
        _presetService = presetService;
        _recentFiles = recentFiles;

        // ── AiChatViewModel（insight-common 共通基盤） ──
        _chatVm = new AiChatViewModel(new AiChatViewModelOptions
        {
            ProductCode = "IAOF",
            ProductName = "Insight AI Office",
            GetLanguage = () => Helpers.LanguageManager.CurrentLanguage == "ja" ? "JA" : "EN",
            GetSystemPrompt = BuildSystemPrompt,
            GetBuiltInPresets = () => Helpers.BuiltInPresets.GetPresetPrompts(),
            LicenseManager = _licenseManager,
            EnableConcierge = true,
        });

        // チャットメッセージフロー: IsSending の変化を監視して ChatMessages を管理
        _chatVm.PropertyChanged += OnChatVmPropertyChanged;

        ChatPanel.DataContext = _chatVm;

        // View イベントハンドラ
        ChatPanel.HelpRequested += (_, _) => Views.HelpWindow.ShowSection(this, "ai-assistant");
        ChatPanel.CloseRequested += (_, _) => CloseRightPanel_Click(this, new RoutedEventArgs());
        ChatPanel.PromptEditorRequested += (_, _) => ChatPromptEditor_Click(this, new RoutedEventArgs());
        ChatPanel.InsertToDocumentRequested += InsertAiResponseText;
        ChatPanel.CopyResponseRequested += CopyAiResponseText;

        UpdatePlanBadge();
        UpdateLicenseBackstage();

        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        VersionLabel.Text = $"v{version?.Major}.{version?.Minor}.{version?.Build}";
        VersionText.Text = $"v{version?.Major}.{version?.Minor}.{version?.Build}";

        Loaded += OnWindowLoaded;
    }

    // ── チャットメッセージフロー管理 ──────────────────────────
    // AiChatViewModel がメッセージ追加・Artifact 検出保存を内部で処理するため、
    // ここでは入力クリア・バッジ更新・ステータスバー更新のみ行う。

    private void OnChatVmPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(AiChatViewModel.IsSending)) return;

        Dispatcher.Invoke(() =>
        {
            if (_chatVm.IsSending && _pendingUserInput == null)
            {
                // 実行開始: 入力をクリア（メッセージ追加は AiChatViewModel が行う）
                _pendingUserInput = _chatVm.AiInput?.Trim();
                _chatVm.AiInput = "";

                // ドキュメント名バッジを更新
                ChatPanel.UpdateDocumentBadge(
                    System.IO.Path.GetFileName(_currentDocPath),
                    !string.IsNullOrEmpty(_currentDocPath));
            }
            else if (!_chatVm.IsSending && _pendingUserInput != null)
            {
                // 実行完了（応答追加・Artifact 処理は AiChatViewModel 内部で完了済み）
                _pendingUserInput = null;
                StatusText.Text = Helpers.LanguageManager.Get("Status_Ready");
            }
        });
    }

    // ── システムプロンプト構築 ────────────────────────────────

    private string BuildSystemPrompt(string lang)
    {
        var isJa = lang != "EN";
        var prompt = isJa
            ? "あなたは Insight AI Office のアシスタントです。ユーザーが開いているドキュメントについて質問に答え、分析・校正・要約などを支援してください。回答は簡潔に、日本語で行ってください。"
            : "You are the Insight AI Office assistant. Answer questions about the user's open document. Provide analysis, proofreading, and summaries. Be concise.";

        // デフォルトプリセットがあればそちらを優先
        var defaultPreset = _presetService.GetDefault();
        if (defaultPreset != null && !string.IsNullOrEmpty(defaultPreset.SystemPrompt))
            prompt = defaultPreset.SystemPrompt;

        // ドキュメントコンテキスト
        var docContent = ExtractDocumentContent();
        if (!string.IsNullOrEmpty(docContent))
            prompt += $"\n\n--- {(isJa ? "現在のドキュメント内容" : "Current Document")} ---\n{docContent}\n--- ---";

        return prompt;
    }

    // ── ウィンドウ初期化 ─────────────────────────────────────

    private void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        InsightCommon.Theme.SyncfusionInitializer.ApplyTheme(
            MainRibbon, WordRibbon, ExcelRibbon, PptxRibbon, Spreadsheet);

        InsightCommon.UI.InsightScaleManager.Instance.ApplyToWindow(this);

        InitializeLanguageRadioButtons();
        ApplyLocalization();
        RefreshRecentFilesList();
    }

    private void ApplyLocalization()
    {
        var L = Helpers.LanguageManager.Get;

        // Title bar
        FileNameLabel.Text = L("App_Tagline");
        ChatToggleBtn.ToolTip = L("Pane_Chat");
        System.Windows.Automation.AutomationProperties.SetName(ChatToggleBtn, L("Pane_Chat"));
        PromptToggleBtn.ToolTip = L("Pane_Prompt");
        PromptToggleText.Text = L("Menu_Prompt");
        MinimizeBtn.ToolTip = L("Window_Minimize");
        System.Windows.Automation.AutomationProperties.SetName(MinimizeBtn, L("Window_Minimize"));
        MaximizeButton.ToolTip = L("Window_Maximize");
        System.Windows.Automation.AutomationProperties.SetName(MaximizeButton, L("Window_Maximize"));
        CloseBtn.ToolTip = L("Window_Close");
        System.Windows.Automation.AutomationProperties.SetName(CloseBtn, L("Window_Close"));

        // Chat panel
        _chatVm.RefreshForLanguageChange();
        ChatPanel.RefreshLocalization();

        // Recent files
        BS_RecentTab.Header = L("BS_Recent");
        BS_RecentTitle.Text = L("BS_Recent");
        RecentFilesEmptyHint.Text = L("BS_RecentEmpty");

        // Backstage
        MainRibbon.BackStageHeader = L("Menu_File");
        BS_Open.Header = L("BS_Open");
        BS_Settings.Header = L("BS_Settings");
        BS_LanguageTab.Header = L("BS_Language");
        BS_LanguageTitle.Text = L("BS_Language");
        BS_LicenseTab.Header = L("BS_License");
        LicenseTitleText.Text = L("License_Title");
        CurrentPlanLabel.Text = L("License_CurrentPlan");
        LicenseOpenBtn.Content = L("License_OpenManager");
        BS_CloseDoc.Header = L("BS_CloseDoc");

        // Welcome panel
        WelcomeTagline.Text = L("Welcome_Tagline");
        WelcomeStartLabel.Text = L("Welcome_Start");
        WelcomeOpenLabel.Text = L("File_Open");
        WelcomeOpenDesc.Text = ".docx / .xlsx / .pptx";
        WelcomeChatLabel.Text = L("Welcome_OpenChat");
        WelcomeChatDesc.Text = L("Welcome_ChatDesc");
        WelcomeDragHint.Text = L("Welcome_DragHint");

        // Status bar
        StatusText.Text = L("Status_Ready");

        // Ribbon localization
        ApplyRibbonLocalization();
    }

    private void ApplyRibbonLocalization()
    {
        var L = Helpers.LanguageManager.Get;

        // ── Default Ribbon ──
        MainRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(MainRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_File"), new[] { L("File_Open") }),
            ("AI", new[] { L("AI_Chat"), L("Pane_AiSettings") }),
            (L("Menu_Help"), new[] { L("Menu_Help") }),
        });
        ApplyBackstage(MainRibbon, L("File_Open"), L("BS_Recent"), L("File_Settings"), null, L("License_Title"));

        // ── Word Ribbon ──
        WordRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(WordRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_File"), new[] { L("File_Open"), L("File_Save"), L("File_SaveAs") }),
            (L("Format_UndoRedo"), new[] { L("Format_Undo"), L("Format_Redo") }),
            (L("Format_Font"), new[] { L("Format_Bold"), L("Format_Italic"), L("Format_Underline"), L("Format_Strikethrough") }),
            (L("Format_Paragraph"), new[] { L("Format_AlignLeft"), L("Format_AlignCenter"), L("Format_AlignRight"), L("Format_BulletList") }),
            (L("Format_Insert"), new[] { L("Format_InsertImage") }),
            (L("Format_Edit"), new[] { L("Format_FindReplace") }),
            (L("File_Print"), new[] { L("File_Print") }),
        });
        ApplyTab(WordRibbon, 1, L("Menu_AI"), new[]
        {
            ("AI", new[] { L("AI_Chat"), L("AI_Summarize"), L("AI_Proofread"), L("AI_Analyze") }),
            (L("Pane_Prompt"), new[] { L("Pane_Prompt") }),
            (L("Pane_AiSettings"), new[] { L("Pane_AiSettings") }),
        });
        ApplyBackstage(WordRibbon, L("File_Open"), L("File_SaveOverwrite"), L("File_Print"), null, L("File_CloseDoc"));

        // ── Excel Ribbon ──
        ExcelRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(ExcelRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_File"), new[] { L("File_Open"), L("File_Save") }),
            (L("Format_UndoRedo"), new[] { L("Format_Undo"), L("Format_Redo") }),
            (L("Excel_Clipboard"), new[] { L("Excel_Paste"), L("Excel_Cut"), L("Excel_Copy") }),
            (L("Format_Font"), new[] { L("Format_Bold"), L("Format_Italic"), L("Format_Underline"), L("Excel_Borders") }),
            (L("Excel_Alignment"), new[] { L("Format_AlignLeft"), L("Format_AlignCenter"), L("Format_AlignRight"), L("Excel_WrapText"), L("Excel_MergeCells") }),
            (L("Excel_Number"), new[] { "%", L("Excel_Comma"), L("Excel_Currency"), L("Excel_DecInc"), L("Excel_DecDec") }),
            (L("Excel_View"), new[] { L("Excel_FreezePanes") }),
        });
        ApplyTab(ExcelRibbon, 1, L("Menu_AI"), new[]
        {
            ("AI", new[] { L("AI_Chat"), L("AI_DataAnalysis"), L("AI_FormulaHelp") }),
            (L("Pane_Prompt"), new[] { L("Pane_Prompt") }),
            (L("Pane_AiSettings"), new[] { L("Pane_AiSettings") }),
        });
        ApplyBackstage(ExcelRibbon, L("File_Open"), L("File_SaveAs"), null, L("File_CloseDoc"));

        // ── PPTX Ribbon ──
        PptxRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(PptxRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_File"), new[] { L("File_Open"), L("Pptx_ExportPdf") }),
            (L("Pptx_SlideOps"), new[] { L("Pptx_Add"), L("Pptx_Duplicate"), L("Pptx_Delete"), L("Pptx_MoveUp"), L("Pptx_MoveDown") }),
            (L("Pptx_SlideInfo"), new[] { L("Pptx_ExtractText") }),
        });
        ApplyTab(PptxRibbon, 1, L("Menu_AI"), new[]
        {
            ("AI", new[] { L("AI_Chat"), L("AI_Summarize"), L("Pptx_Proofread"), L("AI_Analyze") }),
            (L("Pane_Prompt"), new[] { L("Pane_Prompt") }),
            (L("Pane_AiSettings"), new[] { L("Pane_AiSettings") }),
        });
        ApplyBackstage(PptxRibbon, L("File_Open"), L("Pptx_ExportPdf"), null, L("File_CloseDoc"));
    }

    private static void ApplyTab(Syncfusion.Windows.Tools.Controls.Ribbon ribbon, int tabIndex,
        string caption, (string header, string[] labels)[] bars)
    {
        if (tabIndex >= ribbon.Items.Count ||
            ribbon.Items[tabIndex] is not Syncfusion.Windows.Tools.Controls.RibbonTab tab) return;
        tab.Caption = caption;
        for (var i = 0; i < bars.Length && i < tab.Items.Count; i++)
        {
            if (tab.Items[i] is not Syncfusion.Windows.Tools.Controls.RibbonBar bar) continue;
            bar.Header = bars[i].header;
            for (var j = 0; j < bars[i].labels.Length && j < bar.Items.Count; j++)
            {
                if (bar.Items[j] is Syncfusion.Windows.Tools.Controls.RibbonButton btn)
                    btn.Label = bars[i].labels[j];
            }
        }
    }

    private static void ApplyBackstage(Syncfusion.Windows.Tools.Controls.Ribbon ribbon, params string?[] headers)
    {
        if (ribbon.BackStage is not Syncfusion.Windows.Tools.Controls.Backstage backstage) return;
        var idx = 0;
        foreach (var item in backstage.Items)
        {
            if (idx >= headers.Length) break;
            switch (item)
            {
                case Syncfusion.Windows.Tools.Controls.BackStageCommandButton cmd:
                    if (headers[idx] != null) cmd.Header = headers[idx]!;
                    idx++;
                    break;
                case Syncfusion.Windows.Tools.Controls.BackstageTabItem tab:
                    if (headers[idx] != null) tab.Header = headers[idx]!;
                    idx++;
                    break;
                case Syncfusion.Windows.Tools.Controls.BackStageSeparator:
                    if (headers[idx] == null) idx++;
                    break;
            }
        }
    }

    // ── Ribbon Switching ──────────────────────────────────────────

    private void SwitchRibbon(string editorType)
    {
        DefaultRibbonPanel.Visibility = Visibility.Collapsed;
        WordRibbonPanel.Visibility = Visibility.Collapsed;
        ExcelRibbonPanel.Visibility = Visibility.Collapsed;
        PptxRibbonPanel.Visibility = Visibility.Collapsed;

        switch (editorType)
        {
            case "word":
                WordRibbonPanel.Visibility = Visibility.Visible;
                break;
            case "excel":
                ExcelRibbonPanel.Visibility = Visibility.Visible;
                break;
            case "pptx":
                PptxRibbonPanel.Visibility = Visibility.Visible;
                break;
            default:
                DefaultRibbonPanel.Visibility = Visibility.Visible;
                break;
        }
    }
}
