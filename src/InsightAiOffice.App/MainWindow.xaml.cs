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
    private List<Views.AttachedFileInfo> _chatAttachedFiles = new();
    private string _artifactDir = GetDefaultArtifactDir();

    private static string GetDefaultArtifactDir()
    {
        var dir = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "HarmonicInsight", "InsightAiOffice", "Artifacts", "Default");
        System.IO.Directory.CreateDirectory(dir);
        return dir;
    }

    private void UpdateArtifactDir()
    {
        if (!string.IsNullOrEmpty(_currentDocPath))
        {
            var projectName = System.IO.Path.GetFileNameWithoutExtension(_currentDocPath);
            var safeName = string.Join("_", projectName.Split(System.IO.Path.GetInvalidFileNameChars()));
            _artifactDir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "HarmonicInsight", "InsightAiOffice", "Artifacts", safeName);
        }
        else
        {
            _artifactDir = GetDefaultArtifactDir();
        }
        System.IO.Directory.CreateDirectory(_artifactDir);
    }

    // ── Multi-Tab ──
    private readonly Dictionary<string, Models.DocumentTab> _openTabs = new();
    private readonly List<string> _tabOrder = new();

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
            GetToolDefinitions = () => Services.DocumentGeneration.FileGenerationToolDefinitions.GetAllTools(),
            CreateToolExecutor = CreateToolExecutor,
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
        ChatPanel.FilesAttached += OnFilesAttached;
        ChatPanel.OpenArtifactFolderRequested += _ =>
        {
            UpdateArtifactDir();
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(_artifactDir) { UseShellExecute = true });
        };

        UpdatePlanBadge();
        UpdateLicenseBackstage();

        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        VersionLabel.Text = $"v{version?.Major}.{version?.Minor}.{version?.Build}";
        VersionText.Text = $"v{version?.Major}.{version?.Minor}.{version?.Build}";

        Loaded += OnWindowLoaded;
        Closed += OnWindowClosed;
    }

    private void OnWindowClosed(object? sender, EventArgs e)
    {
        _chatVm.PropertyChanged -= OnChatVmPropertyChanged;
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
                // 実行完了
                _pendingUserInput = null;
                _chatAttachedFiles.Clear();
                ChatPanel.ClearAttachments();
                StatusText.Text = Helpers.LanguageManager.Get("Status_Ready");
            }
        });
    }

    private void OnFilesAttached(List<Views.AttachedFileInfo> files)
    {
        _chatAttachedFiles = files;
    }

    // ── ツール実行ハンドラ ───────────────────────────────────

    private InsightCommon.AI.IToolExecutor CreateToolExecutor()
    {
        // SfRichTextBoxAdv の検索・置換は Selection API 経由で行う
        bool DoReplace(string find, string replacement)
        {
            return Dispatcher.Invoke(() =>
            {
                try
                {
                    if (RichTextEditor == null) return false;
                    // FindAndReplace はコマンド経由
                    var searchResult = RichTextEditor.Find(find, Syncfusion.Windows.Controls.RichTextBoxAdv.FindOptions.None);
                    if (searchResult == null) return false;
                    RichTextEditor.Selection.InsertText(replacement);
                    return true;
                }
                catch { return false; }
            });
        }

        var callbacks = new Services.DocumentGeneration.DocumentEditorCallbacks
        {
            MarkCorrection = (original, correction, reason) =>
            {
                var replaced = $"【削除: {original}】→ {correction}" + (reason != null ? $" ({reason})" : "");
                return DoReplace(original, replaced);
            },
            AddComment = (targetText, comment) =>
            {
                return DoReplace(targetText, $"{targetText} 【※{comment}】");
            },
            HighlightText = (targetText, color) =>
            {
                return DoReplace(targetText, $"★{targetText}★");
            },
            FindAndReplace = (find, replace) =>
            {
                int count = 0;
                while (DoReplace(find, replace)) count++;
                return count;
            },
        };

        UpdateArtifactDir();
        return new Services.DocumentGeneration.DocumentGenerationToolExecutor(_artifactDir, callbacks);
    }

    // ── システムプロンプト構築 ────────────────────────────────

    private string BuildSystemPrompt(string lang)
    {
        var isJa = lang != "EN";
        var prompt = isJa
            ? """
              あなたは Insight AI Office のアシスタントです。
              ユーザーが開いているドキュメントについて質問に答え、分析・校正・要約などを支援してください。

              【ドキュメント生成ツール】
              - generate_report: Word/HTMLレポートを生成（theme対応）
              - generate_spreadsheet: Excelを生成（theme対応）
              - generate_presentation: PowerPointを生成
              - rewrite_document: 添付Wordの書式を維持して内容を書き換え

              【ドキュメント編集ツール（開いているファイルに直接操作）】
              - mark_correction: テキストに赤入れ（取り消し線＋修正案）
              - add_comment: テキストにコメント追加
              - highlight_text: テキストを蛍光マーカーで強調
              - find_and_replace: テキストの検索・置換

              【ツール呼び出しルール】
              ユーザーが校正・赤入れ・修正・ドキュメント作成を依頼した場合、
              必ず適切なツールを実際に呼び出してください。説明だけで終わらないこと。
              """
            : "You are the Insight AI Office assistant. Answer questions about the user's open document. Provide analysis, proofreading, and summaries. Be concise.";

        // デフォルトプリセットがあればそちらを優先
        var defaultPreset = _presetService.GetDefault();
        if (defaultPreset != null && !string.IsNullOrEmpty(defaultPreset.SystemPrompt))
            prompt = defaultPreset.SystemPrompt;

        // ドキュメントコンテキスト（エディタで開いているファイル）
        var docContent = ExtractDocumentContent();
        if (!string.IsNullOrEmpty(docContent))
            prompt += $"\n\n--- {(isJa ? "現在のドキュメント内容" : "Current Document")} ---\n{docContent}\n--- ---";

        // 添付ファイルコンテキスト
        if (_chatAttachedFiles.Count > 0)
        {
            prompt += isJa
                ? "\n\n【添付ファイル】\n以下のファイルが添付されています。内容を分析してください。\nテンプレートベース操作では、添付ファイルのフルパスをツールの template_path / source_path に指定してください。\n"
                : "\n\n[Attached Files]\n";

            foreach (var f in _chatAttachedFiles)
            {
                prompt += $"- {f.FileName} (path: {f.FullPath})\n";
                try
                {
                    var ext = System.IO.Path.GetExtension(f.FullPath).ToLowerInvariant();
                    if (ext is ".txt" or ".csv" or ".md")
                    {
                        var text = System.IO.File.ReadAllText(f.FullPath);
                        if (text.Length > 50000) text = text[..50000] + "\n...(truncated)";
                        prompt += $"\n```\n{text}\n```\n";
                    }
                }
                catch { /* ignore read errors */ }
            }
        }

        // テーマカラー
        var theme = ChatPanel?.SelectedTheme ?? "gold";
        prompt += isJa
            ? $"\n\n【カラーテーマ】\n全ドキュメント生成ツールに theme パラメータがあります。現在のテーマ: {theme}"
            : $"\n\n[Color Theme] Current: {theme}";

        return prompt;
    }

    // ── ウィンドウ初期化 ─────────────────────────────────────

    private void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        InsightCommon.Theme.SyncfusionInitializer.ApplyTheme(
            MainRibbon, WordRibbon, ExcelRibbon, PptxRibbon, Spreadsheet, RichTextEditor);

        InsightCommon.UI.InsightScaleManager.Instance.ApplyToWindow(this);

        InitializeLanguageRadioButtons();
        ApplyLocalization();
        RefreshRecentFilesList();

        // バックステージが開くたびに「最近使ったファイル」タブを選択
        if (MainRibbon.BackStage is Syncfusion.Windows.Tools.Controls.Backstage backstage)
            backstage.IsVisibleChanged += (_, _) => { if (backstage.IsVisible) BS_RecentTab.IsSelected = true; };
        if (WordRibbon.BackStage is Syncfusion.Windows.Tools.Controls.Backstage wordBs)
            wordBs.IsVisibleChanged += (_, _) => { if (wordBs.IsVisible) WordRecentTab.IsSelected = true; };
        if (ExcelRibbon.BackStage is Syncfusion.Windows.Tools.Controls.Backstage excelBs)
            excelBs.IsVisibleChanged += (_, _) => { if (excelBs.IsVisible) ExcelRecentTab.IsSelected = true; };
        if (PptxRibbon.BackStage is Syncfusion.Windows.Tools.Controls.Backstage pptxBs)
            pptxBs.IsVisibleChanged += (_, _) => { if (pptxBs.IsVisible) PptxRecentTab.IsSelected = true; };

        // デフォルトで AI コンシェルジュ（右パネル）を開く
        if (!_isRightPanelOpen)
            ToggleRightPanel();
    }

    private void ApplyLocalization()
    {
        var L = Helpers.LanguageManager.Get;

        // Title bar
        FileNameLabel.Text = L("App_Tagline");
        ChatToggleBtn.ToolTip = L("Pane_Chat");
        ChatToggleText.Text = L("Pane_Chat");
        System.Windows.Automation.AutomationProperties.SetName(ChatToggleBtn, L("Pane_Chat"));
        MinimizeBtn.ToolTip = L("Window_Minimize");
        System.Windows.Automation.AutomationProperties.SetName(MinimizeBtn, L("Window_Minimize"));
        MaximizeButton.ToolTip = L("Window_Maximize");
        System.Windows.Automation.AutomationProperties.SetName(MaximizeButton, L("Window_Maximize"));
        CloseBtn.ToolTip = L("Window_Close");
        System.Windows.Automation.AutomationProperties.SetName(CloseBtn, L("Window_Close"));

        // Chat panel
        _chatVm.RefreshForLanguageChange();
        ChatPanel.RefreshLocalization();

        // Recent files (all backstages)
        LocalizeRecentTab(BS_RecentTab, BS_RecentTitle, RecentFilesEmptyHint);
        LocalizeRecentTab(WordRecentTab, WordRecentTitle, WordRecentEmptyHint);
        LocalizeRecentTab(ExcelRecentTab, ExcelRecentTitle, ExcelRecentEmptyHint);
        LocalizeRecentTab(PptxRecentTab, PptxRecentTitle, PptxRecentEmptyHint);

        // Backstage (main)
        MainRibbon.BackStageHeader = L("Menu_File");
        BS_Open.Header = L("BS_Open");
        BS_Settings.Header = L("BS_Settings");
        BS_LanguageTab.Header = L("BS_Language");
        BS_LanguageTitle.Text = L("BS_Language");
        BS_CloseDoc.Header = L("BS_CloseDoc");

        // License + Language tabs (all backstages)
        LocalizeLicenseTab(BS_LicenseTab, LicenseTitleText, CurrentPlanLabel, LicenseOpenBtn);
        LocalizeLicenseTab(WordLicenseTab, WordLicenseTitleText, WordCurrentPlanLabel, WordLicenseOpenBtn);
        LocalizeLicenseTab(ExcelLicenseTab, ExcelLicenseTitleText, ExcelCurrentPlanLabel, ExcelLicenseOpenBtn);
        LocalizeLicenseTab(PptxLicenseTab, PptxLicenseTitleText, PptxCurrentPlanLabel, PptxLicenseOpenBtn);

        WordLanguageTab.Header = L("BS_Language");
        ExcelLanguageTab.Header = L("BS_Language");
        PptxLanguageTab.Header = L("BS_Language");

        // Welcome panel
        WelcomeTagline.Text = L("Welcome_Tagline");
        WelcomeEditTitle.Text = L("Welcome_EditTitle");
        WelcomeEditDesc.Text = L("Welcome_EditDesc");
        WelcomeCreateTitle.Text = L("Welcome_CreateTitle");
        WelcomeCreateDesc.Text = L("Welcome_CreateDesc");
        WelcomeOpenLabel.Text = L("Welcome_OpenLabel");
        WelcomeOpenDesc.Text = L("Welcome_OpenDesc");
        WelcomeChatLabel.Text = L("Welcome_ChatLabel");
        WelcomeChatDesc.Text = L("Welcome_ChatDescNew");
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
        ApplyBackstage(ExcelRibbon, L("File_Open"), L("File_SaveAs"), null, L("File_CloseDoc"));

        // ── PPTX Ribbon ──
        PptxRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(PptxRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_File"), new[] { L("File_Open"), L("Pptx_ExportPdf") }),
            (L("Pptx_SlideOps"), new[] { L("Pptx_Add"), L("Pptx_Duplicate"), L("Pptx_Delete"), L("Pptx_MoveUp"), L("Pptx_MoveDown") }),
            (L("Pptx_SlideInfo"), new[] { L("Pptx_ExtractText") }),
        });
        ApplyBackstage(PptxRibbon, L("File_Open"), L("Pptx_ExportPdf"), null, L("File_CloseDoc"));
    }

    private static void LocalizeRecentTab(
        Syncfusion.Windows.Tools.Controls.BackstageTabItem tab,
        System.Windows.Controls.TextBlock title,
        System.Windows.Controls.TextBlock emptyHint)
    {
        var L = Helpers.LanguageManager.Get;
        tab.Header = L("BS_Recent");
        title.Text = L("BS_Recent");
        emptyHint.Text = L("BS_RecentEmpty");
    }

    private static void LocalizeLicenseTab(
        Syncfusion.Windows.Tools.Controls.BackstageTabItem tab,
        System.Windows.Controls.TextBlock titleText,
        System.Windows.Controls.TextBlock planLabel,
        System.Windows.Controls.Button openBtn)
    {
        var L = Helpers.LanguageManager.Get;
        tab.Header = L("BS_License");
        titleText.Text = L("License_Title");
        planLabel.Text = L("License_CurrentPlan");
        openBtn.Content = L("License_OpenManager");
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
        PdfRibbonPanel.Visibility = Visibility.Collapsed;

        Syncfusion.SfSkinManager.SfSkinManager.SetTheme(MainRibbon, new Syncfusion.SfSkinManager.Theme("Office2019White"));

        switch (editorType)
        {
            case "word":
                WordRibbonPanel.Visibility = Visibility.Visible;
                Syncfusion.SfSkinManager.SfSkinManager.SetTheme(WordRibbon, new Syncfusion.SfSkinManager.Theme("Office2019White"));
                break;
            case "excel":
                ExcelRibbonPanel.Visibility = Visibility.Visible;
                Syncfusion.SfSkinManager.SfSkinManager.SetTheme(ExcelRibbon, new Syncfusion.SfSkinManager.Theme("Office2019White"));
                Syncfusion.SfSkinManager.SfSkinManager.SetTheme(Spreadsheet, new Syncfusion.SfSkinManager.Theme("Office2019White"));
                break;
            case "pptx":
                PptxRibbonPanel.Visibility = Visibility.Visible;
                Syncfusion.SfSkinManager.SfSkinManager.SetTheme(PptxRibbon, new Syncfusion.SfSkinManager.Theme("Office2019White"));
                break;
            case "pdf":
                PdfRibbonPanel.Visibility = Visibility.Visible;
                Syncfusion.SfSkinManager.SfSkinManager.SetTheme(PdfRibbon, new Syncfusion.SfSkinManager.Theme("Office2019White"));
                break;
            default:
                DefaultRibbonPanel.Visibility = Visibility.Visible;
                break;
        }
    }
}
