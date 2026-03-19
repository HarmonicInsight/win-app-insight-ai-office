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
    private string _activeEditorType = ""; // "word", "excel", "pptx", "text"
    private string _currentDocPath = "";
    private bool _textDirty; // テキストエディタの変更検知
    private string _textOriginal = ""; // テキストエディタの初期内容
    private List<Views.AttachedFileInfo> _chatAttachedFiles = new();
    private string _artifactDir = GetDefaultArtifactDir();
    private Services.IaofProjectService? _projectService;

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
        ChatPanel.PopOutRequested += (_, _) => PopOutChatPanel();
        ChatPanel.PromptEditorRequested += (_, _) => ChatPromptEditor_Click(this, new RoutedEventArgs());
        ChatPanel.InsertToDocumentRequested += InsertAiResponseText;
        ChatPanel.CopyResponseRequested += CopyAiResponseText;
        ChatPanel.FilesAttached += OnFilesAttached;
        ChatPanel.OpenArtifactInEditorRequested += path => OpenFileByPath(path);
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
        Closing += OnWindowClosing;
        Closed += OnWindowClosed;
        TextEditor.TextChanged += (_, _) =>
        {
            if (_activeEditorType == "text")
            {
                _textDirty = TextEditor.Text != _textOriginal;
                // 分割モード時はデバウンス付きリアルタイム更新
                if (_isMarkdownFile && MdSplitMode.IsChecked == true)
                    ScheduleMarkdownPreviewUpdate();
            }
        };
    }

    private System.Windows.Threading.DispatcherTimer? _mdDebounceTimer;

    private void ScheduleMarkdownPreviewUpdate()
    {
        if (_mdDebounceTimer == null)
        {
            _mdDebounceTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _mdDebounceTimer.Tick += (_, _) =>
            {
                _mdDebounceTimer.Stop();
                UpdateMarkdownPreview();
            };
        }
        _mdDebounceTimer.Stop();
        _mdDebounceTimer.Start();
    }

    // ═══════════════════════════════════════════════════════════════
    // Markdown エディタ
    // ═══════════════════════════════════════════════════════════════

    private bool _isMarkdownFile;
    private static readonly Markdig.MarkdownPipeline s_mdPipeline =
        Markdig.MarkdownExtensions.UseAdvancedExtensions(new Markdig.MarkdownPipelineBuilder()).Build();

    /// <summary>Markdown ファイルを開いた時にツールバーを表示する</summary>
    private void SetMarkdownMode(bool isMarkdown)
    {
        _isMarkdownFile = isMarkdown;
        MarkdownToolbar.Visibility = isMarkdown ? Visibility.Visible : Visibility.Collapsed;

        if (isMarkdown)
        {
            // デフォルトは編集モード
            MdEditMode.IsChecked = true;
            ApplyMarkdownLayout("edit");
        }
        else
        {
            // 通常テキスト: エディタのみ（カラム幅で制御）
            MdEditColumn.Width = new GridLength(1, GridUnitType.Star);
            MdSplitterColumn.Width = new GridLength(0);
            MdPreviewColumn.Width = new GridLength(0);
            MdSplitter.Visibility = Visibility.Collapsed;
            TextEditor.Visibility = Visibility.Visible;
        }
    }

    private void MdMode_Changed(object sender, RoutedEventArgs e)
    {
        if (!_isMarkdownFile) return;

        if (MdEditMode.IsChecked == true)
            ApplyMarkdownLayout("edit");
        else if (MdPreviewMode.IsChecked == true)
            ApplyMarkdownLayout("preview");
        else if (MdSplitMode.IsChecked == true)
            ApplyMarkdownLayout("split");
    }

    private void ApplyMarkdownLayout(string mode)
    {
        // WebBrowser は Visibility で制御するとレンダリングされない WPF の問題があるため
        // カラム幅のみで表示/非表示を切り替える
        switch (mode)
        {
            case "edit":
                TextEditor.Visibility = Visibility.Visible;
                MdSplitter.Visibility = Visibility.Collapsed;
                MdEditColumn.Width = new GridLength(1, GridUnitType.Star);
                MdSplitterColumn.Width = new GridLength(0);
                MdPreviewColumn.Width = new GridLength(0);
                break;

            case "preview":
                TextEditor.Visibility = Visibility.Collapsed;
                MdSplitter.Visibility = Visibility.Collapsed;
                MdEditColumn.Width = new GridLength(0);
                MdSplitterColumn.Width = new GridLength(0);
                MdPreviewColumn.Width = new GridLength(1, GridUnitType.Star);
                UpdateMarkdownPreview();
                break;

            case "split":
                TextEditor.Visibility = Visibility.Visible;
                MdSplitter.Visibility = Visibility.Visible;
                MdEditColumn.Width = new GridLength(1, GridUnitType.Star);
                MdSplitterColumn.Width = new GridLength(5);
                MdPreviewColumn.Width = new GridLength(1, GridUnitType.Star);
                UpdateMarkdownPreview();
                break;
        }
    }

    private async void UpdateMarkdownPreview()
    {
        try
        {
            // WebView2 の初期化を待つ
            if (MarkdownPreview.CoreWebView2 == null)
            {
                await MarkdownPreview.EnsureCoreWebView2Async();
            }

            var md = TextEditor.Text ?? "";
            var bodyHtml = Markdig.Markdown.ToHtml(md, s_mdPipeline);
            var fullHtml = string.Concat(
                "<!DOCTYPE html><html><head><meta charset=\"utf-8\"/><style>",
                "body{font-family:'Yu Gothic UI','Meiryo','Segoe UI',sans-serif;font-size:14px;line-height:1.7;color:#1C1917;padding:16px 24px;margin:0;background:#FAFAF8}",
                "h1{font-size:1.6em;border-bottom:2px solid #B8942F;padding-bottom:6px}",
                "h2{font-size:1.3em;border-bottom:1px solid #E7E2DA;padding-bottom:4px}",
                "h3{font-size:1.1em;color:#57534E}",
                "code{background:#F5F0E6;padding:2px 6px;border-radius:3px;font-family:Consolas,monospace;font-size:0.9em}",
                "pre{background:#F5F0E6;padding:12px 16px;border-radius:6px;overflow-x:auto;border:1px solid #E7E2DA}",
                "pre code{background:none;padding:0}",
                "blockquote{border-left:3px solid #B8942F;margin:12px 0;padding:8px 16px;background:#FAF8F5;color:#57534E}",
                "table{border-collapse:collapse;width:100%;margin:12px 0}",
                "th,td{border:1px solid #E7E2DA;padding:8px 12px;text-align:left}",
                "th{background:#F5F0E6;font-weight:600}",
                "tr:nth-child(even){background:#FAFAF8}",
                "a{color:#B8942F}img{max-width:100%;height:auto}",
                "ul,ol{padding-left:24px}li{margin:4px 0}",
                "hr{border:none;border-top:1px solid #E7E2DA;margin:20px 0}",
                "</style></head><body>",
                bodyHtml,
                "</body></html>");

            MarkdownPreview.NavigateToString(fullHtml);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Markdown preview: {ex.Message}";
        }
    }

    private void OnWindowClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_activeEditorType == "text" && _textDirty)
        {
            var result = System.Windows.MessageBox.Show(
                "テキストが変更されています。保存しますか？",
                "保存の確認",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel) { e.Cancel = true; return; }
            if (result == MessageBoxResult.Yes)
            {
                try { System.IO.File.WriteAllText(_currentDocPath, TextEditor.Text); }
                catch { /* best effort */ }
            }
        }
    }

    private void OnWindowClosed(object? sender, EventArgs e)
    {
        _chatVm.PropertyChanged -= OnChatVmPropertyChanged;
        // チャット履歴を保存
        Services.ChatHistoryService.Save(_chatVm.ChatMessages);
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

    // ── ライセンスチェック ───────────────────────────────────

    private void CheckLicenseAndRestrict()
    {
        var plan = _licenseManager.CurrentLicense.Plan;
        var isActivated = _licenseManager.IsActivated;

        // プランバッジ更新
        UpdatePlanBadge();

        if (!isActivated)
        {
            // FREE: AI 送信無効（GridSplitter等のレイアウト操作は維持）
            ChatPanel.SetAiEnabled(false);

            // 初回起動時にライセンス案内
            var result = System.Windows.MessageBox.Show(
                "Insight AI Office をご利用いただきありがとうございます。\n\n" +
                "AI コンシェルジュ機能をご利用いただくには、\nトライアルライセンスの登録が必要です。\n\n" +
                "今すぐライセンスを登録しますか？",
                "ライセンス登録のご案内",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Information);

            if (result == System.Windows.MessageBoxResult.Yes)
            {
                ShowLicenseDialog();
                // 再チェック
                if (_licenseManager.IsActivated)
                    ChatPanel.SetAiEnabled(true);
            }
        }

        // 期限切れ警告
        if (_licenseManager.ShouldShowExpiryWarning)
        {
            StatusText.Text = $"ライセンス残り {_licenseManager.DaysRemaining} 日";
        }
    }

    // ── ツール実行ハンドラ ───────────────────────────────────

    private InsightCommon.AI.IToolExecutor CreateToolExecutor()
    {
        // 検索・置換（Word / テキストエディタ 両対応）
        bool DoReplace(string find, string replacement)
        {
            return Dispatcher.Invoke(() =>
            {
                try
                {
                    // テキストエディタの場合
                    if (_activeEditorType == "text" && TextEditor != null)
                    {
                        var idx = TextEditor.Text.IndexOf(find, StringComparison.Ordinal);
                        if (idx < 0) return false;
                        TextEditor.Text = TextEditor.Text.Remove(idx, find.Length).Insert(idx, replacement);
                        return true;
                    }

                    // PPTX の場合（OpenXML 経由）
                    if (_activeEditorType == "pptx" && !string.IsNullOrEmpty(_currentDocPath))
                    {
                        var count = Services.PptxService.FindAndReplace(_currentDocPath, find, replacement);
                        if (count > 0)
                        {
                            var sel = PptxSlideList.SelectedIndex;
                            OpenPptxViewer(_currentDocPath, System.IO.Path.GetFileName(_currentDocPath));
                            if (sel >= 0 && sel < PptxSlideList.Items.Count) PptxSlideList.SelectedIndex = sel;
                        }
                        return count > 0;
                    }

                    // Word エディタの場合
                    if (RichTextEditor == null) return false;
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
            InsertDocumentText = (text, after) =>
            {
                return Dispatcher.Invoke(() =>
                {
                    try
                    {
                        // テキストエディタの場合
                        if (_activeEditorType == "text" && TextEditor != null)
                        {
                            if (after != null)
                            {
                                var idx = TextEditor.Text.IndexOf(after, StringComparison.Ordinal);
                                if (idx < 0) return false;
                                TextEditor.Text = TextEditor.Text.Insert(idx + after.Length, text);
                            }
                            else
                            {
                                TextEditor.Text += "\n" + text;
                            }
                            return true;
                        }

                        // Word エディタの場合
                        if (_activeEditorType != "word" || RichTextEditor == null) return false;
                        if (after != null)
                        {
                            var found = RichTextEditor.Find(after, Syncfusion.Windows.Controls.RichTextBoxAdv.FindOptions.None);
                            if (found == null) return false;
                            RichTextEditor.Selection.InsertText(after + text);
                        }
                        else
                        {
                            System.Windows.Input.ApplicationCommands.SelectAll.Execute(null, RichTextEditor);
                            var endPos = RichTextEditor.Selection.End;
                            RichTextEditor.Selection.Select(endPos, endPos);
                            RichTextEditor.Selection.InsertText("\n" + text);
                        }
                        return true;
                    }
                    catch { return false; }
                });
            },
            EditSpreadsheetCells = (sheetName, cells) =>
            {
                return Dispatcher.Invoke(() =>
                {
                    try
                    {
                        if (_activeEditorType != "excel" || Spreadsheet?.ActiveSheet == null) return 0;

                        var ws = Spreadsheet.ActiveSheet;

                        // シート切替
                        if (!string.IsNullOrEmpty(sheetName))
                        {
                            var targetSheet = Spreadsheet.Workbook?.Worksheets
                                .Cast<dynamic>()
                                .FirstOrDefault(s => (string)s.Name == sheetName);
                            if (targetSheet != null) ws = targetSheet;
                        }

                        int count = 0;
                        foreach (var (cellRef, value) in cells)
                        {
                            try
                            {
                                ws.Range[cellRef].Text = value;
                                count++;
                            }
                            catch { /* skip invalid cell ref */ }
                        }

                        Spreadsheet?.ActiveGrid?.InvalidateCells();
                        return count;
                    }
                    catch { return 0; }
                });
            },
            CreateTextFile = (title, content) =>
            {
                return Dispatcher.Invoke(() =>
                {
                    try
                    {
                        var tempDir = System.IO.Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                            "HarmonicInsight", "InsightAiOffice", "Temp");
                        System.IO.Directory.CreateDirectory(tempDir);
                        var safeName = string.Join("_", title.Split(System.IO.Path.GetInvalidFileNameChars()));
                        var filePath = System.IO.Path.Combine(tempDir, $"{safeName}.txt");
                        System.IO.File.WriteAllText(filePath, content);
                        OpenFileByPath(filePath);
                        return filePath;
                    }
                    catch { return null; }
                });
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
              - insert_document_text: 開いているWordにテキスト挿入（afterで挿入位置指定、省略で末尾追記）
              - edit_spreadsheet_cells: 開いているExcelのセルに値を直接書き込み（cellでA1形式指定）
              - create_text_file: テキストファイルを新規作成して新しいタブで開く

              【重要ルール】
              - ユーザーが開いているExcel/Wordの内容を修正するよう依頼 → edit_spreadsheet_cells / find_and_replace で直接編集
              - PPTXのテキストを「ですます調に変更」「英語に翻訳」等 → find_and_replace を複数回呼び出して各テキストを置換
              - テキストの添削・翻訳・要約など「元を残して新しい版を作る」依頼 → create_text_file で新しいタブに出力
              - 新しいレポート・表を作る依頼 → generate_report / generate_spreadsheet でファイル生成

              【ツール呼び出しルール】
              ユーザーが校正・赤入れ・修正・ドキュメント作成を依頼した場合、
              必ず適切なツールを実際に呼び出してください。説明だけで終わらないこと。
              """
            : "You are the Insight AI Office assistant. Answer questions about the user's open document. Provide analysis, proofreading, and summaries. Be concise.";

        // デフォルトプリセットがあればそちらを優先
        var defaultPreset = _presetService.GetDefault();
        if (defaultPreset != null && !string.IsNullOrEmpty(defaultPreset.SystemPrompt))
            prompt = defaultPreset.SystemPrompt;

        // ドキュメントコンテキスト（エディタで開いているファイル — 構造化圧縮済み）
        var docContent = ExtractDocumentContent();
        if (!string.IsNullOrEmpty(docContent))
        {
            // 圧縮データの場合、AI に「追加データを要求できる」ことを伝える
            var compressionNotice = InsightCommon.AI.DocumentCompressor.ShouldCompress(docContent)
                ? InsightCommon.AI.DocumentCompressor.GetCompressionNotice(isJa ? "ja" : "en")
                : "";
            prompt += $"\n\n--- {(isJa ? "現在のドキュメント内容" : "Current Document")} ---\n{compressionNotice}{docContent}\n--- ---";
        }

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

                        if (InsightCommon.AI.DocumentCompressor.ShouldCompress(text))
                        {
                            if (ext == ".csv")
                            {
                                // CSV: 行パース → CompressSpreadsheet で構造化圧縮
                                var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
                                var rows = lines.Select(line =>
                                    line.TrimEnd('\r').Split(',').Select(c => c.Trim('"', ' ')).ToArray()
                                ).ToList();
                                var headerRow = rows.Count > 0 ? rows[0] : null;
                                var colCount = rows.Count > 0 ? rows.Max(r => r.Length) : 0;
                                text = InsightCommon.AI.DocumentCompressor.CompressSpreadsheet(
                                    f.FileName, rows, rows.Count, colCount, headerRow);
                            }
                            else
                            {
                                // .txt/.md: 段落分割 → CompressDocument で構造化圧縮
                                var paragraphs = text.Split(new[] { "\n\n", "\r\n\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                                var sections = new System.Collections.Generic.List<InsightCommon.AI.DocumentCompressor.DocumentSection>
                                {
                                    new()
                                    {
                                        Heading = f.FileName,
                                        TextContent = text,
                                        Level = 1
                                    }
                                };
                                var wordCount = text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
                                text = InsightCommon.AI.DocumentCompressor.CompressDocument(
                                    f.FileName, sections, paragraphs.Length, wordCount, totalPages: 1);
                            }
                            prompt += $"\n{text}\n";
                        }
                        else
                        {
                            prompt += $"\n```\n{text}\n```\n";
                        }
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
            MainRibbon, WordRibbon, ExcelRibbon, PptxRibbon, PdfRibbon, TextRibbon, Spreadsheet, RichTextEditor);

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

        // チャット履歴を復元
        var savedMessages = Services.ChatHistoryService.Load();
        foreach (var msg in savedMessages)
            _chatVm.ChatMessages.Add(msg);

        // ライセンスチェック: FREE版はAI機能を制限
        CheckLicenseAndRestrict();

        // デフォルトで AI コンシェルジュ（右パネル）を開く
        if (!_isRightPanelOpen && _licenseManager.IsActivated)
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
        LocalizeRecentTab(PdfRecentTab, PdfRecentTitle, PdfRecentEmptyHint);

        // Backstage (main)
        MainRibbon.BackStageHeader = L("Menu_File");
        BS_LanguageTitle.Text = L("BS_Language");

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
        WelcomeOpenLabel.Text = L("Welcome_OpenLabel");
        WelcomeOpenDesc.Text = L("Welcome_OpenDesc");
        WelcomeViewArtifactsLabel.Text = L("Welcome_ViewArtifactsLabel");
        WelcomeViewArtifactsDesc.Text = L("Welcome_ViewArtifactsDesc");
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
            (L("Menu_Project"), new[] { L("Project_New"), L("Project_Open"), L("Project_Save"), L("Project_SaveAs") }),
            (L("Menu_File"), new[] { L("File_Open"), L("File_NewText") }),
            (L("Menu_Help"), new[] { L("Menu_Tutorial"), L("Menu_Help") }),
        });
        // Backstage: 新規プロジェクト / プロジェクトを開く / 上書き保存 / 別名保存 / [最近] / [sep] / 言語 / ライセンス / [sep] / 閉じる
        ApplyBackstage(MainRibbon,
            L("Project_New"), L("Project_Open"), L("Project_Save"), L("Project_SaveAs"),
            L("BS_Recent"), null, L("BS_Language"), L("License_Title"), null, L("File_CloseDoc"));

        // ── Word Ribbon ──
        WordRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(WordRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_Project"), new[] { L("Project_New"), L("Project_Open"), L("Project_Save"), L("Project_SaveAs") }),
            (L("Menu_File"), new[] { L("File_Open"), L("File_Save"), L("File_SaveAs"), L("File_NewText") }),
            (L("Format_UndoRedo"), new[] { L("Format_Undo"), L("Format_Redo") }),
            (L("Format_Font"), new[] { L("Format_Bold"), L("Format_Italic"), L("Format_Underline"), L("Format_Strikethrough") }),
            (L("Format_Paragraph"), new[] { L("Format_AlignLeft"), L("Format_AlignCenter"), L("Format_AlignRight"), L("Format_BulletList") }),
            (L("Format_Edit"), new[] { L("Format_FindReplace") }),
            (L("File_Print"), new[] { L("File_Print") }),
        });
        ApplyBackstage(WordRibbon, L("File_Open"), L("File_SaveOverwrite"), L("File_Print"), null, L("File_CloseDoc"));

        // ── Excel Ribbon ──
        ExcelRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(ExcelRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_Project"), new[] { L("Project_New"), L("Project_Open"), L("Project_Save"), L("Project_SaveAs") }),
            (L("Menu_File"), new[] { L("File_Open"), L("File_SaveAs"), L("File_NewText") }),
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
        // ── PPTX Ribbon ──
        PptxRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(PptxRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_Project"), new[] { L("Project_New"), L("Project_Open"), L("Project_Save"), L("Project_SaveAs") }),
            (L("Menu_File"), new[] { L("File_Open"), L("File_NewText") }),
            (L("Pptx_Slide"), new[] { L("Pptx_Add"), L("Pptx_Duplicate"), L("Pptx_Delete"), L("Pptx_MoveUp"), L("Pptx_MoveDown") }),
            (L("Pdf_Text"), new[] { L("Pdf_ExtractText"), L("Format_FindReplace"), L("Pptx_EditNotes") }),
            (L("Menu_Help"), new[] { L("Menu_Tutorial"), L("Menu_Help") }),
        });
        ApplyBackstage(PptxRibbon, L("File_Open"), null, L("File_CloseDoc"));

        // ── PDF Ribbon ──
        PdfRibbon.BackStageHeader = L("Menu_File");
        ApplyTab(PdfRibbon, 0, L("Menu_Home"), new[]
        {
            (L("Menu_Project"), new[] { L("Project_New"), L("Project_Open"), L("Project_Save"), L("Project_SaveAs") }),
            (L("Menu_File"), new[] { L("File_Open"), L("File_Save"), L("File_SaveAs"), L("Pdf_Print"), L("File_NewText") }),
            (L("Pdf_Annotation"), new[] { L("Pdf_Highlight"), L("Pdf_Underline"), L("Pdf_Strikethrough"), L("Pdf_Ink") }),
            (L("Pdf_Navigation"), new[] { L("Pdf_PrevPage"), L("Pdf_NextPage"), L("Pdf_ZoomIn"), L("Pdf_ZoomOut") }),
            (L("Menu_Help"), new[] { L("Menu_Tutorial"), L("Menu_Help") }),
        });
        ApplyTab(PdfRibbon, 1, L("Pdf_Edit"), new[]
        {
            (L("Pdf_PageOps"), new[] { L("Pdf_RotateRight"), L("Pdf_RotateLeft"), L("Pdf_DeletePage") }),
            (L("Pdf_MergeSplit"), new[] { L("Pdf_Merge"), L("Pdf_Split") }),
            (L("Pdf_Text"), new[] { L("Pdf_ExtractText"), L("Pdf_SearchText") }),
        });
        ApplyBackstage(PdfRibbon, L("File_Open"), L("BS_Recent"), L("BS_Settings"), L("File_CloseDoc"));
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
                switch (bar.Items[j])
                {
                    case Syncfusion.Windows.Tools.Controls.RibbonButton btn:
                        btn.Label = bars[i].labels[j];
                        break;
                    case Syncfusion.Windows.Tools.Controls.SplitButton split:
                        split.Label = bars[i].labels[j];
                        break;
                }
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

    private void ScaleUp_Click(object sender, RoutedEventArgs e) =>
        InsightCommon.UI.InsightScaleManager.Instance.ZoomIn();

    private void ScaleDown_Click(object sender, RoutedEventArgs e) =>
        InsightCommon.UI.InsightScaleManager.Instance.ZoomOut();

    private void SwitchRibbon(string editorType)
    {
        DefaultRibbonPanel.Visibility = Visibility.Collapsed;
        WordRibbonPanel.Visibility = Visibility.Collapsed;
        ExcelRibbonPanel.Visibility = Visibility.Collapsed;
        PptxRibbonPanel.Visibility = Visibility.Collapsed;
        PdfRibbonPanel.Visibility = Visibility.Collapsed;
        TextRibbonPanel.Visibility = Visibility.Collapsed;

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
            case "text":
                // Text はデフォルトリボンを共用
                DefaultRibbonPanel.Visibility = Visibility.Visible;
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
