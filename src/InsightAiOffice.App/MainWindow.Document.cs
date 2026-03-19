using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Media.Imaging;
using Syncfusion.Windows.Controls.RichTextBoxAdv;
using InsightAiOffice.App.ViewModels;

namespace InsightAiOffice.App;

public partial class MainWindow
{
    // ── Document Loading ──────────────────────────────────────────

    public void OpenFileByPath(string filePath)
    {
        if (!File.Exists(filePath)) return;

        _recentFiles.Add(filePath);
        RefreshRecentFilesList();

        // Already open — just switch to it
        if (_openTabs.ContainsKey(filePath))
        {
            SwitchToTab(filePath);
            return;
        }

        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var fileName = Path.GetFileName(filePath);

        // .iaof プロジェクトファイルの場合は専用処理
        if (ext == ".iaof")
        {
            _ = OpenProjectAsync(filePath);
            return;
        }

        var editorType = ext switch
        {
            ".docx" or ".doc" => "word",
            ".xlsx" or ".xls" or ".csv" => "excel",
            ".pptx" or ".ppt" => "pptx",
            ".pdf" => "pdf",
            ".txt" or ".md" or ".log" or ".json" or ".xml" or ".html" or ".css" or ".js" => "text",
            _ => ""
        };
        if (string.IsNullOrEmpty(editorType))
        {
            StatusText.Text = Helpers.LanguageManager.Format("Doc_Unsupported", ext);
            return;
        }

        // Save current tab before opening new
        SaveCurrentTabState();

        // Register new tab
        var tab = new Models.DocumentTab
        {
            FilePath = filePath,
            FileName = fileName,
            EditorType = editorType,
        };
        _openTabs[filePath] = tab;
        _tabOrder.Add(filePath);

        // Open in editor
        _currentDocPath = filePath;
        HideEditorPanels();

        switch (editorType)
        {
            case "word":
                OpenWordEditor(filePath, fileName);
                break;
            case "excel":
                OpenExcelEditor(filePath, fileName);
                break;
            case "pptx":
                OpenPptxViewer(filePath, fileName);
                break;
            case "pdf":
                OpenPdfViewer(filePath, fileName);
                break;
            case "text":
                OpenTextEditor(filePath, fileName);
                break;
        }

        if (DataContext is MainViewModel vm)
        {
            vm.CurrentFilePath = filePath;
            vm.DocumentTitle = fileName;
            vm.IsFileLoaded = true;
        }

        RefreshTabBar();
    }

    private void OpenWordEditor(string filePath, string displayName)
    {
        try
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            WordEditorPanel.Visibility = Visibility.Visible;
            FileNameLabel.Text = displayName;
            FileTypeLabel.Text = "DOCX";
            _activeEditorType = "word";
            SwitchRibbon("word");

            using var stream = File.OpenRead(filePath);
            RichTextEditor.Load(stream, FormatType.Docx);
            StatusText.Text = Helpers.LanguageManager.Format("Doc_Loaded", displayName);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Word: {ex.Message}";
        }
    }

    private void OpenExcelEditor(string filePath, string displayName)
    {
        try
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            ExcelEditorPanel.Visibility = Visibility.Visible;
            FileNameLabel.Text = displayName;
            FileTypeLabel.Text = Path.GetExtension(filePath).TrimStart('.').ToUpperInvariant();
            _activeEditorType = "excel";
            SwitchRibbon("excel");

            Spreadsheet.Open(filePath);
            StatusText.Text = Helpers.LanguageManager.Format("Doc_Loaded", displayName);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Excel: {ex.Message}";
        }
    }

    private void OpenTextEditor(string filePath, string displayName)
    {
        try
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            TextEditorPanel.Visibility = Visibility.Visible;
            FileNameLabel.Text = displayName;
            FileTypeLabel.Text = Path.GetExtension(filePath).TrimStart('.').ToUpperInvariant();
            _activeEditorType = "text";
            SwitchRibbon("text");

            var content = File.ReadAllText(filePath);
            TextEditor.Text = content;
            _textOriginal = content;
            _textDirty = false;

            // Markdown モード判定
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            SetMarkdownMode(ext == ".md");

            StatusText.Text = Helpers.LanguageManager.Format("Doc_Loaded", displayName);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"Text: {ex.Message}";
        }
    }

    private List<BitmapSource> _pptxThumbnails = new();
    private List<BitmapSource> _pptxFullSlides = new();

    private void OpenPptxViewer(string filePath, string displayName)
    {
        try
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            PptxInfoPanel.Visibility = Visibility.Visible;
            PptxFileName.Text = displayName;
            FileNameLabel.Text = displayName;
            FileTypeLabel.Text = "PPTX";
            _activeEditorType = "pptx";
            SwitchRibbon("pptx");

            _pptxThumbnails.Clear();
            _pptxFullSlides.Clear();
            var slides = Services.PptxService.RenderAllSlides(filePath, 280);
            foreach (var (full, thumb) in slides)
            {
                _pptxFullSlides.Add(full);
                _pptxThumbnails.Add(thumb);
            }

            System.Diagnostics.Debug.WriteLine(
                $"[PPTX] Rendered {_pptxFullSlides.Count} slides, " +
                $"first full: {(_pptxFullSlides.Count > 0 ? $"{_pptxFullSlides[0].PixelWidth}x{_pptxFullSlides[0].PixelHeight}" : "N/A")}, " +
                $"first thumb: {(_pptxThumbnails.Count > 0 ? $"{_pptxThumbnails[0].PixelWidth}x{_pptxThumbnails[0].PixelHeight}" : "N/A")}");

            PptxSlideCount.Text = $"{_pptxFullSlides.Count} スライド";

            var items = new List<PptxSlideItem>();
            for (int i = 0; i < _pptxThumbnails.Count; i++)
            {
                items.Add(new PptxSlideItem { Thumbnail = _pptxThumbnails[i], Label = $"スライド {i + 1}", Index = i });
            }
            PptxSlideList.ItemsSource = items;
            if (items.Count > 0)
                PptxSlideList.SelectedIndex = 0;

            StatusText.Text = Helpers.LanguageManager.Format("Doc_Loaded", $"{displayName} ({_pptxFullSlides.Count} slides)");
        }
        catch (Exception ex)
        {
            StatusText.Text = $"PowerPoint: {ex.Message}";
        }
    }

    private void PptxSlideList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (PptxSlideList.SelectedItem is PptxSlideItem item && item.Index < _pptxFullSlides.Count)
        {
            PptxMainSlideImage.Source = _pptxFullSlides[item.Index];
        }
    }

    public class PptxSlideItem
    {
        public BitmapSource? Thumbnail { get; set; }
        public string Label { get; set; } = "";
        public int Index { get; set; }
    }

    private void OpenPdfViewer(string filePath, string displayName)
    {
        try
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            PdfViewerPanel.Visibility = Visibility.Visible;
            PdfFileName.Text = displayName;
            FileNameLabel.Text = displayName;
            FileTypeLabel.Text = "PDF";
            _activeEditorType = "pdf";
            SwitchRibbon("pdf");

            PdfViewer.Load(filePath);
            StatusText.Text = Helpers.LanguageManager.Format("Doc_Loaded", displayName);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"PDF: {ex.Message}";
        }
    }

    private void CloseAllEditors()
    {
        // Close all tabs
        _openTabs.Clear();
        _tabOrder.Clear();

        HideEditorPanels();
        WelcomePanel.Visibility = Visibility.Visible;
        _activeEditorType = "";
        _currentDocPath = "";
        SwitchRibbon("");
        RefreshTabBar();
    }

    private async Task OpenProjectAsync(string projectPath)
    {
        try
        {
            _projectService?.Dispose();
            _projectService = new Services.IaofProjectService();

            var (docPath, editorType, chatHistory) = await _projectService.OpenAsync(projectPath);

            // チャット履歴を復元
            if (chatHistory.Sessions.Count > 0)
            {
                _chatVm.ChatMessages.Clear();
                foreach (var msg in chatHistory.Sessions.SelectMany(s => s.Messages))
                {
                    var role = msg.Role switch
                    {
                        "user" => InsightCommon.AI.ChatRole.User,
                        "assistant" => InsightCommon.AI.ChatRole.Assistant,
                        _ => InsightCommon.AI.ChatRole.User,
                    };
                    _chatVm.ChatMessages.Add(new InsightCommon.AI.ChatMessageVm { Role = role, Content = msg.Content });
                }

                if (!_isRightPanelOpen) ToggleRightPanel();
            }

            // ドキュメントを開く
            if (docPath != null)
                OpenFileByPath(docPath);

            StatusText.Text = $"プロジェクトを開きました: {Path.GetFileName(projectPath)}";
            Title = $"{Path.GetFileNameWithoutExtension(projectPath)} — Insight AI Office";
        }
        catch (Exception ex) { StatusText.Text = $"プロジェクト: {ex.Message}"; }
    }

    private void CloseEditor_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_currentDocPath))
        {
            CloseTab(_currentDocPath);
        }
        else
        {
            CloseAllEditors();
            FileNameLabel.Text = Helpers.LanguageManager.Get("App_Tagline");
            FileTypeLabel.Text = "";
            StatusText.Text = Helpers.LanguageManager.Get("Status_Ready");

            if (DataContext is MainViewModel vm)
            {
                vm.IsFileLoaded = false;
                vm.CurrentFilePath = null;
            }
        }
    }

    // ── DrillDown: AI が詳細範囲を要求した場合のデータ抽出 ─────────

    /// <summary>
    /// 圧縮されたドキュメントデータの特定範囲を非圧縮で返す。
    /// get_document_detail ツールのバックエンド実装。
    /// </summary>
    internal string ExtractDocumentDetailRange(string rangeType, int start, int end)
    {
        try
        {
            return rangeType switch
            {
                "rows" => ExtractExcelDetailRows(start, end),
                "section" => ExtractWordDetailSections(start, end),
                "slides" => ExtractPptxDetailSlides(start, end),
                "pages" => ExtractPdfDetailPages(start, end),
                _ => $"[エラー] 不明な range_type: {rangeType}。rows / section / slides / pages のいずれかを指定してください。"
            };
        }
        catch (Exception ex)
        {
            return $"[詳細データ取得エラー: {ex.Message}]";
        }
    }

    private string ExtractExcelDetailRows(int startRow, int endRow)
    {
        var ws = Spreadsheet?.ActiveSheet;
        if (ws == null) return "(Excel シートが開かれていません)";

        var usedRange = ws.UsedRange;
        if (usedRange == null) return "(空のシート)";

        var maxCols = Math.Min(usedRange.LastColumn, usedRange.Column + 25);

        // ヘッダー行を取得
        string[]? headerRow = null;
        {
            var hCells = new List<string>();
            for (int c = usedRange.Column; c <= maxCols; c++)
                hCells.Add(ws.Range[usedRange.Row, c].DisplayText ?? "");
            headerRow = hCells.ToArray();
        }

        // 全行を読み取り
        var allRows = new List<string[]>();
        for (int r = usedRange.Row; r <= usedRange.LastRow; r++)
        {
            var cells = new List<string>();
            for (int c = usedRange.Column; c <= maxCols; c++)
                cells.Add(ws.Range[r, c].DisplayText ?? "");
            allRows.Add(cells.ToArray());
        }

        return InsightCommon.AI.DocumentCompressor.GetDetailRows(allRows, headerRow, startRow, endRow);
    }

    private string ExtractWordDetailSections(int startSection, int endSection)
    {
        var doc = RichTextEditor?.Document;
        if (doc == null) return "(Word ドキュメントが開かれていません)";

        using var ms = new MemoryStream();
        RichTextEditor.Save(ms, FormatType.Txt);
        ms.Position = 0;
        using var reader = new StreamReader(ms);
        var fullText = reader.ReadToEnd();

        var paragraphs = fullText.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries);
        var sections = new List<InsightCommon.AI.DocumentCompressor.DocumentSection>();
        var currentSection = new InsightCommon.AI.DocumentCompressor.DocumentSection { Heading = "本文", Level = 1 };
        var sectionText = new StringBuilder();

        foreach (var para in paragraphs)
        {
            var trimmed = para.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;

            if (trimmed.Length < 60 && (
                System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^(第[0-9０-９一-九]+[章節条項]|[0-9]+[\.\s]|[IVX]+[\.\s])") ||
                trimmed.StartsWith("■") || trimmed.StartsWith("●") || trimmed.StartsWith("◆")))
            {
                if (sectionText.Length > 0)
                {
                    currentSection.TextContent = sectionText.ToString();
                    sections.Add(currentSection);
                }
                currentSection = new InsightCommon.AI.DocumentCompressor.DocumentSection { Heading = trimmed, Level = 1 };
                sectionText.Clear();
            }
            else
            {
                sectionText.AppendLine(trimmed);
            }
        }
        if (sectionText.Length > 0)
        {
            currentSection.TextContent = sectionText.ToString();
            sections.Add(currentSection);
        }

        return InsightCommon.AI.DocumentCompressor.GetDetailSections(sections, startSection, endSection);
    }

    private string ExtractPptxDetailSlides(int startSlide, int endSlide)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || !File.Exists(_currentDocPath))
            return "(PPTX ファイルが開かれていません)";

        try
        {
            using var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(_currentDocPath, false);
            var presPart = doc.PresentationPart;
            if (presPart == null) return "(プレゼンテーションが空です)";

            var slides = new List<InsightCommon.AI.DocumentCompressor.SlideInfo>();
            var slideIndex = 0;

            foreach (var slideId in presPart.Presentation.SlideIdList?.ChildElements
                         .OfType<DocumentFormat.OpenXml.Presentation.SlideId>() ?? [])
            {
                slideIndex++;
                var slidePart = (DocumentFormat.OpenXml.Packaging.SlidePart)presPart.GetPartById(slideId.RelationshipId!);

                var textParts = new List<string>();
                string title = "";

                foreach (var textBody in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>())
                {
                    foreach (var para in textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
                    {
                        var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            textParts.Add(text.Trim());
                            if (string.IsNullOrEmpty(title)) title = text.Trim();
                        }
                    }
                }

                var notes = "";
                if (slidePart.NotesSlidePart != null)
                {
                    var noteTexts = new List<string>();
                    foreach (var textBody in slidePart.NotesSlidePart.NotesSlide
                                 .Descendants<DocumentFormat.OpenXml.Drawing.TextBody>())
                    {
                        foreach (var para in textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
                        {
                            var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
                            if (!string.IsNullOrWhiteSpace(text))
                                noteTexts.Add(text.Trim());
                        }
                    }
                    notes = string.Join(" ", noteTexts);
                }

                slides.Add(new InsightCommon.AI.DocumentCompressor.SlideInfo
                {
                    Number = slideIndex,
                    Title = title,
                    FullText = string.Join("\n", textParts),
                    Notes = notes
                });
            }

            return InsightCommon.AI.DocumentCompressor.GetDetailSlides(slides, startSlide, endSlide);
        }
        catch (Exception ex)
        {
            return $"[PPTX 詳細取得エラー: {ex.Message}]";
        }
    }

    private string ExtractPdfDetailPages(int startPage, int endPage)
    {
        try
        {
            var doc = PdfViewer?.LoadedDocument;
            if (doc == null) return "(PDF が開かれていません)";

            var pages = new List<InsightCommon.AI.DocumentCompressor.PageInfo>();
            for (int i = 0; i < doc.Pages.Count; i++)
            {
                pages.Add(new InsightCommon.AI.DocumentCompressor.PageInfo
                {
                    Number = i + 1,
                    Text = doc.Pages[i].ExtractText() ?? ""
                });
            }

            return InsightCommon.AI.DocumentCompressor.GetDetailPages(pages, startPage, endPage);
        }
        catch (Exception ex)
        {
            return $"[PDF 詳細取得エラー: {ex.Message}]";
        }
    }

    // ── Document Content Extraction (for AI) ──────────────────────

    private string ExtractDocumentContent()
    {
        try
        {
            return _activeEditorType switch
            {
                "word" => ExtractWordContent(),
                "excel" => ExtractExcelContent(),
                "pptx" => ExtractPptxContent(_currentDocPath),
                "pdf" => ExtractPdfContent(),
                "text" => TextEditor?.Text ?? "",
                _ => ""
            };
        }
        catch (Exception ex)
        {
            return $"[コンテンツ抽出エラー: {ex.Message}]";
        }
    }

    private string ExtractWordContent()
    {
        var doc = RichTextEditor.Document;
        if (doc == null) return "";

        // テキスト全文を取得
        using var ms = new MemoryStream();
        RichTextEditor.Save(ms, FormatType.Txt);
        ms.Position = 0;
        using var reader = new StreamReader(ms);
        var fullText = reader.ReadToEnd();

        // 小さい文書（推定2000トークン以下）は全文を返す
        if (!InsightCommon.AI.DocumentCompressor.ShouldCompress(fullText))
            return fullText;

        // 構造化圧縮: 段落を分割してセクション化
        var paragraphs = fullText.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries);
        var sections = new List<InsightCommon.AI.DocumentCompressor.DocumentSection>();
        var currentSection = new InsightCommon.AI.DocumentCompressor.DocumentSection
        {
            Heading = "本文",
            Level = 1
        };
        var sectionText = new StringBuilder();

        foreach (var para in paragraphs)
        {
            var trimmed = para.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;

            // 見出しらしい行を検出（複数パターン対応）
            // - 日本語: 第N章/節/条/項、N. 見出し
            // - 記号系: ■●◆▼▶◇★☆►、全角数字、丸数字
            // - 英語: I. II. III. 1. 2. A) B) Chapter N Section N
            // - Markdown: # ## ###
            // - 短い行で次の段落と明確に分離されている
            var isHeading = trimmed.Length < 80 && (
                System.Text.RegularExpressions.Regex.IsMatch(trimmed,
                    @"^(第[0-9０-９一二三四五六七八九十百]+[章節条項款号]|[0-9]+[\.\s\)）]|[０-９]+[\.\s]|[IVXivx]+[\.\s\)]|[A-Z][\.\)）]\s)") ||
                System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^#{1,3}\s") || // Markdown 見出し
                System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^[①-⑳⑴-⒇]\s*") || // 丸数字
                System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^(Chapter|Section|Part|CHAPTER|SECTION)\s+\d", System.Text.RegularExpressions.RegexOptions.IgnoreCase) ||
                "■●◆▼▶◇★☆►【〔".Contains(trimmed[0]));
            if (isHeading)
            {
                // 前のセクションを保存
                if (sectionText.Length > 0)
                {
                    currentSection.TextContent = sectionText.ToString();
                    sections.Add(currentSection);
                }
                currentSection = new InsightCommon.AI.DocumentCompressor.DocumentSection
                {
                    Heading = trimmed,
                    Level = 1
                };
                sectionText.Clear();
            }
            else
            {
                sectionText.AppendLine(trimmed);
            }
        }
        // 最後のセクション
        if (sectionText.Length > 0)
        {
            currentSection.TextContent = sectionText.ToString();
            sections.Add(currentSection);
        }

        // セクションが1つだけなら見出し検出できなかった → 段落分割で圧縮
        if (sections.Count <= 1)
        {
            sections.Clear();
            var chunkSize = paragraphs.Length / Math.Max(1, paragraphs.Length / 10);
            for (int i = 0; i < paragraphs.Length; i += Math.Max(1, chunkSize))
            {
                var chunk = string.Join("\n", paragraphs.Skip(i).Take(chunkSize));
                sections.Add(new InsightCommon.AI.DocumentCompressor.DocumentSection
                {
                    Heading = $"パート {sections.Count + 1}",
                    TextContent = chunk,
                    Level = 1
                });
            }
        }

        var fileName = System.IO.Path.GetFileName(_currentDocPath ?? "document.docx");
        var wordCount = fullText.Split(new[] { ' ', '\n', '\r', '　' }, StringSplitOptions.RemoveEmptyEntries).Length;
        var pageEstimate = Math.Max(1, wordCount / 400); // 概算

        return InsightCommon.AI.DocumentCompressor.CompressDocument(
            fileName, sections, paragraphs.Length, wordCount, pageEstimate);
    }

    private string ExtractExcelContent()
    {
        if (Spreadsheet?.Workbook == null) return "";

        var sb = new StringBuilder();
        var worksheets = Spreadsheet.Workbook.Worksheets;
        var activeSheetName = Spreadsheet.ActiveSheet?.Name;

        // 全シートの概要を先に表示
        sb.AppendLine($"【ブック構成】シート数: {worksheets.Count}, アクティブ: {activeSheetName ?? "N/A"}");
        for (int s = 0; s < worksheets.Count; s++)
        {
            var sheet = worksheets[s];
            var usedRange = sheet.UsedRange;
            var rows = usedRange != null ? usedRange.LastRow - usedRange.Row + 1 : 0;
            var cols = usedRange != null ? usedRange.LastColumn - usedRange.Column + 1 : 0;
            var marker = sheet.Name == activeSheetName ? " ★" : "";
            sb.AppendLine($"  {s + 1}. {sheet.Name} ({rows}行×{cols}列){marker}");
        }
        sb.AppendLine();

        // 各シートを構造化圧縮
        for (int s = 0; s < worksheets.Count; s++)
        {
            var ws = worksheets[s];
            var usedRange = ws.UsedRange;
            if (usedRange == null) continue;

            var totalRows = usedRange.LastRow - usedRange.Row + 1;
            var totalCols = usedRange.LastColumn - usedRange.Column + 1;
            if (totalRows <= 0 || totalCols <= 0) continue;

            var maxRows = Math.Min(usedRange.LastRow, usedRange.Row + 500);
            var maxCols = Math.Min(usedRange.LastColumn, usedRange.Column + 25);

            var rows = new List<string[]>();
            string[]? headerRow = null;
            for (int r = usedRange.Row; r <= maxRows; r++)
            {
                var cells = new List<string>();
                for (int c = usedRange.Column; c <= maxCols; c++)
                {
                    var val = ws.Range[r, c].DisplayText ?? "";
                    cells.Add(val);
                }
                var row = cells.ToArray();
                if (r == usedRange.Row) headerRow = row;
                rows.Add(row);
            }

            sb.AppendLine(InsightCommon.AI.DocumentCompressor.CompressSpreadsheet(
                ws.Name ?? $"Sheet{s + 1}", rows, totalRows, totalCols, headerRow));
            sb.AppendLine();
        }

        return sb.ToString();
    }

    private static string ExtractPptxContent(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return "";

        try
        {
            using var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(filePath, false);
            var presPart = doc.PresentationPart;
            if (presPart == null) return "";

            var slides = new List<InsightCommon.AI.DocumentCompressor.SlideInfo>();
            var slideIndex = 0;

            foreach (var slideId in presPart.Presentation.SlideIdList?.ChildElements
                         .OfType<DocumentFormat.OpenXml.Presentation.SlideId>() ?? [])
            {
                slideIndex++;
                var slidePart = (DocumentFormat.OpenXml.Packaging.SlidePart)presPart.GetPartById(slideId.RelationshipId!);

                var textParts = new List<string>();
                string title = "";

                foreach (var textBody in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>())
                {
                    foreach (var para in textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
                    {
                        var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            textParts.Add(text.Trim());
                            // 最初のテキストをタイトルとみなす
                            if (string.IsNullOrEmpty(title)) title = text.Trim();
                        }
                    }
                }

                var notes = "";
                if (slidePart.NotesSlidePart != null)
                {
                    var noteTexts = new List<string>();
                    foreach (var textBody in slidePart.NotesSlidePart.NotesSlide
                                 .Descendants<DocumentFormat.OpenXml.Drawing.TextBody>())
                    {
                        foreach (var para in textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
                        {
                            var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
                            if (!string.IsNullOrWhiteSpace(text))
                                noteTexts.Add(text.Trim());
                        }
                    }
                    notes = string.Join(" ", noteTexts);
                }

                slides.Add(new InsightCommon.AI.DocumentCompressor.SlideInfo
                {
                    Number = slideIndex,
                    Title = title,
                    FullText = string.Join("\n", textParts),
                    Notes = notes
                });
            }

            // 構造化圧縮
            var fileName = Path.GetFileName(filePath);
            return InsightCommon.AI.DocumentCompressor.CompressPresentation(fileName, slides);
        }
        catch (Exception ex)
        {
            return $"[PPTX テキスト抽出エラー: {ex.Message}]";
        }
    }

    private string ExtractPdfContent()
    {
        try
        {
            var doc = PdfViewer.LoadedDocument;
            if (doc == null) return "";

            var pages = new List<InsightCommon.AI.DocumentCompressor.PageInfo>();
            for (int i = 0; i < doc.Pages.Count; i++)
            {
                var page = doc.Pages[i];
                pages.Add(new InsightCommon.AI.DocumentCompressor.PageInfo
                {
                    Number = i + 1,
                    Text = page.ExtractText() ?? ""
                });
            }

            var fileName = Path.GetFileName(_currentDocPath ?? "document.pdf");
            return InsightCommon.AI.DocumentCompressor.CompressPdf(fileName, pages);
        }
        catch (Exception ex)
        {
            return $"[PDF テキスト抽出エラー: {ex.Message}]";
        }
    }
}
