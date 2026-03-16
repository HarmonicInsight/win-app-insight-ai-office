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

        using var ms = new MemoryStream();
        RichTextEditor.Save(ms, FormatType.Txt);
        ms.Position = 0;
        using var reader = new StreamReader(ms);
        var text = reader.ReadToEnd();

        if (text.Length > 8000)
            text = text[..8000] + "\n\n[...以降省略...]";

        return text;
    }

    private string ExtractExcelContent()
    {
        var ws = Spreadsheet?.ActiveSheet;
        if (ws == null) return "";

        var lines = new StringBuilder();
        var usedRange = ws.UsedRange;
        if (usedRange == null) return "";

        var maxRows = Math.Min(usedRange.LastRow, 100);
        var maxCols = Math.Min(usedRange.LastColumn, 20);

        for (int r = usedRange.Row; r <= maxRows; r++)
        {
            var cells = new List<string>();
            for (int c = usedRange.Column; c <= maxCols; c++)
            {
                var val = ws.Range[r, c].DisplayText ?? "";
                cells.Add(val);
            }
            lines.AppendLine(string.Join("\t", cells));
        }

        var text = lines.ToString();
        if (text.Length > 8000)
            text = text[..8000] + "\n\n[...以降省略...]";

        return text;
    }

    private static string ExtractPptxContent(string filePath)
    {
        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return "";

        try
        {
            using var doc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(filePath, false);
            var presPart = doc.PresentationPart;
            if (presPart == null) return "";

            var sb = new StringBuilder();
            var slideIndex = 0;

            foreach (var slideId in presPart.Presentation.SlideIdList?.ChildElements
                         .OfType<DocumentFormat.OpenXml.Presentation.SlideId>() ?? [])
            {
                slideIndex++;
                var slidePart = (DocumentFormat.OpenXml.Packaging.SlidePart)presPart.GetPartById(slideId.RelationshipId!);
                sb.AppendLine($"--- スライド {slideIndex} ---");

                foreach (var textBody in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>())
                {
                    foreach (var para in textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
                    {
                        var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
                        if (!string.IsNullOrWhiteSpace(text))
                            sb.AppendLine(text.Trim());
                    }
                }

                if (slidePart.NotesSlidePart != null)
                {
                    foreach (var textBody in slidePart.NotesSlidePart.NotesSlide
                                 .Descendants<DocumentFormat.OpenXml.Drawing.TextBody>())
                    {
                        foreach (var para in textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
                        {
                            var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
                            if (!string.IsNullOrWhiteSpace(text))
                                sb.AppendLine($"[ノート] {text.Trim()}");
                        }
                    }
                }

                sb.AppendLine();
            }

            var result = sb.ToString();
            if (result.Length > 8000)
                result = result[..8000] + "\n\n[...以降省略...]";

            return result;
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

            var sb = new StringBuilder();
            var pageCount = doc.Pages.Count;
            for (int i = 0; i < pageCount; i++)
            {
                sb.AppendLine($"--- ページ {i + 1} ---");
                var page = doc.Pages[i];
                sb.AppendLine(page.ExtractText());
                sb.AppendLine();
            }

            var text = sb.ToString();
            if (text.Length > 8000)
                text = text[..8000] + "\n\n[...以降省略...]";

            return text;
        }
        catch (Exception ex)
        {
            return $"[PDF テキスト抽出エラー: {ex.Message}]";
        }
    }
}
