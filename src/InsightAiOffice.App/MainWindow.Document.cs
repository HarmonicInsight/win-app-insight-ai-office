using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Media.Imaging;
using Syncfusion.Windows.Controls.RichTextBoxAdv;
using InsightAiOffice.App.ViewModels;
using InsightAiOffice.Data.Repositories;

namespace InsightAiOffice.App;

public partial class MainWindow
{
    private ProjectArchiveAdapter? _currentProject;

    // ── Document Loading ──────────────────────────────────────────

    public void OpenFileByPath(string filePath)
    {
        if (!File.Exists(filePath)) return;

        _currentDocPath = filePath;
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var fileName = Path.GetFileName(filePath);

        CloseAllEditors();

        switch (ext)
        {
            case ".iaof":
                OpenProjectFile(filePath);
                return;
            case ".docx" or ".doc":
                OpenWordEditor(filePath, fileName);
                break;
            case ".xlsx" or ".xls" or ".csv":
                OpenExcelEditor(filePath, fileName);
                break;
            case ".pptx" or ".ppt":
                OpenPptxViewer(filePath, fileName);
                break;
            default:
                StatusText.Text = Helpers.LanguageManager.Format("Doc_Unsupported", ext);
                return;
        }

        if (DataContext is MainViewModel vm)
        {
            vm.CurrentFilePath = filePath;
            vm.DocumentTitle = fileName;
            vm.IsFileLoaded = true;
        }
    }

    private void OpenWordEditor(string filePath, string displayName)
    {
        try
        {
            WelcomePanel.Visibility = Visibility.Collapsed;
            WordEditorPanel.Visibility = Visibility.Visible;
            WordFileName.Text = displayName;
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
            ExcelFileName.Text = displayName;
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
            var slides = InsightCommon.Services.PresentationRenderingService.RenderAllSlides(filePath, 280);
            foreach (var (full, thumb) in slides)
            {
                _pptxFullSlides.Add(full);
                _pptxThumbnails.Add(thumb);
            }

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

    // ── .iaof Project File ──────────────────────────────────────────

    private void OpenProjectFile(string iaofPath)
    {
        try
        {
            _currentProject?.Dispose();
            var project = new ProjectArchiveAdapter();
            project.Open(iaofPath);
            _currentProject = project;

            if (project.DocumentPath == null)
            {
                StatusText.Text = Helpers.LanguageManager.Get("Doc_ProjectNotFound");
                return;
            }

            // Open the inner document
            OpenFileByPath(project.DocumentPath);

            // Override the path to the .iaof file
            _currentDocPath = iaofPath;
            var displayName = Path.GetFileName(iaofPath);
            FileNameLabel.Text = displayName;

            if (DataContext is MainViewModel vm)
            {
                vm.CurrentFilePath = iaofPath;
                vm.DocumentTitle = displayName;
            }

            StatusText.Text = Helpers.LanguageManager.Format("Doc_ProjectOpened", displayName);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"{Helpers.LanguageManager.Get("Error_Title")}: {ex.Message}";
        }
    }

    private void SaveAsProject_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_currentDocPath) || _activeEditorType == "")
        {
            StatusText.Text = Helpers.LanguageManager.Get("Doc_SaveFirst");
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "IAOF Project|*.iaof",
            DefaultExt = ".iaof",
            FileName = Path.GetFileNameWithoutExtension(_currentDocPath) + ".iaof",
        };

        if (dialog.ShowDialog() != true) return;

        try
        {
            // Save current document to temp first if it's a Word doc
            var docPath = _currentDocPath;
            if (_activeEditorType == "word" && !_currentDocPath.EndsWith(".iaof", StringComparison.OrdinalIgnoreCase))
            {
                docPath = Path.Combine(Path.GetTempPath(), "iaof_save_" + Path.GetFileName(_currentDocPath));
                using var stream = File.Create(docPath);
                RichTextEditor.Save(stream, FormatType.Docx);
            }

            ProjectArchiveAdapter.CreateFromDocument(docPath, dialog.FileName);
            _currentDocPath = dialog.FileName;

            var displayName = Path.GetFileName(dialog.FileName);
            FileNameLabel.Text = displayName;
            if (DataContext is MainViewModel vm)
            {
                vm.CurrentFilePath = dialog.FileName;
                vm.DocumentTitle = displayName;
            }

            StatusText.Text = Helpers.LanguageManager.Format("Doc_ProjectSaved", displayName);
        }
        catch (Exception ex)
        {
            StatusText.Text = $"{Helpers.LanguageManager.Get("Error_Title")}: {ex.Message}";
        }
    }

    private void CloseAllEditors()
    {
        WordEditorPanel.Visibility = Visibility.Collapsed;
        ExcelEditorPanel.Visibility = Visibility.Collapsed;
        PptxInfoPanel.Visibility = Visibility.Collapsed;
        WelcomePanel.Visibility = Visibility.Visible;
        _activeEditorType = "";
        SwitchRibbon("");
    }

    private void CloseEditor_Click(object sender, RoutedEventArgs e)
    {
        CloseAllEditors();

        FileNameLabel.Text = Helpers.LanguageManager.Get("App_Tagline");
        FileTypeLabel.Text = "";
        _currentDocPath = "";
        StatusText.Text = Helpers.LanguageManager.Get("Status_Ready");

        if (DataContext is MainViewModel vm)
        {
            vm.IsFileLoaded = false;
            vm.CurrentFilePath = null;
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
}
