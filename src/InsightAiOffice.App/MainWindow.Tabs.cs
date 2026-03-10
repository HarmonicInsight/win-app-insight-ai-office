using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Syncfusion.Windows.Controls.RichTextBoxAdv;
using WpfBorder = System.Windows.Controls.Border;
using InsightAiOffice.App.Helpers;

namespace InsightAiOffice.App;

public partial class MainWindow
{
    // ── Tab State Save / Restore ─────────────────────────────────

    private void SaveCurrentTabState()
    {
        if (string.IsNullOrEmpty(_currentDocPath) || !_openTabs.TryGetValue(_currentDocPath, out var tab))
            return;

        switch (tab.EditorType)
        {
            case "word":
                try
                {
                    using var ms = new MemoryStream();
                    RichTextEditor.Save(ms, FormatType.Docx);
                    tab.WordContent = ms.ToArray();
                }
                catch { /* save failed — keep previous state */ }
                break;

            case "pptx":
                tab.PptxSelectedSlide = PptxSlideList.SelectedIndex;
                break;

            // Excel: SfSpreadsheet auto-saves changes to the open file internally,
            // no explicit state capture needed.
        }
    }

    private void RestoreTabState(Models.DocumentTab tab)
    {
        HideEditorPanels();

        switch (tab.EditorType)
        {
            case "word":
                WelcomePanel.Visibility = Visibility.Collapsed;
                WordEditorPanel.Visibility = Visibility.Visible;
                WordFileName.Text = tab.FileName;
                _activeEditorType = "word";
                SwitchRibbon("word");

                if (tab.WordContent != null)
                {
                    using var ms = new MemoryStream(tab.WordContent);
                    RichTextEditor.Load(ms, FormatType.Docx);
                }
                else
                {
                    using var fs = File.OpenRead(tab.FilePath);
                    RichTextEditor.Load(fs, FormatType.Docx);
                }
                break;

            case "excel":
                WelcomePanel.Visibility = Visibility.Collapsed;
                ExcelEditorPanel.Visibility = Visibility.Visible;
                ExcelFileName.Text = tab.FileName;
                _activeEditorType = "excel";
                SwitchRibbon("excel");
                Spreadsheet.Open(tab.FilePath);
                break;

            case "pptx":
                OpenPptxViewer(tab.FilePath, tab.FileName);
                if (tab.PptxSelectedSlide >= 0)
                    PptxSlideList.SelectedIndex = tab.PptxSelectedSlide;
                break;

            case "pdf":
                OpenPdfViewer(tab.FilePath, tab.FileName);
                break;
        }

        _currentDocPath = tab.FilePath;
        FileNameLabel.Text = tab.FileName;
        FileTypeLabel.Text = tab.EditorType switch
        {
            "word" => "DOCX",
            "excel" => Path.GetExtension(tab.FilePath).TrimStart('.').ToUpperInvariant(),
            "pptx" => "PPTX",
            "pdf" => "PDF",
            _ => ""
        };

        StatusText.Text = LanguageManager.Format("Doc_Loaded", tab.FileName);
    }

    private void HideEditorPanels()
    {
        WordEditorPanel.Visibility = Visibility.Collapsed;
        ExcelEditorPanel.Visibility = Visibility.Collapsed;
        PptxInfoPanel.Visibility = Visibility.Collapsed;
        PdfViewerPanel.Visibility = Visibility.Collapsed;
        WelcomePanel.Visibility = Visibility.Collapsed;
        _pptxThumbnails.Clear();
        _pptxFullSlides.Clear();
    }

    // ── Tab Switch ───────────────────────────────────────────────

    private void SwitchToTab(string filePath)
    {
        if (filePath == _currentDocPath) return;
        if (!_openTabs.ContainsKey(filePath)) return;

        SaveCurrentTabState();
        RestoreTabState(_openTabs[filePath]);
        RefreshTabBar();

        if (DataContext is ViewModels.MainViewModel vm)
        {
            vm.CurrentFilePath = filePath;
            vm.DocumentTitle = Path.GetFileName(filePath);
            vm.IsFileLoaded = true;
        }
    }

    // ── Tab Close ────────────────────────────────────────────────

    private void CloseTab(string filePath)
    {
        if (!_openTabs.Remove(filePath)) return;
        _tabOrder.Remove(filePath);

        if (filePath == _currentDocPath)
        {
            // Closing active tab — switch to another or show welcome
            if (_tabOrder.Count > 0)
            {
                var nextTab = _tabOrder[^1];
                _currentDocPath = ""; // clear so SwitchToTab doesn't try to save closed tab
                SwitchToTab(nextTab);
            }
            else
            {
                _currentDocPath = "";
                _activeEditorType = "";
                HideEditorPanels();
                WelcomePanel.Visibility = Visibility.Visible;
                SwitchRibbon("");
                FileNameLabel.Text = LanguageManager.Get("App_Tagline");
                FileTypeLabel.Text = "";
                StatusText.Text = LanguageManager.Get("Status_Ready");

                if (DataContext is ViewModels.MainViewModel vm)
                {
                    vm.IsFileLoaded = false;
                    vm.CurrentFilePath = null;
                }
            }
        }

        RefreshTabBar();
    }

    // ── Tab Bar UI ───────────────────────────────────────────────

    private void RefreshTabBar()
    {
        TabBarItems.Children.Clear();

        if (_tabOrder.Count == 0)
        {
            TabBarPanel.Visibility = Visibility.Collapsed;
            return;
        }

        TabBarPanel.Visibility = Visibility.Visible;

        foreach (var path in _tabOrder)
        {
            if (!_openTabs.TryGetValue(path, out var tab)) continue;
            var isActive = path == _currentDocPath;

            var tabPanel = new StackPanel { Orientation = Orientation.Horizontal };

            // File type icon
            var icon = tab.EditorType switch
            {
                "word" => "\uE8A5",
                "excel" => "\uE80A",
                "pptx" => "\uE7B5",
                "pdf" => "\uEA90",
                _ => "\uE7C3",
            };
            tabPanel.Children.Add(new TextBlock
            {
                Text = icon,
                FontFamily = new FontFamily("Segoe MDL2 Assets"),
                FontSize = 11,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = isActive
                    ? (Brush)FindResource("PrimaryBrush")
                    : (Brush)FindResource("TextTertiaryBrush"),
                Margin = new Thickness(0, 0, 5, 0),
            });

            // File name
            tabPanel.Children.Add(new TextBlock
            {
                Text = tab.FileName,
                FontSize = 11,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = isActive
                    ? (Brush)FindResource("TextPrimaryBrush")
                    : (Brush)FindResource("TextSecondaryBrush"),
                FontWeight = isActive ? FontWeights.SemiBold : FontWeights.Normal,
                MaxWidth = 160,
                TextTrimming = TextTrimming.CharacterEllipsis,
            });

            // Close button
            var closeBtn = new Button
            {
                Content = "\uE711",
                FontFamily = new FontFamily("Segoe MDL2 Assets"),
                FontSize = 8,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(4, 2, 4, 2),
                Margin = new Thickness(4, 0, 0, 0),
                Cursor = Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = (Brush)FindResource("TextTertiaryBrush"),
                Tag = path,
            };
            closeBtn.Click += TabCloseButton_Click;
            tabPanel.Children.Add(closeBtn);

            // Tab container
            var tabBorder = new WpfBorder
            {
                Child = tabPanel,
                Padding = new Thickness(10, 6, 6, 6),
                Cursor = Cursors.Hand,
                Background = isActive
                    ? (Brush)FindResource("BgCardBrush")
                    : Brushes.Transparent,
                BorderBrush = isActive
                    ? (Brush)FindResource("PrimaryBrush")
                    : Brushes.Transparent,
                BorderThickness = new Thickness(0, 0, 0, isActive ? 2 : 0),
                Tag = path,
            };
            tabBorder.MouseLeftButtonDown += TabItem_Click;
            TabBarItems.Children.Add(tabBorder);
        }
    }

    private void TabItem_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is WpfBorder { Tag: string path })
            SwitchToTab(path);
    }

    private void TabCloseButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string path })
            CloseTab(path);
    }
}
