using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using InsightCommon.License;

namespace InsightAiOffice.App;

public partial class MainWindow
{
    // ── Reference Materials Management (Left Panel) ─────────────

    private void ToggleReferencePane_Click(object sender, RoutedEventArgs e)
    {
        ToggleLeftPanel();
    }

    private void ToggleLeftPanel()
    {
        if (_isLeftPanelOpen)
        {
            LeftPanelCol.Width = new GridLength(0);
            LeftPanelCol.MinWidth = 0;
            _isLeftPanelOpen = false;
        }
        else
        {
            LeftPanelCol.Width = new GridLength(280);
            LeftPanelCol.MinWidth = 180;
            _isLeftPanelOpen = true;
            RefreshReferenceList();
        }
    }

    private void CloseLeftPanel_Click(object sender, RoutedEventArgs e)
    {
        LeftPanelCol.Width = new GridLength(0);
        LeftPanelCol.MinWidth = 0;
        _isLeftPanelOpen = false;
    }

    private void RefreshReferenceList()
    {
        ReferenceList.Items.Clear();

        var isEmpty = _referenceService.Count == 0;
        RefEmptyState.Visibility = isEmpty ? Visibility.Visible : Visibility.Collapsed;
        ReferenceList.Visibility = isEmpty ? Visibility.Collapsed : Visibility.Visible;

        if (isEmpty)
        {
            RefCountText.Text = "";
            return;
        }

        foreach (var doc in _referenceService.Documents)
        {
            var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(4, 3, 4, 3) };

            var iconChar = doc.FileType switch
            {
                "pdf" => "\uEA90",
                "docx" or "doc" => "\uE8A5",
                "xlsx" or "xls" or "csv" => "\uE9F9",
                _ => "\uE8A5",
            };

            panel.Children.Add(new TextBlock
            {
                Text = iconChar,
                FontFamily = new FontFamily("Segoe MDL2 Assets"),
                FontSize = 14,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = FindResource("TextSecondaryBrush") as Brush ?? Brushes.Gray,
                Margin = new Thickness(0, 0, 8, 0)
            });

            var nameBlock = new TextBlock
            {
                Text = doc.FileName,
                FontSize = 12.5,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = FindResource("TextPrimaryBrush") as Brush ?? Brushes.Black,
                TextTrimming = TextTrimming.CharacterEllipsis,
                MaxWidth = 180,
            };
            panel.Children.Add(nameBlock);

            panel.Children.Add(new TextBlock
            {
                Text = $" ({doc.ExtractedText.Length:N0} 文字)",
                FontSize = 10,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = FindResource("TextTertiaryBrush") as Brush ?? Brushes.Gray,
            });

            var delBtn = new Button
            {
                Content = new TextBlock
                {
                    Text = "\uE711",
                    FontFamily = new FontFamily("Segoe MDL2 Assets"),
                    FontSize = 11,
                    Foreground = FindResource("TextTertiaryBrush") as Brush ?? Brushes.Gray,
                },
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Padding = new Thickness(4, 2, 4, 2),
                Margin = new Thickness(8, 0, 0, 0),
                Tag = doc.Id,
                ToolTip = Helpers.LanguageManager.Get("Doc_Delete"),
            };
            delBtn.Click += RemoveReferenceInline_Click;
            panel.Children.Add(delBtn);

            var item = new ListBoxItem { Content = panel, Tag = doc.Id };
            ReferenceList.Items.Add(item);
        }

        RefCountText.Text = $"{_referenceService.Count}";
        RefStatusText.Text = Helpers.LanguageManager.Format("Doc_RefUsed", _referenceService.Count);
    }

    private async void AddReference_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "参考資料|*.pdf;*.docx;*.doc;*.xlsx;*.xls;*.csv;*.txt;*.md|All Files|*.*",
            Multiselect = true,
        };

        if (dialog.ShowDialog() != true) return;

        foreach (var filePath in dialog.FileNames)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            var fileType = ext switch
            {
                ".pdf" => "pdf",
                ".docx" or ".doc" => "docx",
                ".xlsx" or ".xls" or ".csv" => "text",
                _ => "text",
            };

            var addingMsg = Helpers.LanguageManager.Format("Doc_Adding", Path.GetFileName(filePath));
            RefStatusText.Text = addingMsg;
            StatusText.Text = addingMsg;
            var result = await _referenceService.AttachAsync(filePath, fileType);
            if (!result.Success)
            {
                var errMsg = Helpers.LanguageManager.Format("Doc_AddError", result.Error ?? "Unknown");
                RefStatusText.Text = errMsg;
                StatusText.Text = errMsg;
            }
        }

        StatusText.Text = Helpers.LanguageManager.Format("Doc_RefAdded", _referenceService.Count);
        RefreshReferenceList();
    }

    private void AddReferenceFolder_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFolderDialog
        {
            Title = Helpers.LanguageManager.Get("Ref_AddFolder"),
        };

        if (dialog.ShowDialog() != true) return;

        var files = Directory.GetFiles(dialog.FolderName, "*.*", SearchOption.AllDirectories)
            .Where(f =>
            {
                var ext = Path.GetExtension(f).ToLowerInvariant();
                return ext is ".pdf" or ".docx" or ".doc" or ".xlsx" or ".xls" or ".csv" or ".txt" or ".md";
            })
            .ToArray();

        if (files.Length == 0)
        {
            RefStatusText.Text = Helpers.LanguageManager.Get("Doc_NoValidFiles");
            return;
        }

        _ = AddReferenceFilesAsync(files);
    }

    private async Task AddReferenceFilesAsync(string[] filePaths)
    {
        foreach (var filePath in filePaths)
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            var fileType = ext switch
            {
                ".pdf" => "pdf",
                ".docx" or ".doc" => "docx",
                _ => "text",
            };

            RefStatusText.Text = Helpers.LanguageManager.Format("Doc_Adding", Path.GetFileName(filePath));
            await _referenceService.AttachAsync(filePath, fileType);
        }

        StatusText.Text = Helpers.LanguageManager.Format("Doc_RefAdded", _referenceService.Count);
        RefreshReferenceList();
    }

    private void RemoveReferenceInline_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: string refId }) return;
        _referenceService.Remove(refId);
        RefreshReferenceList();
        StatusText.Text = Helpers.LanguageManager.Get("Doc_RefRemoved");
    }

    private void OnReferenceDragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private async void OnReferenceDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is not string[] files || files.Length == 0) return;

        var validFiles = files.Where(f =>
        {
            var ext = Path.GetExtension(f).ToLowerInvariant();
            return ext is ".pdf" or ".docx" or ".doc" or ".xlsx" or ".xls" or ".csv" or ".txt" or ".md";
        }).ToArray();

        if (validFiles.Length > 0)
            await AddReferenceFilesAsync(validFiles);
    }

    // ── Window Chrome ─────────────────────────────────────────────

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2) ToggleMaximize();
        else DragMove();
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
    private void MaximizeButton_Click(object sender, RoutedEventArgs e) => ToggleMaximize();
    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

    private void ToggleMaximize()
    {
        if (WindowState == WindowState.Maximized)
        {
            WindowState = WindowState.Normal;
            MaximizeButton.Content = "\uE739";
        }
        else
        {
            WindowState = WindowState.Maximized;
            MaximizeButton.Content = "\uE923";
        }
    }

    // ── Drag & Drop ───────────────────────────────────────────────

    private void OnDragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void OnFileDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            OpenFileByPath(files[0]);
    }

    // ── Right Panel Toggle (AI Chat) ─────────────────────────────

    private void ToggleChatPane_Click(object sender, RoutedEventArgs e) =>
        ToggleRightPanel();

    private void CloseRightPanel_Click(object sender, RoutedEventArgs e)
    {
        RightPanelCol.Width = new GridLength(0);
        RightPanelCol.MinWidth = 0;
        _isRightPanelOpen = false;
    }

    private void ToggleRightPanel()
    {
        if (_isRightPanelOpen)
        {
            RightPanelCol.Width = new GridLength(0);
            RightPanelCol.MinWidth = 0;
            _isRightPanelOpen = false;
        }
        else
        {
            RightPanelCol.Width = new GridLength(380);
            RightPanelCol.MinWidth = 200;
            _isRightPanelOpen = true;
        }
    }

    private void ChatPromptEditor_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Views.PromptEditorDialog(_presetService, _aiService.Config) { Owner = this };
        dialog.ShowDialog();

        if (dialog.HasChanges)
            _chatVm.LoadPresetGroups();

        if (!string.IsNullOrEmpty(dialog.ExecutePromptText))
        {
            _chatVm.AiInput = dialog.ExecutePromptText;
            if (!_isRightPanelOpen)
                ToggleRightPanel();
        }
    }

    // ── License ──────────────────────────────────────────────────

    private void LicenseButton_Click(object sender, MouseButtonEventArgs e)
    {
        ShowLicenseDialog();
        UpdatePlanBadge();
        UpdateLicenseBackstage();
    }

    private void ShowLicenseDialog()
    {
        var contentPackManager = new InsightCommon.TemplatePack.ContentPackManager(
            "IAOF", Path.Combine(AppContext.BaseDirectory, "assets", "content-packs"), "1.0.0");
        var isJa = Helpers.LanguageManager.CurrentLanguage == "ja";
        var dialog = new InsightCommon.UI.InsightLicenseDialog(new InsightCommon.UI.LicenseDialogOptions
        {
            ProductCode = "IAOF",
            ProductName = "Insight AI Office",
            LicenseManager = _licenseManager,
            Locale = isJa ? "ja" : "en",
            ContentPackManager = contentPackManager,
            Features =
            [
                new("multi_format", isJa ? "マルチフォーマット表示・編集" : "Multi-Format Viewing & Editing"),
                new("ai_assistant", isJa ? "AIアシスタント (BYOK)" : "AI Assistant (BYOK)"),
                new("prompt_management", isJa ? "プロンプト保存・管理・配信" : "Prompt Management & Distribution"),
                new("reference_materials", isJa ? "参考資料登録" : "Reference Materials"),
                new("ai_analysis", isJa ? "AI ドキュメント分析" : "AI Document Analysis"),
                new("export", isJa ? "各形式エクスポート" : "Multi-Format Export"),
            ],
            FeatureMatrix = new()
            {
                ["export"] = [PlanCode.Trial, PlanCode.Biz, PlanCode.Ent],
            },
        });
        dialog.Owner = this;
        dialog.ShowDialog();
        contentPackManager.Dispose();
    }

    private void UpdatePlanBadge()
    {
        var license = _licenseManager.CurrentLicense;
        PlanBadgeText.Text = license.PlanDisplayName;
        var (bg, fg) = license.Plan switch
        {
            PlanCode.Biz => ("#DBEAFE", "#2563EB"),
            PlanCode.Ent => ("#EDE9FE", "#7C3AED"),
            PlanCode.Trial => ("#FEF3C7", "#D97706"),
            _ => ("#F5F5F4", "#A8A29E"),
        };
        try
        {
            PlanBadge.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(bg));
            PlanBadgeText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(fg));
        }
        catch (FormatException) { /* invalid color string — keep default */ }
    }

    private void UpdateLicenseBackstage()
    {
        var license = _licenseManager.CurrentLicense;
        LicensePlanText.Text = license.PlanDisplayName;
        LicenseExpiryText.Text = license.Plan == PlanCode.Free ? "" : Helpers.LanguageManager.Format("Doc_Expire", license.ExpiresAt ?? (object)"—");
    }
}
