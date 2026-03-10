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
        var dialog = new Views.PromptEditorDialog(_presetService, _chatVm.AiService.Config) { Owner = this };
        dialog.ShowDialog();

        if (dialog.HasChanges)
            _chatVm.RefreshUserPromptGroups();

        if (!string.IsNullOrEmpty(dialog.ExecutePromptText))
        {
            _chatVm.AiInput = dialog.ExecutePromptText;
            if (!_isRightPanelOpen)
                ToggleRightPanel();
        }
    }

    // ── Recent Files ───────────────────────────────────────────────

    private void RefreshRecentFilesList()
    {
        var entries = _recentFiles.Entries;
        RecentFilesList.ItemsSource = entries;
        RecentFilesEmptyHint.Visibility = entries.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void RecentFile_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string path } && File.Exists(path))
        {
            HideAllBackstages();
            OpenFileByPath(path);
        }
    }

    private void RecentFile_DoubleClick(object sender, MouseButtonEventArgs e)
    {
        // Handled by button click inside the ListBox
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
                new("formats", isJa
                    ? "Word / Excel / PowerPoint / PDF 対応"
                    : "Word / Excel / PowerPoint / PDF Support"),
                new("editing", isJa
                    ? "Word・Excel の編集・書式設定・保存"
                    : "Word & Excel Editing, Formatting & Save"),
                new("ai_concierge", isJa
                    ? "AI コンシェルジュ — 要約・校正・分析 (BYOK)"
                    : "AI Concierge — Summarize, Proofread & Analyze (BYOK)"),
                new("prompt_management", isJa
                    ? "プロンプト保存・管理"
                    : "Prompt Management"),
                new("export", isJa
                    ? "Word / Excel エクスポート"
                    : "Word / Excel Export"),
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
