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
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
        {
            foreach (var file in files)
                OpenFileByPath(file);
        }
    }

    // ── Right Panel Toggle (AI Chat) ─────────────────────────────

    private void ToggleChatPane_Click(object sender, RoutedEventArgs e) =>
        ToggleRightPanel();

    private void CloseRightPanel_Click(object sender, RoutedEventArgs e)
    {
        RightPanelCol.Width = new GridLength(0);
        RightPanelCol.MinWidth = 0;
        _isRightPanelOpen = false;
        UpdateChatToggleColor();
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
        UpdateChatToggleColor();
    }

    private void UpdateChatToggleColor()
    {
        var brush = _isRightPanelOpen
            ? (System.Windows.Media.Brush)FindResource("PrimaryBrush")
            : (System.Windows.Media.Brush)FindResource("TextSecondaryBrush");
        ChatToggleIcon.Foreground = brush;
        ChatToggleText.Foreground = brush;
    }

    private void ChatPromptEditor_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Views.PromptEditorDialog(_presetService, _chatVm.AiService.Config) { Owner = this };
        dialog.ShowDialog();

        if (dialog.HasChanges)
            _chatVm.RefreshUserPromptGroups();

        if (!string.IsNullOrEmpty(dialog.ExecutePromptText))
        {
            // チャットパネルを開く
            if (!_isRightPanelOpen)
                ToggleRightPanel();

            // プロンプトをセットして即AI実行
            _chatVm.AiInput = dialog.ExecutePromptText;
            if (_chatVm.ExecuteFromInputCommand.CanExecute(null))
                _chatVm.ExecuteFromInputCommand.Execute(null);
        }
    }

    // ── Tutorial ──────────────────────────────────────────────────

    private void WelcomeOpenProject_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "IAOF Project|*.iaof",
            Title = "プロジェクトを開く",
        };
        if (dialog.ShowDialog() == true)
            OpenFileByPath(dialog.FileName);
    }

    private void WelcomeOpen_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (DataContext is ViewModels.MainViewModel vm && vm.OpenDocumentCommand.CanExecute(null))
            vm.OpenDocumentCommand.Execute(null);
    }

    private void WelcomeChat_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (!_isRightPanelOpen) ToggleRightPanel();
    }

    private void WelcomeRecentFile_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement el && el.Tag is string path && System.IO.File.Exists(path))
            OpenFileByPath(path);
    }

    private void WelcomeTutorial_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        => Tutorial_Click(sender, new RoutedEventArgs());

    private void Tutorial_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Views.TutorialDialog { Owner = this };
        dialog.TutorialExecuteRequested += (files, prompt, outputFormat) =>
        {
            // チャットパネルを開く
            if (!_isRightPanelOpen)
                ToggleRightPanel();

            // 添付ファイルをセット
            ChatPanel.ClearAttachments();
            foreach (var file in files)
                ChatPanel.AttachedFiles.Add(new Views.AttachedFileInfo(file));
            _chatAttachedFiles = files.Select(f => new Views.AttachedFileInfo(f)).ToList();

            // プロンプトを入力欄にセット（実行はユーザーに委ねる）
            if (!string.IsNullOrEmpty(prompt))
                _chatVm.AiInput = prompt;

            StatusText.Text = "チュートリアルをセットしました — 送信ボタンで実行してください";
        };
        dialog.ShowDialog();
    }

    // ── Recent Files ───────────────────────────────────────────────

    private void RefreshRecentFilesList()
    {
        var entries = _recentFiles.Entries;
        var vis = entries.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

        RecentFilesList.ItemsSource = entries;
        RecentFilesEmptyHint.Visibility = vis;

        WordRecentFilesList.ItemsSource = entries;
        WordRecentEmptyHint.Visibility = vis;

        ExcelRecentFilesList.ItemsSource = entries;
        ExcelRecentEmptyHint.Visibility = vis;

        PptxRecentFilesList.ItemsSource = entries;
        PptxRecentEmptyHint.Visibility = vis;

        PdfRecentFilesList.ItemsSource = entries;
        PdfRecentEmptyHint.Visibility = vis;

        // ウェルカム画面
        WelcomeRecentList.ItemsSource = entries;
        WelcomeRecentEmpty.Visibility = vis;
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
        using var contentPackManager = new InsightCommon.TemplatePack.ContentPackManager(
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
                ["ai_concierge"] = [PlanCode.Trial, PlanCode.Biz, PlanCode.Ent],
                ["prompt_management"] = [PlanCode.Trial, PlanCode.Biz, PlanCode.Ent],
                ["export"] = [PlanCode.Trial, PlanCode.Biz, PlanCode.Ent],
            },
        });
        dialog.Owner = this;
        dialog.ShowDialog();

        // ライセンス変更後にAI機能の有効/無効を更新
        UpdatePlanBadge();
        ChatPanel.SetAiEnabled(_licenseManager.IsActivated);
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
        var planName = license.PlanDisplayName;
        var expiry = license.Plan == PlanCode.Free ? "" : Helpers.LanguageManager.Format("Doc_Expire", license.ExpiresAt ?? (object)"—");

        SetLicenseDisplay(LicensePlanText, LicenseExpiryText, planName, expiry);
        SetLicenseDisplay(WordLicensePlanText, WordLicenseExpiryText, planName, expiry);
        SetLicenseDisplay(ExcelLicensePlanText, ExcelLicenseExpiryText, planName, expiry);
        SetLicenseDisplay(PptxLicensePlanText, PptxLicenseExpiryText, planName, expiry);
    }

    private static void SetLicenseDisplay(
        System.Windows.Controls.TextBlock planText,
        System.Windows.Controls.TextBlock expiryText,
        string planName, string expiry)
    {
        planText.Text = planName;
        expiryText.Text = expiry;
    }
}
