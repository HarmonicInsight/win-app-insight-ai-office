using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;
using InsightCommon.AI;

namespace InsightAiOffice.App.Views;

public partial class ChatPanelView : UserControl
{
    public event EventHandler? HelpRequested;
    public event EventHandler? CloseRequested;
    public event EventHandler? PopOutRequested;
    public event EventHandler? PromptEditorRequested;
    public event Action<string>? InsertToDocumentRequested;
    public event Action<string>? CopyResponseRequested;
    public event Action<string>? ThemeColorChanged;
    public event Action<List<AttachedFileInfo>>? FilesAttached;
    public event Action<string>? OpenArtifactFolderRequested;
    public event Action<string>? OpenArtifactInEditorRequested;

    /// <summary>現在選択中のテーマカラー名</summary>
    public string SelectedTheme { get; private set; } = "gold";

    /// <summary>
    /// AI 送信機能の有効/無効を切り替える。
    /// GridSplitter やスクロールなどのレイアウト操作は常に有効のまま、
    /// 入力・送信系コントロールのみを無効化する。
    /// </summary>
    public void SetAiEnabled(bool enabled)
    {
        ChatInput.IsEnabled = enabled;
        InputAreaBorder.Opacity = enabled ? 1.0 : 0.5;
    }

    /// <summary>添付ファイル一覧</summary>
    public ObservableCollection<AttachedFileInfo> AttachedFiles { get; } = new();

    private Storyboard? _loadingStoryboard;

    /// <summary>選択中のモデルID</summary>
    public string SelectedModelId { get; private set; } = InsightCommon.AI.ClaudeModels.DefaultStandardModel;

    public ChatPanelView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        ChatInput.TextChanged += ChatInput_TextChanged;
        AttachedFilesList.ItemsSource = AttachedFiles;

        Loaded += (_, _) =>
        {
            UpdatePlaceholder();
            InitModelCombo();
        };
    }

    private void InitModelCombo()
    {
        // モデル特性（ToolTip で表示）
        var tips = new Dictionary<string, string>
        {
            [InsightCommon.AI.ClaudeModels.HaikuId] = "高速・低コスト — 簡単な質問・要約向け",
            [InsightCommon.AI.ClaudeModels.SonnetId] = "バランス型（推奨）— 文書生成・Excel編集・分析",
            [InsightCommon.AI.ClaudeModels.OpusId] = "最高品質 — 複雑な分析・大量データ処理・高精度生成",
        };

        ModelCombo.Items.Clear();
        foreach (var m in InsightCommon.AI.ClaudeModels.Registry)
        {
            var item = new ComboBoxItem
            {
                Content = $"{m.CostIndicator} {m.Label}",
                Tag = m.Id,
                ToolTip = tips.TryGetValue(m.Id, out var tip) ? tip : m.Label,
            };
            ModelCombo.Items.Add(item);
            if (m.Id == InsightCommon.AI.ClaudeModels.DefaultStandardModel)
                ModelCombo.SelectedItem = item;
        }
    }

    private void ModelCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ModelCombo?.SelectedItem is ComboBoxItem item && item.Tag is string modelId)
        {
            SelectedModelId = modelId;
            // ViewModelにも反映
            if (DataContext is AiChatViewModel vm)
                vm.SelectedModelId = modelId;
        }
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (e.OldValue is AiChatViewModel oldVm)
        {
            oldVm.ChatMessages.CollectionChanged -= OnChatMessagesChanged;
            oldVm.PropertyChanged -= OnViewModelPropertyChanged;
        }
        if (e.NewValue is AiChatViewModel newVm)
        {
            newVm.ChatMessages.CollectionChanged += OnChatMessagesChanged;
            newVm.PropertyChanged += OnViewModelPropertyChanged;
        }
    }

    private void OnViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(AiChatViewModel.IsSending))
        {
            Dispatcher.InvokeAsync(() =>
            {
                if (sender is AiChatViewModel vm)
                {
                    if (vm.IsSending)
                        StartLoadingAnimation();
                    else
                        StopLoadingAnimation();
                }
            });
        }
    }

    private void StartLoadingAnimation()
    {
        _loadingStoryboard ??= (Storyboard)FindResource("LoadingDotsStoryboard");
        try { _loadingStoryboard.Begin(this, true); } catch { /* storyboard target may not be visible */ }
    }

    private void StopLoadingAnimation()
    {
        if (_loadingStoryboard != null)
        {
            try { _loadingStoryboard.Stop(this); } catch { }
        }
    }

    private void OnChatMessagesChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.Action == NotifyCollectionChangedAction.Add)
        {
            Dispatcher.InvokeAsync(() => ChatScrollViewer.ScrollToEnd(),
                System.Windows.Threading.DispatcherPriority.Background);
        }
    }

    // ── ドキュメントバッジ更新（MainWindow から呼び出し） ──

    public void UpdateDocumentBadge(string? docName, bool hasDocument)
    {
        DocBadgeText.Text = docName ?? "";
        DocBadge.Visibility = hasDocument ? Visibility.Visible : Visibility.Collapsed;
    }

    // ── ローカライゼーション ──

    public void RefreshLocalization()
    {
        var L = Helpers.LanguageManager.Get;
        CancelLabel.Text = L("Chat_Cancel");
        PlaceholderText.Text = L("Chat_Placeholder");
        ChatWelcomeText.Text = L("Chat_WelcomeMessage");
        PromptEditorLabel.Text = L("Pane_Prompt");
        UpdatePlaceholder();
    }

    private void UpdatePlaceholder()
    {
        PlaceholderText.Visibility =
            string.IsNullOrEmpty(ChatInput.Text) ? Visibility.Visible : Visibility.Collapsed;
    }

    private void ChatInput_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdatePlaceholder();
    }

    // ── Event handlers ──

    private void Help_Click(object sender, RoutedEventArgs e) =>
        HelpRequested?.Invoke(this, EventArgs.Empty);

    private void Close_Click(object sender, RoutedEventArgs e) =>
        CloseRequested?.Invoke(this, EventArgs.Empty);

    private void PopOut_Click(object sender, RoutedEventArgs e) =>
        PopOutRequested?.Invoke(this, EventArgs.Empty);

    private void NewThread_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not AiChatViewModel vm) return;
        vm.NewThread();
    }

    private void RenameThread_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not AiChatViewModel vm || vm.CurrentThread == null) return;
        var thread = vm.CurrentThread;

        var inputBox = new System.Windows.Controls.TextBox
        {
            Text = thread.Title, FontSize = 12, Margin = new Thickness(0, 8, 0, 0),
            SelectionStart = 0, SelectionLength = thread.Title.Length,
        };
        var okBtn = new System.Windows.Controls.Button
        {
            Content = "OK", Width = 70, IsDefault = true,
            Margin = new Thickness(0, 8, 0, 0),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
        };
        var dlg = new Window
        {
            Title = "スレッド名の変更",
            Width = 320, Height = 150,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = Window.GetWindow(this),
            ResizeMode = ResizeMode.NoResize,
        };
        okBtn.Click += (_, _) => dlg.DialogResult = true;
        dlg.Content = new System.Windows.Controls.StackPanel
        {
            Margin = new Thickness(12),
            Children =
            {
                new System.Windows.Controls.TextBlock { Text = "スレッド名:", FontSize = 11 },
                inputBox,
                okBtn,
            }
        };
        inputBox.KeyDown += (_, ke) =>
        {
            if (ke.Key == System.Windows.Input.Key.Enter) dlg.DialogResult = true;
        };
        inputBox.Focus();
        if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(inputBox.Text))
            thread.Title = inputBox.Text.Trim();
    }

    private void DeleteThread_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not AiChatViewModel vm || vm.CurrentThread == null) return;
        if (vm.Threads.Count <= 1) return;

        var result = MessageBox.Show(
            $"「{vm.CurrentThread.Title}」を削除しますか？",
            "スレッドの削除",
            MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result != MessageBoxResult.Yes) return;

        vm.DeleteCurrentThread();
    }

    private void PromptEditor_Click(object sender, RoutedEventArgs e) =>
        PromptEditorRequested?.Invoke(this, EventArgs.Empty);

    private void ChatInput_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && Keyboard.Modifiers == ModifierKeys.None)
        {
            e.Handled = true;
            if (DataContext is AiChatViewModel vm && vm.ExecuteFromInputCommand.CanExecute(null))
                vm.ExecuteFromInputCommand.Execute(null);
        }
    }

    private void InsertAiResponse_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string text })
            InsertToDocumentRequested?.Invoke(text);
    }

    private void CopyMessage_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { Tag: string content })
        {
            try { Clipboard.SetText(content); }
            catch { /* clipboard may be locked */ }
            CopyResponseRequested?.Invoke(content);
        }
    }

    // ── ファイル添付 ──

    private void AttachFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "All Supported|*.xlsx;*.docx;*.pptx;*.csv;*.pdf;*.txt;*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp|All Files|*.*",
            Multiselect = true,
        };
        if (dialog.ShowDialog() == true)
        {
            foreach (var file in dialog.FileNames)
                AddAttachment(file);
        }
    }

    private void InputArea_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void InputArea_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
        {
            foreach (var file in files)
                AddAttachment(file);
        }
    }

    private void AddAttachment(string filePath)
    {
        if (AttachedFiles.Any(f => f.FullPath == filePath)) return;
        AttachedFiles.Add(new AttachedFileInfo(filePath));
        FilesAttached?.Invoke(AttachedFiles.ToList());
    }

    private void OpenAttachment_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2 && sender is FrameworkElement el && el.Tag is string path && File.Exists(path))
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
        }
    }

    private void RemoveAttachment_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string path)
        {
            var item = AttachedFiles.FirstOrDefault(f => f.FullPath == path);
            if (item != null)
            {
                AttachedFiles.Remove(item);
                FilesAttached?.Invoke(AttachedFiles.ToList());
            }
        }
    }

    /// <summary>送信後に添付ファイルをクリア</summary>
    public void ClearAttachments() => AttachedFiles.Clear();

    private void ThemeColor_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ThemeColorCombo?.SelectedItem is ComboBoxItem item && item.Tag is string theme)
        {
            SelectedTheme = theme;
            ThemeColorChanged?.Invoke(theme);
        }
    }

    // ── 成果物パネル ──

    private readonly ObservableCollection<ArtifactFileInfo> _artifacts = new();

    public void AddArtifact(string filePath)
    {
        _artifacts.Insert(0, new ArtifactFileInfo(filePath));
        ArtifactList.ItemsSource = _artifacts;
        ArtifactSection.Visibility = Visibility.Visible;
    }

    private void OpenArtifactFolder_Click(object sender, RoutedEventArgs e)
        => OpenArtifactFolderRequested?.Invoke("");

    private void ArtifactItem_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement el && el.Tag is string path && File.Exists(path))
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext is ".docx" or ".xlsx" or ".pptx" or ".pdf")
                OpenArtifactInEditorRequested?.Invoke(path);
            else
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
        }
    }

    private void RetryMessage_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not AiChatViewModel vm) return;

        for (int i = vm.ChatMessages.Count - 1; i >= 0; i--)
        {
            if (vm.ChatMessages[i].Role == ChatRole.User)
            {
                vm.AiInput = vm.ChatMessages[i].Content;
                vm.ExecuteFromInputCommand.Execute(null);
                break;
            }
        }
    }
}

/// <summary>成果物ファイル情報</summary>
public class ArtifactFileInfo
{
    public string FileName { get; set; }
    public string FilePath { get; set; }
    public string Icon { get; set; }

    public ArtifactFileInfo(string filePath)
    {
        FilePath = filePath;
        FileName = Path.GetFileName(filePath);
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        Icon = ext switch
        {
            ".xlsx" or ".csv" => "📊",
            ".docx" => "📝",
            ".pptx" => "📽️",
            ".pdf" => "📕",
            ".html" => "📄",
            _ => "📁",
        };
    }
}

/// <summary>添付ファイル情報</summary>
public class AttachedFileInfo
{
    public string FileName { get; set; }
    public string FullPath { get; set; }
    public string Icon { get; set; }

    public AttachedFileInfo(string filePath)
    {
        FullPath = filePath;
        FileName = Path.GetFileName(filePath);
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        Icon = ext switch
        {
            ".xlsx" or ".xlsm" or ".csv" => "📊",
            ".docx" => "📝",
            ".pptx" => "📽️",
            ".pdf" => "📕",
            ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".webp" => "🖼️",
            ".txt" or ".md" => "📄",
            _ => "📁",
        };
    }
}
