using System.Collections.Specialized;
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
    public event EventHandler? PromptEditorRequested;
    public event Action<string>? InsertToDocumentRequested;
    public event Action<string>? CopyResponseRequested;

    private Storyboard? _loadingStoryboard;

    public ChatPanelView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        ChatInput.TextChanged += ChatInput_TextChanged;

        Loaded += (_, _) => UpdatePlaceholder();
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

    private void RetryMessage_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not AiChatViewModel vm) return;

        // Find the last user message and re-send it
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
