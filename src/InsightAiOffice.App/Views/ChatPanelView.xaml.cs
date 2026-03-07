using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using InsightAiOffice.App.ViewModels;

namespace InsightAiOffice.App.Views;

public partial class ChatPanelView : UserControl
{
    public event EventHandler? HelpRequested;
    public event EventHandler? CloseRequested;
    public event EventHandler? PromptEditorRequested;
    public event Action<string>? InsertToDocumentRequested;
    public event Action<string>? CopyResponseRequested;

    public ChatPanelView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (e.OldValue is ChatPanelViewModel oldVm)
            oldVm.ChatMessages.CollectionChanged -= OnChatMessagesChanged;
        if (e.NewValue is ChatPanelViewModel newVm)
            newVm.ChatMessages.CollectionChanged += OnChatMessagesChanged;
    }

    private void OnChatMessagesChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.Action == NotifyCollectionChangedAction.Add)
        {
            Dispatcher.InvokeAsync(() => ChatScrollViewer.ScrollToEnd(),
                System.Windows.Threading.DispatcherPriority.Background);
        }
    }

    public void RefreshLocalization()
    {
        var L = Helpers.LanguageManager.Get;
        PresetToggleLabel.Text = L("Chat_Presets");
        PromptEditorLabel.Text = L("Chat_PromptEditor");
        CancelLabel.Text = L("Chat_Cancel");
    }

    private void Help_Click(object sender, RoutedEventArgs e) =>
        HelpRequested?.Invoke(this, EventArgs.Empty);

    private void Close_Click(object sender, RoutedEventArgs e) =>
        CloseRequested?.Invoke(this, EventArgs.Empty);

    private void PromptEditor_Click(object sender, RoutedEventArgs e) =>
        PromptEditorRequested?.Invoke(this, EventArgs.Empty);

    private void PresetChip_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.DataContext is PresetPromptItem item
            && DataContext is ChatPanelViewModel vm)
        {
            vm.AiInput = item.PromptText;
            ChatInput.Focus();
        }
    }

    private async void ChatInput_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && Keyboard.Modifiers == ModifierKeys.None)
        {
            e.Handled = true;
            if (DataContext is ChatPanelViewModel vm)
                await vm.ExecuteSelectedPromptCommand.ExecuteAsync(null);
        }
    }

    private void InsertAiResponse_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string text })
            InsertToDocumentRequested?.Invoke(text);
    }

    private void CopyAiResponse_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string text })
            CopyResponseRequested?.Invoke(text);
    }
}
