using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using InsightCommon.AI;
using InsightCommon.Addon;
using InsightAiOffice.App.Helpers;
using InsightAiOffice.App.Services;

namespace InsightAiOffice.App.ViewModels;

// ── Models ───────────────────────────────────────────────────

public class ChatMessage(string content, string role, bool isThinking = false)
{
    public string Content { get; } = content;
    public string Role { get; } = role; // "User", "Assistant", "System"
    public bool IsThinking { get; } = isThinking;
    public Visibility ActionVisibility => Role == "User" || IsThinking ? Visibility.Collapsed : Visibility.Visible;
}

public class PresetPromptGroup
{
    public string CategoryName { get; set; } = "";
    public ObservableCollection<PresetPromptItem> Prompts { get; set; } = [];
}

public class PresetPromptItem
{
    public string Label { get; set; } = "";
    public string Id { get; set; } = "";
    public string PromptText { get; set; } = "";
}

// ── ViewModel ────────────────────────────────────────────────

public partial class ChatPanelViewModel : ObservableObject
{
    private readonly AiService _aiService;
    private readonly PromptPresetService _presetService;
    private readonly ReferenceMaterialsService _referenceService;
    private CancellationTokenSource? _cts;
    private DocumentToolExecutor? _toolExecutor;

    // ── Callbacks to MainWindow ──
    public Func<string>? GetDocumentContent { get; set; }
    public Action<string>? SetStatusText { get; set; }
    public Func<string, Task<bool>>? OnBeforeSend { get; set; }

    // ── Observable collections ──
    public ObservableCollection<ChatMessage> ChatMessages { get; } = [];
    public ObservableCollection<PresetPromptGroup> PresetPromptGroups { get; } = [];

    // ── Properties ──
    [ObservableProperty] private string _panelTitle = "AI コンシェルジュ";
    [ObservableProperty] private bool _isPresetsExpanded;
    [ObservableProperty] private string _aiInput = "";
    [ObservableProperty] private bool _isSending;
    [ObservableProperty] private string _aiProcessingModelName = "";
    [ObservableProperty] private string _aiProcessingText = "";

    public bool IsApiKeySet => _aiService.IsConfigured;

    public ChatPanelViewModel(
        AiService aiService,
        PromptPresetService presetService,
        ReferenceMaterialsService referenceService)
    {
        _aiService = aiService;
        _presetService = presetService;
        _referenceService = referenceService;
    }

    public void SetToolExecutor(DocumentToolExecutor executor) => _toolExecutor = executor;

    // ── Localization ──

    public void RefreshLocalization()
    {
        PanelTitle = LanguageManager.Get("Pane_Chat");
    }

    // ── Preset Prompts ──

    public void LoadPresetGroups()
    {
        PresetPromptGroups.Clear();
        foreach (var group in _presetService.LoadAll().GroupBy(p => p.Category ?? ""))
        {
            PresetPromptGroups.Add(new PresetPromptGroup
            {
                CategoryName = group.Key,
                Prompts = new ObservableCollection<PresetPromptItem>(
                    group.Select(p => new PresetPromptItem
                    {
                        Label = p.Name,
                        Id = p.Id,
                        PromptText = p.SystemPrompt,
                    }))
            });
        }
    }

    // ── Message helpers ──

    public void AddUserMessage(string text) =>
        ChatMessages.Add(new ChatMessage(text, "User"));

    public void AddAssistantMessage(string text) =>
        ChatMessages.Add(new ChatMessage(text, "Assistant"));

    public void AddSystemMessage(string text) =>
        ChatMessages.Add(new ChatMessage(text, "System"));

    // ── Commands ──

    [RelayCommand]
    private async Task ExecuteSelectedPromptAsync()
    {
        var input = AiInput?.Trim();
        if (string.IsNullOrEmpty(input)) return;
        AiInput = "";
        await SendMessageAsync(input);
    }

    [RelayCommand]
    private void CancelSend() => _cts?.Cancel();

    [RelayCommand]
    private void ClearChat()
    {
        ChatMessages.Clear();
        SetStatusText?.Invoke(LanguageManager.Get("Doc_ChatCleared"));
    }

    [RelayCommand]
    private void OpenAiSettings()
    {
        var owner = Application.Current.Windows
            .OfType<Window>().FirstOrDefault(w => w.IsActive)
            ?? Application.Current.MainWindow;
        if (owner == null) return;

        var theme = InsightCommon.Theme.InsightTheme.Create();
        _aiService.ShowSettingsDialog(owner, theme, "Insight AI Office",
            LanguageManager.CurrentLanguage);
        SetStatusText?.Invoke(_aiService.IsConfigured
            ? LanguageManager.Format("AI_ConfigDone", _aiService.CurrentModel)
            : LanguageManager.Get("AI_ConfigNotSet"));
        OnPropertyChanged(nameof(IsApiKeySet));
    }

    // ── Core send logic ──

    public async Task SendMessageAsync(string input)
    {
        if (string.IsNullOrEmpty(input) || IsSending) return;

        // Allow MainWindow to intercept (e.g. _pendingDocGeneration)
        if (OnBeforeSend != null && await OnBeforeSend(input))
            return;

        if (!_aiService.IsConfigured)
        {
            AddAssistantMessage(LanguageManager.Get("AI_ConfigRequired"));
            return;
        }

        AddUserMessage(input);
        IsSending = true;
        AiProcessingModelName = _aiService.CurrentModel;
        AiProcessingText = LanguageManager.Get("AI_Thinking");

        var thinkingMsg = new ChatMessage(LanguageManager.Get("AI_Thinking"), "Assistant", isThinking: true);
        ChatMessages.Add(thinkingMsg);
        SetStatusText?.Invoke(LanguageManager.Get("AI_Responding"));

        try
        {
            _cts = new CancellationTokenSource(TimeSpan.FromSeconds(120));

            var docContent = GetDocumentContent?.Invoke() ?? "";
            var refContext = _referenceService.BuildAutoContext(input, 3);
            var systemPrompt = BuildSystemPrompt(docContent, refContext);

            var history = ChatMessages
                .Where(m => !m.IsThinking)
                .Select(m => new AiMessage
                {
                    Content = m.Content,
                    Role = m.Role == "User" ? AiMessageRole.User : AiMessageRole.Assistant,
                })
                .ToList();

            if (history.Count > 0) history.RemoveAt(history.Count - 1);

            var response = await _aiService.SendChatAsync(history, input, systemPrompt);

            ChatMessages.Remove(thinkingMsg);
            if (_toolExecutor != null)
            {
                var (cleaned, execCount) = _toolExecutor.ParseAndExecute(response);
                AddAssistantMessage(cleaned);
                SetStatusText?.Invoke(execCount > 0
                    ? LanguageManager.Format("AI_ToolExecuted", execCount)
                    : LanguageManager.Get("Status_Ready"));
            }
            else
            {
                AddAssistantMessage(response);
                SetStatusText?.Invoke(LanguageManager.Get("Status_Ready"));
            }
        }
        catch (OperationCanceledException)
        {
            ChatMessages.Remove(thinkingMsg);
            AddSystemMessage(LanguageManager.Get("AI_Timeout"));
            SetStatusText?.Invoke(LanguageManager.Get("AI_Error"));
        }
        catch (Exception ex)
        {
            ChatMessages.Remove(thinkingMsg);
            AddAssistantMessage($"{LanguageManager.Get("Error_Title")}: {ex.Message}");
            SetStatusText?.Invoke(LanguageManager.Get("AI_Error"));
        }
        finally
        {
            IsSending = false;
            _cts?.Dispose();
            _cts = null;
        }
    }

    /// <summary>Execute a prompt from ribbon AI action buttons.</summary>
    public async Task ExecuteActionAsync(string prompt)
    {
        if (!_aiService.IsConfigured)
        {
            SetStatusText?.Invoke(LanguageManager.Get("AI_ConfigShort"));
            return;
        }

        AddUserMessage(prompt);
        IsSending = true;
        AiProcessingModelName = _aiService.CurrentModel;
        AiProcessingText = LanguageManager.Get("AI_Processing");

        var thinkingMsg = new ChatMessage(LanguageManager.Get("AI_Thinking"), "Assistant", isThinking: true);
        ChatMessages.Add(thinkingMsg);
        SetStatusText?.Invoke(LanguageManager.Get("AI_Processing"));

        try
        {
            var docContent = GetDocumentContent?.Invoke() ?? "";
            var refContext = _referenceService.Count > 0
                ? _referenceService.BuildAutoContext(prompt, 3)
                : "";
            var systemPrompt = BuildSystemPrompt(docContent, refContext);
            var response = await _aiService.SendMessageAsync(prompt, systemPrompt);

            ChatMessages.Remove(thinkingMsg);
            AddAssistantMessage(response);
            SetStatusText?.Invoke(LanguageManager.Get("Status_Ready"));
        }
        catch (Exception ex)
        {
            ChatMessages.Remove(thinkingMsg);
            AddAssistantMessage($"{LanguageManager.Get("Error_Title")}: {ex.Message}");
            SetStatusText?.Invoke(LanguageManager.Get("AI_Error"));
        }
        finally
        {
            IsSending = false;
        }
    }

    /// <summary>Execute with a custom system prompt (for specialized actions).</summary>
    public async Task ExecuteWithSystemPromptAsync(string userPrompt, string systemPrompt)
    {
        if (!_aiService.IsConfigured)
        {
            SetStatusText?.Invoke(LanguageManager.Get("AI_ConfigShort"));
            return;
        }

        AddUserMessage(userPrompt);
        IsSending = true;
        AiProcessingModelName = _aiService.CurrentModel;
        AiProcessingText = LanguageManager.Get("AI_Processing");

        var thinkingMsg = new ChatMessage(LanguageManager.Get("AI_Thinking"), "Assistant", isThinking: true);
        ChatMessages.Add(thinkingMsg);
        SetStatusText?.Invoke(LanguageManager.Get("AI_Processing"));

        try
        {
            var response = await _aiService.SendMessageAsync(userPrompt, systemPrompt);
            ChatMessages.Remove(thinkingMsg);
            AddAssistantMessage(response);
            SetStatusText?.Invoke(LanguageManager.Get("Status_Ready"));
            return;
        }
        catch (Exception ex)
        {
            ChatMessages.Remove(thinkingMsg);
            AddAssistantMessage($"{LanguageManager.Get("Error_Title")}: {ex.Message}");
            SetStatusText?.Invoke(LanguageManager.Get("AI_Error"));
        }
        finally
        {
            IsSending = false;
        }
    }

    /// <summary>Get the last assistant message content (for document generation etc.).</summary>
    public string? GetLastAssistantMessage() =>
        ChatMessages.LastOrDefault(m => m.Role == "Assistant" && !m.IsThinking)?.Content;

    public void NotifyApiKeyChanged() => OnPropertyChanged(nameof(IsApiKeySet));

    private string BuildSystemPrompt(string docContent, string refContext = "")
    {
        var prompt = "あなたは Insight AI Office のアシスタントです。" +
            "ユーザーが開いているドキュメントについて質問に答え、分析・校正・要約などを支援してください。" +
            "回答は簡潔に、日本語で行ってください。";

        var defaultPreset = _presetService.GetDefault();
        if (defaultPreset != null && !string.IsNullOrEmpty(defaultPreset.SystemPrompt))
            prompt = defaultPreset.SystemPrompt;

        if (!string.IsNullOrEmpty(docContent))
            prompt += $"\n\n--- 現在のドキュメント内容 ---\n{docContent}\n--- ここまで ---";

        if (!string.IsNullOrEmpty(refContext))
            prompt += $"\n\n{refContext}";

        return prompt;
    }
}
