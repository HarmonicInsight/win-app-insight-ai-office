namespace InsightAiOffice.Core.Services;

/// <summary>
/// プロンプト管理のインターフェース。
/// 実装は InsightCommon.AI.PromptPresetService を DI 経由で使用。
/// </summary>
public interface IPromptService
{
    IReadOnlyList<PromptEntry> GetPrompts();
    void AddPrompt(string title, string content, string category);
    void DeletePrompt(string id);
}

public record PromptEntry(string Id, string Title, string Content, string Category, DateTime CreatedAt);
