using System.Text.Json;
using System.Text.RegularExpressions;

namespace InsightAiOffice.App.Services;

/// <summary>
/// Parses and executes structured AI tool-call commands against the active document.
/// AI can output JSON tool blocks like: {"tool":"insert_text","args":{"text":"Hello"}}
/// </summary>
public partial class DocumentToolExecutor
{
    public record ToolCall(string Tool, JsonElement Args);
    public record ToolResult(bool Success, string Message);

    private readonly Action<string> _insertText;
    private readonly Func<string> _getSelectedText;
    private readonly Action<string> _setStatus;

    public DocumentToolExecutor(
        Action<string> insertText,
        Func<string> getSelectedText,
        Action<string> setStatus)
    {
        _insertText = insertText;
        _getSelectedText = getSelectedText;
        _setStatus = setStatus;
    }

    /// <summary>
    /// Parses AI response for tool-call JSON blocks and executes them.
    /// Returns (cleanedResponse, executedToolCount).
    /// </summary>
    public (string CleanedResponse, int ExecutedCount) ParseAndExecute(string aiResponse)
    {
        var toolCalls = ExtractToolCalls(aiResponse);
        if (toolCalls.Count == 0)
            return (aiResponse, 0);

        var executed = 0;
        foreach (var call in toolCalls)
        {
            var result = Execute(call);
            if (result.Success) executed++;
            _setStatus(result.Message);
        }

        // Remove tool-call JSON blocks from the displayed response
        var cleaned = ToolBlockRegex().Replace(aiResponse, "").Trim();
        if (string.IsNullOrWhiteSpace(cleaned))
            cleaned = $"[{executed} 件のドキュメント操作を実行しました]";

        return (cleaned, executed);
    }

    private List<ToolCall> ExtractToolCalls(string text)
    {
        var calls = new List<ToolCall>();
        foreach (Match match in ToolBlockRegex().Matches(text))
        {
            try
            {
                var json = match.Value;
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                if (root.TryGetProperty("tool", out var toolProp) &&
                    root.TryGetProperty("args", out var argsProp))
                {
                    calls.Add(new ToolCall(toolProp.GetString() ?? "", argsProp.Clone()));
                }
            }
            catch (JsonException) { /* skip malformed JSON — AI may produce partial blocks */ }
        }
        return calls;
    }

    private ToolResult Execute(ToolCall call)
    {
        return call.Tool switch
        {
            "insert_text" => ExecuteInsertText(call.Args),
            "replace_selection" => ExecuteReplaceSelection(call.Args),
            _ => new ToolResult(false, $"未対応のツール: {call.Tool}")
        };
    }

    private ToolResult ExecuteInsertText(JsonElement args)
    {
        if (!args.TryGetProperty("text", out var textProp))
            return new ToolResult(false, "insert_text: text パラメータが必要です");

        var text = textProp.GetString() ?? "";
        _insertText(text);
        return new ToolResult(true, $"テキストを挿入しました（{text.Length} 文字）");
    }

    private ToolResult ExecuteReplaceSelection(JsonElement args)
    {
        if (!args.TryGetProperty("text", out var textProp))
            return new ToolResult(false, "replace_selection: text パラメータが必要です");

        var text = textProp.GetString() ?? "";
        _insertText(text); // InsertText on selection replaces it
        return new ToolResult(true, $"選択テキストを置換しました（{text.Length} 文字）");
    }

    [GeneratedRegex(@"\{""tool""\s*:\s*""[^""]+"".*?""args""\s*:\s*\{[^}]*\}\s*\}", RegexOptions.Compiled | RegexOptions.Singleline)]
    private static partial Regex ToolBlockRegex();
}
