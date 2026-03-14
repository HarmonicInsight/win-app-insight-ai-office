using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using InsightCommon.AI;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// IToolExecutor 実装 — ドキュメント生成ツールの実行ハンドラ
/// AiChatViewModel から呼ばれる。
/// </summary>
public class DocumentGenerationToolExecutor : IToolExecutor
{
    private readonly string? _outputDir;

    public DocumentGenerationToolExecutor(string? outputDir = null)
    {
        _outputDir = outputDir;
    }

    public async Task<ToolExecutionResult> ExecuteAsync(
        string toolName, JsonElement input, CancellationToken ct)
    {
        try
        {
            var (result, outputPath) = await FileGenerationExecutor.ExecuteAsync(
                toolName, input, _outputDir, ct);

            return new ToolExecutionResult
            {
                Content = result,
                IsError = false,
            };
        }
        catch (System.Exception ex)
        {
            return new ToolExecutionResult
            {
                Content = $"{{\"success\":false,\"message\":\"{ex.Message}\"}}",
                IsError = true,
            };
        }
    }
}
