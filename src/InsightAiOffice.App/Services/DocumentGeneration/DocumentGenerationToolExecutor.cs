using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using InsightCommon.AI;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// ドキュメント編集操作のコールバック
/// MainWindow から渡されて、Syncfusion エディタを操作する。
/// </summary>
public class DocumentEditorCallbacks
{
    /// <summary>テキストを検索して赤い取り消し線 + 修正案を挿入</summary>
    public Func<string, string, string?, bool>? MarkCorrection { get; set; }

    /// <summary>テキストを検索してコメントを挿入</summary>
    public Func<string, string, bool>? AddComment { get; set; }

    /// <summary>テキストを検索して蛍光マーカーで強調</summary>
    public Func<string, string, bool>? HighlightText { get; set; }

    /// <summary>テキストを検索して置換</summary>
    public Func<string, string, int>? FindAndReplace { get; set; }

    /// <summary>現在のドキュメントの全テキストを取得</summary>
    public Func<string>? GetDocumentText { get; set; }

    /// <summary>Wordドキュメントにテキストを挿入（after=null で末尾追記）</summary>
    public Func<string, string?, bool>? InsertDocumentText { get; set; }

    /// <summary>Excelセルに値を書き込み（sheetName, cellRef, value）→ 成功数</summary>
    public Func<string?, List<(string cell, string value)>, int>? EditSpreadsheetCells { get; set; }

    /// <summary>テキストファイルを新規作成して新しいタブで開く（title, content）→ filePath</summary>
    public Func<string, string, string?>? CreateTextFile { get; set; }
}

/// <summary>
/// IToolExecutor 実装 — ドキュメント生成 + エディタ操作ツールの実行ハンドラ
/// </summary>
public class DocumentGenerationToolExecutor : IToolExecutor
{
    private readonly string? _outputDir;
    private readonly DocumentEditorCallbacks? _editorCallbacks;

    public DocumentGenerationToolExecutor(string? outputDir = null, DocumentEditorCallbacks? editorCallbacks = null)
    {
        _outputDir = outputDir;
        _editorCallbacks = editorCallbacks;
    }

    public async Task<ToolExecutionResult> ExecuteAsync(
        string toolName, JsonElement input, CancellationToken ct)
    {
        try
        {
            // エディタ操作ツール
            var editorResult = toolName switch
            {
                "mark_correction" => ExecuteMarkCorrection(input),
                "add_comment" => ExecuteAddComment(input),
                "highlight_text" => ExecuteHighlightText(input),
                "find_and_replace" => ExecuteFindAndReplace(input),
                "insert_document_text" => ExecuteInsertDocumentText(input),
                "edit_spreadsheet_cells" => ExecuteEditSpreadsheetCells(input),
                "create_text_file" => ExecuteCreateTextFile(input),
                _ => (string?)null,
            };

            if (editorResult != null)
                return new ToolExecutionResult { Content = editorResult, IsError = false };

            // ファイル生成ツール
            var (result, outputPath) = await FileGenerationExecutor.ExecuteAsync(
                toolName, input, _outputDir, ct);

            return new ToolExecutionResult { Content = result, IsError = false };
        }
        catch (Exception ex)
        {
            return new ToolExecutionResult
            {
                Content = JsonSerializer.Serialize(new { success = false, message = ex.Message }),
                IsError = true,
            };
        }
    }

    private string ExecuteMarkCorrection(JsonElement input)
    {
        var original = input.GetProperty("original_text").GetString() ?? "";
        var correction = input.GetProperty("correction").GetString() ?? "";
        var reason = input.TryGetProperty("reason", out var r) ? r.GetString() : null;

        if (_editorCallbacks?.MarkCorrection == null)
            return JsonSerializer.Serialize(new { success = false, message = "エディタが開かれていません" });

        var ok = _editorCallbacks.MarkCorrection(original, correction, reason);
        return JsonSerializer.Serialize(new
        {
            success = ok,
            original_text = original,
            correction,
            reason,
            message = ok ? $"修正マーク: 「{Truncate(original)}」→「{Truncate(correction)}」" : $"テキストが見つかりません: 「{Truncate(original)}」",
        });
    }

    private string ExecuteAddComment(JsonElement input)
    {
        var targetText = input.GetProperty("target_text").GetString() ?? "";
        var comment = input.GetProperty("comment").GetString() ?? "";

        if (_editorCallbacks?.AddComment == null)
            return JsonSerializer.Serialize(new { success = false, message = "エディタが開かれていません" });

        var ok = _editorCallbacks.AddComment(targetText, comment);
        return JsonSerializer.Serialize(new
        {
            success = ok,
            target_text = targetText,
            comment,
            message = ok ? $"コメント追加: 「{Truncate(targetText)}」" : $"テキストが見つかりません: 「{Truncate(targetText)}」",
        });
    }

    private string ExecuteHighlightText(JsonElement input)
    {
        var targetText = input.GetProperty("target_text").GetString() ?? "";
        var color = input.TryGetProperty("color", out var c) ? c.GetString() ?? "yellow" : "yellow";

        if (_editorCallbacks?.HighlightText == null)
            return JsonSerializer.Serialize(new { success = false, message = "エディタが開かれていません" });

        var ok = _editorCallbacks.HighlightText(targetText, color);
        return JsonSerializer.Serialize(new
        {
            success = ok,
            target_text = targetText,
            color,
            message = ok ? $"ハイライト: 「{Truncate(targetText)}」" : $"テキストが見つかりません: 「{Truncate(targetText)}」",
        });
    }

    private string ExecuteFindAndReplace(JsonElement input)
    {
        var find = input.GetProperty("find").GetString() ?? "";
        var replace = input.GetProperty("replace").GetString() ?? "";

        if (_editorCallbacks?.FindAndReplace == null)
            return JsonSerializer.Serialize(new { success = false, message = "エディタが開かれていません" });

        var count = _editorCallbacks.FindAndReplace(find, replace);
        return JsonSerializer.Serialize(new
        {
            success = count > 0,
            find,
            replace,
            replaced_count = count,
            message = count > 0 ? $"{count} 箇所を置換しました" : $"「{Truncate(find)}」が見つかりません",
        });
    }

    private string ExecuteInsertDocumentText(JsonElement input)
    {
        var text = input.GetProperty("text").GetString() ?? "";
        var after = input.TryGetProperty("after", out var a) ? a.GetString() : null;

        if (_editorCallbacks?.InsertDocumentText == null)
            return JsonSerializer.Serialize(new { success = false, message = "Wordドキュメントが開かれていません" });

        var ok = _editorCallbacks.InsertDocumentText(text, after);
        return JsonSerializer.Serialize(new
        {
            success = ok,
            message = ok
                ? (after != null ? $"「{Truncate(after)}」の後にテキストを挿入しました" : "ドキュメント末尾にテキストを追記しました")
                : (after != null ? $"「{Truncate(after)}」が見つかりません" : "挿入に失敗しました"),
        });
    }

    private string ExecuteEditSpreadsheetCells(JsonElement input)
    {
        var sheet = input.TryGetProperty("sheet", out var s) ? s.GetString() : null;
        var cellsJson = input.GetProperty("cells");
        var cells = new List<(string cell, string value)>();
        foreach (var item in cellsJson.EnumerateArray())
        {
            var cell = item.GetProperty("cell").GetString() ?? "";
            var value = item.GetProperty("value").GetString() ?? "";
            cells.Add((cell, value));
        }

        if (_editorCallbacks?.EditSpreadsheetCells == null)
            return JsonSerializer.Serialize(new { success = false, message = "Excelスプレッドシートが開かれていません" });

        var count = _editorCallbacks.EditSpreadsheetCells(sheet, cells);
        return JsonSerializer.Serialize(new
        {
            success = count > 0,
            updated_count = count,
            total = cells.Count,
            message = $"{count}/{cells.Count} セルを更新しました",
        });
    }

    private string ExecuteCreateTextFile(JsonElement input)
    {
        var title = input.GetProperty("title").GetString() ?? "新規テキスト";
        var content = input.GetProperty("content").GetString() ?? "";

        if (_editorCallbacks?.CreateTextFile == null)
            return JsonSerializer.Serialize(new { success = false, message = "テキストファイル作成機能が利用できません" });

        var filePath = _editorCallbacks.CreateTextFile(title, content);
        return JsonSerializer.Serialize(new
        {
            success = filePath != null,
            file_path = filePath,
            message = filePath != null ? $"新しいタブで開きました: {title}.txt" : "作成に失敗しました",
        });
    }

    private static string Truncate(string s, int max = 30) =>
        s.Length <= max ? s : s[..max] + "...";
}
