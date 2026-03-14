using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using InsightAiOffice.App.Services.DocumentGeneration;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// AI ツール実行 — ファイル生成系ツールのディスパッチャ
/// </summary>
public static class FileGenerationExecutor
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    /// <summary>
    /// ツール名に応じたファイル生成を実行
    /// </summary>
    public static async Task<(string result, string? outputPath)> ExecuteAsync(
        string toolName, JsonElement input, string? fallbackDir = null,
        CancellationToken ct = default)
    {
        return toolName switch
        {
            "generate_report" => ExecuteGenerateReport(input, fallbackDir),
            "generate_presentation" => await ExecuteGeneratePresentationAsync(input, fallbackDir, null, ct),
            "generate_spreadsheet" => ExecuteGenerateSpreadsheet(input, fallbackDir),
            "generate_presentation_from_template" => await ExecuteGeneratePresentationAsync(input, fallbackDir, GetTemplatePath(input), ct),
            "rewrite_document" => ExecuteRewriteDocument(input, fallbackDir),
            "batch_generate" => ExecuteBatchGenerate(input, fallbackDir),
            _ => ($"Unknown tool: {toolName}", null),
        };
    }



    // --- generate_report ---

    private static (string result, string? outputPath) ExecuteGenerateReport(
        JsonElement input, string? fallbackDir)
    {
        var title = input.GetProperty("title").GetString() ?? "Report";
        var format = input.TryGetProperty("format", out var fmt) ? fmt.GetString() ?? "docx" : "docx";
        var theme = input.TryGetProperty("theme", out var th) ? th.GetString() : null;

        var report = new ReportStructure
        {
            Title = title,
            Author = input.TryGetProperty("author", out var a) ? a.GetString() ?? "" : "",
            Date = input.TryGetProperty("date", out var d) ? d.GetString() ?? "" : "",
        };

        if (input.TryGetProperty("sections", out var sectionsJson))
        {
            report.Sections = JsonSerializer.Deserialize<List<ReportSection>>(
                sectionsJson.GetRawText(), JsonOptions) ?? new();
        }

        var ext = format == "html" ? ".html" : ".docx";
        var outputPath = ResolveOutputPath(input, title, ext, fallbackDir);

        if (format == "html")
            ReportRendererService.RenderToHtml(report, outputPath, theme);
        else
            ReportRendererService.Render(report, outputPath, theme);

        return (JsonSerializer.Serialize(new
        {
            success = true,
            output_path = outputPath,
            format,
            message = $"レポートを生成しました: {Path.GetFileName(outputPath)}"
        }), outputPath);
    }

    // --- generate_presentation / generate_presentation_from_template ---

    private static async Task<(string result, string? outputPath)> ExecuteGeneratePresentationAsync(
        JsonElement input, string? fallbackDir, string? templatePath, CancellationToken ct)
    {
        var title = input.GetProperty("title").GetString() ?? "Presentation";
        var slidesJson = input.GetProperty("slides");

        var slides = JsonSerializer.Deserialize<List<SlideSpecItem>>(
            slidesJson.GetRawText(), JsonOptions)
            ?? throw new InvalidOperationException("Failed to deserialize slides.");

        var outputPath = ResolveOutputPath(input, title, ".pptx", fallbackDir);

        await PptxRendererService.RenderAsync(slides, outputPath, templatePath, ct);

        var mode = templatePath != null ? "テンプレートベース" : "新規";
        return (JsonSerializer.Serialize(new
        {
            success = true,
            output_path = outputPath,
            slide_count = slides.Count,
            template_used = templatePath != null,
            message = $"{mode}プレゼンテーションを生成しました: {Path.GetFileName(outputPath)}"
        }), outputPath);
    }

    // --- generate_spreadsheet ---

    private static (string result, string? outputPath) ExecuteGenerateSpreadsheet(
        JsonElement input, string? fallbackDir)
    {
        var title = input.GetProperty("title").GetString() ?? "Spreadsheet";
        var theme = input.TryGetProperty("theme", out var th) ? th.GetString() : null;
        var sheetsJson = input.GetProperty("sheets");

        var spec = new SpreadsheetStructure
        {
            Title = title,
            Sheets = JsonSerializer.Deserialize<List<SimpleSheetData>>(
                sheetsJson.GetRawText(), JsonOptions)
                ?? throw new InvalidOperationException("Failed to deserialize sheets.")
        };

        var outputPath = ResolveOutputPath(input, title, ".xlsx", fallbackDir);
        SpreadsheetRendererService.Render(spec, outputPath, theme);

        return (JsonSerializer.Serialize(new
        {
            success = true,
            output_path = outputPath,
            sheet_count = spec.Sheets.Count,
            message = $"スプレッドシートを生成しました: {Path.GetFileName(outputPath)}"
        }), outputPath);
    }

    // --- batch_generate (差し込み印刷) ---

    private static (string result, string? outputPath) ExecuteBatchGenerate(
        JsonElement input, string? fallbackDir)
    {
        // TODO: MailMergeService 統合時に実装
        return (JsonSerializer.Serialize(new
        {
            success = false,
            message = "差し込み印刷は現在準備中です"
        }), null);
    }

    // --- Helpers ---

    private static string? GetTemplatePath(JsonElement input)
    {
        return input.TryGetProperty("template_path", out var tp) ? tp.GetString() : null;
    }

    // --- rewrite_document ---

    private static (string result, string? outputPath) ExecuteRewriteDocument(
        JsonElement input, string? fallbackDir)
    {
        var sourcePath = input.GetProperty("source_path").GetString()
            ?? throw new InvalidOperationException("source_path is required");

        if (!File.Exists(sourcePath))
            return (JsonSerializer.Serialize(new { success = false, message = $"ファイルが見つかりません: {sourcePath}" }), null);

        var title = input.TryGetProperty("title", out var t) ? t.GetString() ?? Path.GetFileNameWithoutExtension(sourcePath) : Path.GetFileNameWithoutExtension(sourcePath);
        var outputPath = ResolveOutputPath(input, title, ".docx", fallbackDir);

        // 元ファイルをコピー
        File.Copy(sourcePath, outputPath, overwrite: true);

        // テキスト置換
        using var doc = new Syncfusion.DocIO.DLS.WordDocument(outputPath, Syncfusion.DocIO.FormatType.Docx);

        if (input.TryGetProperty("replacements", out var replacements))
        {
            foreach (var r in replacements.EnumerateArray())
            {
                var find = r.GetProperty("find").GetString() ?? "";
                var replace = r.GetProperty("replace").GetString() ?? "";
                if (!string.IsNullOrEmpty(find))
                    doc.Replace(find, replace, false, false);
            }
        }

        doc.Save(outputPath, Syncfusion.DocIO.FormatType.Docx);

        return (JsonSerializer.Serialize(new
        {
            success = true,
            output_path = outputPath,
            message = $"文書を書き換えました: {Path.GetFileName(outputPath)}"
        }), outputPath);
    }

    // --- Helpers ---

    private static string ResolveOutputPath(JsonElement input, string title, string extension, string? fallbackDir)
    {
        var outputPath = input.TryGetProperty("output_path", out var op) ? op.GetString() : null;
        if (!string.IsNullOrEmpty(outputPath)) return outputPath;

        var dir = fallbackDir ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "HarmonicInsight", "InsightAiOffice", "Artifacts", "Default");
        var safeName = string.Join("_", title.Split(Path.GetInvalidFileNameChars()));
        if (string.IsNullOrWhiteSpace(safeName)) safeName = "Output";
        return Path.Combine(dir, $"{safeName}{extension}");
    }
}
