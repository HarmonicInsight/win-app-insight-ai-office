using InsightCommon.AI;
using Xunit;

namespace InsightAiOffice.Core.Tests;

public class DocumentCompressorTests
{
    // ─── EstimateTokens ───────────────────────────────────────

    [Fact]
    public void EstimateTokens_EmptyString_ReturnsZero()
    {
        Assert.Equal(0, DocumentCompressor.EstimateTokens(""));
        Assert.Equal(0, DocumentCompressor.EstimateTokens(null!));
    }

    [Fact]
    public void EstimateTokens_JapaneseText_UsesHigherRatio()
    {
        // 日本語10文字 ≈ 15 tokens
        var tokens = DocumentCompressor.EstimateTokens("あいうえおかきくけこ");
        Assert.True(tokens >= 10 && tokens <= 20);
    }

    [Fact]
    public void EstimateTokens_EnglishText_UsesLowerRatio()
    {
        // 英語40文字 ≈ 10 tokens
        var tokens = DocumentCompressor.EstimateTokens("Hello World this is a test for tokens.");
        Assert.True(tokens >= 5 && tokens <= 20);
    }

    // ─── ShouldCompress ───────────────────────────────────────

    [Fact]
    public void ShouldCompress_SmallText_ReturnsFalse()
    {
        Assert.False(DocumentCompressor.ShouldCompress("短いテキスト"));
    }

    [Fact]
    public void ShouldCompress_LargeText_ReturnsTrue()
    {
        var largeText = new string('あ', 2000); // 2000文字 × 1.5 = 3000 tokens > 2000
        Assert.True(DocumentCompressor.ShouldCompress(largeText));
    }

    // ─── CompressSpreadsheet ──────────────────────────────────

    [Fact]
    public void CompressSpreadsheet_EmptyRows_ReturnsEmpty()
    {
        var result = DocumentCompressor.CompressSpreadsheet("Sheet1", [], 0, 0);
        Assert.Contains("空のシート", result);
    }

    [Fact]
    public void CompressSpreadsheet_IncludesHeaderAndStats()
    {
        var rows = new System.Collections.Generic.List<string[]>
        {
            new[] { "名前", "金額", "日付" },
            new[] { "田中", "10000", "2025-01-01" },
            new[] { "鈴木", "20000", "2025-01-02" },
            new[] { "佐藤", "30000", "2025-01-03" },
            new[] { "伊藤", "15000", "2025-01-04" },
            new[] { "渡辺", "25000", "2025-01-05" },
        };
        var header = rows[0];

        var result = DocumentCompressor.CompressSpreadsheet(
            "売上", rows, 6, 3, header);

        // 構造情報が含まれる
        Assert.Contains("売上", result);
        Assert.Contains("行数: 6", result);
        Assert.Contains("名前", result);

        // 統計が含まれる
        Assert.Contains("金額", result);

        // サンプルデータが含まれる
        Assert.Contains("田中", result);
    }

    [Fact]
    public void CompressSpreadsheet_LargeData_OnlyShowsSamples()
    {
        var rows = new System.Collections.Generic.List<string[]>();
        rows.Add(new[] { "ID", "Value" }); // header
        for (int i = 1; i <= 100; i++)
            rows.Add(new[] { i.ToString(), (i * 100).ToString() });

        var result = DocumentCompressor.CompressSpreadsheet(
            "Data", rows, 101, 2, rows[0]);

        // 先頭5行 + 末尾3行のサンプル
        Assert.Contains("row 1:", result);  // header
        Assert.Contains("row 5:", result);  // last of head sample
        Assert.Contains("row 101:", result); // tail sample

        // 中間行は含まれない
        Assert.DoesNotContain("row 50:", result);
    }

    // ─── CompressDocument ─────────────────────────────────────

    [Fact]
    public void CompressDocument_IncludesStructure()
    {
        var sections = new System.Collections.Generic.List<DocumentCompressor.DocumentSection>
        {
            new() { Heading = "第1章 総則", TextContent = "甲と乙は以下の契約を締結する。", Level = 1 },
            new() { Heading = "第2章 業務内容", TextContent = "乙は甲の指示に基づき業務を遂行する。", Level = 1 },
        };

        var result = DocumentCompressor.CompressDocument(
            "契約書.docx", sections, 50, 500, 5);

        Assert.Contains("契約書.docx", result);
        Assert.Contains("全5ページ", result);
        Assert.Contains("第1章 総則", result);
        Assert.Contains("第2章 業務内容", result);
        Assert.Contains("甲と乙", result);
    }

    // ─── CompressPresentation ─────────────────────────────────

    [Fact]
    public void CompressPresentation_IncludesToc()
    {
        var slides = new System.Collections.Generic.List<DocumentCompressor.SlideInfo>
        {
            new() { Number = 1, Title = "タイトルスライド", FullText = "2025年度 営業戦略", Notes = "" },
            new() { Number = 2, Title = "市場分析", FullText = "市場規模 5000億円...", Notes = "ここで詳細を説明" },
            new() { Number = 3, Title = "まとめ", FullText = "以上が提案の概要です。", Notes = "" },
        };

        var result = DocumentCompressor.CompressPresentation("提案書.pptx", slides);

        Assert.Contains("全3スライド", result);
        Assert.Contains("タイトルスライド", result);
        Assert.Contains("市場分析", result);
        Assert.Contains("[ノート]", result);
    }

    // ─── CompressPdf ──────────────────────────────────────────

    [Fact]
    public void CompressPdf_IncludesPageInfo()
    {
        var pages = new System.Collections.Generic.List<DocumentCompressor.PageInfo>
        {
            new() { Number = 1, Text = "表紙: 月次レポート 2025年1月" },
            new() { Number = 2, Text = "売上サマリー: 総売上 1億2300万円..." },
        };

        var result = DocumentCompressor.CompressPdf("report.pdf", pages);

        Assert.Contains("全2ページ", result);
        Assert.Contains("ページ 1", result);
        Assert.Contains("表紙", result);
    }

    // ─── DrillDown ────────────────────────────────────────────

    [Fact]
    public void GetDetailRows_ReturnsCorrectRange()
    {
        var rows = new System.Collections.Generic.List<string[]>
        {
            new[] { "ID", "Name" },
            new[] { "1", "Alice" },
            new[] { "2", "Bob" },
            new[] { "3", "Charlie" },
            new[] { "4", "Diana" },
        };

        var result = DocumentCompressor.GetDetailRows(rows, rows[0], 2, 3);

        Assert.Contains("行 2〜3", result);
        Assert.Contains("Alice", result);
        Assert.Contains("Bob", result);
        Assert.DoesNotContain("Charlie", result);
    }

    [Fact]
    public void GetDetailRows_OutOfRange_ReturnsMessage()
    {
        var rows = new System.Collections.Generic.List<string[]>
        {
            new[] { "1", "A" },
        };

        var result = DocumentCompressor.GetDetailRows(rows, null, 5, 10);
        Assert.Contains("範囲外", result);
    }

    // ─── GetCompressionNotice ─────────────────────────────────

    [Fact]
    public void GetCompressionNotice_Ja_ReturnsJapanese()
    {
        var notice = DocumentCompressor.GetCompressionNotice("ja");
        Assert.Contains("圧縮", notice);
    }

    [Fact]
    public void GetCompressionNotice_En_ReturnsEnglish()
    {
        var notice = DocumentCompressor.GetCompressionNotice("en");
        Assert.Contains("Compressed", notice);
    }

    // ─── GetDrillDownToolDefinition ───────────────────────────

    [Fact]
    public void GetDrillDownToolDefinition_ReturnsValidTool()
    {
        var tool = DocumentCompressor.GetDrillDownToolDefinition();
        Assert.Equal("get_document_detail", tool.Name);
        Assert.NotNull(tool.InputSchema);
    }
}
