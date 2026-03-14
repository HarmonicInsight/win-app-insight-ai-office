using System;
using System.Collections.Generic;
using System.IO;
using InsightAiOffice.App.Services.DocumentGeneration;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using HAlign = Syncfusion.DocIO.DLS.HorizontalAlignment;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// ReportStructure（AI 生成 JSON）→ .docx プレミアムレンダリング（Syncfusion DocIO）
///
/// Ivory &amp; Gold テーマ — 法人向けプロフェッショナル品質
/// </summary>
public static class ReportRendererService
{
    private static readonly System.Drawing.Color SuccessGreen = System.Drawing.Color.FromArgb(0x16, 0xA3, 0x4A);
    private static readonly System.Drawing.Color ErrorRed = System.Drawing.Color.FromArgb(0xDC, 0x26, 0x26);

    private const string FontBody = "Yu Gothic UI";
    private const string FontHeading = "Yu Gothic UI";

    // 現在のテーマ（スレッドローカル的に使用）
    [ThreadStatic] private static DocumentColorTheme? _currentTheme;

    // =========================================================================
    // Public API
    // =========================================================================

    public static string Render(ReportStructure report, string outputPath, string? themeName = null)
    {
        _currentTheme = DocumentColorTheme.FromName(themeName);
        using var doc = new WordDocument();
        SetupDocument(doc);
        var section = doc.AddSection();
        SetupSection(section);

        foreach (var s in report.Sections)
            RenderSection(section, s);

        AddFooter(section);
        doc.Save(outputPath, FormatType.Docx);
        return outputPath;
    }

    public static string RenderToHtml(ReportStructure report, string outputPath, string? themeName = null)
    {
        _currentTheme = DocumentColorTheme.FromName(themeName);
        using var doc = new WordDocument();
        SetupDocument(doc);
        var section = doc.AddSection();
        SetupSection(section);

        foreach (var s in report.Sections)
            RenderSection(section, s);

        doc.Save(outputPath, FormatType.Html);
        return outputPath;
    }

    // テーマカラー取得ヘルパー
    private static DocumentColorTheme T => _currentTheme ?? DocumentColorTheme.IvoryGold;

    // =========================================================================
    // Document / Section setup
    // =========================================================================

    private static void SetupDocument(WordDocument doc)
    {
        // Syncfusion DocIO — ドキュメント設定
    }

    private static void SetupSection(IWSection section)
    {
        section.PageSetup.Margins.Top = 72;
        section.PageSetup.Margins.Bottom = 56;
        section.PageSetup.Margins.Left = 72;
        section.PageSetup.Margins.Right = 72;
    }

    private static void AddFooter(IWSection section)
    {
        // フッター: 左にゴールドライン + 右にページ番号
        var footer = section.HeadersFooters.Footer;
        var footerPara = footer.AddParagraph();
        footerPara.ParagraphFormat.HorizontalAlignment = HAlign.Right;
        footerPara.ParagraphFormat.Borders.Top.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Single;
        footerPara.ParagraphFormat.Borders.Top.Color = T.BorderColor;
        footerPara.ParagraphFormat.Borders.Top.Space = 6;

        var prefixRun = footerPara.AppendText("Insight Agent  —  ");
        prefixRun.CharacterFormat.FontName = FontBody;
        prefixRun.CharacterFormat.FontSize = 8;
        prefixRun.CharacterFormat.TextColor = T.TextSecondary;

        footerPara.AppendField("Page", FieldType.FieldPage);
        var slashRun = footerPara.AppendText(" / ");
        slashRun.CharacterFormat.FontSize = 8;
        slashRun.CharacterFormat.TextColor = T.TextSecondary;
        footerPara.AppendField("NumPages", FieldType.FieldNumPages);
    }

    // =========================================================================
    // Section router
    // =========================================================================

    private static void RenderSection(IWSection section, ReportSection s)
    {
        switch (s.Type)
        {
            case "title": RenderTitle(section, s); break;
            case "heading": RenderHeading(section, s); break;
            case "summary": RenderSummaryBox(section, s); break;
            case "recommendation": RenderRecommendationBox(section, s); break;
            case "text": RenderText(section, s); break;
            case "bullet_list": RenderBulletList(section, s); break;
            case "table":
            case "comparison": RenderTable(section, s); break;
            case "chart": RenderChartAsTable(section, s); break;
            case "key_metrics": RenderKeyMetrics(section, s); break;
            case "page_break":
                section.AddParagraph().AppendBreak(BreakType.PageBreak);
                break;
            default:
                if (!string.IsNullOrEmpty(s.Content)) RenderText(section, s);
                break;
        }

        foreach (var child in s.Sections)
            RenderSection(section, child);
    }

    // =========================================================================
    // Title — プレミアム表紙
    // =========================================================================

    private static void RenderTitle(IWSection section, ReportSection s)
    {
        // 上部スペース
        for (int i = 0; i < 6; i++)
        {
            var sp = section.AddParagraph();
            sp.AppendText(" ").CharacterFormat.FontSize = 14;
        }

        // ゴールド太帯
        var band = section.AddParagraph();
        band.ParagraphFormat.BackColor = T.Primary;
        band.ParagraphFormat.BeforeSpacing = 0;
        band.ParagraphFormat.AfterSpacing = 0;
        var bandRun = band.AppendText(" ");
        bandRun.CharacterFormat.FontSize = 8;

        // タイトル
        var titlePara = section.AddParagraph();
        titlePara.ParagraphFormat.HorizontalAlignment = HAlign.Center;
        titlePara.ParagraphFormat.BeforeSpacing = 36;
        titlePara.ParagraphFormat.AfterSpacing = 12;
        var titleRun = titlePara.AppendText(s.Title);
        titleRun.CharacterFormat.FontName = FontHeading;
        titleRun.CharacterFormat.FontSize = 28;
        titleRun.CharacterFormat.Bold = true;
        titleRun.CharacterFormat.TextColor = T.Primary;

        // サブタイトル
        if (!string.IsNullOrEmpty(s.Content))
        {
            var sub = section.AddParagraph();
            sub.ParagraphFormat.HorizontalAlignment = HAlign.Center;
            sub.ParagraphFormat.AfterSpacing = 8;
            var subRun = sub.AppendText(s.Content);
            subRun.CharacterFormat.FontName = FontBody;
            subRun.CharacterFormat.FontSize = 13;
            subRun.CharacterFormat.TextColor = T.TextSecondary;
        }

        // ゴールド細帯
        var band2 = section.AddParagraph();
        band2.ParagraphFormat.BackColor = T.Primary;
        band2.ParagraphFormat.BeforeSpacing = 24;
        band2.AppendText(" ").CharacterFormat.FontSize = 3;

        // 日付
        var datePara = section.AddParagraph();
        datePara.ParagraphFormat.HorizontalAlignment = HAlign.Center;
        datePara.ParagraphFormat.BeforeSpacing = 24;
        var dateRun = datePara.AppendText(DateTime.Now.ToString("yyyy年M月d日"));
        dateRun.CharacterFormat.FontName = FontBody;
        dateRun.CharacterFormat.FontSize = 11;
        dateRun.CharacterFormat.TextColor = T.TextSecondary;

        // 表紙後の改ページ
        section.AddParagraph().AppendBreak(BreakType.PageBreak);
    }

    // =========================================================================
    // Heading — ゴールド左ボーダー付き見出し
    // =========================================================================

    private static void RenderHeading(IWSection section, ReportSection s)
    {
        var para = section.AddParagraph();

        var (fontSize, level) = s.Level switch
        {
            1 => (16f, 1),
            2 => (14f, 2),
            3 => (12f, 3),
            _ => (12f, 3),
        };

        para.ParagraphFormat.BeforeSpacing = level == 1 ? 24 : 16;
        para.ParagraphFormat.AfterSpacing = 8;

        // H1: ゴールド左ボーダー + 背景
        if (level == 1)
        {
            para.ParagraphFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Single;
            para.ParagraphFormat.Borders.Left.Color = T.Primary;
            para.ParagraphFormat.Borders.Left.LineWidth = 4f;
            para.ParagraphFormat.Borders.Left.Space = 8;
            para.ParagraphFormat.BackColor = T.SummaryBg;
            para.ParagraphFormat.BeforeSpacing = 24;
            para.ParagraphFormat.AfterSpacing = 12;
        }

        // H2: ゴールド下線
        if (level == 2)
        {
            para.ParagraphFormat.Borders.Bottom.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Single;
            para.ParagraphFormat.Borders.Bottom.Color = T.PrimaryLight;
            para.ParagraphFormat.Borders.Bottom.LineWidth = 1f;
            para.ParagraphFormat.Borders.Bottom.Space = 4;
        }

        var run = para.AppendText(s.Title);
        run.CharacterFormat.FontName = FontHeading;
        run.CharacterFormat.FontSize = fontSize;
        run.CharacterFormat.Bold = true;
        run.CharacterFormat.TextColor = level == 1 ? T.PrimaryDark : T.PrimaryDark;

        if (!string.IsNullOrEmpty(s.Content))
            RenderText(section, s);
    }

    // =========================================================================
    // Summary Box — ゴールド左ボーダー + 背景のハイライト
    // =========================================================================

    private static void RenderSummaryBox(IWSection section, ReportSection s)
    {
        if (string.IsNullOrEmpty(s.Content) && string.IsNullOrEmpty(s.Title)) return;

        // ラベル
        if (!string.IsNullOrEmpty(s.Title))
        {
            var label = section.AddParagraph();
            label.ParagraphFormat.BeforeSpacing = 12;
            label.ParagraphFormat.AfterSpacing = 4;
            var labelRun = label.AppendText(s.Title);
            labelRun.CharacterFormat.FontName = FontHeading;
            labelRun.CharacterFormat.FontSize = 11;
            labelRun.CharacterFormat.Bold = true;
            labelRun.CharacterFormat.TextColor = T.PrimaryDark;
        }

        // ボックス
        var para = section.AddParagraph();
        para.ParagraphFormat.BackColor = T.SummaryBg;
        para.ParagraphFormat.Borders.Left.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Single;
        para.ParagraphFormat.Borders.Left.Color = T.Primary;
        para.ParagraphFormat.Borders.Left.LineWidth = 3f;
        para.ParagraphFormat.Borders.Left.Space = 10;
        para.ParagraphFormat.BeforeSpacing = 0;
        para.ParagraphFormat.AfterSpacing = 12;
        para.ParagraphFormat.LineSpacing = 20;

        var run = para.AppendText(s.Content ?? "");
        run.CharacterFormat.FontName = FontBody;
        run.CharacterFormat.FontSize = 11;
        run.CharacterFormat.TextColor = T.TextPrimary;
    }

    // =========================================================================
    // Recommendation Box — 強調ボックス（ゴールド背景）
    // =========================================================================

    private static void RenderRecommendationBox(IWSection section, ReportSection s)
    {
        if (string.IsNullOrEmpty(s.Content) && string.IsNullOrEmpty(s.Title)) return;

        var table = section.AddTable();
        table.ResetCells(1, 1);
        var cell = table.Rows[0].Cells[0];
        cell.CellFormat.BackColor = T.SummaryBg;
        cell.CellFormat.Borders.Color = T.PrimaryLight;
        cell.CellFormat.Borders.LineWidth = 1.5f;
        cell.Width = 430;

        // アイコン + ラベル
        var headerPara = cell.AddParagraph();
        headerPara.ParagraphFormat.AfterSpacing = 6;
        var iconRun = headerPara.AppendText("💡 ");
        iconRun.CharacterFormat.FontSize = 14;
        var headerRun = headerPara.AppendText(s.Title ?? "提案・推奨事項");
        headerRun.CharacterFormat.FontName = FontHeading;
        headerRun.CharacterFormat.FontSize = 12;
        headerRun.CharacterFormat.Bold = true;
        headerRun.CharacterFormat.TextColor = T.PrimaryDark;

        // 本文
        if (!string.IsNullOrEmpty(s.Content))
        {
            var bodyPara = cell.AddParagraph();
            bodyPara.ParagraphFormat.LineSpacing = 20;
            var bodyRun = bodyPara.AppendText(s.Content);
            bodyRun.CharacterFormat.FontName = FontBody;
            bodyRun.CharacterFormat.FontSize = 11;
            bodyRun.CharacterFormat.TextColor = T.TextPrimary;
        }

        var spacer = section.AddParagraph();
        spacer.ParagraphFormat.AfterSpacing = 12;
    }

    // =========================================================================
    // Text
    // =========================================================================

    private static void RenderText(IWSection section, ReportSection s)
    {
        if (string.IsNullOrEmpty(s.Content)) return;

        var para = section.AddParagraph();
        para.ParagraphFormat.AfterSpacing = 8;
        para.ParagraphFormat.LineSpacing = 20;
        para.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
        var run = para.AppendText(s.Content);
        run.CharacterFormat.FontName = FontBody;
        run.CharacterFormat.FontSize = 10.5f;
        run.CharacterFormat.TextColor = T.TextPrimary;
    }

    // =========================================================================
    // Bullet List — ゴールドビュレット
    // =========================================================================

    private static void RenderBulletList(IWSection section, ReportSection s)
    {
        if (!string.IsNullOrEmpty(s.Title))
        {
            var titlePara = section.AddParagraph();
            titlePara.ParagraphFormat.BeforeSpacing = 12;
            titlePara.ParagraphFormat.AfterSpacing = 4;
            var titleRun = titlePara.AppendText(s.Title);
            titleRun.CharacterFormat.FontName = FontHeading;
            titleRun.CharacterFormat.FontSize = 11;
            titleRun.CharacterFormat.Bold = true;
            titleRun.CharacterFormat.TextColor = T.PrimaryDark;
        }

        foreach (var item in s.Items)
        {
            var para = section.AddParagraph();
            para.ParagraphFormat.LeftIndent = 18;
            para.ParagraphFormat.FirstLineIndent = -12;
            para.ParagraphFormat.AfterSpacing = 4;
            para.ParagraphFormat.LineSpacing = 18;
            para.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;

            // ゴールドのビュレット
            var bulletRun = para.AppendText("■ ");
            bulletRun.CharacterFormat.FontName = FontBody;
            bulletRun.CharacterFormat.FontSize = 7;
            bulletRun.CharacterFormat.TextColor = T.Primary;

            var textRun = para.AppendText(item);
            textRun.CharacterFormat.FontName = FontBody;
            textRun.CharacterFormat.FontSize = 10.5f;
            textRun.CharacterFormat.TextColor = T.TextPrimary;
        }

        // リスト後のスペース
        var spacer = section.AddParagraph();
        spacer.ParagraphFormat.AfterSpacing = 6;
    }

    // =========================================================================
    // Table — ゴールドヘッダー + ストライプ行
    // =========================================================================

    private static void RenderTable(IWSection section, ReportSection s)
    {
        if (s.TableData == null || s.TableData.Headers.Count == 0) return;

        if (!string.IsNullOrEmpty(s.Title))
        {
            var tp = section.AddParagraph();
            tp.ParagraphFormat.BeforeSpacing = 14;
            tp.ParagraphFormat.AfterSpacing = 6;
            var tr = tp.AppendText(s.Title);
            tr.CharacterFormat.FontName = FontHeading;
            tr.CharacterFormat.FontSize = 11;
            tr.CharacterFormat.Bold = true;
            tr.CharacterFormat.TextColor = T.PrimaryDark;
        }

        int colCount = s.TableData.Headers.Count;
        int rowCount = s.TableData.Rows.Count + 1;

        var table = section.AddTable();
        table.ResetCells(rowCount, colCount);
        table.TableFormat.IsAutoResized = true;

        // ヘッダー行（ゴールド背景 + 白文字）
        for (int c = 0; c < colCount; c++)
        {
            var cell = table.Rows[0].Cells[c];
            cell.CellFormat.BackColor = T.TableHeaderBg;
            cell.CellFormat.Borders.Color = T.PrimaryDark;
            cell.CellFormat.Borders.LineWidth = 0.5f;
            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            var para = cell.AddParagraph();
            para.ParagraphFormat.HorizontalAlignment = HAlign.Center;
            var run = para.AppendText(s.TableData.Headers[c]);
            run.CharacterFormat.FontName = FontHeading;
            run.CharacterFormat.FontSize = 10;
            run.CharacterFormat.Bold = true;
            run.CharacterFormat.TextColor = T.TableHeaderText;
        }
        table.Rows[0].Height = 28;
        table.Rows[0].HeightType = TableRowHeightType.AtLeast;

        // データ行（ストライプ）
        for (int r = 0; r < s.TableData.Rows.Count; r++)
        {
            var rowData = s.TableData.Rows[r];
            bool isStripe = r % 2 == 1;

            for (int c = 0; c < Math.Min(rowData.Count, colCount); c++)
            {
                var cell = table.Rows[r + 1].Cells[c];
                cell.CellFormat.Borders.Color = T.BorderColor;
                cell.CellFormat.Borders.LineWidth = 0.5f;
                cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                if (isStripe)
                    cell.CellFormat.BackColor = T.TableStripeBg;

                var para = cell.AddParagraph();
                var run = para.AppendText(rowData[c]);
                run.CharacterFormat.FontName = FontBody;
                run.CharacterFormat.FontSize = 10;
                run.CharacterFormat.TextColor = T.TextPrimary;
            }

            table.Rows[r + 1].Height = 22;
            table.Rows[r + 1].HeightType = TableRowHeightType.AtLeast;
        }

        var spacer = section.AddParagraph();
        spacer.ParagraphFormat.AfterSpacing = 14;
    }

    // =========================================================================
    // Chart → Table
    // =========================================================================

    private static void RenderChartAsTable(IWSection section, ReportSection s)
    {
        if (s.ChartData == null) return;

        var ts = new ReportSection
        {
            Type = "table",
            Title = s.ChartData.Title,
            TableData = new ReportTableData
            {
                Headers = new List<string> { "項目" },
                Rows = new List<List<string>>(),
            }
        };

        foreach (var series in s.ChartData.Series)
            ts.TableData.Headers.Add(series.Name);

        for (int i = 0; i < s.ChartData.Categories.Count; i++)
        {
            var row = new List<string> { s.ChartData.Categories[i] };
            foreach (var series in s.ChartData.Series)
                row.Add(i < series.Values.Count ? series.Values[i].ToString("N0") : "");
            ts.TableData.Rows.Add(row);
        }

        RenderTable(section, ts);
    }

    // =========================================================================
    // Key Metrics — KPI ダッシュボードカード
    // =========================================================================

    private static void RenderKeyMetrics(IWSection section, ReportSection s)
    {
        if (s.Metrics.Count == 0) return;

        if (!string.IsNullOrEmpty(s.Title))
        {
            var tp = section.AddParagraph();
            tp.ParagraphFormat.BeforeSpacing = 14;
            tp.ParagraphFormat.AfterSpacing = 8;
            var tr = tp.AppendText(s.Title);
            tr.CharacterFormat.FontName = FontHeading;
            tr.CharacterFormat.FontSize = 12;
            tr.CharacterFormat.Bold = true;
            tr.CharacterFormat.TextColor = T.PrimaryDark;
        }

        // KPI テーブル（ラベル行 + 値行）
        var table = section.AddTable();
        table.ResetCells(2, s.Metrics.Count);
        table.TableFormat.IsAutoResized = true;

        for (int i = 0; i < s.Metrics.Count; i++)
        {
            var m = s.Metrics[i];

            // ラベル行（ゴールド背景）
            var labelCell = table.Rows[0].Cells[i];
            labelCell.CellFormat.BackColor = T.TableHeaderBg;
            labelCell.CellFormat.Borders.Color = T.PrimaryDark;
            labelCell.CellFormat.Borders.LineWidth = 0.5f;
            labelCell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            var labelPara = labelCell.AddParagraph();
            labelPara.ParagraphFormat.HorizontalAlignment = HAlign.Center;
            var labelRun = labelPara.AppendText(m.Label);
            labelRun.CharacterFormat.FontName = FontHeading;
            labelRun.CharacterFormat.FontSize = 9;
            labelRun.CharacterFormat.Bold = true;
            labelRun.CharacterFormat.TextColor = T.TableHeaderText;

            // 値行
            var valueCell = table.Rows[1].Cells[i];
            valueCell.CellFormat.Borders.Color = T.BorderColor;
            valueCell.CellFormat.Borders.LineWidth = 0.5f;
            valueCell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            var valuePara = valueCell.AddParagraph();
            valuePara.ParagraphFormat.HorizontalAlignment = HAlign.Center;
            valuePara.ParagraphFormat.BeforeSpacing = 6;
            var valueRun = valuePara.AppendText(m.Value);
            valueRun.CharacterFormat.FontName = FontHeading;
            valueRun.CharacterFormat.FontSize = 20;
            valueRun.CharacterFormat.Bold = true;
            valueRun.CharacterFormat.TextColor = T.Primary;

            if (!string.IsNullOrEmpty(m.Change))
            {
                var changePara = valueCell.AddParagraph();
                changePara.ParagraphFormat.HorizontalAlignment = HAlign.Center;
                changePara.ParagraphFormat.AfterSpacing = 4;
                var trendColor = m.Trend switch
                {
                    "positive" => SuccessGreen,
                    "negative" => ErrorRed,
                    _ => T.TextSecondary,
                };
                var arrow = m.Trend switch
                {
                    "positive" => "▲ ",
                    "negative" => "▼ ",
                    _ => "",
                };
                var changeRun = changePara.AppendText($"{arrow}{m.Change}");
                changeRun.CharacterFormat.FontName = FontBody;
                changeRun.CharacterFormat.FontSize = 9;
                changeRun.CharacterFormat.Bold = true;
                changeRun.CharacterFormat.TextColor = trendColor;
            }
        }

        table.Rows[0].Height = 26;
        table.Rows[0].HeightType = TableRowHeightType.AtLeast;
        table.Rows[1].Height = 52;
        table.Rows[1].HeightType = TableRowHeightType.AtLeast;

        var spacer = section.AddParagraph();
        spacer.ParagraphFormat.AfterSpacing = 14;
    }
}
