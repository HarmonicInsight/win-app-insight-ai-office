using System;
using System.Collections.Generic;
using System.IO;
using InsightAiOffice.App.Services.DocumentGeneration;

namespace InsightAiOffice.App.Tools;

/// <summary>
/// チュートリアル用サンプル出力ファイル（Word / Excel）を再生成するツール。
/// Syncfusion ライセンス登録済みの状態で実行することでトライアルスタンプを除去。
///
/// 使い方: dotnet run --project src/InsightAiOffice.App -- --generate-samples
/// </summary>
public static class SampleOutputGenerator
{
    public static void Generate()
    {
        var outputDir = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "assets", "tutorials", "sample-outputs");
        Directory.CreateDirectory(outputDir);

        var themes = new[] { "gold", "blue", "navy" };

        foreach (var theme in themes)
        {
            var wordPath = Path.Combine(outputDir, $"word_{theme}.docx");
            var excelPath = Path.Combine(outputDir, $"excel_{theme}.xlsx");

            ReportRendererService.Render(CreateSampleReport(), wordPath, theme);
            SpreadsheetRendererService.Render(CreateSampleSpreadsheet(), excelPath, theme);

            Console.WriteLine($"  [OK] {wordPath}");
            Console.WriteLine($"  [OK] {excelPath}");
        }

        Console.WriteLine();
        Console.WriteLine($"出力先: {outputDir}");
    }

    /// <summary>アプリの assets/tutorials/sample-outputs にコピーする</summary>
    public static void GenerateToSource(string sourceAssetsDir)
    {
        var outputDir = Path.Combine(sourceAssetsDir, "tutorials", "sample-outputs");
        Directory.CreateDirectory(outputDir);

        var themes = new[] { "gold", "blue", "navy" };

        foreach (var theme in themes)
        {
            var wordPath = Path.Combine(outputDir, $"word_{theme}.docx");
            var excelPath = Path.Combine(outputDir, $"excel_{theme}.xlsx");

            ReportRendererService.Render(CreateSampleReport(), wordPath, theme);
            SpreadsheetRendererService.Render(CreateSampleSpreadsheet(), excelPath, theme);

            Console.WriteLine($"  [OK] {Path.GetFileName(wordPath)}");
            Console.WriteLine($"  [OK] {Path.GetFileName(excelPath)}");
        }

        Console.WriteLine($"\n出力先: {outputDir}");
    }

    private static ReportStructure CreateSampleReport() => new()
    {
        Title = "月次売上分析レポート",
        Author = "Insight AI Office",
        Date = "2026年3月",
        Sections = new List<ReportSection>
        {
            new()
            {
                Type = "title",
                Title = "月次売上分析レポート",
                Content = "2026年3月度 — 営業部門パフォーマンス分析"
            },
            new()
            {
                Type = "key_metrics",
                Title = "主要KPI",
                Metrics = new List<ReportMetric>
                {
                    new() { Label = "月間売上高", Value = "¥12.8億", Change = "+8.5%", Trend = "positive" },
                    new() { Label = "受注件数", Value = "156件", Change = "+12件", Trend = "positive" },
                    new() { Label = "平均単価", Value = "¥820万", Change = "-2.1%", Trend = "negative" },
                    new() { Label = "粗利率", Value = "34.2%", Change = "+1.3pt", Trend = "positive" },
                }
            },
            new()
            {
                Type = "heading",
                Title = "1. エグゼクティブサマリー",
                Level = 1,
            },
            new()
            {
                Type = "summary",
                Title = "概要",
                Content = "2026年3月度の売上高は前年同月比 +8.5% の12.8億円となり、四半期目標の達成率は102%に到達しました。特に新規顧客からの受注が好調で、DX関連案件が全体の42%を占めています。"
            },
            new()
            {
                Type = "heading",
                Title = "2. 部門別売上実績",
                Level = 1,
            },
            new()
            {
                Type = "table",
                Title = "部門別売上比較",
                TableData = new ReportTableData
                {
                    Headers = new List<string> { "部門", "今月実績", "前月実績", "前年同月", "前年比" },
                    Rows = new List<List<string>>
                    {
                        new() { "コンサルティング事業部", "¥5.2億", "¥4.8億", "¥4.6億", "+13.0%" },
                        new() { "ソリューション事業部", "¥4.1億", "¥3.9億", "¥4.0億", "+2.5%" },
                        new() { "クラウド事業部", "¥2.3億", "¥2.1億", "¥1.8億", "+27.8%" },
                        new() { "保守サポート事業部", "¥1.2億", "¥1.2億", "¥1.4億", "-14.3%" },
                    }
                }
            },
            new()
            {
                Type = "heading",
                Title = "3. 重点施策の進捗",
                Level = 1,
            },
            new()
            {
                Type = "bullet_list",
                Title = "進行中のプロジェクト",
                Items = new List<string>
                {
                    "大手製造業A社 DX推進プロジェクト — フェーズ2完了、ROI 180% 達成",
                    "金融機関B社 AI導入支援 — PoC 完了、本番導入フェーズへ移行決定",
                    "小売業C社 基幹システム刷新 — 要件定義完了、4月より開発フェーズ開始",
                    "自治体D市 デジタル化推進 — 住民サービス向上プラン策定中"
                }
            },
            new()
            {
                Type = "heading",
                Title = "4. 課題と対策",
                Level = 1,
            },
            new()
            {
                Type = "recommendation",
                Title = "来月のアクションプラン",
                Content = "保守サポート事業部の売上減少傾向に対し、既存顧客へのアップセル施策を強化します。具体的には、AI活用支援パッケージの提案を重点的に行い、月次レビュー会議での進捗確認を実施します。"
            },
        }
    };

    private static SpreadsheetStructure CreateSampleSpreadsheet() => new()
    {
        Title = "月次売上管理表",
        Sheets = new List<SimpleSheetData>
        {
            new()
            {
                Name = "月次売上サマリー",
                Headers = new List<string> { "月", "売上高（万円）", "原価（万円）", "粗利（万円）", "粗利率", "受注件数" },
                Rows = new List<List<string>>
                {
                    new() { "2026年1月", "108500", "72300", "36200", "33.4", "142" },
                    new() { "2026年2月", "115200", "76800", "38400", "33.3", "148" },
                    new() { "2026年3月", "128000", "84200", "43800", "34.2", "156" },
                }
            },
            new()
            {
                Name = "部門別実績",
                Headers = new List<string> { "部門名", "売上高（万円）", "目標（万円）", "達成率", "前年比" },
                Rows = new List<List<string>>
                {
                    new() { "コンサルティング事業部", "52000", "50000", "104.0", "113.0" },
                    new() { "ソリューション事業部", "41000", "42000", "97.6", "102.5" },
                    new() { "クラウド事業部", "23000", "20000", "115.0", "127.8" },
                    new() { "保守サポート事業部", "12000", "15000", "80.0", "85.7" },
                }
            },
            new()
            {
                Name = "案件一覧",
                Headers = new List<string> { "案件名", "顧客名", "金額（万円）", "ステータス", "担当部門" },
                Rows = new List<List<string>>
                {
                    new() { "DX推進プロジェクト Phase2", "A製造株式会社", "18500", "進行中", "コンサルティング事業部" },
                    new() { "AI導入支援 本番移行", "B金融株式会社", "12000", "受注確定", "ソリューション事業部" },
                    new() { "基幹システム刷新", "C小売株式会社", "24000", "開発準備中", "ソリューション事業部" },
                    new() { "クラウド移行支援", "D通信株式会社", "8500", "進行中", "クラウド事業部" },
                    new() { "デジタル化推進計画", "E市役所", "6200", "要件定義中", "コンサルティング事業部" },
                    new() { "セキュリティ強化", "F保険株式会社", "4800", "提案中", "クラウド事業部" },
                }
            }
        }
    };
}
