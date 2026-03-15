using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ClosedXML.Excel;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Border = System.Windows.Controls.Border;

namespace InsightAiOffice.App.Views;

public partial class TutorialDialog : Window
{
    /// <summary>カードクリック時に発火。(サンプルファイルパス一覧, プロンプトテキスト, 出力形式)</summary>
    public event Action<List<string>, string, string>? TutorialExecuteRequested;

    private readonly string _tutorialDir;

    public TutorialDialog()
    {
        InitializeComponent();
        _tutorialDir = Path.Combine(Path.GetTempPath(), "IAOF_Tutorial");
        Directory.CreateDirectory(_tutorialDir);
        BuildCards();
    }

    private void BuildCards()
    {
        var cards = new[]
        {
            new TutorialCard
            {
                Icon = "📦",
                Title = "提案フルセット（3ファイル同時生成）",
                Desc = "会議メモ＋売上データから\n提案書(.docx) + 見積書(.xlsx) + プレゼン(.pptx) を一括生成",
                Tag = "営業部向け",
                Color = "#B8942F",
                SampleGenerator = GenerateSalesData,
                Prompt = @"添付の会議メモと売上実績データを参考に、新製品の提案資料一式を作成してください。

1. generate_report ツールで提案書（Word）を生成:
   - 表紙、エグゼクティブサマリー、市場分析、製品コンセプト、導入効果・ROI、スケジュール、投資計画

2. generate_spreadsheet ツールで収支計画書（Excel）を生成:
   - 初年度〜3年目の月別売上・原価・粗利・営業利益の推移表

3. generate_presentation ツールで経営会議用プレゼン（PowerPoint）を生成:
   - 8枚構成: 表紙/市場機会/製品概要/競合比較/収支計画/ロードマップ/体制/まとめ",
                OutputFormat = "auto",
            },
            new TutorialCard
            {
                Icon = "📋",
                Title = "議事録 → 報告書",
                Desc = "会議メモから正式な議事録と\nアクションアイテム管理表を自動作成",
                Tag = "総務部向け",
                Color = "#2563EB",
                SampleGenerator = GenerateMeetingMemo,
                Prompt = @"添付の会議メモをもとに以下の2つを作成してください。

1. generate_report ツールで正式議事録（Word）を生成:
   - 会議情報（日時・場所・出席者）
   - 各議題の議論内容と決定事項
   - 質疑応答
   - 次回予定

2. generate_spreadsheet ツールでアクションアイテム管理表（Excel）を生成:
   - 列: No / アクション内容 / 担当者 / 期限 / 優先度(高/中/低) / ステータス / 備考",
                OutputFormat = "auto",
            },
            new TutorialCard
            {
                Icon = "💹",
                Title = "売上データ → 月次レポート",
                Desc = "Excel売上データから\n月次決算報告書と分析グラフを自動生成",
                Tag = "経理部向け",
                Color = "#16A34A",
                SampleGenerator = GenerateSalesOnly,
                Prompt = @"添付の売上実績データを分析し、以下を作成してください。

1. generate_report ツールで月次決算報告書（Word）を生成:
   - 当期業績サマリー
   - 月別売上推移と前年比分析
   - 粗利率の推移と要因分析
   - 顧客数推移と新規獲得動向
   - 来期の見通しと課題

theme: blue を指定してください。",
                OutputFormat = "word",
            },
            new TutorialCard
            {
                Icon = "👤",
                Title = "入社手続きセット",
                Desc = "入社チェックリスト（Excel）と\n歓迎案内文（Word）を一括生成",
                Tag = "人事部向け",
                Color = "#8B2252",
                SampleGenerator = null,
                Prompt = @"4月1日入社の新入社員（営業部配属、大卒）の入社手続きセットを作成してください。

1. generate_spreadsheet ツールで入社手続きチェックリスト（Excel）を生成:
   列: No / カテゴリ / タスク名 / 担当 / 期限目安 / 完了 / 備考
   - 書類系（雇用契約書、秘密保持、住民票、口座届出等）
   - 設備系（PC、メール、入館証、名刺等）
   - 研修系（コンプライアンス、安全衛生、OJT等）

2. generate_report ツールで歓迎案内文（Word）を生成:
   - ウェルカムメッセージ、初日スケジュール、オフィスガイド、社内ルール、連絡先一覧",
                OutputFormat = "auto",
            },
            new TutorialCard
            {
                Icon = "🏗️",
                Title = "施工計画書＋工程表",
                Desc = "工事概要から施工計画書（Word）と\n工程表（Excel）を自動生成",
                Tag = "工事部向け",
                Color = "#1E293B",
                SampleGenerator = null,
                Prompt = @"以下の工事概要で施工計画書と工程表を作成してください。

工事名: ○○マンション新築工事（RC造 地上10階 延床面積3,500㎡）
工期: 2026年4月1日〜2027年3月31日
発注者: △△不動産株式会社

1. generate_report ツールで施工計画書（Word）を生成:
   工事概要、施工体制、工種別施工方法、品質管理計画、安全衛生管理、環境対策、仮設計画

2. generate_spreadsheet ツールで工程表（Excel）を生成:
   列: 工種 / 細別 / 数量 / 4月〜3月の12ヶ月
   主要工種15行＋マイルストーン（着工/中間検査/竣工）",
                OutputFormat = "auto",
            },
            new TutorialCard
            {
                Icon = "🔄",
                Title = "BPO業務マニュアル",
                Desc = "業務マニュアル（Word）と\n日次チェックリスト（Excel）を生成",
                Tag = "BPO事業者向け",
                Color = "#CA8A04",
                SampleGenerator = null,
                Prompt = @"データ入力業務のBPOオペレーションマニュアルとチェックリストを作成してください。

1. generate_report ツールで業務マニュアル（Word）を生成:
   - 業務概要（受託範囲、SLA: 処理件数500件/日、エラー率0.1%以下）
   - 業務フロー（受付→データ入力→ダブルチェック→納品）
   - 各工程の作業手順
   - エスカレーション基準
   - 品質管理体制

2. generate_spreadsheet ツールで日次チェックリスト（Excel）を生成:
   時間帯 / 作業項目 / 担当者 / チェック / 処理件数 / エラー件数 / 備考",
                OutputFormat = "auto",
            },
        };

        foreach (var card in cards)
            CardPanel.Children.Add(CreateCardUI(card));
    }

    private Border CreateCardUI(TutorialCard card)
    {
        var border = new Border
        {
            Width = 250,
            Margin = new Thickness(0, 0, 10, 10),
            Padding = new Thickness(14, 12, 14, 12),
            CornerRadius = new CornerRadius(8),
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(card.Color)),
            BorderThickness = new Thickness(1),
            Cursor = Cursors.Hand,
        };

        var stack = new StackPanel();

        // Tag badge
        var tagBorder = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(card.Color)) { Opacity = 0.12 },
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(6, 2, 6, 2),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
            Margin = new Thickness(0, 0, 0, 6),
        };
        tagBorder.Child = new TextBlock
        {
            Text = card.Tag,
            FontSize = 9,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(card.Color)),
        };
        stack.Children.Add(tagBorder);

        // Icon + Title
        stack.Children.Add(new TextBlock
        {
            Text = $"{card.Icon} {card.Title}",
            FontSize = 12,
            FontWeight = FontWeights.SemiBold,
            Foreground = (Brush)FindResource("TextPrimaryBrush"),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 4),
        });

        // Description
        stack.Children.Add(new TextBlock
        {
            Text = card.Desc,
            FontSize = 10,
            Foreground = (Brush)FindResource("TextSecondaryBrush"),
            TextWrapping = TextWrapping.Wrap,
            LineHeight = 16,
            Margin = new Thickness(0, 0, 0, 8),
        });

        // Run button
        var runBtn = new TextBlock
        {
            Text = "▶ クリックしてセット",
            FontSize = 10,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(card.Color)),
        };
        stack.Children.Add(runBtn);

        border.Child = stack;

        // Hover effect
        border.MouseEnter += (_, _) => border.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(card.Color)) { Opacity = 0.05 };
        border.MouseLeave += (_, _) => border.Background = Brushes.White;

        // Click
        border.MouseLeftButtonUp += (_, _) =>
        {
            var files = new List<string>();
            if (card.SampleGenerator != null)
                files = card.SampleGenerator(_tutorialDir);
            TutorialExecuteRequested?.Invoke(files, card.Prompt, card.OutputFormat);
            runBtn.Text = "✔ セット完了";
        };

        return border;
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    // ── サンプルデータ生成 ──

    private static List<string> GenerateSalesData(string dir)
    {
        var files = new List<string>();
        files.Add(GenerateMeetingMemoFile(dir));
        files.Add(GenerateSalesFile(dir));
        return files;
    }

    private static List<string> GenerateMeetingMemo(string dir)
    {
        return new List<string> { GenerateMeetingMemoFile(dir) };
    }

    private static List<string> GenerateSalesOnly(string dir)
    {
        return new List<string> { GenerateSalesFile(dir) };
    }

    private static string GenerateMeetingMemoFile(string dir)
    {
        var path = Path.Combine(dir, "会議メモ_新製品企画.docx");
        if (File.Exists(path)) return path;

        using var doc = new WordDocument();
        var sec = doc.AddSection() as WSection;
        sec!.PageSetup.Margins.All = 72;

        void H1(string t) { var p = sec.AddParagraph() as WParagraph; var r = p!.AppendText(t); r.CharacterFormat.FontSize = 18; r.CharacterFormat.Bold = true; r.CharacterFormat.FontName = "Yu Gothic"; p.ParagraphFormat.AfterSpacing = 12; }
        void H2(string t) { var p = sec.AddParagraph() as WParagraph; var r = p!.AppendText(t); r.CharacterFormat.FontSize = 13; r.CharacterFormat.Bold = true; r.CharacterFormat.FontName = "Yu Gothic"; p.ParagraphFormat.AfterSpacing = 6; }
        void P(string t) { var p = sec.AddParagraph() as WParagraph; if (!string.IsNullOrEmpty(t)) { var r = p!.AppendText(t); r.CharacterFormat.FontSize = 11; r.CharacterFormat.FontName = "Yu Mincho"; } }

        H1("新製品企画会議 議事メモ");
        P("日時: 2026年3月10日 14:00-15:30");
        P("場所: 本社 3F 会議室A");
        P("出席者: 田中部長、佐藤課長、鈴木（企画）、高橋（開発）、山田（営業）");
        P("");
        H2("1. 市場調査結果の報告（佐藤）");
        P("・国内市場規模: 約850億円（前年比12%成長）");
        P("・競合3社の動向: A社はクラウド版を今秋リリース予定、B社は価格改定（15%値下げ）、C社は大企業向けに特化");
        P("・ターゲット顧客の声: 「既存ツールは操作が複雑」「導入に3ヶ月以上かかる」「カスタマイズ費用が高い」");
        P("");
        H2("2. 新製品コンセプト案（鈴木）");
        P("・コンセプト: 「誰でも30分で使い始められる業務効率化ツール」");
        P("・差別化: AI搭載による自動設定、テンプレート豊富、初期費用ゼロ");
        P("・想定価格: 月額4,980円/ユーザー（年間契約で20%割引）");
        P("・ターゲット: 従業員50-300名の中堅企業");
        P("");
        H2("3. 開発スケジュール案（高橋）");
        P("・Phase1（MVP）: 2026年6月末 — 基本機能");
        P("・Phase2（正式版）: 2026年9月末 — AI機能、API連携");
        P("・Phase3（拡張）: 2026年12月末 — 多言語、エンタープライズ");
        P("");
        H2("4. 営業計画（山田）");
        P("・初年度売上目標: 1億2,000万円（200社 × 月額5万円 × 12ヶ月）");
        P("・営業戦略: 展示会3回、ウェビナー月2回、代理店5社");
        P("");
        H2("5. 決定事項・次回アクション");
        P("【決定】コンセプト案を経営会議に上程（4月第1週）");
        P("【田中部長】投資対効果の試算を財務部と調整（3/20まで）");
        P("【鈴木】競合比較表を詳細化（3/17まで）");
        P("【高橋】技術検証（AI基盤の選定）（3/24まで）");
        P("【山田】ターゲット企業リスト50社を作成（3/20まで）");
        P("次回会議: 3月25日 14:00");

        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
        doc.Save(stream, FormatType.Docx);
        return path;
    }

    private static string GenerateSalesFile(string dir)
    {
        var path = Path.Combine(dir, "売上実績_2025年度.xlsx");
        if (File.Exists(path)) return path;

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("月別売上");
        string[] headers = { "月", "売上高（万円）", "原価（万円）", "粗利（万円）", "粗利率(%)", "顧客数", "新規獲得" };
        for (int c = 0; c < headers.Length; c++) { ws.Cell(1, c + 1).Value = headers[c]; ws.Cell(1, c + 1).Style.Font.Bold = true; }

        var months = new[] { "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月" };
        var sales = new[] { 4200, 3800, 5100, 4600, 3500, 4900, 5800, 5200, 6100, 4800, 5500, 7200 };
        var costs = new[] { 2520, 2280, 2950, 2760, 2100, 2840, 3190, 2860, 3350, 2640, 3025, 3960 };
        for (int i = 0; i < 12; i++)
        {
            int r = i + 2;
            ws.Cell(r, 1).Value = months[i];
            ws.Cell(r, 2).Value = sales[i];
            ws.Cell(r, 3).Value = costs[i];
            ws.Cell(r, 4).Value = sales[i] - costs[i];
            ws.Cell(r, 5).Value = Math.Round((double)(sales[i] - costs[i]) / sales[i] * 100, 1);
            ws.Cell(r, 6).Value = 45 + i * 3;
            ws.Cell(r, 7).Value = 3 + (i % 4);
        }
        ws.Range(1, 1, 1, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#F5F0E8");
        ws.Columns().AdjustToContents();
        wb.SaveAs(path);
        return path;
    }

    private sealed class TutorialCard
    {
        public string Icon { get; init; } = "";
        public string Title { get; init; } = "";
        public string Desc { get; init; } = "";
        public string Tag { get; init; } = "";
        public string Color { get; init; } = "#B8942F";
        public Func<string, List<string>>? SampleGenerator { get; init; }
        public string Prompt { get; init; } = "";
        public string OutputFormat { get; init; } = "auto";
    }
}
