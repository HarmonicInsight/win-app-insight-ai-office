using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Border = System.Windows.Controls.Border;

namespace InsightAiOffice.App.Views;

public partial class TutorialDialog : Window
{
    /// <summary>カードクリック時に発火。(サンプルファイルパス一覧, プロンプトテキスト, 出力形式)</summary>
    public event Action<List<string>, string, string>? TutorialExecuteRequested;

    private readonly string _assetsDir;
    private readonly string _workDir;

    public TutorialDialog()
    {
        InitializeComponent();
        _assetsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "tutorials");
        _workDir = Path.Combine(Path.GetTempPath(), "IAOF_Tutorial");
        Directory.CreateDirectory(_workDir);
        BuildCards();
    }

    private void BuildCards()
    {
        var cards = new[]
        {
            new TutorialCard
            {
                Icon = "📄",
                Title = "契約書 → 入出金 Excel",
                Desc = "下請負契約書を分析し\n支払い予定のExcel管理表を自動生成",
                Tag = "全般",
                Color = "#B8942F",
                FolderName = "01_契約書から入出金",
            },
            new TutorialCard
            {
                Icon = "📋",
                Title = "会議メモ → 議事録",
                Desc = "走り書きの会議メモから\n正式な議事録とアクションアイテム表を生成",
                Tag = "総務部向け",
                Color = "#2563EB",
                FolderName = "02_会議メモから議事録",
            },
            new TutorialCard
            {
                Icon = "📊",
                Title = "売上データ → 分析レポート",
                Desc = "月次売上CSVから\n分析レポート(Word) + 集計Excel を作成",
                Tag = "経理部向け",
                Color = "#16A34A",
                FolderName = "03_売上分析レポート",
            },
            new TutorialCard
            {
                Icon = "📽️",
                Title = "企画書 → プレゼン 10枚",
                Desc = "テキストの企画書から\n経営会議向けプレゼンを自動生成",
                Tag = "経営企画",
                Color = "#8B2252",
                FolderName = "04_企画書からプレゼン",
            },
            new TutorialCard
            {
                Icon = "💰",
                Title = "請求書チェック → 支払管理",
                Desc = "請求書一覧CSVから\n支払管理表と月次レポートを作成",
                Tag = "経理部向け",
                Color = "#1E293B",
                FolderName = "07_経理_請求書チェック",
            },
            new TutorialCard
            {
                Icon = "👤",
                Title = "応募者情報 → 選考比較表",
                Desc = "応募者CSVから\n強み・推奨度付きの比較表を作成",
                Tag = "人事部向け",
                Color = "#CA8A04",
                FolderName = "08_人事_採用選考表",
            },
        };

        foreach (var card in cards)
            CardPanel.Children.Add(CreateCardUI(card));

        // サンプル出力を見るボタン
        CardPanel.Children.Add(CreateSampleOutputCard());
    }

    private Border CreateSampleOutputCard()
    {
        var border = new Border
        {
            Width = 250,
            Margin = new Thickness(0, 0, 10, 10),
            Padding = new Thickness(14, 12, 14, 12),
            CornerRadius = new CornerRadius(8),
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#16A34A")),
            BorderThickness = new Thickness(2),
            Cursor = Cursors.Hand,
        };

        var stack = new StackPanel();
        var tagBorder = new Border
        {
            Background = new SolidColorBrush(Color.FromArgb(30, 22, 163, 74)),
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(6, 2, 6, 2),
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 0, 0, 6),
        };
        tagBorder.Child = new TextBlock { Text = "APIキー不要", FontSize = 9, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#16A34A")) };
        stack.Children.Add(tagBorder);
        stack.Children.Add(new TextBlock { Text = "📂 サンプル出力を見る", FontSize = 12, FontWeight = FontWeights.SemiBold, Foreground = (Brush)FindResource("TextPrimaryBrush"), Margin = new Thickness(0, 0, 0, 4) });
        stack.Children.Add(new TextBlock { Text = "AIが生成するExcel・Wordの\n品質をすぐに確認できます\n(Gold / Blue / Navy)", FontSize = 10, Foreground = (Brush)FindResource("TextSecondaryBrush"), TextWrapping = TextWrapping.Wrap, LineHeight = 16, Margin = new Thickness(0, 0, 0, 8) });
        stack.Children.Add(new TextBlock { Text = "▶ フォルダを開く", FontSize = 10, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#16A34A")) });
        border.Child = stack;

        border.MouseEnter += (_, _) => border.Background = new SolidColorBrush(Color.FromArgb(13, 22, 163, 74));
        border.MouseLeave += (_, _) => border.Background = Brushes.White;

        border.MouseLeftButtonUp += (_, _) =>
        {
            var sampleDir = Path.Combine(_assetsDir, "sample-outputs");
            if (Directory.Exists(sampleDir))
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sampleDir) { UseShellExecute = true });
        };

        return border;
    }

    /// <summary>tutorialsフォルダからデータファイルとプロンプトを読み込む</summary>
    private (List<string> files, string prompt) LoadTutorialData(string folderName)
    {
        var srcDir = Path.Combine(_assetsDir, folderName);
        var files = new List<string>();
        var prompt = "";

        if (!Directory.Exists(srcDir))
            return (files, "このチュートリアルのデータが見つかりません。");

        // プロンプト読み込み
        var promptPath = Path.Combine(srcDir, "プロンプト例.txt");
        if (File.Exists(promptPath))
            prompt = File.ReadAllText(promptPath).Trim();

        // データファイルをワークディレクトリにコピー
        foreach (var srcFile in Directory.GetFiles(srcDir))
        {
            var fileName = Path.GetFileName(srcFile);
            if (fileName == "プロンプト例.txt") continue; // プロンプトは除外

            var destFile = Path.Combine(_workDir, fileName);
            try
            {
                File.Copy(srcFile, destFile, overwrite: true);
                files.Add(destFile);
            }
            catch { /* ファイルロック等のエラーは無視 */ }
        }

        return (files, prompt);
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
            HorizontalAlignment = HorizontalAlignment.Left,
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
            var (files, prompt) = LoadTutorialData(card.FolderName);
            TutorialExecuteRequested?.Invoke(files, prompt, "auto");
            runBtn.Text = "✔ セット完了";
        };

        return border;
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private sealed class TutorialCard
    {
        public string Icon { get; init; } = "";
        public string Title { get; init; } = "";
        public string Desc { get; init; } = "";
        public string Tag { get; init; } = "";
        public string Color { get; init; } = "#B8942F";
        public string FolderName { get; init; } = "";
    }
}
