using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace InsightAiOffice.App.Views;

public partial class TutorialDialog : Window
{
    /// <summary>選択されたチュートリアルのプロンプト</summary>
    public string? SelectedPrompt { get; private set; }

    /// <summary>選択されたチュートリアルの添付ファイルパス一覧</summary>
    public List<string> SelectedFiles { get; private set; } = new();

    public TutorialDialog()
    {
        InitializeComponent();
        TutorialList.ItemsSource = BuildTutorials();
    }

    private void TutorialCard_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement el || el.Tag is not string id) return;
        var tutorials = TutorialList.ItemsSource as List<TutorialItem>;
        var tutorial = tutorials?.FirstOrDefault(t => t.Id == id);
        if (tutorial == null) return;

        SelectedPrompt = tutorial.Prompt;
        SelectedFiles = tutorial.Files.Where(File.Exists).ToList();
        DialogResult = true;
    }

    private static List<TutorialItem> BuildTutorials()
    {
        var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "tutorials");

        return new List<TutorialItem>
        {
            new()
            {
                Id = "01", Icon = "📄", DepartmentTag = "全般",
                Title = "契約書 → 入出金スケジュール Excel",
                Description = "下請負契約書を分析し、支払い予定のExcel管理表を自動生成します。",
                Prompt = LoadPrompt(baseDir, "01_契約書から入出金"),
                Files = FindDataFiles(baseDir, "01_契約書から入出金"),
            },
            new()
            {
                Id = "02", Icon = "📋", DepartmentTag = "全般",
                Title = "会議メモ → 正式な議事録",
                Description = "走り書きの会議メモから、決定事項・アクションアイテム付きの議事録を生成します。",
                Prompt = LoadPrompt(baseDir, "02_会議メモから議事録"),
                Files = FindDataFiles(baseDir, "02_会議メモから議事録"),
            },
            new()
            {
                Id = "03", Icon = "📊", DepartmentTag = "経営企画",
                Title = "売上データ → 分析レポート",
                Description = "月次売上CSVから部門別・前年比の分析レポートと集計Excelを作成します。",
                Prompt = LoadPrompt(baseDir, "03_売上分析レポート"),
                Files = FindDataFiles(baseDir, "03_売上分析レポート"),
            },
            new()
            {
                Id = "04", Icon = "📽️", DepartmentTag = "経営企画",
                Title = "企画書 → プレゼン資料 10枚",
                Description = "テキストの企画書から、経営会議向けの10枚プレゼンを自動生成します。",
                Prompt = LoadPrompt(baseDir, "04_企画書からプレゼン"),
                Files = FindDataFiles(baseDir, "04_企画書からプレゼン"),
            },
            new()
            {
                Id = "05", Icon = "📝", DepartmentTag = "営業",
                Title = "見積書テンプレートの書き換え",
                Description = "テンプレートの宛先・金額・担当者を指定して正式な見積書を作成します。",
                Prompt = LoadPrompt(baseDir, "05_書類書き換え"),
                Files = FindDataFiles(baseDir, "05_書類書き換え"),
            },
            new()
            {
                Id = "07", Icon = "💰", DepartmentTag = "経理",
                Title = "請求書チェック → 支払管理表",
                Description = "請求書一覧CSVから支払期日順の管理表と月次支払レポートを作成します。",
                Prompt = LoadPrompt(baseDir, "07_経理_請求書チェック"),
                Files = FindDataFiles(baseDir, "07_経理_請求書チェック"),
            },
            new()
            {
                Id = "08", Icon = "👤", DepartmentTag = "人事",
                Title = "応募者情報 → 採用選考比較表",
                Description = "応募者CSVから強み・懸念点・推奨度付きの比較表を作成します。",
                Prompt = LoadPrompt(baseDir, "08_人事_採用選考表"),
                Files = FindDataFiles(baseDir, "08_人事_採用選考表"),
            },
            new()
            {
                Id = "09", Icon = "🏢", DepartmentTag = "総務",
                Title = "社内通知文の作成",
                Description = "オフィス移転のお知らせ等、正式な社内通知文をWord形式で作成します。",
                Prompt = LoadPrompt(baseDir, "09_総務_社内通知文"),
                Files = new(),
            },
            new()
            {
                Id = "10", Icon = "💼", DepartmentTag = "営業",
                Title = "顧客向け提案書プレゼン",
                Description = "顧客情報を指定して、10枚の提案プレゼン資料を自動生成します。",
                Prompt = LoadPrompt(baseDir, "10_営業_提案書作成"),
                Files = new(),
            },
        };
    }

    private static string LoadPrompt(string baseDir, string folder)
    {
        var path = Path.Combine(baseDir, folder, "プロンプト例.txt");
        return File.Exists(path) ? File.ReadAllText(path).Trim() : "";
    }

    private static List<string> FindDataFiles(string baseDir, string folder)
    {
        var dir = Path.Combine(baseDir, folder);
        if (!Directory.Exists(dir)) return new();
        return Directory.GetFiles(dir)
            .Where(f => !Path.GetFileName(f).StartsWith("プロンプト") && Path.GetExtension(f) != ".md")
            .ToList();
    }
}

public class TutorialItem
{
    public string Id { get; set; } = "";
    public string Icon { get; set; } = "📄";
    public string Title { get; set; } = "";
    public string Description { get; set; } = "";
    public string DepartmentTag { get; set; } = "";
    public string Prompt { get; set; } = "";
    public List<string> Files { get; set; } = new();
}
