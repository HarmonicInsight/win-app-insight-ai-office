using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using InsightAiOffice.App.Helpers;

namespace InsightAiOffice.App.Views;

public partial class HelpWindow : Window
{
    private readonly string[] _sectionIds;
    private readonly string[] _sectionNames;

    public HelpWindow(string? initialSection = null)
    {
        InitializeComponent();

        var isEn = LanguageManager.CurrentLanguage == "en";
        Title = isEn ? "Insight AI Office - Help" : "Insight AI Office - ヘルプ";

        _sectionIds =
        [
            "overview", "ui-layout", "file-ops", "multi-tab",
            "word-edit", "excel-edit", "pptx-view", "pdf-view",
            "ai-assistant", "prompt-presets",
            "shortcuts", "license", "system-req", "support"
        ];

        _sectionNames = isEn
            ?
            [
                "Overview", "UI Layout", "File Operations", "Multi-Tab",
                "Word Editing", "Excel Editing", "PowerPoint Viewer", "PDF Viewer",
                "AI Concierge", "Prompt Presets",
                "Keyboard Shortcuts", "License", "System Requirements", "Support"
            ]
            :
            [
                "はじめに", "画面構成", "ファイル操作", "マルチタブ",
                "Word 編集", "Excel 編集", "PowerPoint ビューア", "PDF ビューア",
                "AI コンシェルジュ", "プロンプトプリセット",
                "キーボードショートカット", "ライセンス", "システム要件", "お問い合わせ"
            ];

        for (int i = 0; i < _sectionNames.Length; i++)
        {
            NavList.Items.Add(new ListBoxItem
            {
                Content = _sectionNames[i],
                Tag = _sectionIds[i],
                Padding = new Thickness(12, 6, 12, 6),
                FontSize = 13,
            });
        }

        var html = GenerateHtml(isEn);
        HelpBrowser.NavigateToString(html);

        HelpBrowser.LoadCompleted += (_, _) =>
        {
            if (initialSection != null)
            {
                for (int i = 0; i < _sectionIds.Length; i++)
                {
                    if (_sectionIds[i] == initialSection)
                    {
                        NavList.SelectedIndex = i;
                        break;
                    }
                }
            }
            else
            {
                NavList.SelectedIndex = 0;
            }
        };
    }

    public static void ShowSection(Window owner, string sectionId)
    {
        var w = new HelpWindow(sectionId) { Owner = owner };
        w.ShowDialog();
    }

    private void OnNavSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (NavList.SelectedItem is not ListBoxItem item || item.Tag is not string sectionId) return;
        try
        {
            HelpBrowser.InvokeScript("scrollTo", new object[] { sectionId });
        }
        catch { /* WebBrowser not ready */ }
    }

    private void HelpBrowser_Navigating(object sender, System.Windows.Navigation.NavigatingCancelEventArgs e)
    {
        if (e.Uri == null) return;
        var scheme = e.Uri.Scheme;
        if (scheme is "http" or "https" or "mailto")
        {
            e.Cancel = true;
            try { Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true }); }
            catch { /* ignore */ }
        }
    }

    private string GenerateHtml(bool isEn)
    {
        var sb = new StringBuilder(64000);
        sb.Append(HtmlHead());
        sb.Append("<body>");

        if (isEn)
            AppendEnglishContent(sb);
        else
            AppendJapaneseContent(sb);

        sb.Append("</body></html>");
        return sb.ToString();
    }

    private static string HtmlHead() => """
        <!DOCTYPE html>
        <html><head><meta charset="utf-8"/>
        <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', 'Yu Gothic UI', sans-serif; background: #FAF8F5; color: #1C1917; padding: 32px 40px; line-height: 1.7; }
        h1 { color: #B8942F; font-size: 28px; font-weight: 300; margin-bottom: 8px; padding-top: 24px; }
        h2 { color: #B8942F; font-size: 20px; font-weight: 600; margin: 32px 0 12px; padding-top: 16px; border-bottom: 2px solid #E7E2DA; padding-bottom: 6px; }
        h3 { color: #57534E; font-size: 15px; font-weight: 600; margin: 20px 0 8px; }
        p { margin: 8px 0; font-size: 14px; }
        ul, ol { margin: 8px 0 8px 24px; font-size: 14px; }
        li { margin: 4px 0; }
        table { width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 13px; }
        th { background: #B8942F; color: white; text-align: left; padding: 8px 12px; }
        td { border-bottom: 1px solid #E7E2DA; padding: 7px 12px; }
        tr:nth-child(even) { background: #FFFFFF; }
        .hero { text-align: center; padding: 16px 0 24px; }
        .hero h1 { font-size: 32px; padding-top: 0; }
        .hero p { color: #57534E; font-size: 15px; }
        .tip { background: #FBF7F0; border-left: 4px solid #B8942F; padding: 12px 16px; margin: 12px 0; border-radius: 0 6px 6px 0; font-size: 13px; }
        .note { background: #F3F0FF; border-left: 4px solid #7C3AED; padding: 12px 16px; margin: 12px 0; border-radius: 0 6px 6px 0; font-size: 13px; }
        .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
        .badge-free { background: #F5F5F4; color: #A8A29E; }
        .badge-biz { background: #DBEAFE; color: #2563EB; }
        .badge-ent { background: #EDE9FE; color: #7C3AED; }
        .badge-trial { background: #FEF3C7; color: #D97706; }
        .feature-grid { display: flex; flex-wrap: wrap; gap: 12px; margin: 16px 0; }
        .feature-card { flex: 1 1 45%; background: white; border: 1px solid #E7E2DA; border-radius: 8px; padding: 16px; }
        .feature-card h3 { margin-top: 0; color: #B8942F; }
        .section-divider { border: none; border-top: 1px solid #E7E2DA; margin: 24px 0; }
        kbd { background: #F5F5F4; border: 1px solid #D6D3D1; border-radius: 3px; padding: 2px 6px; font-family: 'Segoe UI', monospace; font-size: 12px; }
        .step { display: flex; align-items: flex-start; margin: 8px 0; }
        .step-num { background: #B8942F; color: white; width: 24px; height: 24px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 600; flex-shrink: 0; margin-right: 10px; margin-top: 2px; }
        .step-text { flex: 1; font-size: 14px; }
        </style>
        <script type="text/javascript">
        function scrollTo(id) {
            var el = document.getElementById(id);
            if (el) el.scrollIntoView(true);
        }
        </script>
        </head>
        """;

    private void AppendJapaneseContent(StringBuilder sb)
    {
        var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var verStr = $"{ver?.Major}.{ver?.Minor}.{ver?.Build}";

        sb.Append($"""
            <div class="hero" id="overview">
                <h1>Insight AI Office</h1>
                <p>AI の民主化：誰もが業務で AI を使えるように</p>
                <p style="font-size:13px;color:#57534E;margin-top:4px;">Word / Excel / PowerPoint を AI で分析・編集する統合オフィスツール</p>
                <p style="font-size:12px;color:#A8A29E;">バージョン {verStr}</p>
            </div>

            <div class="feature-grid">
                <div class="feature-card"><h3>オフィスを編集する</h3><p>Word・Excel を直接編集、PowerPoint のスライドを操作、PDF を表示。ファイルを開くだけですぐに作業開始。</p></div>
                <div class="feature-card"><h3>新しい資料を作る</h3><p>開いたファイル（Word / Excel / PowerPoint / PDF）を AI が読み取り、要約・分析・翻訳。レポート・グラフ・表を自動生成してアーティファクトとして保存。</p></div>
            </div>

            <div class="tip">
                <strong>Insight AI Office の 2 つの価値</strong><br/>
                1. Word・Excel・PowerPoint を開いて直接編集できます。<br/>
                2. 開いたファイルの内容を AI が読み取り、要約・校正・分析・翻訳・統計などを実行。結果は HTML レポート、グラフ（Chart.js）、表、Mermaid 図、SVG などのアーティファクトとして自動保存され、リンクからいつでも閲覧できます。<br/>
                <strong>編集するだけでなく、新しい成果物を生み出す</strong> — それが Insight AI Office です。
            </div>

            <div class="feature-grid">
                <div class="feature-card"><h3>Word 編集</h3><p>Word (.docx/.doc) の表示・編集・保存。太字・斜体・下線・取り消し線、段落配置、画像挿入、検索・置換、印刷に対応。</p></div>
                <div class="feature-card"><h3>Excel 編集</h3><p>Excel (.xlsx/.xls/.csv) の表示・編集・保存。クリップボード操作、書式設定、セル結合、数値書式、ウィンドウ枠の固定に対応。</p></div>
                <div class="feature-card"><h3>PowerPoint ビューア</h3><p>PowerPoint (.pptx/.ppt) のスライド表示・操作。スライドの追加・複製・削除・並べ替え、テキスト抽出、PDF エクスポートに対応。</p></div>
                <div class="feature-card"><h3>AI コンシェルジュ</h3><p>Claude / OpenAI / Gemini 対応。要約・校正・分析をワンクリックで。248 種類以上のプロンプトプリセット搭載。</p></div>
            </div>

            <hr class="section-divider"/>

            <h2 id="ui-layout">画面構成</h2>
            <h3>タイトルバー</h3>
            <ul>
                <li>製品名・バージョン・ライセンスプラン（バッジ）・現在のファイル名を表示</li>
                <li>右側に「AI コンシェルジュ」ボタン — クリックでチャットパネルの開閉</li>
                <li>ウィンドウ操作ボタン（最小化・最大化・閉じる）</li>
                <li>タイトルバーのダブルクリックでウィンドウ最大化/復元</li>
            </ul>

            <h3>リボン</h3>
            <p>開いているファイル形式に応じてリボンが自動で切り替わります：</p>
            <ul>
                <li><strong>デフォルト</strong> — ファイル未オープン時。「開く」「ヘルプ」</li>
                <li><strong>Word</strong> — ファイル操作、元に戻す/やり直し、フォント書式、段落配置、画像挿入、検索・置換、印刷、ヘルプ</li>
                <li><strong>Excel</strong> — ファイル操作、元に戻す/やり直し、クリップボード、フォント書式、配置、数値書式、表示設定、ヘルプ</li>
                <li><strong>PowerPoint</strong> — ファイル操作、スライド操作（追加・複製・削除・移動）、テキスト抽出、ヘルプ</li>
            </ul>

            <h3>バックステージ</h3>
            <p>リボン左端の「ファイル」タブをクリックすると開きます：</p>
            <ul>
                <li><strong>開く / 名前を付けて保存</strong> — ファイル操作</li>
                <li><strong>最近使ったファイル</strong> — 直近に開いたファイルの一覧</li>
                <li><strong>設定</strong> — AI 設定ダイアログを開く</li>
                <li><strong>言語</strong> — 日本語 / English の切り替え</li>
                <li><strong>ライセンス</strong> — 現在のプラン・有効期限の確認</li>
                <li><strong>ドキュメントを閉じる</strong> — 現在のタブを閉じる</li>
            </ul>

            <h3>3 ペインレイアウト</h3>
            <ul>
                <li><strong>左パネル</strong> — 参考資料パネル。ファイルをドラッグ&ドロップまたは「＋」で追加。AI が自動的にコンテキストとして使用。</li>
                <li><strong>中央</strong> — ドキュメントエディタ / ビューア（タブバー付き）</li>
                <li><strong>右パネル</strong> — AI コンシェルジュ（チャット）。ドラッグで幅を調整可能。</li>
            </ul>

            <h3>ステータスバー</h3>
            <p>画面下部に操作状態、ファイル形式、バージョン、ズーム倍率を表示します。</p>

            <hr class="section-divider"/>

            <h2 id="file-ops">ファイル操作</h2>
            <h3>対応フォーマット</h3>
            <table>
                <tr><th>形式</th><th>拡張子</th><th>操作</th></tr>
                <tr><td>Word</td><td>.docx, .doc</td><td>表示・編集・保存・印刷</td></tr>
                <tr><td>Excel</td><td>.xlsx, .xls, .csv</td><td>表示・編集・保存</td></tr>
                <tr><td>PowerPoint</td><td>.pptx, .ppt</td><td>表示・スライド操作・テキスト抽出・PDF エクスポート</td></tr>
                <tr><td>PDF</td><td>.pdf</td><td>表示・AI 分析・要約・翻訳・資料作成</td></tr>
            </table>

            <h3>ファイルを開く</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">リボンの「開く」ボタンをクリック、または <kbd>Ctrl</kbd>+<kbd>O</kbd></div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">ファイル選択ダイアログからファイルを選択</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">ファイル形式に応じたエディタ / ビューアが自動で開きます</div></div>

            <div class="tip">ファイルをウィンドウにドラッグ&ドロップしても開けます。複数ファイルの同時ドロップにも対応しています。</div>

            <h3>保存・エクスポート</h3>
            <ul>
                <li><strong>上書き保存</strong> — <kbd>Ctrl</kbd>+<kbd>S</kbd> またはリボンの「保存」ボタン</li>
                <li><strong>名前を付けて保存</strong> — バックステージの「名前を付けて保存」</li>
                <li><strong>Word</strong> — DOCX 形式でエクスポート</li>
                <li><strong>Excel</strong> — XLSX 形式でエクスポート</li>
                <li><strong>PowerPoint</strong> — PDF 形式でエクスポート</li>
            </ul>

            <h3>最近使ったファイル</h3>
            <p>バックステージの「最近使ったファイル」タブから、直近に開いたファイルをワンクリックで再度開けます。</p>

            <hr class="section-divider"/>

            <h2 id="multi-tab">マルチタブ</h2>
            <p>複数のファイルを同時に開いてタブで切り替えることができます。</p>

            <h3>タブの操作</h3>
            <ul>
                <li><strong>新しいタブで開く</strong> — ファイルを開くと自動的に新しいタブが追加されます</li>
                <li><strong>タブの切り替え</strong> — タブバーのタブをクリックして切り替え</li>
                <li><strong>タブを閉じる</strong> — タブの「×」ボタン、またはバックステージの「ドキュメントを閉じる」</li>
                <li><strong>同じファイルを再度開く</strong> — 既に開いているファイルを開くと、そのタブに切り替わります</li>
            </ul>

            <h3>タブの状態保持</h3>
            <p>タブを切り替えても編集内容は保持されます：</p>
            <ul>
                <li><strong>Word</strong> — 編集内容をメモリに自動保存・復元</li>
                <li><strong>Excel</strong> — ファイルから自動再読み込み</li>
                <li><strong>PowerPoint</strong> — 選択中のスライド位置を記憶</li>
            </ul>

            <div class="tip">タブバーにはファイル形式のアイコン（Word/Excel/PowerPoint）が表示されるため、一目でファイルの種類が分かります。</div>

            <hr class="section-divider"/>

            <h2 id="word-edit">Word 編集</h2>
            <p>Word (.docx/.doc) ファイルの本格的な編集機能を提供します。</p>

            <h3>フォント書式</h3>
            <table>
                <tr><th>操作</th><th>リボンボタン</th><th>ショートカット</th></tr>
                <tr><td>太字</td><td>B</td><td><kbd>Ctrl</kbd>+<kbd>B</kbd></td></tr>
                <tr><td>斜体</td><td>I</td><td><kbd>Ctrl</kbd>+<kbd>I</kbd></td></tr>
                <tr><td>下線</td><td>U</td><td><kbd>Ctrl</kbd>+<kbd>U</kbd></td></tr>
                <tr><td>取り消し線</td><td>abc</td><td>—</td></tr>
            </table>

            <h3>段落配置</h3>
            <ul>
                <li><strong>左揃え</strong> — テキストを左に配置</li>
                <li><strong>中央揃え</strong> — テキストを中央に配置</li>
                <li><strong>右揃え</strong> — テキストを右に配置</li>
                <li><strong>箇条書き</strong> — 箇条書きリストの挿入/解除</li>
            </ul>

            <h3>画像挿入</h3>
            <p>リボンの「画像」ボタンから、PNG / JPG / BMP / GIF / TIFF 形式の画像をドキュメントに挿入できます。</p>

            <h3>検索・置換</h3>
            <p>リボンの「検索と置換」ボタン、または <kbd>Ctrl</kbd>+<kbd>H</kbd> で検索・置換ダイアログを開きます。</p>

            <h3>元に戻す / やり直し</h3>
            <p><kbd>Ctrl</kbd>+<kbd>Z</kbd> で元に戻す、<kbd>Ctrl</kbd>+<kbd>Y</kbd> でやり直しができます。</p>

            <h3>印刷</h3>
            <p>リボンの「印刷」ボタン、または <kbd>Ctrl</kbd>+<kbd>P</kbd> でドキュメントを印刷できます。</p>

            <hr class="section-divider"/>

            <h2 id="excel-edit">Excel 編集</h2>
            <p>Excel (.xlsx/.xls/.csv) ファイルの編集機能を提供します。</p>

            <h3>クリップボード操作</h3>
            <table>
                <tr><th>操作</th><th>ショートカット</th></tr>
                <tr><td>切り取り</td><td><kbd>Ctrl</kbd>+<kbd>X</kbd></td></tr>
                <tr><td>コピー</td><td><kbd>Ctrl</kbd>+<kbd>C</kbd></td></tr>
                <tr><td>貼り付け</td><td><kbd>Ctrl</kbd>+<kbd>V</kbd></td></tr>
            </table>

            <h3>フォント書式</h3>
            <ul>
                <li><strong>太字 / 斜体 / 下線</strong> — 選択セルのテキスト書式を変更</li>
                <li><strong>罫線</strong> — 選択範囲にセル罫線を追加</li>
            </ul>

            <h3>配置</h3>
            <ul>
                <li><strong>左揃え / 中央揃え / 右揃え</strong> — セル内の水平配置を変更</li>
                <li><strong>折り返して全体を表示</strong> — セル内テキストの折り返し切り替え</li>
                <li><strong>セルの結合</strong> — 選択範囲のセル結合 / 結合解除</li>
            </ul>

            <h3>数値書式</h3>
            <table>
                <tr><th>ボタン</th><th>書式</th><th>例</th></tr>
                <tr><td>%</td><td>パーセント</td><td>0.5 → 50%</td></tr>
                <tr><td>,</td><td>桁区切り</td><td>1000 → 1,000</td></tr>
                <tr><td>¥</td><td>通貨</td><td>1000 → ¥1,000</td></tr>
                <tr><td>.0 / .00</td><td>小数点以下の増減</td><td>表示桁数を調整</td></tr>
            </table>

            <h3>表示</h3>
            <ul>
                <li><strong>ウィンドウ枠の固定</strong> — スクロール時にヘッダー行/列を固定表示</li>
            </ul>

            <h3>元に戻す / やり直し</h3>
            <p><kbd>Ctrl</kbd>+<kbd>Z</kbd> で元に戻す、<kbd>Ctrl</kbd>+<kbd>Y</kbd> でやり直しができます。</p>

            <hr class="section-divider"/>

            <h2 id="pptx-view">PowerPoint ビューア</h2>
            <p>PowerPoint (.pptx/.ppt) のスライド表示と操作機能を提供します。</p>

            <h3>スライド表示</h3>
            <ul>
                <li>左側にサムネイル一覧、中央に選択スライドの拡大表示</li>
                <li>サムネイルをクリックしてスライドを切り替え</li>
                <li>スライド番号 / 総数を表示</li>
            </ul>

            <h3>スライド操作</h3>
            <table>
                <tr><th>操作</th><th>説明</th></tr>
                <tr><td>スライド追加</td><td>選択スライドの後に新しいスライドを挿入</td></tr>
                <tr><td>スライド複製</td><td>選択スライドのコピーを直後に挿入</td></tr>
                <tr><td>スライド削除</td><td>選択スライドを削除（最後の 1 枚は削除不可）</td></tr>
                <tr><td>上へ移動</td><td>選択スライドを 1 つ前に移動</td></tr>
                <tr><td>下へ移動</td><td>選択スライドを 1 つ後に移動</td></tr>
            </table>

            <h3>テキスト抽出</h3>
            <p>リボンの「テキスト抽出」ボタンで、全スライドのテキストとノートをクリップボードにコピーします。</p>

            <h3>PDF エクスポート</h3>
            <p>バックステージまたはリボンの「PDF エクスポート」から、プレゼンテーション全体を PDF に変換・保存できます。</p>

            <hr class="section-divider"/>

            <h2 id="pdf-view">PDF ビューア</h2>
            <p>PDF (.pdf) ファイルを表示し、AI コンシェルジュと連携して活用できます。</p>

            <h3>PDF の表示</h3>
            <ul>
                <li>Syncfusion PDF ビューアによる高品質なレンダリング</li>
                <li>ページ送り、ズーム、スクロール操作</li>
                <li>マルチタブで他のファイルと同時に開ける</li>
            </ul>

            <h3>AI との連携</h3>
            <p>PDF を開くと、テキスト内容が自動的に AI コンテキストとして抽出されます。これにより：</p>
            <ul>
                <li><strong>要約</strong> — PDF の内容を要約</li>
                <li><strong>翻訳</strong> — PDF を英訳・和訳</li>
                <li><strong>資料作成</strong> — PDF を元にレポート・企画書・報告書を新規作成</li>
                <li><strong>分析</strong> — PDF のデータからグラフ・表を生成</li>
                <li><strong>Q&amp;A</strong> — PDF の内容に関する質問応答</li>
            </ul>

            <div class="tip">
                <strong>活用例：</strong>契約書の PDF を開いて「この契約のリスクポイントを一覧にまとめて」と指示すると、AI が内容を読み取り、リスク分析レポートをアーティファクトとして自動生成します。
            </div>

            <hr class="section-divider"/>

            <h2 id="ai-assistant">AI コンシェルジュ</h2>
            <p>AI コンシェルジュは、ドキュメントの内容を理解し、要約・校正・分析・翻訳などを行う AI チャット機能です。</p>

            <h3>セットアップ</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">バックステージの「設定」、またはリボンの「AI 設定」をクリック</div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">プロバイダー（Claude / OpenAI / Gemini）を選択し、API キーを入力</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">モデルを選択して「保存」</div></div>

            <div class="note">API キーはお客様ご自身で取得してください（BYOK 方式）。キーは端末内に暗号化保存され、外部に送信されることはありません。</div>

            <h3>チャットパネル</h3>
            <ul>
                <li>タイトルバーの「AI コンシェルジュ」ボタンで開閉</li>
                <li>開いているドキュメントの内容を自動的にコンテキストとして AI に送信</li>
                <li>AI の回答をドキュメントに挿入（Word の場合）またはクリップボードにコピー</li>
                <li>パネル幅はドラッグで自由に調整可能</li>
            </ul>

            <h3>ドキュメントコンテキストの自動取得</h3>
            <p>AI コンシェルジュは開いているドキュメントの内容を自動で読み取ります：</p>
            <ul>
                <li><strong>Word</strong> — テキスト全文（最大 8,000 文字）</li>
                <li><strong>Excel</strong> — セルデータ（最大 100 行 × 20 列、8,000 文字）</li>
                <li><strong>PowerPoint</strong> — 全スライドのテキスト＋ノート（最大 8,000 文字）</li>
            </ul>

            <h3>参考資料</h3>
            <p>左パネルに参考資料を追加すると、AI が自動的にコンテキストとして活用します。Word / Excel / PowerPoint / テキストファイルに対応。</p>

            <h3>AI の回答をドキュメントに挿入</h3>
            <p>AI の回答は以下の方法でドキュメントに反映できます：</p>
            <ul>
                <li><strong>Word</strong> — カーソル位置にテキストを直接挿入</li>
                <li><strong>Excel / PowerPoint</strong> — クリップボードにコピーして手動で貼り付け</li>
            </ul>

            <h3>アーティファクト（成果物の自動保存）</h3>
            <p>AI の回答に含まれるレポート・グラフ・表などは<strong>アーティファクト</strong>として自動保存されます：</p>
            <table>
                <tr><th>種類</th><th>説明</th></tr>
                <tr><td>HTML レポート</td><td>要約・分析結果をリッチな HTML で保存</td></tr>
                <tr><td>Chart.js グラフ</td><td>棒グラフ・折れ線・円グラフなど</td></tr>
                <tr><td>HTML テーブル</td><td>データの一覧表</td></tr>
                <tr><td>Mermaid 図</td><td>フローチャート・シーケンス図など</td></tr>
                <tr><td>SVG 画像</td><td>ベクター画像</td></tr>
                <tr><td>Markdown</td><td>マークダウン形式のドキュメント</td></tr>
            </table>
            <p>保存されたアーティファクトは成果物一覧からリンクでいつでも閲覧できます。30 日以上経過したアーティファクトは自動的にクリーンアップされます。</p>

            <div class="note">
                <strong>活用例：</strong>Excel を開いて「売上データを分析してグラフ付きレポートを作成して」と指示すると、AI がデータを読み取り、Chart.js グラフ付きの HTML レポートをアーティファクトとして自動生成・保存します。
            </div>

            <hr class="section-divider"/>

            <h2 id="prompt-presets">プロンプトプリセット</h2>
            <p>248 種類以上のプロンプトプリセットを搭載。製品別（Word / Excel / PowerPoint）に整理されています。</p>

            <h3>プリセットの使い方</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">AI コンシェルジュパネルの「プロンプトエディタ」ボタンをクリック</div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">ツリーからプリセットを選択</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">「実行」ボタンで AI にプロンプトを送信</div></div>

            <h3>内蔵プリセットカテゴリ</h3>
            <table>
                <tr><th>グループ</th><th>カテゴリ例</th></tr>
                <tr><td>Office 共通</td><td>分析・要約、校正・チェック、翻訳、文書作成</td></tr>
                <tr><td>Word</td><td>品質レビュー、文書改善、翻訳、分析・要約、ビジュアル・HTML、AI・自動化、業種特化</td></tr>
                <tr><td>Excel</td><td>品質レビュー、分析・レポート、翻訳、AI・自動化、業種特化</td></tr>
                <tr><td>PowerPoint</td><td>品質レビュー、プレゼン改善、翻訳、分析・要約、ビジュアル・HTML、AI・自動化、業種特化</td></tr>
            </table>

            <h3>カスタムプリセット</h3>
            <ul>
                <li><strong>新規作成</strong> — 「＋」ボタンでカスタムプリセットを作成</li>
                <li><strong>編集</strong> — 名前、カテゴリ、モデル、プロンプトテキストを自由に編集</li>
                <li><strong>カスタマイズ</strong> — 内蔵プリセットを元にカスタムコピーを作成</li>
                <li><strong>エクスポート / インポート</strong> — カスタムプリセットを JSON ファイルで共有</li>
            </ul>

            <hr class="section-divider"/>

            <h2 id="shortcuts">キーボードショートカット</h2>
            <h3>全般</h3>
            <table>
                <tr><th>キー</th><th>操作</th></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>N</kbd></td><td>新規作成</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>O</kbd></td><td>ファイルを開く</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>S</kbd></td><td>上書き保存</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>Z</kbd></td><td>元に戻す</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>Y</kbd></td><td>やり直し</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>+</kbd></td><td>UI 拡大</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>-</kbd></td><td>UI 縮小</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>0</kbd></td><td>UI 100% にリセット</td></tr>
                <tr><td><kbd>F1</kbd></td><td>ヘルプを開く</td></tr>
            </table>

            <h3>Word</h3>
            <table>
                <tr><th>キー</th><th>操作</th></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>B</kbd></td><td>太字</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>I</kbd></td><td>斜体</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>U</kbd></td><td>下線</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>H</kbd></td><td>検索と置換</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>P</kbd></td><td>印刷</td></tr>
            </table>

            <h3>Excel</h3>
            <table>
                <tr><th>キー</th><th>操作</th></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>X</kbd></td><td>切り取り</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>C</kbd></td><td>コピー</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>V</kbd></td><td>貼り付け</td></tr>
            </table>

            <div class="tip">Windows 標準の <kbd>Win</kbd>+<kbd>H</kbd> で音声入力が利用できます。AI との対話にも活用できます。</div>

            <hr class="section-divider"/>

            <h2 id="license">ライセンス</h2>
            <table>
                <tr><th>プラン</th><th>説明</th><th>有効期限</th></tr>
                <tr><td><span class="badge badge-free">FREE</span></td><td>全機能利用可能（エクスポート制限あり）</td><td>無期限</td></tr>
                <tr><td><span class="badge badge-trial">TRIAL</span></td><td>全機能利用可能（評価用）</td><td>30日間</td></tr>
                <tr><td><span class="badge badge-biz">BIZ</span></td><td>法人向け全機能</td><td>365日</td></tr>
                <tr><td><span class="badge badge-ent">ENT</span></td><td>法人向け全機能 + API/SSO/監査ログ</td><td>要相談</td></tr>
            </table>

            <h3>ライセンスキー形式</h3>
            <p><code>IAOF-[Plan]-[YYMM]-[HASH]-[SIG1]-[SIG2]</code></p>

            <h3>ライセンス管理</h3>
            <p>タイトルバーのプランバッジをクリックするとライセンス管理画面が開きます。ライセンスキーの入力・確認・変更が行えます。</p>
            <p>バックステージの「ライセンス」タブからも現在のプラン・有効期限を確認できます。</p>

            <hr class="section-divider"/>

            <h2 id="system-req">システム要件</h2>
            <table>
                <tr><th>項目</th><th>要件</th></tr>
                <tr><td>OS</td><td>Windows 10 (64-bit) / Windows 11</td></tr>
                <tr><td>ランタイム</td><td>.NET 8.0 Desktop Runtime</td></tr>
                <tr><td>メモリ</td><td>4 GB 以上（8 GB 推奨）</td></tr>
                <tr><td>ディスク</td><td>500 MB 以上の空き容量</td></tr>
                <tr><td>インターネット</td><td>AI 機能の使用時に必要</td></tr>
            </table>

            <hr class="section-divider"/>

            <h2 id="support">お問い合わせ</h2>
            <table>
                <tr><th>項目</th><th>内容</th></tr>
                <tr><td>開発元</td><td>HARMONIC insight</td></tr>
                <tr><td>製品名</td><td>Insight AI Office (IAOF)</td></tr>
                <tr><td>メール</td><td><a href="mailto:support@h-insight.jp">support@h-insight.jp</a></td></tr>
            </table>
            <p style="margin-top:24px;font-size:12px;color:#A8A29E;">Copyright &copy; 2025-2026 HARMONIC insight. All rights reserved.</p>
            """);
    }

    private void AppendEnglishContent(StringBuilder sb)
    {
        var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var verStr = $"{ver?.Major}.{ver?.Minor}.{ver?.Build}";

        sb.Append($"""
            <div class="hero" id="overview">
                <h1>Insight AI Office</h1>
                <p>Democratizing AI: Empowering Everyone to Use AI at Work</p>
                <p style="font-size:13px;color:#57534E;margin-top:4px;">An integrated office tool for analyzing and editing Word / Excel / PowerPoint with AI</p>
                <p style="font-size:12px;color:#A8A29E;">Version {verStr}</p>
            </div>

            <div class="feature-grid">
                <div class="feature-card"><h3>Edit Office Documents</h3><p>Edit Word & Excel directly, manage PowerPoint slides, view PDFs. Open a file and start working immediately.</p></div>
                <div class="feature-card"><h3>Create New Materials</h3><p>AI reads your open files (Word / Excel / PowerPoint / PDF), then summarizes, analyzes, and translates. Auto-generates reports, charts, and tables saved as artifacts.</p></div>
            </div>

            <div class="tip">
                <strong>Two Core Values of Insight AI Office</strong><br/>
                1. Open and directly edit Word, Excel, and PowerPoint files.<br/>
                2. AI reads the content of your open files and performs summarization, proofreading, analysis, translation, and statistics. Results are auto-saved as artifacts (HTML reports, Chart.js graphs, tables, Mermaid diagrams, SVG) and accessible via links anytime.<br/>
                <strong>Not just editing — creating new deliverables</strong> — that's Insight AI Office.
            </div>

            <div class="feature-grid">
                <div class="feature-card"><h3>Word Editing</h3><p>View, edit, and save Word (.docx/.doc) files. Bold, italic, underline, strikethrough, paragraph alignment, image insertion, find & replace, and printing.</p></div>
                <div class="feature-card"><h3>Excel Editing</h3><p>View, edit, and save Excel (.xlsx/.xls/.csv) files. Clipboard operations, formatting, cell merge, number formats, and freeze panes.</p></div>
                <div class="feature-card"><h3>PowerPoint Viewer</h3><p>View and manage PowerPoint (.pptx/.ppt) slides. Add, duplicate, delete, reorder slides, extract text, and export to PDF.</p></div>
                <div class="feature-card"><h3>AI Concierge</h3><p>Supports Claude / OpenAI / Gemini. One-click summarize, proofread, and analyze. 248+ built-in prompt presets.</p></div>
            </div>

            <hr class="section-divider"/>

            <h2 id="ui-layout">UI Layout</h2>
            <h3>Title Bar</h3>
            <ul>
                <li>Displays product name, version, license plan badge, and current file name</li>
                <li>"AI Concierge" button on the right — click to toggle the chat panel</li>
                <li>Window control buttons (minimize, maximize, close)</li>
                <li>Double-click the title bar to maximize/restore the window</li>
            </ul>

            <h3>Ribbon</h3>
            <p>The ribbon automatically switches based on the open file format:</p>
            <ul>
                <li><strong>Default</strong> — No file open. "Open" and "Help"</li>
                <li><strong>Word</strong> — File operations, undo/redo, font formatting, paragraph alignment, image insertion, find & replace, print, help</li>
                <li><strong>Excel</strong> — File operations, undo/redo, clipboard, font formatting, alignment, number formats, view settings, help</li>
                <li><strong>PowerPoint</strong> — File operations, slide operations (add/duplicate/delete/move), text extraction, help</li>
            </ul>

            <h3>Backstage</h3>
            <p>Click the "File" tab at the left of the ribbon to access:</p>
            <ul>
                <li><strong>Open / Save As</strong> — File operations</li>
                <li><strong>Recent Files</strong> — List of recently opened files</li>
                <li><strong>Settings</strong> — Open AI settings dialog</li>
                <li><strong>Language</strong> — Switch between Japanese / English</li>
                <li><strong>License</strong> — Check current plan and expiry date</li>
                <li><strong>Close Document</strong> — Close the current tab</li>
            </ul>

            <h3>3-Pane Layout</h3>
            <ul>
                <li><strong>Left Panel</strong> — Reference materials. Drag & drop or click "+" to add files. AI automatically uses them as context.</li>
                <li><strong>Center</strong> — Document editor / viewer (with tab bar)</li>
                <li><strong>Right Panel</strong> — AI Concierge (chat). Drag the edge to resize.</li>
            </ul>

            <h3>Status Bar</h3>
            <p>Displays operation status, file format, version, and zoom level at the bottom of the window.</p>

            <hr class="section-divider"/>

            <h2 id="file-ops">File Operations</h2>
            <h3>Supported Formats</h3>
            <table>
                <tr><th>Format</th><th>Extensions</th><th>Operations</th></tr>
                <tr><td>Word</td><td>.docx, .doc</td><td>View, Edit, Save, Print</td></tr>
                <tr><td>Excel</td><td>.xlsx, .xls, .csv</td><td>View, Edit, Save</td></tr>
                <tr><td>PowerPoint</td><td>.pptx, .ppt</td><td>View, Slide Operations, Text Extraction, PDF Export</td></tr>
                <tr><td>PDF</td><td>.pdf</td><td>View, AI Analysis, Summarize, Translate, Create Materials</td></tr>
            </table>

            <h3>Opening Files</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">Click "Open" in the ribbon, or press <kbd>Ctrl</kbd>+<kbd>O</kbd></div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">Select a file from the dialog</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">The appropriate editor/viewer opens automatically</div></div>

            <div class="tip">You can also drag & drop files onto the window. Multiple files can be dropped simultaneously.</div>

            <h3>Saving & Exporting</h3>
            <ul>
                <li><strong>Save</strong> — <kbd>Ctrl</kbd>+<kbd>S</kbd> or the "Save" ribbon button</li>
                <li><strong>Save As</strong> — Via the backstage "Save As" option</li>
                <li><strong>Word</strong> — Export as DOCX</li>
                <li><strong>Excel</strong> — Export as XLSX</li>
                <li><strong>PowerPoint</strong> — Export as PDF</li>
            </ul>

            <h3>Recent Files</h3>
            <p>Access the "Recent Files" tab in the backstage to quickly reopen recently used files.</p>

            <hr class="section-divider"/>

            <h2 id="multi-tab">Multi-Tab</h2>
            <p>Open multiple files simultaneously and switch between them using tabs.</p>

            <h3>Tab Operations</h3>
            <ul>
                <li><strong>Open in New Tab</strong> — Opening a file automatically adds a new tab</li>
                <li><strong>Switch Tabs</strong> — Click a tab in the tab bar to switch</li>
                <li><strong>Close Tab</strong> — Click the "x" button on the tab, or use "Close Document" in the backstage</li>
                <li><strong>Reopen Same File</strong> — Opening an already-open file switches to its tab</li>
            </ul>

            <h3>Tab State Preservation</h3>
            <p>Your editing state is preserved when switching tabs:</p>
            <ul>
                <li><strong>Word</strong> — Content is auto-saved and restored in memory</li>
                <li><strong>Excel</strong> — Automatically reloaded from file</li>
                <li><strong>PowerPoint</strong> — Selected slide position is remembered</li>
            </ul>

            <div class="tip">File type icons (Word/Excel/PowerPoint) are shown in the tab bar for quick identification.</div>

            <hr class="section-divider"/>

            <h2 id="word-edit">Word Editing</h2>
            <p>Full-featured editing for Word (.docx/.doc) files.</p>

            <h3>Font Formatting</h3>
            <table>
                <tr><th>Action</th><th>Ribbon Button</th><th>Shortcut</th></tr>
                <tr><td>Bold</td><td>B</td><td><kbd>Ctrl</kbd>+<kbd>B</kbd></td></tr>
                <tr><td>Italic</td><td>I</td><td><kbd>Ctrl</kbd>+<kbd>I</kbd></td></tr>
                <tr><td>Underline</td><td>U</td><td><kbd>Ctrl</kbd>+<kbd>U</kbd></td></tr>
                <tr><td>Strikethrough</td><td>abc</td><td>—</td></tr>
            </table>

            <h3>Paragraph Alignment</h3>
            <ul>
                <li><strong>Align Left</strong> — Align text to the left</li>
                <li><strong>Align Center</strong> — Center-align text</li>
                <li><strong>Align Right</strong> — Align text to the right</li>
                <li><strong>Bullet List</strong> — Toggle bullet list formatting</li>
            </ul>

            <h3>Image Insertion</h3>
            <p>Insert images (PNG / JPG / BMP / GIF / TIFF) into the document via the "Image" ribbon button.</p>

            <h3>Find & Replace</h3>
            <p>Open the Find & Replace dialog with the ribbon button or <kbd>Ctrl</kbd>+<kbd>H</kbd>.</p>

            <h3>Undo / Redo</h3>
            <p><kbd>Ctrl</kbd>+<kbd>Z</kbd> to undo, <kbd>Ctrl</kbd>+<kbd>Y</kbd> to redo.</p>

            <h3>Print</h3>
            <p>Print the document via the ribbon button or <kbd>Ctrl</kbd>+<kbd>P</kbd>.</p>

            <hr class="section-divider"/>

            <h2 id="excel-edit">Excel Editing</h2>
            <p>Full-featured editing for Excel (.xlsx/.xls/.csv) files.</p>

            <h3>Clipboard Operations</h3>
            <table>
                <tr><th>Action</th><th>Shortcut</th></tr>
                <tr><td>Cut</td><td><kbd>Ctrl</kbd>+<kbd>X</kbd></td></tr>
                <tr><td>Copy</td><td><kbd>Ctrl</kbd>+<kbd>C</kbd></td></tr>
                <tr><td>Paste</td><td><kbd>Ctrl</kbd>+<kbd>V</kbd></td></tr>
            </table>

            <h3>Font Formatting</h3>
            <ul>
                <li><strong>Bold / Italic / Underline</strong> — Toggle text formatting for selected cells</li>
                <li><strong>Borders</strong> — Add cell borders to the selected range</li>
            </ul>

            <h3>Alignment</h3>
            <ul>
                <li><strong>Left / Center / Right Align</strong> — Change horizontal alignment within cells</li>
                <li><strong>Wrap Text</strong> — Toggle text wrapping within cells</li>
                <li><strong>Merge Cells</strong> — Merge or unmerge the selected cell range</li>
            </ul>

            <h3>Number Formats</h3>
            <table>
                <tr><th>Button</th><th>Format</th><th>Example</th></tr>
                <tr><td>%</td><td>Percentage</td><td>0.5 → 50%</td></tr>
                <tr><td>,</td><td>Comma separator</td><td>1000 → 1,000</td></tr>
                <tr><td>¥</td><td>Currency</td><td>1000 → ¥1,000</td></tr>
                <tr><td>.0 / .00</td><td>Decimal places</td><td>Adjust display precision</td></tr>
            </table>

            <h3>View</h3>
            <ul>
                <li><strong>Freeze Panes</strong> — Keep header rows/columns visible while scrolling</li>
            </ul>

            <h3>Undo / Redo</h3>
            <p><kbd>Ctrl</kbd>+<kbd>Z</kbd> to undo, <kbd>Ctrl</kbd>+<kbd>Y</kbd> to redo.</p>

            <hr class="section-divider"/>

            <h2 id="pptx-view">PowerPoint Viewer</h2>
            <p>View and manage PowerPoint (.pptx/.ppt) presentations.</p>

            <h3>Slide Display</h3>
            <ul>
                <li>Thumbnail list on the left, enlarged selected slide in the center</li>
                <li>Click a thumbnail to switch slides</li>
                <li>Slide number and total count displayed</li>
            </ul>

            <h3>Slide Operations</h3>
            <table>
                <tr><th>Action</th><th>Description</th></tr>
                <tr><td>Add Slide</td><td>Insert a new slide after the selected one</td></tr>
                <tr><td>Duplicate Slide</td><td>Create a copy of the selected slide</td></tr>
                <tr><td>Delete Slide</td><td>Remove the selected slide (cannot delete the last slide)</td></tr>
                <tr><td>Move Up</td><td>Move the selected slide one position earlier</td></tr>
                <tr><td>Move Down</td><td>Move the selected slide one position later</td></tr>
            </table>

            <h3>Text Extraction</h3>
            <p>Click "Extract Text" in the ribbon to copy all slide text and notes to the clipboard.</p>

            <h3>PDF Export</h3>
            <p>Export the entire presentation to PDF via the backstage or ribbon "Export PDF" button.</p>

            <hr class="section-divider"/>

            <h2 id="pdf-view">PDF Viewer</h2>
            <p>View PDF (.pdf) files and leverage the AI Concierge for content analysis and creation.</p>

            <h3>Viewing PDFs</h3>
            <ul>
                <li>High-quality rendering with Syncfusion PDF Viewer</li>
                <li>Page navigation, zoom, and scroll</li>
                <li>Open alongside other files in multi-tab mode</li>
            </ul>

            <h3>AI Integration</h3>
            <p>When a PDF is opened, its text content is automatically extracted as AI context. This enables:</p>
            <ul>
                <li><strong>Summarize</strong> — Summarize the PDF content</li>
                <li><strong>Translate</strong> — Translate the PDF to English or Japanese</li>
                <li><strong>Create Materials</strong> — Generate reports, proposals, and documents based on the PDF</li>
                <li><strong>Analyze</strong> — Create charts and tables from PDF data</li>
                <li><strong>Q&amp;A</strong> — Ask questions about the PDF content</li>
            </ul>

            <div class="tip">
                <strong>Example:</strong> Open a contract PDF and ask "List all risk points in this contract." The AI reads the content and auto-generates a risk analysis report as an artifact.
            </div>

            <hr class="section-divider"/>

            <h2 id="ai-assistant">AI Concierge</h2>
            <p>The AI Concierge understands your document content and provides summarization, proofreading, analysis, translation, and more via an AI chat interface.</p>

            <h3>Setup</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">Click "Settings" in the backstage, or "AI Settings" in the ribbon</div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">Select a provider (Claude / OpenAI / Gemini) and enter your API key</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">Choose a model and click "Save"</div></div>

            <div class="note">You need to obtain your own API key (BYOK model). Keys are stored encrypted on your device and never sent externally.</div>

            <h3>Chat Panel</h3>
            <ul>
                <li>Toggle with the "AI Concierge" button in the title bar</li>
                <li>Automatically sends the open document's content as context to the AI</li>
                <li>Insert AI responses into the document (Word) or copy to clipboard</li>
                <li>Panel width is adjustable by dragging the edge</li>
            </ul>

            <h3>Automatic Document Context</h3>
            <p>The AI Concierge automatically reads the content of the open document:</p>
            <ul>
                <li><strong>Word</strong> — Full text (up to 8,000 characters)</li>
                <li><strong>Excel</strong> — Cell data (up to 100 rows x 20 columns, 8,000 characters)</li>
                <li><strong>PowerPoint</strong> — All slide text + notes (up to 8,000 characters)</li>
            </ul>

            <h3>Reference Materials</h3>
            <p>Add reference materials to the left panel, and AI will automatically use them as context. Supports Word / Excel / PowerPoint / text files.</p>

            <h3>Inserting AI Responses</h3>
            <p>AI responses can be applied to your document:</p>
            <ul>
                <li><strong>Word</strong> — Insert text directly at the cursor position</li>
                <li><strong>Excel / PowerPoint</strong> — Copy to clipboard for manual paste</li>
            </ul>

            <h3>Artifacts (Auto-Saved Deliverables)</h3>
            <p>Reports, charts, and tables in AI responses are automatically saved as <strong>artifacts</strong>:</p>
            <table>
                <tr><th>Type</th><th>Description</th></tr>
                <tr><td>HTML Report</td><td>Rich HTML summaries and analysis results</td></tr>
                <tr><td>Chart.js Graph</td><td>Bar charts, line graphs, pie charts, etc.</td></tr>
                <tr><td>HTML Table</td><td>Data tables and listings</td></tr>
                <tr><td>Mermaid Diagram</td><td>Flowcharts, sequence diagrams, etc.</td></tr>
                <tr><td>SVG Image</td><td>Vector graphics</td></tr>
                <tr><td>Markdown</td><td>Markdown-formatted documents</td></tr>
            </table>
            <p>Saved artifacts can be viewed anytime via links in the deliverables list. Artifacts older than 30 days are automatically cleaned up.</p>

            <div class="note">
                <strong>Example:</strong> Open an Excel file and ask "Analyze the sales data and create a report with charts." The AI reads the data, then auto-generates and saves an HTML report with Chart.js charts as an artifact.
            </div>

            <hr class="section-divider"/>

            <h2 id="prompt-presets">Prompt Presets</h2>
            <p>248+ built-in prompt presets organized by product (Word / Excel / PowerPoint).</p>

            <h3>How to Use Presets</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">Click the "Prompt Editor" button in the AI Concierge panel</div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">Select a preset from the tree view</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">Click "Execute" to send the prompt to the AI</div></div>

            <h3>Built-in Preset Categories</h3>
            <table>
                <tr><th>Group</th><th>Example Categories</th></tr>
                <tr><td>Office Common</td><td>Analysis & Summary, Proofreading, Translation, Document Creation</td></tr>
                <tr><td>Word</td><td>Quality Review, Document Improvement, Translation, Analysis, Visual/HTML, AI/Automation, Industry-Specific</td></tr>
                <tr><td>Excel</td><td>Quality Review, Analysis & Reports, Translation, AI/Automation, Industry-Specific</td></tr>
                <tr><td>PowerPoint</td><td>Quality Review, Presentation Improvement, Translation, Analysis, Visual/HTML, AI/Automation, Industry-Specific</td></tr>
            </table>

            <h3>Custom Presets</h3>
            <ul>
                <li><strong>Create New</strong> — Click "+" to create a custom preset</li>
                <li><strong>Edit</strong> — Customize name, category, model, and prompt text</li>
                <li><strong>Customize</strong> — Create a custom copy based on a built-in preset</li>
                <li><strong>Export / Import</strong> — Share custom presets via JSON files</li>
            </ul>

            <hr class="section-divider"/>

            <h2 id="shortcuts">Keyboard Shortcuts</h2>
            <h3>General</h3>
            <table>
                <tr><th>Key</th><th>Action</th></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>N</kbd></td><td>New Document</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>O</kbd></td><td>Open File</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>S</kbd></td><td>Save</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>Z</kbd></td><td>Undo</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>Y</kbd></td><td>Redo</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>+</kbd></td><td>Zoom In UI</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>-</kbd></td><td>Zoom Out UI</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>0</kbd></td><td>Reset UI to 100%</td></tr>
                <tr><td><kbd>F1</kbd></td><td>Open Help</td></tr>
            </table>

            <h3>Word</h3>
            <table>
                <tr><th>Key</th><th>Action</th></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>B</kbd></td><td>Bold</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>I</kbd></td><td>Italic</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>U</kbd></td><td>Underline</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>H</kbd></td><td>Find & Replace</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>P</kbd></td><td>Print</td></tr>
            </table>

            <h3>Excel</h3>
            <table>
                <tr><th>Key</th><th>Action</th></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>X</kbd></td><td>Cut</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>C</kbd></td><td>Copy</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>V</kbd></td><td>Paste</td></tr>
            </table>

            <div class="tip">Use <kbd>Win</kbd>+<kbd>H</kbd> for Windows voice input. It also works well for AI chat conversations.</div>

            <hr class="section-divider"/>

            <h2 id="license">License</h2>
            <table>
                <tr><th>Plan</th><th>Description</th><th>Duration</th></tr>
                <tr><td><span class="badge badge-free">FREE</span></td><td>All features available (export limited)</td><td>Unlimited</td></tr>
                <tr><td><span class="badge badge-trial">TRIAL</span></td><td>All features available (evaluation)</td><td>30 days</td></tr>
                <tr><td><span class="badge badge-biz">BIZ</span></td><td>Full business features</td><td>365 days</td></tr>
                <tr><td><span class="badge badge-ent">ENT</span></td><td>Full features + API/SSO/Audit Log</td><td>Negotiable</td></tr>
            </table>

            <h3>License Key Format</h3>
            <p><code>IAOF-[Plan]-[YYMM]-[HASH]-[SIG1]-[SIG2]</code></p>

            <h3>License Management</h3>
            <p>Click the plan badge in the title bar to open the license management screen. Enter, verify, or change license keys.</p>
            <p>You can also check your current plan and expiry date in the backstage "License" tab.</p>

            <hr class="section-divider"/>

            <h2 id="system-req">System Requirements</h2>
            <table>
                <tr><th>Item</th><th>Requirement</th></tr>
                <tr><td>OS</td><td>Windows 10 (64-bit) / Windows 11</td></tr>
                <tr><td>Runtime</td><td>.NET 8.0 Desktop Runtime</td></tr>
                <tr><td>Memory</td><td>4 GB minimum (8 GB recommended)</td></tr>
                <tr><td>Disk</td><td>500 MB free space</td></tr>
                <tr><td>Internet</td><td>Required for AI features</td></tr>
            </table>

            <hr class="section-divider"/>

            <h2 id="support">Support</h2>
            <table>
                <tr><th>Item</th><th>Details</th></tr>
                <tr><td>Developer</td><td>HARMONIC insight</td></tr>
                <tr><td>Product</td><td>Insight AI Office (IAOF)</td></tr>
                <tr><td>Email</td><td><a href="mailto:support@h-insight.jp">support@h-insight.jp</a></td></tr>
            </table>
            <p style="margin-top:24px;font-size:12px;color:#A8A29E;">Copyright &copy; 2025-2026 HARMONIC insight. All rights reserved.</p>
            """);
    }
}
