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
            "overview", "ui-layout", "file-ops", "ai-assistant",
            "shortcuts", "license", "system-req", "support"
        ];

        _sectionNames = isEn
            ?
            [
                "Overview", "UI Layout", "File Operations", "AI Assistant",
                "Keyboard Shortcuts", "License", "System Requirements", "Support"
            ]
            :
            [
                "はじめに", "画面構成", "ファイル操作", "AIアシスタント",
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
            HelpBrowser.InvokeScript("eval",
                $"document.getElementById('{sectionId}').scrollIntoView({{behavior:'smooth',block:'start'}})");
        }
        catch (InvalidOperationException) { /* WebBrowser not ready for script execution */ }
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
        var sb = new StringBuilder(32000);
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
        </style></head>
        """;

    private void AppendJapaneseContent(StringBuilder sb)
    {
        var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var verStr = $"{ver?.Major}.{ver?.Minor}.{ver?.Build}";

        sb.Append($"""
            <div class="hero" id="overview">
                <h1>Insight AI Office</h1>
                <p>Word / Excel / PowerPoint を AI で分析する統合ツール</p>
                <p style="font-size:12px;color:#A8A29E;">バージョン {verStr}</p>
            </div>

            <div class="feature-grid">
                <div class="feature-card"><h3>Word 編集</h3><p>Word (.docx/.doc) ファイルの表示・編集・保存。書式設定、検索・置換に対応。</p></div>
                <div class="feature-card"><h3>Excel 編集</h3><p>Excel (.xlsx/.xls/.csv) ファイルの表示・編集・保存。数式、書式、セル操作に対応。</p></div>
                <div class="feature-card"><h3>PowerPoint 表示</h3><p>PowerPoint (.pptx/.ppt) のスライドサムネイル表示とテキスト抽出。</p></div>
                <div class="feature-card"><h3>AI アシスタント</h3><p>Claude / OpenAI / Gemini 対応の AI チャット。要約・校正・分析をワンクリックで。</p></div>
            </div>

            <hr class="section-divider"/>

            <h2 id="ui-layout">画面構成</h2>
            <h3>タイトルバー</h3>
            <p>製品名、バージョン、ライセンスプラン、ファイル名を表示。右側に AI / プロンプト / 参考資料の切替ボタン。</p>

            <h3>リボン</h3>
            <p>開いているファイル形式に応じてリボンが自動で切り替わります：</p>
            <ul>
                <li><strong>デフォルト</strong> — ファイル未オープン時。「開く」「AI チャット」「AI 設定」</li>
                <li><strong>Word</strong> — フォント書式、段落配置、検索・置換、エクスポート</li>
                <li><strong>Excel</strong> — クリップボード、フォント、配置、数値書式</li>
                <li><strong>PowerPoint</strong> — テキスト抽出</li>
            </ul>

            <h3>3 ペインレイアウト</h3>
            <ul>
                <li><strong>左パネル</strong> — 参考資料。ファイルをドラッグ&ドロップまたは「＋」で追加。AI が自動的にコンテキストとして使用。</li>
                <li><strong>中央</strong> — ドキュメントエディタ / ビューア</li>
                <li><strong>右パネル</strong> — AI コンシェルジュ（チャット）</li>
            </ul>

            <hr class="section-divider"/>

            <h2 id="file-ops">ファイル操作</h2>
            <h3>対応フォーマット</h3>
            <table>
                <tr><th>形式</th><th>拡張子</th><th>操作</th></tr>
                <tr><td>Word</td><td>.docx, .doc</td><td>表示・編集・保存</td></tr>
                <tr><td>Excel</td><td>.xlsx, .xls, .csv</td><td>表示・編集・保存</td></tr>
                <tr><td>PowerPoint</td><td>.pptx, .ppt</td><td>表示・テキスト抽出</td></tr>
            </table>

            <h3>ファイルを開く</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">リボンの「開く」ボタンをクリック、または <kbd>Ctrl</kbd>+<kbd>O</kbd></div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">ファイル選択ダイアログからファイルを選択</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">ファイル形式に応じたエディタ / ビューアが自動で開きます</div></div>

            <div class="tip">ファイルをウィンドウにドラッグ&ドロップしても開けます。</div>

            <hr class="section-divider"/>

            <h2 id="ai-assistant">AI アシスタント</h2>
            <h3>セットアップ</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">リボンの「AI 設定」をクリック</div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">プロバイダー（Claude / OpenAI / Gemini）を選択し、API キーを入力</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">モデルを選択して「保存」</div></div>

            <div class="note">API キーはお客様ご自身で取得してください（BYOK 方式）。キーは端末内に暗号化保存され、外部に送信されることはありません。</div>

            <h3>ワンクリック分析</h3>
            <table>
                <tr><th>ボタン</th><th>機能</th></tr>
                <tr><td>要約</td><td>ドキュメント全体の要約を生成</td></tr>
                <tr><td>校正</td><td>誤字脱字・文法チェック・改善提案</td></tr>
                <tr><td>分析</td><td>文書構造・内容を多角的に分析</td></tr>
                <tr><td>数式ヘルプ (Excel)</td><td>Excel 数式・関数の提案</td></tr>
            </table>

            <h3>プロンプトプリセット</h3>
            <p>内蔵 8 種類のプリセット（要約・分析・校正・英訳・和訳・数式ヘルプ・データインサイト・議事録作成）に加え、カスタムプリセットを作成・管理できます。</p>

            <h3>参考資料</h3>
            <p>左パネルに参考資料を追加すると、AI が自動的にコンテキストとして活用します。Word / Excel / PowerPoint / テキストファイルに対応。</p>

            <hr class="section-divider"/>

            <h2 id="shortcuts">キーボードショートカット</h2>
            <table>
                <tr><th>キー</th><th>操作</th></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>N</kbd></td><td>新規作成</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>O</kbd></td><td>ファイルを開く</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>S</kbd></td><td>上書き保存</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>P</kbd></td><td>印刷</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>+</kbd></td><td>UI 拡大</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>-</kbd></td><td>UI 縮小</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>0</kbd></td><td>UI 100% にリセット</td></tr>
                <tr><td><kbd>F1</kbd></td><td>ヘルプを開く</td></tr>
                <tr><td><kbd>Win</kbd>+<kbd>H</kbd></td><td>音声入力（Windows 標準）</td></tr>
            </table>

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
            <p>タイトルバーのプランバッジをクリックするとライセンス管理画面が開きます。</p>

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
                <tr><td>メール</td><td><a href="mailto:support@harmonic-insight.com">support@harmonic-insight.com</a></td></tr>
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
                <p>An integrated tool for analyzing Word / Excel / PowerPoint with AI</p>
                <p style="font-size:12px;color:#A8A29E;">Version {verStr}</p>
            </div>

            <div class="feature-grid">
                <div class="feature-card"><h3>Word Editing</h3><p>View, edit, and save Word (.docx/.doc) files. Supports formatting, find & replace.</p></div>
                <div class="feature-card"><h3>Excel Editing</h3><p>View, edit, and save Excel (.xlsx/.xls/.csv) files. Supports formulas, formatting, cell operations.</p></div>
                <div class="feature-card"><h3>PowerPoint Viewer</h3><p>View PowerPoint (.pptx/.ppt) slide thumbnails and extract text content.</p></div>
                <div class="feature-card"><h3>AI Assistant</h3><p>AI chat supporting Claude / OpenAI / Gemini. One-click summarize, proofread, and analyze.</p></div>
            </div>

            <hr class="section-divider"/>

            <h2 id="ui-layout">UI Layout</h2>
            <h3>Title Bar</h3>
            <p>Displays product name, version, license plan, and file name. Toggle buttons for AI / Prompts / References on the right.</p>

            <h3>Ribbon</h3>
            <p>The ribbon automatically switches based on the open file format:</p>
            <ul>
                <li><strong>Default</strong> — No file open. "Open", "AI Chat", "AI Settings"</li>
                <li><strong>Word</strong> — Font formatting, paragraph alignment, find & replace, export</li>
                <li><strong>Excel</strong> — Clipboard, font, alignment, number formatting</li>
                <li><strong>PowerPoint</strong> — Text extraction</li>
            </ul>

            <h3>3-Pane Layout</h3>
            <ul>
                <li><strong>Left Panel</strong> — Reference materials. Drag & drop or click "+" to add. AI automatically uses them as context.</li>
                <li><strong>Center</strong> — Document editor / viewer</li>
                <li><strong>Right Panel</strong> — AI Concierge (chat)</li>
            </ul>

            <hr class="section-divider"/>

            <h2 id="file-ops">File Operations</h2>
            <h3>Supported Formats</h3>
            <table>
                <tr><th>Format</th><th>Extensions</th><th>Operations</th></tr>
                <tr><td>Word</td><td>.docx, .doc</td><td>View, Edit, Save</td></tr>
                <tr><td>Excel</td><td>.xlsx, .xls, .csv</td><td>View, Edit, Save</td></tr>
                <tr><td>PowerPoint</td><td>.pptx, .ppt</td><td>View, Text Extraction</td></tr>
            </table>

            <h3>Opening Files</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">Click "Open" in the ribbon, or press <kbd>Ctrl</kbd>+<kbd>O</kbd></div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">Select a file from the dialog</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">The appropriate editor/viewer opens automatically</div></div>

            <div class="tip">You can also drag & drop files onto the window to open them.</div>

            <hr class="section-divider"/>

            <h2 id="ai-assistant">AI Assistant</h2>
            <h3>Setup</h3>
            <div class="step"><div class="step-num">1</div><div class="step-text">Click "AI Settings" in the ribbon</div></div>
            <div class="step"><div class="step-num">2</div><div class="step-text">Select a provider (Claude / OpenAI / Gemini) and enter your API key</div></div>
            <div class="step"><div class="step-num">3</div><div class="step-text">Choose a model and click "Save"</div></div>

            <div class="note">You need to obtain your own API key (BYOK model). Keys are stored encrypted on your device and never sent externally.</div>

            <h3>One-Click Analysis</h3>
            <table>
                <tr><th>Button</th><th>Function</th></tr>
                <tr><td>Summarize</td><td>Generate a summary of the entire document</td></tr>
                <tr><td>Proofread</td><td>Check for typos, grammar, and suggest improvements</td></tr>
                <tr><td>Analyze</td><td>Multi-faceted analysis of document structure and content</td></tr>
                <tr><td>Formula Help (Excel)</td><td>Suggest Excel formulas and functions</td></tr>
            </table>

            <h3>Prompt Presets</h3>
            <p>8 built-in presets (Summarize, Analyze, Proofread, Translate EN/JA, Formula Help, Data Insight, Meeting Minutes) plus custom preset creation and management.</p>

            <h3>Reference Materials</h3>
            <p>Add reference materials to the left panel, and AI will automatically use them as context. Supports Word / Excel / PowerPoint / text files.</p>

            <hr class="section-divider"/>

            <h2 id="shortcuts">Keyboard Shortcuts</h2>
            <table>
                <tr><th>Key</th><th>Action</th></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>N</kbd></td><td>New Document</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>O</kbd></td><td>Open File</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>S</kbd></td><td>Save</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>P</kbd></td><td>Print</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>+</kbd></td><td>Zoom In UI</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>-</kbd></td><td>Zoom Out UI</td></tr>
                <tr><td><kbd>Ctrl</kbd>+<kbd>0</kbd></td><td>Reset UI to 100%</td></tr>
                <tr><td><kbd>F1</kbd></td><td>Open Help</td></tr>
                <tr><td><kbd>Win</kbd>+<kbd>H</kbd></td><td>Voice Input (Windows built-in)</td></tr>
            </table>

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
            <p>Click the plan badge in the title bar to open the license management screen.</p>

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
                <tr><td>Email</td><td><a href="mailto:support@harmonic-insight.com">support@harmonic-insight.com</a></td></tr>
            </table>
            <p style="margin-top:24px;font-size:12px;color:#A8A29E;">Copyright &copy; 2025-2026 HARMONIC insight. All rights reserved.</p>
            """);
    }
}
