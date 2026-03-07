# IAOF (Insight AI Office) 製品評価レポート

**評価日**: 2026-03-07 (最終更新)
**対象バージョン**: v1.0.0
**価格**: 49,800 円/端末・年（税抜）

---

## 1. 現在の実装状況

### 実装済み機能

| 機能 | 状態 | 品質 |
|------|:----:|:----:|
| Word (.docx/.doc) 編集 | 完了 | B |
| Excel (.xlsx/.xls/.csv) 編集 | 完了 | B |
| PowerPoint (.pptx/.ppt) スライド表示 | 完了 | A |
| ファイル形式別リボン自動切替 (4種) | 完了 | A |
| AI チャット (Claude/OpenAI/Gemini) | 完了 | B |
| AI ワンクリック分析 (要約/校正/分析) | 完了 | B |
| AI → ドキュメント書き戻し | 完了 | B |
| AI 文書生成 (AI Generate Doc) | 完了 | B |
| AI レポート生成 (Excel → Word) | 完了 | B |
| クロスフォーマット AI 分析 | 完了 | B |
| DocumentToolExecutor (AI ツール実行) | 完了 | B |
| プロンプトプリセット管理 (8個内蔵) | 完了 | A |
| 参考資料パネル (左ペイン) | 完了 | B |
| ライセンスシステム (共通) | 完了 | A |
| カスタムタイトルバー | 完了 | A |
| ドラッグ & ドロップ | 完了 | B |
| Word エクスポート (.docx) | 完了 | B |
| Syncfusion 初期化共通化 | 完了 | A |
| PPTX レンダリング共通化 | 完了 | A |
| HelpWindow (F1 対応) | 完了 | A |
| 設定画面 (言語切替) | 完了 | B |
| InsightScaleManager 統合 (Ctrl+/-) | 完了 | A |
| MainWindow 分割 (5 partial class) | 完了 | B |
| AI サービス DI 化 | 完了 | A |
| ローカライゼーション (JA/EN) | 完了 | B |
| プロジェクトファイル (.iaof) | 完了 | B |
| コンテンツパック (デモ用) | 完了 | B |
| テスト (17件) | 完了 | B |

### 変更履歴 (2026-03-07)

| 変更 | 理由 |
|------|------|
| PDF 対応を全面削除 | Syncfusion PDF テキスト抽出が未実装で AI 連携不可。PDF は参考資料としてのみ利用可能 |
| ウェルカム画面を削除 | BtoB 製品のためウェルカム画面は不要。起動後すぐにリボンからファイルを開く |
| 対応フォーマット: 4→3 | Word / Excel / PowerPoint の 3 フォーマットに集中 |
| MainWindow 分割実施 | 1200行→5ファイル (xaml.cs/AI/Document/Ribbon/UI) に分割 |
| AI サービス DI 化完了 | ServiceConfiguration で AiService/PromptPresetService/ReferenceMaterialsService を登録 |
| ローカライゼーション実装 | LanguageManager (80キー JA/EN) + ApplyLocalization() + Ribbon 全ラベル動的設定 |
| テスト整備 | LanguageManager (7件) + DocumentToolExecutor (4件) + ProjectArchive (4件) + PromptEntry (2件) |
| DocumentToolExecutor 正規表現修正 | ネストされた JSON ブロック（args内オブジェクト）のパース修正 |

### 残タスク

| 機能 | 重要度 | 備考 |
|------|:------:|------|
| 隠しコマンド (7タップ) | 低 | ライセンス画面に未実装 |
| バージョン管理 | 低 | IOSH/IOSD にはある機能 |
| AI コンシェルジュペルソナ | 低 | AiConciergeConfig に IAOF 定義はある |
| PPTX 編集機能 | 中 | 現在は閲覧のみ（INSS で編集可能） |
| DPAPI 暗号化保存 | 中 | API キー・ライセンスの暗号化保存 |

---

## 2. マーケター目線の評価

### 製品ポジショニング

**企画意図**: Word/Excel/PowerPoint を1つのアプリで AI 分析できる統合ツール
**競合**: Microsoft 365 Copilot, Google Workspace + Gemini, Notion AI

### 強み (Strengths)

1. **3フォーマット統合**: Word/Excel/PPTX を切替不要で扱える — 競合にない独自価値
2. **BYOK モデル**: クライアントの API キーを使うため、月額 AI 課金が不要
3. **マルチプロバイダー**: Claude/OpenAI/Gemini 全対応 — ベンダーロックインなし
4. **参考資料付き AI**: 参照ドキュメントをコンテキストとして AI に渡せる
5. **プロンプトプリセット**: 業務特化プロンプトの保存・配信が可能
6. **AI → ドキュメント書き戻し**: AI 提案を1クリックでドキュメントに反映
7. **クロスフォーマット分析**: 複数フォーマットの資料を横断的に AI で分析
8. **AI 文書・レポート生成**: チャットから Word ドキュメント/レポートを自動生成

### 弱み (Weaknesses) — 残課題

1. **PPTX は閲覧のみ**: スライド編集ができない（INSS で編集可能なのに）
2. **デモ時の "Wow Factor" が限定的**: AI 文書生成はあるが、テンプレート不足

### 改善案 (マーケティング観点) — 対応状況

| # | 施策 | 状態 | 備考 |
|---|------|:----:|------|
| M1 | AI ドキュメント生成 | ✅ 完了 | チャット経由で Word 文書を自動生成 |
| M2 | AI → ドキュメント書き戻し | ✅ 完了 | 1クリックで AI 応答をドキュメントに挿入 |
| M3 | クロスフォーマット AI 分析 | ✅ 完了 | 複数ファイル + 参考資料を横断分析 |
| M4 | AI レポート生成 (Excel → Word) | ✅ 完了 | Excel データから Word レポートを AI 生成 |

---

## 3. エンジニア目線の評価

### アーキテクチャ

| 項目 | 評価 | コメント |
|------|:----:|---------|
| プロジェクト構造 | B | App/Core/Data 分離。Core/Data はインターフェース中心 |
| DI 設計 | A | ServiceConfiguration で全サービスを DI 登録済み |
| MVVM 遵守度 | C | MainWindow 5分割で改善。ViewModel は Commands 中心 |
| 共通化 | A | Syncfusion 初期化/PPTX レンダリング/ライセンスは共通化済み |
| テスト | B | 17テスト（LanguageManager/DocumentToolExecutor/ProjectArchive/PromptEntry） |
| エラーハンドリング | B | App レベルの3層例外ハンドラーあり |
| セキュリティ | B | BYOK。DPAPI は未実装 |
| ローカライゼーション | B | LanguageManager (80キー JA/EN)。Ribbon 全ラベル動的設定 |

### コード品質 — 改善実施済み

| 問題 | 対応状況 |
|------|---------|
| ~~God Object: MainWindow.xaml.cs 1200行以上~~ | ✅ 5 partial class に分割 (xaml.cs/AI/Document/Ribbon/UI) |
| ~~AI サービスが DI されていない~~ | ✅ ServiceConfiguration で AiService/PromptPresetService/ReferenceMaterialsService を登録 |
| ~~テストがない~~ | ✅ 17テスト追加 |
| ~~ローカライゼーション未対応~~ | ✅ LanguageManager + ApplyLocalization + Ribbon 動的ラベル |
| ~~Core のスタブ実装残存~~ | ✅ インターフェースのみに整理 |
| ViewModel が Commands のみ | 残。段階的に機能移行 |

### 改善案 (エンジニアリング観点) — 対応状況

| # | 施策 | 状態 | 備考 |
|---|------|:----:|------|
| E1 | MainWindow 分割 | ✅ 完了 | 5 partial class に分割 |
| E2 | AI サービス DI 化 | ✅ 完了 | ServiceConfiguration 登録済み |
| E3 | HelpWindow 実装 | ✅ 完了 | F1 対応、ShowDialog() |
| E4 | DocumentToolExecutor | ✅ 完了 | insert_text/replace_selection 対応 |
| E5 | InsightScaleManager 統合 | ✅ 完了 | Ctrl+/-/0 対応 |
| E6 | AutomationProperties 追加 | ✅ 完了 | ApplyLocalization() で動的設定 |

---

## 4. 実装ロードマップ — 完了状況

### Phase 1: 最低限リリース品質 — ✅ 全完了

| # | 施策 | 状態 |
|---|------|:----:|
| 1 | HelpWindow 実装 | ✅ |
| 2 | AI サービス DI 化 | ✅ |
| 3 | InsightScaleManager 統合 | ✅ |
| 4 | AutomationProperties 追加 | ✅ |
| 5 | ローカライゼーション（日英） | ✅ |

### Phase 2: 差別化・競争力 — ✅ 全完了

| # | 施策 | 状態 |
|---|------|:----:|
| 6 | AI → ドキュメント書き戻し | ✅ |
| 7 | クロスフォーマット AI 分析 | ✅ |
| 8 | 設定画面 (言語切替) | ✅ |
| 9 | MainWindow 分割リファクタ | ✅ |

### Phase 3: Wow Factor — ✅ 全完了

| # | 施策 | 状態 |
|---|------|:----:|
| 10 | AI ドキュメント自動生成 | ✅ |
| 11 | AI レポート生成 (Excel → Word) | ✅ |
| 12 | DocumentToolExecutor | ✅ |
| 13 | プロジェクトファイル (.iaof) | ✅ |
| 14 | コンテンツパック統合 | ✅ |

---

## 5. 共通化実施記録 (2026-03-07)

### 新規作成 (InsightCommon)

| ファイル | 役割 |
|---------|------|
| `Theme/SyncfusionInitializer.cs` | Syncfusion ライセンス登録 + テーマ適用ヘルパー |
| `Services/PresentationRenderingService.cs` | PPTX スライド画像レンダリング共通サービス |

### IAOF 新規ファイル

| ファイル | 役割 |
|---------|------|
| `MainWindow.AI.cs` | AI 機能 (チャット/ワンクリック/設定/書き戻し/横断分析/文書生成/レポート生成) |
| `MainWindow.Document.cs` | ドキュメント管理 (開く/閉じる/抽出/.iaof プロジェクト) |
| `MainWindow.Ribbon.cs` | リボン操作 (書式/段落/Excel/PPTX/エクスポート/バックステージ) |
| `MainWindow.UI.cs` | UI 操作 (参考資料/ウィンドウ/D&D/パネル/ライセンス) |
| `Services/DocumentToolExecutor.cs` | AI ツール呼び出し実行 (insert_text/replace_selection) |
| `Views/HelpWindow.xaml(.cs)` | ヘルプウィンドウ (F1 対応) |
| `Views/SettingsWindow.xaml(.cs)` | 設定画面 (言語切替) |
| `Helpers/LanguageManager.cs` | 多言語管理 (80キー JA/EN) |
| `Helpers/BuiltInPresets.cs` | 組み込みプロンプトプリセット (8個) |
| `assets/content-packs/demo-iaof/` | デモ用コンテンツパック |

### 修正ファイル

| ファイル | 変更内容 |
|---------|---------|
| `InsightCommon.csproj` | SfSkinManager.WPF + System.Drawing.Common 追加、NoWarn 統一 |
| `License/InsightLicenseManager.cs` | 製品コード正規表現に IAOF 追加 |
| `config/third-party-licenses.json` | usedBy + components に IAOF 追加（PdfViewer 削除） |
| IAOF `App.xaml.cs` | SyncfusionInitializer.Initialize() に切替 |
| IAOF `MainWindow.xaml.cs` | 5分割 + ApplyLocalization() + ApplyRibbonLocalization() |
| IAOF `MainWindow.xaml` | 全ラベル x:Name 化、DataTemplate に x:Static バインディング |
| IAOF `ServiceConfiguration.cs` | AiService/PromptPresetService/ReferenceMaterialsService DI 登録 |
| IAOF `InsightAiOffice.App.csproj` | Syncfusion.PdfViewer.WPF 削除、コンテンツパックアセット追加 |
| IAOF `ViewModels/MainViewModel.cs` | .iaof フィルタ追加 |
| IAOF `Data/ProjectArchiveAdapter.cs` | ZIP プロジェクトファイル完全実装 |
| IAOF `Core/Services/*.cs` | スタブ実装削除、インターフェースのみに整理 |

### 1箇所修正で全製品に反映される項目

| 変更シナリオ | 修正ファイル |
|------------|-------------|
| Syncfusion ライセンスキー更新 | third-party-licenses.json のみ |
| Syncfusion テーマ変更 | SyncfusionInitializer.cs のみ |
| PPTX レンダリングバグ修正 | PresentationRenderingService.cs のみ |
| 新製品コード追加 | InsightLicenseManager.cs の正規表現のみ |

---

## 6. テストカバレッジ

| テストファイル | テスト数 | 対象 |
|--------------|:-------:|------|
| LanguageManagerTests.cs | 7 | JA/EN 切替、フォールバック、Format、静的プロパティ、キー整合性 |
| DocumentToolExecutorTests.cs | 4 | ツール呼び出しなし/insert_text/replace_selection/複数実行 |
| ProjectArchiveAdapterTests.cs | 4 | ZIP 作成/読込/ドキュメントタイプ/保存タイムスタンプ更新 |
| PromptServiceTests.cs | 2 | PromptEntry レコード プロパティ/等値性 |
| **合計** | **17** | |
