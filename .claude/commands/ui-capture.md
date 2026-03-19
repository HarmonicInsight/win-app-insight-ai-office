# UI キャプチャ & レビューコマンド

WPF デスクトップアプリのスクリーンショットを撮影し、UI レビューを自動実行します。

## 動作モード

| 優先度 | モード | 条件 | 動作 |
|:------:|--------|------|------|
| 1 | **Smart（デフォルト）** | アプリが起動中 | プロジェクト名でウィンドウを自動検出→キャプチャ |
| 2 | **Launch** | アプリが未起動 | ビルド→起動→ウィンドウ待機→キャプチャ |
| 3 | **WindowHint** | `-WindowHint` 指定 | 指定名でウィンドウ検索→キャプチャ |

> **重要**: ターミナル/CLI ウィンドウは自動除外される（PowerShell, cmd, Terminal, VSCode 等）。
> Claude Code の実行ウィンドウがキャプチャされる心配はない。

## 実行手順

### Step 1: 対象プロジェクトの特定

1. `$ARGUMENTS` が指定されている場合はそのディレクトリを対象にする
2. 未指定の場合はカレントディレクトリを対象にする
3. 対象プロジェクトの製品コード（INSS/IOSH/IOSD/INMV 等）を特定する
4. カラーテーマ（Ivory & Gold / Cool Blue & Slate）を CLAUDE.md の製品定義から判定する

### Step 2: スクリーンショット撮影

**基本コマンド（Smart モード — アプリ起動中ならそのままキャプチャ、未起動ならビルド&起動）:**

```powershell
powershell -ExecutionPolicy Bypass -File "C:/dev/cross-lib-insight-common/tools/ScreenCapture/auto-capture.ps1" -ProjectDir "${ARGUMENTS:-.}"
```

**オプション:**

```powershell
# ウィンドウ名を明示指定
powershell -ExecutionPolicy Bypass -File "C:/dev/cross-lib-insight-common/tools/ScreenCapture/auto-capture.ps1" -ProjectDir "${ARGUMENTS:-.}" -WindowHint "Insight"

# 強制的にビルド&起動（アプリが起動していても新しく起動）
powershell -ExecutionPolicy Bypass -File "C:/dev/cross-lib-insight-common/tools/ScreenCapture/auto-capture.ps1" -ProjectDir "${ARGUMENTS:-.}" -Launch -KeepRunning

# ディレイ付き（メニューやダイアログを開いてからキャプチャ）
powershell -ExecutionPolicy Bypass -File "C:/dev/cross-lib-insight-common/tools/ScreenCapture/auto-capture.ps1" -ProjectDir "${ARGUMENTS:-.}" -Delay 3
```

### Step 3: スクリーンショット読み取り

出力の最終行 `SCREENSHOT_PATH=<path>` からパスを取得し、Read ツールで画像を読み取る。

### Step 4: UI レビュー（自動チェック）

画像を分析し、以下の観点でレビューする。問題があれば具体的な修正案を提示する。

#### 4-1. カラー標準

**Ivory & Gold テーマ（INSS/IOSH/IOSD/INMV/INIG/INPY/ISOF/IAOF/INAG）:**
- プライマリカラーが Gold (#B8942F) であること
- 背景色が Ivory (#FAF8F5) であること
- Blue (#2563EB) がプライマリとして使われていないこと

**Cool Blue & Slate テーマ（INBT/INCA/IVIN）:**
- プライマリカラーが Blue (#2563EB) であること
- 背景色が Slate (#F8FAFC) であること
- Gold (#B8942F) がプライマリとして使われていないこと

#### 4-2. レイアウト標準

- 余白（Margin/Padding）が均等であること
- 要素の整列（左揃え・中央揃え）が一貫していること
- テキストの切れ・はみ出しがないこと
- ボタンやコントロールのサイズが統一されていること
- スクロールバーが不要な場所に表示されていないこと

#### 4-3. Ribbon / ツールバー標準

- Ribbon が Syncfusion SfSkinManager のテーマに従っていること
- タイトルバーのフォーマット: `InsightOffice` + 製品名 + バージョン + プランバッジ
- BackStage に必須コマンド（新規作成・開く・上書き保存・名前を付けて保存・印刷・閉じる）があること

#### 4-4. AI パネル標準（AI 搭載製品のみ）

- AI パネルが右サイドパネルに配置されていること
- パネルヘッダーが「AI コンシェルジュ」であること（「AI アシスタント」は NG）
- AI パネルがデフォルトで表示されていること

#### 4-5. フォント・テキスト標準

- 日本語フォントが読みやすいこと（メイリオ / Yu Gothic UI 推奨）
- フォントサイズが小さすぎないこと（本文 12px 以上推奨）
- テキストのコントラスト比が十分であること

#### 4-6. アクセシビリティ

- ステータスバーにスケール倍率が表示されていること
- ToolTip がアイコンボタンに設定されているか（見た目で判断可能な範囲）

### Step 5: レビュー結果の報告

チェックリスト形式で結果を報告する:

```
## UI レビュー結果

**製品**: INSS (Insight Deck Quality Gate)
**テーマ**: Ivory & Gold
**スクリーンショット**: <path>

### チェック結果
- [x] カラー標準: Gold プライマリ ✅
- [x] 背景色: Ivory ✅
- [ ] レイアウト: ⚠️ 左パネルの余白が不均等
- [x] Ribbon 標準 ✅
- [x] AI パネル配置 ✅

### 修正提案
1. **左パネルの Margin を統一**: `Margin="12"` → `Margin="12,8,12,8"`
   - ファイル: `Views/MainWindow.xaml` L.45
```

修正が必要な場合は、ユーザーに確認の上、コード修正 → 再キャプチャ → 再レビューのサイクルを提案する。

## 注意事項

- ビルドエラーが発生した場合は `scripts/build-doctor.sh` の実行を提案する
- キャプチャ画像は `.gitignore` に `ui-capture_*.png` / `screenshot_*.png` を追加推奨
- アプリが起動しない場合は `-RunningOnly` オプションでユーザーに手動起動を依頼する
- ダイアログやメニューの確認が必要な場合は `screenshot <WindowTitle> 3` の手動コマンドを案内する
