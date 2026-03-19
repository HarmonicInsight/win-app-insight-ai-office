# Ribbon メニュー監査コマンド

Syncfusion Ribbon の全ボタンを解析し、XAML のラベルと C# の実装の紐づけを監査します。

## 実行手順

### Step 1: 監査スクリプトを実行

```powershell
powershell -ExecutionPolicy Bypass -File "C:/dev/cross-lib-insight-common/tools/RibbonAuditor/audit-ribbon.ps1" -ProjectDir "${ARGUMENTS:-.}"
```

### Step 2: レポートを読み取る

出力の最終行 `AUDIT_REPORT_PATH=<path>` からパスを取得し、Read ツールでレポートを読み取る。

### Step 3: 問題の分析と修正提案

レポートの内容を分析し、以下の問題を特定・修正案を提示する:

#### 3-1. MISSING（ハンドラ/コマンド未実装）

XAML で Command や Click が指定されているが、C# に対応するメソッドやプロパティが存在しない。

**修正方法:**
- Command バインディングの場合: ViewModel に ICommand プロパティを追加
- Click ハンドラの場合: コードビハインド(.xaml.cs) にイベントハンドラを追加

#### 3-2. EMPTY / NOT_IMPL / STUB（空のハンドラ）

C# にメソッドは存在するが、処理本体が空、または NotImplementedException をスロー。
→ ボタンを押しても何も起きない状態。

**修正方法:**
- Syncfusion API を使った実装を追加
- または IsEnabled=false にして UI 上で無効化し、ToolTip で「今後対応予定」と表示

#### 3-3. NO_BINDING（バインディングなし）

XAML にボタンがあるが、Command も Click も設定されていない。
→ ボタンは表示されるが押しても何も起きない。

#### 3-4. Symmetry Issues（機能の対称性不足）

「結合」があるのに「解除」がない、「元に戻す」があるのに「やり直し」がない等。
ユーザーが操作を元に戻せない問題。

**修正方法:**
- 対になる機能のボタンを XAML に追加
- 対応する C# 実装を追加

### Step 4: 修正の優先度

| 優先度 | 問題 | 理由 |
|:------:|------|------|
| 1 | MISSING | ボタンがあるのにクラッシュまたは無反応 |
| 2 | EMPTY/NOT_IMPL | ボタンが無反応、ユーザーが混乱 |
| 3 | Symmetry | 操作を元に戻せない、基本操作が欠けている |
| 4 | NO_BINDING | 表示だけのボタン |
| 5 | TODO | 開発中 |

### Step 5: 修正の実施

ユーザーの確認後、以下の順で修正:
1. MISSING → ハンドラ/コマンドの実装を追加
2. EMPTY → Syncfusion API を使った実装を追加
3. Symmetry → 対になるボタン + 実装を追加
4. 修正後、再度 `/audit-ribbon` を実行して確認

## 注意事項

- レポートは `{ProjectDir}/ribbon-audit-report.md` に保存される
- ラベルが `{DynamicResource ...}` の場合、リソースファイルから実際の表示テキストを確認する
- Syncfusion の Ribbon API: https://help.syncfusion.com/wpf/ribbon/overview
- `.gitignore` に `ribbon-audit-report.md` を追加推奨
