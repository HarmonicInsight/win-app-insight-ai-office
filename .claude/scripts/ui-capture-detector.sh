#!/bin/bash
# =============================================================
# HARMONIC insight - UI 自動検知フック（UserPromptSubmit）
# =============================================================
#
# ユーザーのメッセージにキーワードが含まれている場合、
# 該当するスキルの実行指示を Claude のコンテキストに注入する。
#
# 対応スキル:
#   1. /ui-capture  — 画面確認・レイアウトチェック
#   2. /audit-ribbon — Ribbon メニュー ↔ C# 実装の整合性チェック
#

INPUT=$(cat)
PROMPT=$(echo "$INPUT" | jq -r '.prompt // ""' 2>/dev/null)

if [ -z "$PROMPT" ]; then
  exit 0
fi

# ---------------------------------------------------------------
# Ribbon 監査キーワード（先にチェック — より具体的）
# ---------------------------------------------------------------
RIBBON_KEYWORDS=(
  "ボタンが動かない"
  "ボタン動かない"
  "メニュー動かない"
  "メニューが動かない"
  "ハンドラ"
  "結合セル"
  "Ribbon監査"
  "リボン監査"
  "ボタン紐づけ"
  "ボタンが何もしない"
  "クリックしても反応しない"
  "未実装ボタン"
  "audit-ribbon"
  "ボタンが反応しない"
  "メニューが反応しない"
  "動かないボタン"
  "紐付け確認"
  "ボタン一覧"
  "メニュー一覧"
)

for keyword in "${RIBBON_KEYWORDS[@]}"; do
  if echo "$PROMPT" | grep -qi "$keyword" 2>/dev/null; then
    cat <<'HOOK_OUTPUT'
<user-prompt-submit-hook>
IMPORTANT: ユーザーが Ribbon メニューの動作問題を報告しています。以下の手順で /audit-ribbon スキルを実行してください:

1. Skill ツールで skill: "audit-ribbon" を実行する
2. スキルの指示に従い、audit-ribbon.ps1 で Ribbon 全ボタンの監査を実行する
3. レポートを Read ツールで読み取り、問題を分析する

audit-ribbon.ps1 のパス: C:/dev/cross-lib-insight-common/tools/RibbonAuditor/audit-ribbon.ps1

特に以下を重点チェック:
- MISSING: ハンドラ/コマンドが C# に存在しない
- EMPTY/STUB: ハンドラが空、または NotImplementedException
- Symmetry: 対になる操作（結合↔解除、元に戻す↔やり直し等）の欠如
- NO_BINDING: Command も Click も設定されていないボタン
</user-prompt-submit-hook>
HOOK_OUTPUT
    exit 0
  fi
done

# ---------------------------------------------------------------
# UI キャプチャキーワード
# ---------------------------------------------------------------
CAPTURE_KEYWORDS=(
  "画面確認"
  "UIチェック"
  "UI チェック"
  "レイアウト確認"
  "見た目確認"
  "表示確認"
  "画面撮って"
  "スクショ"
  "スクリーンショット"
  "screenshot"
  "レイアウト崩れ"
  "余白がおかしい"
  "配置がずれ"
  "UIがおかしい"
  "UI がおかしい"
  "デザイン確認"
  "画面チェック"
  "ui-capture"
)

for keyword in "${CAPTURE_KEYWORDS[@]}"; do
  if echo "$PROMPT" | grep -qi "$keyword" 2>/dev/null; then
    cat <<'HOOK_OUTPUT'
<user-prompt-submit-hook>
IMPORTANT: ユーザーが UI 確認を要求しています。以下の手順で /ui-capture スキルを実行してください:

1. Skill ツールで skill: "ui-capture" を実行する
2. スキルの指示に従い、auto-capture.ps1 でスクリーンショットを撮影する
3. 撮影した画像を Read ツールで読み取り、UI レビューを行う

auto-capture.ps1 のパス: C:/dev/cross-lib-insight-common/tools/ScreenCapture/auto-capture.ps1

ユーザーのメッセージに含まれる具体的な UI の問題点（レイアウト、余白、配置等）にも注目して回答すること。
</user-prompt-submit-hook>
HOOK_OUTPUT
    exit 0
  fi
done

exit 0
