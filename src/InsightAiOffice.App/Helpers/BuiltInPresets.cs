using InsightCommon.AI;

namespace InsightAiOffice.App.Helpers;

/// <summary>
/// IAOF ビルトインプリセット定義
/// - GetAll(): PromptPresetService 用（ユーザー保存形式、ディスク永続化）
/// - GetPresetPrompts(): AiChatViewModel 用（多言語プリセットツリー UI）
/// </summary>
public static class BuiltInPresets
{
    /// <summary>PromptPresetService 向け — ユーザープリセットとしてディスクに永続化される</summary>
    public static List<UserPromptPreset> GetAll() =>
    [
        // ── 分析・要約 ──
        new() { Id = "builtin_summarize", Name = "ドキュメント要約", Category = "分析・要約", Icon = "📋",
            SystemPrompt = "以下のドキュメントの内容を要約してください。重要なポイントを箇条書きで示してください。",
            Description = "ドキュメント全体の要約を生成", RequiresContextData = true },
        new() { Id = "builtin_analyze", Name = "構造分析", Category = "分析・要約", Icon = "🔍",
            SystemPrompt = "以下のドキュメントの構造と内容を分析してください。文書の目的、主要なセクション、データの傾向、改善すべき点を挙げてください。",
            Description = "文書構造・内容を多角的に分析", RequiresContextData = true },
        new() { Id = "builtin_data_insight", Name = "データインサイト", Category = "分析・要約", Icon = "📈",
            SystemPrompt = "以下のスプレッドシートのデータを分析し、重要なインサイト（傾向、異常値、改善提案）を提示してください。数値の根拠を明示してください。",
            Description = "表データから傾向・異常値を検出", RequiresContextData = true },

        // ── 校正・チェック ──
        new() { Id = "builtin_proofread", Name = "文章校正", Category = "校正・チェック", Icon = "✏️",
            SystemPrompt = "以下のドキュメントを校正してください。誤字脱字、文法の誤り、不自然な表現を指摘し、修正案を示してください。",
            Description = "誤字脱字・文法チェック・改善提案", RequiresContextData = true },

        // ── 翻訳 ──
        new() { Id = "builtin_translate_en", Name = "英訳", Category = "翻訳", Icon = "🌐",
            SystemPrompt = "以下のドキュメントの内容を自然なビジネス英語に翻訳してください。原文の意図を保ちながら、英語圏のビジネス文書として適切な表現にしてください。",
            Description = "ドキュメントを英語に翻訳", RequiresContextData = true },
        new() { Id = "builtin_translate_ja", Name = "和訳", Category = "翻訳", Icon = "🇯🇵",
            SystemPrompt = "以下のドキュメントの内容を自然な日本語に翻訳してください。原文の意図を保ちながら、日本語のビジネス文書として適切な表現にしてください。",
            Description = "ドキュメントを日本語に翻訳", RequiresContextData = true },

        // ── Excel ──
        new() { Id = "builtin_excel_formula", Name = "Excel 数式ヘルプ", Category = "Excel", Icon = "📊",
            SystemPrompt = "ユーザーが求める計算・集計を実現する Excel 数式を提案してください。関数名、引数、使用例を示し、注意点があれば補足してください。",
            Description = "Excel 数式・関数の提案", RequiresContextData = false },

        // ── 文書作成 ──
        new() { Id = "builtin_meeting_minutes", Name = "議事録作成", Category = "文書作成", Icon = "📝",
            SystemPrompt = "以下の内容から議事録を作成してください。日時、参加者、議題、決定事項、アクションアイテム（担当者・期限付き）の形式でまとめてください。",
            Description = "テキストから議事録フォーマットを生成", RequiresContextData = true },
    ];

    /// <summary>AiChatViewModel 向け — プリセットツリー UI（多言語対応）</summary>
    public static IReadOnlyList<PresetPrompt> GetPresetPrompts() =>
    [
        // ═══════════════════════════════════════════════════════════════
        // ■ 分析・要約（Analysis & Summary）
        // ═══════════════════════════════════════════════════════════════
        new PresetPrompt
        {
            Id = "preset_summarize",
            CategoryJa = "分析・要約", CategoryEn = "Analysis",
            LabelJa = "ドキュメント要約", LabelEn = "Summarize Document",
            PromptJa = "以下のドキュメントの内容を要約してください。\n\n【出力フォーマット】\n■ 概要（2〜3文）\n■ 重要ポイント（箇条書き5項目以内）\n■ 結論・次のアクション",
            PromptEn = "Summarize the following document.\n\nFormat:\n■ Overview (2-3 sentences)\n■ Key Points (up to 5 bullets)\n■ Conclusion & Next Steps",
            Icon = "📋", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_analyze",
            CategoryJa = "分析・要約", CategoryEn = "Analysis",
            LabelJa = "構造・内容分析", LabelEn = "Structure Analysis",
            PromptJa = "以下のドキュメントを多角的に分析してください。\n\n分析項目:\n1. 文書の目的・対象読者\n2. 構成（セクション一覧と各セクションの役割）\n3. 内容の強み・弱み\n4. データや根拠の妥当性\n5. 改善提案（具体的に3つ以上）",
            PromptEn = "Analyze the following document comprehensively.\n\n1. Purpose & target audience\n2. Structure (section overview)\n3. Strengths & weaknesses\n4. Data validity\n5. Improvement suggestions (3+)",
            Icon = "🔍", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_data_insight",
            CategoryJa = "分析・要約", CategoryEn = "Analysis",
            LabelJa = "データインサイト抽出", LabelEn = "Extract Data Insights",
            PromptJa = "以下のデータを分析し、ビジネスインサイトを報告してください。\n\n【出力フォーマット】\n■ データ概要（件数・期間・カラム説明）\n■ 主要トレンド（数値根拠付き）\n■ 異常値・注意点\n■ アクション提案（優先度付き）",
            PromptEn = "Analyze the following data and report business insights.\n\n■ Data Overview\n■ Key Trends (with numbers)\n■ Anomalies & Alerts\n■ Action Items (prioritized)",
            Icon = "📈", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_compare",
            CategoryJa = "分析・要約", CategoryEn = "Analysis",
            LabelJa = "比較・差分分析", LabelEn = "Comparison Analysis",
            PromptJa = "このドキュメント内のデータや情報を比較分析してください。\n\n・項目間の共通点と相違点を表形式で整理\n・数値がある場合は増減率・差分を計算\n・注目すべきギャップや改善機会を指摘",
            PromptEn = "Compare and analyze the data in this document.\n\n- Similarities & differences in table format\n- Calculate changes/deltas for numerical data\n- Highlight gaps and improvement opportunities",
            Icon = "⚖️", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_key_figures",
            CategoryJa = "分析・要約", CategoryEn = "Analysis",
            LabelJa = "重要数値の抽出", LabelEn = "Extract Key Figures",
            PromptJa = "このドキュメントから重要な数値・金額・日付・KPIをすべて抽出し、一覧表にまとめてください。\n\n【出力フォーマット】\n| 項目名 | 数値 | 単位 | 文脈・備考 |\n（数値が見当たらない場合はその旨を報告）",
            PromptEn = "Extract all key figures, amounts, dates, and KPIs from this document and list them in a table.\n\n| Item | Value | Unit | Context |",
            Icon = "🔢", RequiresContextData = true,
        },

        // ═══════════════════════════════════════════════════════════════
        // ■ 校正・品質チェック（Proofreading & QC）
        // ═══════════════════════════════════════════════════════════════
        new PresetPrompt
        {
            Id = "preset_proofread",
            CategoryJa = "校正・チェック", CategoryEn = "Proofreading",
            LabelJa = "文章校正", LabelEn = "Proofread",
            PromptJa = "以下のドキュメントを校正してください。\n\n検出項目:\n1. 誤字脱字\n2. 文法・助詞の誤り\n3. 不自然な表現・冗長な言い回し\n4. 表記ゆれ（漢字/ひらがな、全角/半角）\n\n【出力フォーマット】\n| # | 箇所 | 原文 | 修正案 | 理由 |",
            PromptEn = "Proofread the following document.\n\nCheck for:\n1. Typos & spelling\n2. Grammar errors\n3. Awkward phrasing\n4. Inconsistent formatting\n\nFormat: | # | Location | Original | Correction | Reason |",
            Icon = "✏️", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_consistency_check",
            CategoryJa = "校正・チェック", CategoryEn = "Proofreading",
            LabelJa = "表記・スタイル統一チェック", LabelEn = "Consistency Check",
            PromptJa = "このドキュメント内の表記ゆれ・スタイルの不統一を検出してください。\n\nチェック項目:\n・漢字/ひらがな表記の揺れ（例: 致します/いたします）\n・敬語レベルの統一\n・数字の全角/半角\n・用語・略語の統一\n・句読点の使い方\n\n一覧表で報告し、推奨する統一ルールも提案してください。",
            PromptEn = "Check this document for inconsistencies in terminology, formatting, number styles, abbreviations, and tone. Report findings in a table and suggest unified rules.",
            Icon = "🔄", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_presubmit_check",
            CategoryJa = "校正・チェック", CategoryEn = "Proofreading",
            LabelJa = "提出前 最終チェック", LabelEn = "Pre-Submission Review",
            PromptJa = "このドキュメントを提出前の最終チェックとしてレビューしてください。\n\nチェック項目:\n✅ 誤字脱字・文法\n✅ 数値の整合性（合計・計算結果）\n✅ 日付・名前・固有名詞の正確性\n✅ 宛名・敬称の適切さ\n✅ 添付ファイルへの言及漏れ\n✅ 全体の論理構成\n\n問題があれば緊急度（高/中/低）付きで報告してください。",
            PromptEn = "Review this document as a pre-submission check.\n\n✅ Typos & grammar\n✅ Number consistency\n✅ Dates, names, proper nouns\n✅ Addressing & honorifics\n✅ Attachment references\n✅ Logical structure\n\nReport issues with severity (High/Medium/Low).",
            Icon = "✅", RequiresContextData = true,
        },

        // ═══════════════════════════════════════════════════════════════
        // ■ 翻訳（Translation）
        // ═══════════════════════════════════════════════════════════════
        new PresetPrompt
        {
            Id = "preset_translate_en",
            CategoryJa = "翻訳", CategoryEn = "Translation",
            LabelJa = "英語に翻訳", LabelEn = "Translate to English",
            PromptJa = "以下のドキュメントを自然なビジネス英語に翻訳してください。\n\n注意点:\n・原文の意図とニュアンスを保持\n・英語圏のビジネス文書として適切な表現に\n・固有名詞や専門用語は原文をカッコ書きで併記",
            PromptEn = "Translate the following document into natural business English, preserving intent and nuance.",
            Icon = "🌐", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_translate_ja",
            CategoryJa = "翻訳", CategoryEn = "Translation",
            LabelJa = "日本語に翻訳", LabelEn = "Translate to Japanese",
            PromptJa = "以下のドキュメントを自然な日本語に翻訳してください。ビジネス文書として適切な敬語レベルで書いてください。",
            PromptEn = "Translate the following document into natural Japanese. Use appropriate business-level keigo.",
            Icon = "🇯🇵", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_translate_zh",
            CategoryJa = "翻訳", CategoryEn = "Translation",
            LabelJa = "中国語に翻訳", LabelEn = "Translate to Chinese",
            PromptJa = "以下のドキュメントを自然な中国語（簡体字）に翻訳してください。ビジネス文書として適切な表現にしてください。",
            PromptEn = "Translate the following document into Simplified Chinese in a professional business tone.",
            Icon = "🇨🇳", RequiresContextData = true,
        },

        // ═══════════════════════════════════════════════════════════════
        // ■ ビジネス文書作成（Document Creation）
        // ═══════════════════════════════════════════════════════════════
        new PresetPrompt
        {
            Id = "preset_meeting_minutes",
            CategoryJa = "文書作成", CategoryEn = "Document Creation",
            LabelJa = "議事録を作成", LabelEn = "Create Meeting Minutes",
            PromptJa = "以下の内容から議事録を作成してください。\n\n【フォーマット】\n■ 会議名:\n■ 日時:\n■ 参加者:\n■ 議題と決定事項（番号付き）\n■ アクションアイテム:\n  | 担当者 | タスク | 期限 |\n■ 次回予定:",
            PromptEn = "Create meeting minutes from the following content.\n\nFormat:\n■ Meeting Title\n■ Date & Time\n■ Attendees\n■ Agenda & Decisions\n■ Action Items: | Owner | Task | Deadline |\n■ Next Meeting",
            Icon = "📝", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_report_draft",
            CategoryJa = "文書作成", CategoryEn = "Document Creation",
            LabelJa = "レポート下書き", LabelEn = "Draft Report",
            PromptJa = "以下のデータ・情報をもとに、ビジネスレポートの下書きを作成してください。\n\n【構成】\n1. エグゼクティブサマリー（経営層が30秒で把握できる要約）\n2. 背景・目的\n3. 分析結果（図表の提案含む）\n4. 考察・インサイト\n5. 提言・次のステップ",
            PromptEn = "Draft a business report.\n\n1. Executive Summary\n2. Background & Purpose\n3. Analysis (suggest charts/tables)\n4. Discussion & Insights\n5. Recommendations & Next Steps",
            Icon = "📄", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_email_draft",
            CategoryJa = "文書作成", CategoryEn = "Document Creation",
            LabelJa = "ビジネスメール作成", LabelEn = "Draft Business Email",
            PromptJa = "以下の要件に基づいて、ビジネスメールの文面を作成してください。\n\n・丁寧かつ簡潔\n・件名も提案\n・必要に応じて複数パターン（フォーマル / カジュアル）を提示",
            PromptEn = "Draft a business email based on the requirements below.\n\n- Professional and concise tone\n- Suggest a subject line\n- Provide formal and casual variants if appropriate",
            Icon = "📧", RequiresContextData = false,
        },
        new PresetPrompt
        {
            Id = "preset_proposal",
            CategoryJa = "文書作成", CategoryEn = "Document Creation",
            LabelJa = "提案書ドラフト", LabelEn = "Draft Proposal",
            PromptJa = "以下の情報をもとに、クライアント向け提案書のドラフトを作成してください。\n\n【構成】\n1. 表紙情報（タイトル・日付・提出先）\n2. 課題の整理\n3. 提案内容（解決策）\n4. 期待効果・メリット\n5. スケジュール案\n6. 概算費用\n7. 次のステップ",
            PromptEn = "Draft a client proposal.\n\n1. Title & Cover Info\n2. Problem Statement\n3. Proposed Solution\n4. Expected Benefits\n5. Timeline\n6. Budget Estimate\n7. Next Steps",
            Icon = "💼", RequiresContextData = false,
        },
        new PresetPrompt
        {
            Id = "preset_contract_review",
            CategoryJa = "文書作成", CategoryEn = "Document Creation",
            LabelJa = "契約書レビュー", LabelEn = "Contract Review",
            PromptJa = "以下の契約書をレビューしてください。\n\nチェック項目:\n・不利な条項や一方的な責任条項\n・曖昧な表現・定義不足\n・期間・解約条件の妥当性\n・損害賠償・免責条項のリスク\n・不足している条項の提案\n\nリスクレベル（高/中/低）付きで報告してください。",
            PromptEn = "Review the following contract.\n\nCheck:\n- Unfavorable or one-sided clauses\n- Ambiguous language\n- Term & termination conditions\n- Liability & indemnification risks\n- Missing clauses\n\nReport with risk level (High/Medium/Low).",
            Icon = "⚖️", RequiresContextData = true,
        },

        // ═══════════════════════════════════════════════════════════════
        // ■ Excel 分析・数式（Excel & Data）
        // ═══════════════════════════════════════════════════════════════
        new PresetPrompt
        {
            Id = "preset_excel_formula",
            CategoryJa = "Excel", CategoryEn = "Excel",
            LabelJa = "数式・関数ヘルプ", LabelEn = "Formula Help",
            PromptJa = "ユーザーが求める計算・集計を実現する Excel 数式を提案してください。\n\n【出力フォーマット】\n■ 推奨数式: =関数(...)\n■ 各引数の説明\n■ 使用例（具体的なセル参照）\n■ 注意点・よくあるミス\n■ 代替案があれば提示",
            PromptEn = "Suggest Excel formulas.\n\n■ Recommended formula\n■ Argument explanation\n■ Usage example with cell references\n■ Common pitfalls\n■ Alternative approaches",
            Icon = "📊", RequiresContextData = false,
        },
        new PresetPrompt
        {
            Id = "preset_excel_data_summary",
            CategoryJa = "Excel", CategoryEn = "Excel",
            LabelJa = "データ概要・集計", LabelEn = "Data Summary",
            PromptJa = "開いているスプレッドシートのデータを分析し、以下を報告してください。\n\n■ データ概要（行数・列数・期間・カテゴリ）\n■ 基本統計（合計・平均・最大・最小・中央値）\n■ カテゴリ別集計\n■ 構成比（上位5件）\n■ データ品質の所見（欠損値・外れ値）",
            PromptEn = "Analyze the open spreadsheet and report:\n\n■ Data overview (rows, columns, period)\n■ Basic statistics (sum, avg, max, min, median)\n■ Category breakdown\n■ Composition ratios (top 5)\n■ Data quality notes",
            Icon = "📊", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_excel_variance",
            CategoryJa = "Excel", CategoryEn = "Excel",
            LabelJa = "予実差異分析", LabelEn = "Budget vs Actual",
            PromptJa = "このスプレッドシートのデータから予算と実績の差異を分析してください。\n\n■ 差異額と差異率を計算\n■ 予算超過・未達の主要項目を特定\n■ 要因分析（なぜ差が出たか仮説を提示）\n■ 来期への改善提案",
            PromptEn = "Analyze budget vs actual variances.\n\n■ Calculate variance amount and percentage\n■ Identify major over/under items\n■ Root cause hypotheses\n■ Improvement suggestions for next period",
            Icon = "📉", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_excel_quality",
            CategoryJa = "Excel", CategoryEn = "Excel",
            LabelJa = "データ品質チェック", LabelEn = "Data Quality Check",
            PromptJa = "このスプレッドシートのデータ品質をチェックしてください。\n\n検出項目:\n・空白セル・欠損値\n・重複データ\n・数値の外れ値（統計的に異常な値）\n・表記ゆれ（同じ意味の異なる表記）\n・日付フォーマットの不整合\n\n件数と具体例を含めて報告してください。",
            PromptEn = "Check data quality.\n\n- Missing values\n- Duplicates\n- Outliers\n- Inconsistent text/formats\n- Date format issues\n\nReport counts and examples.",
            Icon = "🔍", RequiresContextData = true,
        },

        // ═══════════════════════════════════════════════════════════════
        // ■ プレゼン支援（Presentation）
        // ═══════════════════════════════════════════════════════════════
        new PresetPrompt
        {
            Id = "preset_pptx_proofread",
            CategoryJa = "プレゼン", CategoryEn = "Presentation",
            LabelJa = "スライドテキスト校正", LabelEn = "Proofread Slides",
            PromptJa = "このプレゼンテーションの全スライドテキストを校正してください。\n\n・誤字脱字\n・スライド間の表記ゆれ\n・文字数が多すぎるスライドの指摘\n・改善案を「スライド番号 → 修正内容」の形式で一覧化",
            PromptEn = "Proofread all slide text.\n\n- Typos & errors\n- Cross-slide inconsistencies\n- Text-heavy slides flagged\n- Corrections listed as: Slide # → Fix",
            Icon = "🎯", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_pptx_structure",
            CategoryJa = "プレゼン", CategoryEn = "Presentation",
            LabelJa = "構成アドバイス", LabelEn = "Structure Advice",
            PromptJa = "このプレゼンテーションの構成をレビューし、改善案を提案してください。\n\n分析項目:\n1. ストーリーフロー（起承転結が明確か）\n2. スライド順序の最適化\n3. 情報の過不足\n4. 聞き手を惹きつけるオープニング案\n5. 効果的なクロージング案",
            PromptEn = "Review presentation structure and suggest improvements.\n\n1. Story flow clarity\n2. Slide order optimization\n3. Information gaps/redundancy\n4. Engaging opening suggestion\n5. Effective closing suggestion",
            Icon = "🏗️", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_pptx_simplify",
            CategoryJa = "プレゼン", CategoryEn = "Presentation",
            LabelJa = "スライドを簡潔に", LabelEn = "Simplify Slides",
            PromptJa = "各スライドのテキストを簡潔化してください。\n\n・1スライド1メッセージを徹底\n・箇条書きは1行20文字以内を目標\n・冗長な表現を体言止め・名詞句に変換\n・元の意味を損なわずに文字数を削減\n\n【出力】スライドごとに「Before → After」形式で提示",
            PromptEn = "Simplify each slide's text.\n\n- One message per slide\n- Target: under 20 chars per bullet\n- Convert to noun phrases\n- Reduce text without losing meaning\n\nShow Before → After for each slide.",
            Icon = "✨", RequiresContextData = true,
        },

        // ═══════════════════════════════════════════════════════════════
        // ■ 業務効率化（Productivity）
        // ═══════════════════════════════════════════════════════════════
        new PresetPrompt
        {
            Id = "preset_todo_extract",
            CategoryJa = "業務効率化", CategoryEn = "Productivity",
            LabelJa = "タスク・TODO 抽出", LabelEn = "Extract Tasks & TODOs",
            PromptJa = "このドキュメントから、タスク・TODO・アクションアイテムをすべて抽出してください。\n\n【出力フォーマット】\n| # | タスク | 担当者 | 期限 | 優先度 |\n\n・暗黙的なタスク（〜する必要がある、〜を検討、等）も含める\n・期限が明記されていない場合は「要確認」と記載",
            PromptEn = "Extract all tasks, TODOs, and action items from this document.\n\n| # | Task | Owner | Deadline | Priority |\n\nInclude implicit tasks. Mark unspecified deadlines as 'TBD'.",
            Icon = "📌", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_explain",
            CategoryJa = "業務効率化", CategoryEn = "Productivity",
            LabelJa = "内容をわかりやすく解説", LabelEn = "Explain Simply",
            PromptJa = "このドキュメントの内容を、専門知識がない人にもわかるように解説してください。\n\n・専門用語には簡単な説明を付加\n・具体的な例やたとえ話を使用\n・重要な部分を「ポイント」として強調",
            PromptEn = "Explain this document in simple terms for non-experts.\n\n- Define technical terms\n- Use examples and analogies\n- Highlight key points",
            Icon = "💡", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_qa_generate",
            CategoryJa = "業務効率化", CategoryEn = "Productivity",
            LabelJa = "想定Q&A 生成", LabelEn = "Generate Q&A",
            PromptJa = "このドキュメントの内容について、読者・聞き手から想定される質問とその回答を生成してください。\n\n・基本的な質問 3〜5件\n・鋭い指摘・反論 2〜3件\n・それぞれに対する模範回答を用意\n\nプレゼンや会議の事前準備として使えるレベルで作成してください。",
            PromptEn = "Generate anticipated Q&A for this document.\n\n- 3-5 basic questions with answers\n- 2-3 challenging questions with answers\n\nPrepare as if for a meeting or presentation.",
            Icon = "❓", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_what_can_ai_do",
            CategoryJa = "業務効率化", CategoryEn = "Productivity",
            LabelJa = "AI にできることを提案", LabelEn = "Suggest AI Actions",
            PromptJa = "現在開いているドキュメントの内容を確認し、AI アシスタントとしてお手伝いできることを提案してください。\n\n・このデータ/文書に対して実行可能な分析\n・改善・効率化できるポイント\n・作成可能な追加ドキュメント\n\n各提案に対して、「やってみますか？」と聞いてください。",
            PromptEn = "Review the open document and suggest what I can help with.\n\n- Possible analyses\n- Improvement opportunities\n- Additional documents I can create\n\nAsk 'Shall I do this?' for each suggestion.",
            Icon = "🤖", RequiresContextData = true,
        },

        // ═══════════════════════════════════════════════════════════════
        // ■ 文章改善（Writing Enhancement）
        // ═══════════════════════════════════════════════════════════════
        new PresetPrompt
        {
            Id = "preset_business_tone",
            CategoryJa = "文章改善", CategoryEn = "Writing",
            LabelJa = "ビジネス文体に変換", LabelEn = "Convert to Business Tone",
            PromptJa = "以下の文章をプロフェッショナルなビジネス文体に変換してください。\n\n・敬語レベルを適切に調整\n・曖昧な表現を具体的に\n・冗長な部分を簡潔に\n・原文の内容は変更しない\n\n変更箇所がわかるように Before → After で示してください。",
            PromptEn = "Convert the text to professional business tone.\n\n- Appropriate formality level\n- Replace vague language with specifics\n- Reduce wordiness\n- Don't change content meaning\n\nShow Before → After for changes.",
            Icon = "👔", RequiresContextData = true,
        },
        new PresetPrompt
        {
            Id = "preset_exec_summary",
            CategoryJa = "文章改善", CategoryEn = "Writing",
            LabelJa = "エグゼクティブサマリー作成", LabelEn = "Create Executive Summary",
            PromptJa = "このドキュメントのエグゼクティブサマリーを作成してください。\n\n条件:\n・経営層が1分以内で読める分量（200〜400字）\n・結論を最初に述べる\n・重要数値を含める\n・次のアクションを明記\n・箇条書きではなく文章形式で",
            PromptEn = "Create an executive summary.\n\n- Readable in under 1 minute (100-200 words)\n- Lead with the conclusion\n- Include key numbers\n- State next actions\n- Prose format, not bullets",
            Icon = "📑", RequiresContextData = true,
        },
    ];
}
