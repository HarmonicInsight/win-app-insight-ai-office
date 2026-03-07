using InsightCommon.AI;

namespace InsightAiOffice.App.Helpers;

public static class BuiltInPresets
{
    public static List<UserPromptPreset> GetAll() =>
    [
        new()
        {
            Id = "builtin_summarize",
            Name = "ドキュメント要約",
            SystemPrompt = "以下のドキュメントの内容を要約してください。重要なポイントを箇条書きで示してください。",
            Description = "ドキュメント全体の要約を生成",
            Category = "分析",
            Icon = "📋",
            RequiresContextData = true,
        },
        new()
        {
            Id = "builtin_analyze",
            Name = "構造分析",
            SystemPrompt = "以下のドキュメントの構造と内容を分析してください。文書の目的、主要なセクション、データの傾向、改善すべき点を挙げてください。",
            Description = "文書構造・内容を多角的に分析",
            Category = "分析",
            Icon = "🔍",
            RequiresContextData = true,
        },
        new()
        {
            Id = "builtin_proofread",
            Name = "文章校正",
            SystemPrompt = "以下のドキュメントを校正してください。誤字脱字、文法の誤り、不自然な表現を指摘し、修正案を示してください。",
            Description = "誤字脱字・文法チェック・改善提案",
            Category = "校正",
            Icon = "✏️",
            RequiresContextData = true,
        },
        new()
        {
            Id = "builtin_translate_en",
            Name = "英訳",
            SystemPrompt = "以下のドキュメントの内容を自然なビジネス英語に翻訳してください。原文の意図を保ちながら、英語圏のビジネス文書として適切な表現にしてください。",
            Description = "ドキュメントを英語に翻訳",
            Category = "翻訳",
            Icon = "🌐",
            RequiresContextData = true,
        },
        new()
        {
            Id = "builtin_translate_ja",
            Name = "和訳",
            SystemPrompt = "以下のドキュメントの内容を自然な日本語に翻訳してください。原文の意図を保ちながら、日本語のビジネス文書として適切な表現にしてください。",
            Description = "ドキュメントを日本語に翻訳",
            Category = "翻訳",
            Icon = "🇯🇵",
            RequiresContextData = true,
        },
        new()
        {
            Id = "builtin_excel_formula",
            Name = "Excel 数式ヘルプ",
            SystemPrompt = "ユーザーが求める計算・集計を実現する Excel 数式を提案してください。関数名、引数、使用例を示し、注意点があれば補足してください。",
            Description = "Excel 数式・関数の提案",
            Category = "Excel",
            Icon = "📊",
            RequiresContextData = false,
        },
        new()
        {
            Id = "builtin_data_insight",
            Name = "データインサイト",
            SystemPrompt = "以下のスプレッドシートのデータを分析し、重要なインサイト（傾向、異常値、改善提案）を提示してください。数値の根拠を明示してください。",
            Description = "表データから傾向・異常値を検出",
            Category = "分析",
            Icon = "📈",
            RequiresContextData = true,
        },
        new()
        {
            Id = "builtin_meeting_minutes",
            Name = "議事録作成",
            SystemPrompt = "以下の内容から議事録を作成してください。日時、参加者、議題、決定事項、アクションアイテム（担当者・期限付き）の形式でまとめてください。",
            Description = "テキストから議事録フォーマットを生成",
            Category = "文書作成",
            Icon = "📝",
            RequiresContextData = true,
        },
    ];
}
