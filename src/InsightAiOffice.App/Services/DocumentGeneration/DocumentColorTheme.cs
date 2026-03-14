namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// ドキュメント生成用のカラーテーマ。
/// AI がユーザーの指示に応じてテーマを選択できる。
/// </summary>
public class DocumentColorTheme
{
    public System.Drawing.Color Primary { get; init; }
    public System.Drawing.Color PrimaryDark { get; init; }
    public System.Drawing.Color PrimaryLight { get; init; }
    public System.Drawing.Color Background { get; init; }
    public System.Drawing.Color TextPrimary { get; init; }
    public System.Drawing.Color TextSecondary { get; init; }
    public System.Drawing.Color TableHeaderBg { get; init; }
    public System.Drawing.Color TableHeaderText { get; init; }
    public System.Drawing.Color TableStripeBg { get; init; }
    public System.Drawing.Color BorderColor { get; init; }
    public System.Drawing.Color SummaryBg { get; init; }

    private static System.Drawing.Color C(byte r, byte g, byte b) =>
        System.Drawing.Color.FromArgb(r, g, b);

    /// <summary>Ivory & Gold（デフォルト）</summary>
    public static DocumentColorTheme IvoryGold { get; } = new()
    {
        Primary = C(0xB8, 0x94, 0x2F),
        PrimaryDark = C(0x8A, 0x6F, 0x23),
        PrimaryLight = C(0xD4, 0xB9, 0x4A),
        Background = C(0xFA, 0xF8, 0xF5),
        TextPrimary = C(0x1C, 0x19, 0x17),
        TextSecondary = C(0x57, 0x53, 0x4E),
        TableHeaderBg = C(0xB8, 0x94, 0x2F),
        TableHeaderText = C(0xFF, 0xFF, 0xFF),
        TableStripeBg = C(0xFA, 0xF8, 0xF5),
        BorderColor = C(0xE7, 0xE2, 0xDA),
        SummaryBg = C(0xF5, 0xF0, 0xE8),
    };

    /// <summary>ブルー系（ビジネス・コーポレート）</summary>
    public static DocumentColorTheme Blue { get; } = new()
    {
        Primary = C(0x25, 0x63, 0xEB),
        PrimaryDark = C(0x1D, 0x4E, 0xD8),
        PrimaryLight = C(0x60, 0xA5, 0xFA),
        Background = C(0xF8, 0xFA, 0xFC),
        TextPrimary = C(0x0F, 0x17, 0x2A),
        TextSecondary = C(0x47, 0x55, 0x69),
        TableHeaderBg = C(0x25, 0x63, 0xEB),
        TableHeaderText = C(0xFF, 0xFF, 0xFF),
        TableStripeBg = C(0xF1, 0xF5, 0xF9),
        BorderColor = C(0xE2, 0xE8, 0xF0),
        SummaryBg = C(0xEF, 0xF6, 0xFF),
    };

    /// <summary>グリーン系（環境・ヘルスケア）</summary>
    public static DocumentColorTheme Green { get; } = new()
    {
        Primary = C(0x16, 0xA3, 0x4A),
        PrimaryDark = C(0x15, 0x80, 0x3D),
        PrimaryLight = C(0x4A, 0xDE, 0x80),
        Background = C(0xF0, 0xFD, 0xF4),
        TextPrimary = C(0x14, 0x53, 0x2D),
        TextSecondary = C(0x3F, 0x6F, 0x52),
        TableHeaderBg = C(0x16, 0xA3, 0x4A),
        TableHeaderText = C(0xFF, 0xFF, 0xFF),
        TableStripeBg = C(0xF0, 0xFD, 0xF4),
        BorderColor = C(0xBB, 0xF7, 0xD0),
        SummaryBg = C(0xDC, 0xFC, 0xE7),
    };

    /// <summary>レッド系（警告・重要文書）</summary>
    public static DocumentColorTheme Red { get; } = new()
    {
        Primary = C(0xDC, 0x26, 0x26),
        PrimaryDark = C(0xB9, 0x1C, 0x1C),
        PrimaryLight = C(0xF8, 0x71, 0x71),
        Background = C(0xFE, 0xF2, 0xF2),
        TextPrimary = C(0x45, 0x0A, 0x0A),
        TextSecondary = C(0x7F, 0x1D, 0x1D),
        TableHeaderBg = C(0xDC, 0x26, 0x26),
        TableHeaderText = C(0xFF, 0xFF, 0xFF),
        TableStripeBg = C(0xFE, 0xF2, 0xF2),
        BorderColor = C(0xFE, 0xCA, 0xCA),
        SummaryBg = C(0xFE, 0xE2, 0xE2),
    };

    /// <summary>ダークネイビー（法務・金融）</summary>
    public static DocumentColorTheme Navy { get; } = new()
    {
        Primary = C(0x1E, 0x29, 0x3B),
        PrimaryDark = C(0x0F, 0x17, 0x2A),
        PrimaryLight = C(0x33, 0x45, 0x5B),
        Background = C(0xF8, 0xFA, 0xFC),
        TextPrimary = C(0x0F, 0x17, 0x2A),
        TextSecondary = C(0x47, 0x55, 0x69),
        TableHeaderBg = C(0x1E, 0x29, 0x3B),
        TableHeaderText = C(0xFF, 0xFF, 0xFF),
        TableStripeBg = C(0xF1, 0xF5, 0xF9),
        BorderColor = C(0xCB, 0xD5, 0xE1),
        SummaryBg = C(0xE2, 0xE8, 0xF0),
    };

    /// <summary>モノクロ（シンプル）</summary>
    public static DocumentColorTheme Mono { get; } = new()
    {
        Primary = C(0x33, 0x33, 0x33),
        PrimaryDark = C(0x1A, 0x1A, 0x1A),
        PrimaryLight = C(0x66, 0x66, 0x66),
        Background = C(0xFF, 0xFF, 0xFF),
        TextPrimary = C(0x1A, 0x1A, 0x1A),
        TextSecondary = C(0x66, 0x66, 0x66),
        TableHeaderBg = C(0x33, 0x33, 0x33),
        TableHeaderText = C(0xFF, 0xFF, 0xFF),
        TableStripeBg = C(0xF5, 0xF5, 0xF5),
        BorderColor = C(0xD4, 0xD4, 0xD4),
        SummaryBg = C(0xF0, 0xF0, 0xF0),
    };

    /// <summary>テーマ名から取得</summary>
    public static DocumentColorTheme FromName(string? name) => (name?.ToLowerInvariant()) switch
    {
        "blue" or "ブルー" or "青" or "青色" => Blue,
        "green" or "グリーン" or "緑" or "緑色" => Green,
        "red" or "レッド" or "赤" or "赤色" => Red,
        "navy" or "ネイビー" or "紺" or "紺色" => Navy,
        "mono" or "モノクロ" or "白黒" or "モノ" => Mono,
        "gold" or "ゴールド" or "金" or "ivory" or "アイボリー" => IvoryGold,
        _ => IvoryGold,
    };
}
