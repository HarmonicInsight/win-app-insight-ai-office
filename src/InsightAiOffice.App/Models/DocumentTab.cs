namespace InsightAiOffice.App.Models;

/// <summary>
/// 開いているドキュメントタブの状態を保持する。
/// タブ切り替え時にエディタ内容の保存・復元に使用。
/// </summary>
public sealed class DocumentTab
{
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public string EditorType { get; set; } = ""; // "word", "excel", "pptx"

    // Word: SfRichTextBoxAdv の内容を byte[] で保持
    public byte[]? WordContent { get; set; }

    // PPTX: 選択中のスライドインデックス
    public int PptxSelectedSlide { get; set; }

    // Text: テキストエディタの内容
    public string? TextContent { get; set; }
}
