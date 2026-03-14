using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// スライド種別
/// </summary>
public enum SlideType
{
    Title, Purpose, Agenda, Overview, Step, Decision,
    Checklist, FAQ, Summary, Data, Blank
}

/// <summary>
/// 1 スライドの中間表現（AI ツール generate_presentation の入出力）
/// </summary>
public class SlideSpecItem
{
    [JsonPropertyName("order")]
    public int Order { get; set; }

    [JsonPropertyName("slideType")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public SlideType SlideType { get; set; } = SlideType.Blank;

    [JsonPropertyName("title")]
    public string Title { get; set; } = string.Empty;

    [JsonPropertyName("keyMessage")]
    public string KeyMessage { get; set; } = string.Empty;

    [JsonPropertyName("bullets")]
    public List<string> Bullets { get; set; } = new();

    [JsonPropertyName("speakerNotes")]
    public string SpeakerNotes { get; set; } = string.Empty;

    [JsonPropertyName("durationSec")]
    public int DurationSec { get; set; }
}
