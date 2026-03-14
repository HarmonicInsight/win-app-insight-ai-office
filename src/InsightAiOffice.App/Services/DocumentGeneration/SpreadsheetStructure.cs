using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// AI が生成するスプレッドシートの中間表現
/// </summary>
public class SpreadsheetStructure
{
    [JsonPropertyName("title")]
    public string Title { get; set; } = string.Empty;

    [JsonPropertyName("sheets")]
    public List<SimpleSheetData> Sheets { get; set; } = new();
}

public class SimpleSheetData
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "Sheet1";

    [JsonPropertyName("headers")]
    public List<string> Headers { get; set; } = new();

    [JsonPropertyName("rows")]
    public List<List<string>> Rows { get; set; } = new();
}
