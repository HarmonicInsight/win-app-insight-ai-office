using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// AI が生成するレポートの中間表現 → ReportRendererService で Word に変換
/// </summary>
public class ReportStructure
{
    [JsonPropertyName("title")]
    public string Title { get; set; } = "";

    [JsonPropertyName("author")]
    public string Author { get; set; } = "";

    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("sections")]
    public List<ReportSection> Sections { get; set; } = new();
}

public class ReportSection
{
    /// <summary>
    /// title, heading, summary, text, recommendation,
    /// bullet_list, table, comparison, chart, key_metrics, page_break
    /// </summary>
    [JsonPropertyName("type")]
    public string Type { get; set; } = "text";

    [JsonPropertyName("title")]
    public string Title { get; set; } = "";

    [JsonPropertyName("level")]
    public int Level { get; set; } = 1;

    [JsonPropertyName("content")]
    public string Content { get; set; } = "";

    [JsonPropertyName("items")]
    public List<string> Items { get; set; } = new();

    [JsonPropertyName("tableData")]
    public ReportTableData? TableData { get; set; }

    [JsonPropertyName("chartData")]
    public ReportChartData? ChartData { get; set; }

    [JsonPropertyName("metrics")]
    public List<ReportMetric> Metrics { get; set; } = new();

    [JsonPropertyName("sections")]
    public List<ReportSection> Sections { get; set; } = new();
}

public class ReportTableData
{
    [JsonPropertyName("headers")]
    public List<string> Headers { get; set; } = new();

    [JsonPropertyName("rows")]
    public List<List<string>> Rows { get; set; } = new();
}

public class ReportChartData
{
    [JsonPropertyName("title")]
    public string Title { get; set; } = "";

    [JsonPropertyName("chartType")]
    public string ChartType { get; set; } = "bar";

    [JsonPropertyName("categories")]
    public List<string> Categories { get; set; } = new();

    [JsonPropertyName("series")]
    public List<ReportChartSeries> Series { get; set; } = new();
}

public class ReportChartSeries
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("values")]
    public List<double> Values { get; set; } = new();
}

public class ReportMetric
{
    [JsonPropertyName("label")]
    public string Label { get; set; } = "";

    [JsonPropertyName("value")]
    public string Value { get; set; } = "";

    [JsonPropertyName("change")]
    public string Change { get; set; } = "";

    /// <summary>positive / negative / neutral</summary>
    [JsonPropertyName("trend")]
    public string Trend { get; set; } = "neutral";
}
