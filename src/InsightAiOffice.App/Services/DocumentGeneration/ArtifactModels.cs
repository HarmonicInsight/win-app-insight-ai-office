using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace InsightAiOffice.App.Services.DocumentGeneration;

public enum ArtifactType
{
    Html, Chart, Table, Mermaid, Markdown, Svg, Word, PowerPoint, Excel
}

/// <summary>永続化されるアーティファクトメタデータ</summary>
public class ArtifactEntry
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";
    [JsonPropertyName("title")]
    public string Title { get; set; } = "";
    [JsonPropertyName("type")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ArtifactType Type { get; set; }
    [JsonPropertyName("fileName")]
    public string FileName { get; set; } = "";
    [JsonPropertyName("createdAt")]
    public DateTime CreatedAt { get; set; }
    [JsonPropertyName("fileSizeBytes")]
    public long FileSizeBytes { get; set; }
    [JsonPropertyName("model")]
    public string Model { get; set; } = "";
    [JsonPropertyName("prompt")]
    public string Prompt { get; set; } = "";
}

/// <summary>パース中間表現</summary>
public class ArtifactBlock
{
    public ArtifactType Type { get; set; }
    public string Title { get; set; } = "";
    public string Content { get; set; } = "";
}

/// <summary>パース結果</summary>
public class ArtifactParseResult
{
    public string DisplayText { get; set; } = "";
    public ArtifactBlock[] Artifacts { get; set; } = Array.Empty<ArtifactBlock>();
}

/// <summary>保存結果</summary>
public class ArtifactSaveResult
{
    public string DisplayText { get; set; } = "";
    public ArtifactEntry[] Entries { get; set; } = Array.Empty<ArtifactEntry>();
    public bool HasArtifacts { get; set; }
}

/// <summary>チャット内のリンク表示用</summary>
public class ArtifactLink
{
    public string Id { get; set; } = "";
    public string Title { get; set; } = "";
    public string Icon { get; set; } = "";
    public ArtifactType Type { get; set; }
}
