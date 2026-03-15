using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using InsightCommon.AI;

namespace InsightAiOffice.App.Services;

/// <summary>
/// チャット履歴の永続化サービス。
/// %APPDATA%/HarmonicInsight/InsightAiOffice/chat_history.json に保存。
/// </summary>
public static class ChatHistoryService
{
    private static readonly JsonSerializerOptions s_opts = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    private static string GetPath() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "HarmonicInsight", "InsightAiOffice", "chat_history.json");

    public static void Save(IEnumerable<ChatMessageVm> messages)
    {
        try
        {
            var entries = messages
                .Where(m => !m.IsWelcome)
                .Select(m => new ChatEntry
                {
                    Role = m.Role.ToString(),
                    Content = m.Content,
                    Timestamp = m.Timestamp,
                })
                .ToList();

            // 最新100件のみ保存
            if (entries.Count > 100)
                entries = entries.Skip(entries.Count - 100).ToList();

            var dir = Path.GetDirectoryName(GetPath())!;
            Directory.CreateDirectory(dir);
            File.WriteAllText(GetPath(), JsonSerializer.Serialize(entries, s_opts));
        }
        catch { /* 保存失敗は無視 */ }
    }

    public static List<ChatMessageVm> Load()
    {
        try
        {
            var path = GetPath();
            if (!File.Exists(path)) return new();

            var json = File.ReadAllText(path);
            var entries = JsonSerializer.Deserialize<List<ChatEntry>>(json, s_opts) ?? new();

            return entries.Select(e => new ChatMessageVm
            {
                Role = Enum.TryParse<ChatRole>(e.Role, out var r) ? r : ChatRole.System,
                Content = e.Content,
                Timestamp = e.Timestamp,
            }).ToList();
        }
        catch { return new(); }
    }

    private class ChatEntry
    {
        public string Role { get; set; } = "";
        public string Content { get; set; } = "";
        public DateTime Timestamp { get; set; }
    }
}
