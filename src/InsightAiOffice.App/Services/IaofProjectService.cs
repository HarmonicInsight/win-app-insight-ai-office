using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using InsightCommon.AI;
using InsightCommon.ProjectFile;

namespace InsightAiOffice.App.Services;

/// <summary>
/// IAOF プロジェクトファイル（.iaof）の保存・読み込みサービス。
///
/// .iaof は ZIP パッケージ:
///   metadata.json          — プロジェクトメタデータ
///   document.{ext}         — メインドキュメント（docx/xlsx/txt 等）
///   ai_chat_history.json   — AI チャット履歴
///   references/            — 添付ファイル
/// </summary>
public class IaofProjectService : IDisposable
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
    };

    private string? _workDir;
    private bool _disposed;

    /// <summary>現在開いている .iaof ファイルパス</summary>
    public string? ProjectPath { get; private set; }

    /// <summary>プロジェクトが開いているか</summary>
    public bool IsOpen => _workDir != null;

    /// <summary>メタデータ</summary>
    public ProjectFileMetadata? Metadata { get; private set; }

    // =========================================================================
    // 新規作成
    // =========================================================================

    /// <summary>
    /// 現在開いているドキュメントから .iaof プロジェクトを新規作成
    /// </summary>
    public async Task CreateAsync(
        string projectPath,
        string? documentPath,
        string documentEditorType,
        IEnumerable<ChatMessageVm> chatMessages,
        string appVersion)
    {
        ProjectPath = projectPath;
        _workDir = CreateTempDir();

        // ディレクトリ構造
        Directory.CreateDirectory(Path.Combine(_workDir, "references", "files"));
        Directory.CreateDirectory(Path.Combine(_workDir, "ai_memory_deep"));
        Directory.CreateDirectory(Path.Combine(_workDir, "history", "snapshots"));

        // ドキュメントコピー
        var innerDocName = ResolveInnerDocName(documentPath, documentEditorType);
        if (documentPath != null && File.Exists(documentPath))
            File.Copy(documentPath, Path.Combine(_workDir, innerDocName));

        // メタデータ
        var now = DateTime.UtcNow.ToString("o");
        Metadata = new ProjectFileMetadata
        {
            SchemaVersion = ProjectFilePaths.SchemaVersion,
            ProductCode = "IAOF",
            AppVersion = appVersion,
            Title = Path.GetFileNameWithoutExtension(projectPath),
            CreatedAt = now,
            UpdatedAt = now,
            OriginalFileName = documentPath != null ? Path.GetFileName(documentPath) : "",
        };

        await WriteJsonAsync(ProjectFilePaths.Metadata, Metadata);
        await WriteJsonAsync(ProjectFilePaths.AiChatHistory, BuildChatHistory(chatMessages));
        await WriteJsonAsync(ProjectFilePaths.AiMemory, new { version = "1.0", entries = Array.Empty<object>() });
        await WriteJsonAsync(ProjectFilePaths.HistoryIndex, new HistoryIndex());
        await WriteJsonAsync(ProjectFilePaths.ReferencesIndex, new ReferencesIndex());

        // ZIP 保存
        await PackAsync();
    }

    // =========================================================================
    // 開く
    // =========================================================================

    /// <summary>.iaof プロジェクトを開く</summary>
    public async Task<(string? documentPath, string editorType, AiChatHistory chatHistory)> OpenAsync(string projectPath)
    {
        ProjectPath = projectPath;
        _workDir = CreateTempDir();

        ZipFile.ExtractToDirectory(projectPath, _workDir);

        // メタデータ読み込み
        Metadata = await ReadJsonAsync<ProjectFileMetadata>(ProjectFilePaths.Metadata);

        // ドキュメント検出
        string? docPath = null;
        string editorType = "text";
        foreach (var ext in new[] { ".docx", ".xlsx", ".pptx", ".txt" })
        {
            var candidate = Path.Combine(_workDir, $"document{ext}");
            if (File.Exists(candidate))
            {
                docPath = candidate;
                editorType = ext switch
                {
                    ".docx" => "word",
                    ".xlsx" => "excel",
                    ".pptx" => "pptx",
                    _ => "text",
                };
                break;
            }
        }

        // チャット履歴読み込み
        var chatHistory = await ReadJsonAsync<AiChatHistory>(ProjectFilePaths.AiChatHistory)
                          ?? new AiChatHistory();

        return (docPath, editorType, chatHistory);
    }

    // =========================================================================
    // 保存（上書き）
    // =========================================================================

    /// <summary>プロジェクトを上書き保存</summary>
    public async Task SaveAsync(
        string? documentPath,
        string documentEditorType,
        IEnumerable<ChatMessageVm> chatMessages)
    {
        if (_workDir == null || ProjectPath == null)
            throw new InvalidOperationException("プロジェクトが開かれていません");

        // ドキュメント更新
        var innerDocName = ResolveInnerDocName(documentPath, documentEditorType);
        if (documentPath != null && File.Exists(documentPath))
        {
            var dest = Path.Combine(_workDir, innerDocName);
            File.Copy(documentPath, dest, overwrite: true);
        }

        // メタデータ更新
        if (Metadata != null)
        {
            Metadata.UpdatedAt = DateTime.UtcNow.ToString("o");
            await WriteJsonAsync(ProjectFilePaths.Metadata, Metadata);
        }

        // チャット履歴更新
        await WriteJsonAsync(ProjectFilePaths.AiChatHistory, BuildChatHistory(chatMessages));

        // ZIP 再パッケージ
        await PackAsync();
    }

    // =========================================================================
    // ヘルパー
    // =========================================================================

    private static string ResolveInnerDocName(string? documentPath, string editorType)
    {
        if (documentPath != null)
        {
            var ext = Path.GetExtension(documentPath).ToLowerInvariant();
            return $"document{ext}";
        }
        return editorType switch
        {
            "word" => "document.docx",
            "excel" => "document.xlsx",
            _ => "document.txt",
        };
    }

    private static AiChatHistory BuildChatHistory(IEnumerable<ChatMessageVm> messages)
    {
        var session = new ChatSession
        {
            StartedAt = DateTime.UtcNow.ToString("o"),
            Messages = messages.Select(m => new InsightCommon.ProjectFile.ChatMessage
            {
                Role = m.Role.ToString().ToLowerInvariant(),
                Content = m.Content,
                Timestamp = m.Timestamp.ToString("o"),
            }).ToList(),
        };

        return new AiChatHistory
        {
            Sessions = session.Messages.Count > 0 ? [session] : [],
        };
    }

    private async Task WriteJsonAsync<T>(string entryPath, T data)
    {
        var fullPath = Path.Combine(_workDir!, entryPath);
        var dir = Path.GetDirectoryName(fullPath);
        if (dir != null && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
        var json = JsonSerializer.Serialize(data, JsonOpts);
        await File.WriteAllTextAsync(fullPath, json);
    }

    private async Task<T?> ReadJsonAsync<T>(string entryPath) where T : class
    {
        var fullPath = Path.Combine(_workDir!, entryPath);
        if (!File.Exists(fullPath)) return null;
        var json = await File.ReadAllTextAsync(fullPath);
        return JsonSerializer.Deserialize<T>(json, JsonOpts);
    }

    private async Task PackAsync()
    {
        if (_workDir == null || ProjectPath == null) return;
        var tmpPath = ProjectPath + ".tmp";
        if (File.Exists(tmpPath)) File.Delete(tmpPath);
        await Task.Run(() => ZipFile.CreateFromDirectory(_workDir, tmpPath));
        File.Move(tmpPath, ProjectPath, overwrite: true);
    }

    private static string CreateTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "IAOF_Project", Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(dir);
        return dir;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        if (_workDir != null && Directory.Exists(_workDir))
        {
            try { Directory.Delete(_workDir, true); } catch { }
        }
    }
}
