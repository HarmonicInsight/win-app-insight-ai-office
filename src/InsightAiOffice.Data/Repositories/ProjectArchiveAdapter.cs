using System.IO.Compression;
using System.Text.Json;

namespace InsightAiOffice.Data.Repositories;

/// <summary>
/// .iaof プロジェクトファイル (ZIP) の読み書きアダプター。
/// ZIP 内部: metadata.json + document.{docx|xlsx|pptx} + ai_chat_history.json + references/
/// </summary>
public class ProjectArchiveAdapter : IDisposable
{
    private string? _tempDir;

    /// <summary>Maximum total extracted size (200 MB) — ZIP bomb protection.</summary>
    private const long MaxExtractedSize = 200 * 1024 * 1024;

    /// <summary>Maximum number of entries in a project file.</summary>
    private const int MaxEntryCount = 500;

    public ProjectMetadata? Metadata { get; private set; }
    public string? DocumentPath { get; private set; }
    public string? ChatHistoryPath { get; private set; }

    /// <summary>Opens a .iaof file by extracting to temp directory.</summary>
    public void Open(string iaofPath)
    {
        if (!File.Exists(iaofPath))
            throw new FileNotFoundException("Project file not found", iaofPath);

        _tempDir = Path.Combine(Path.GetTempPath(), "IAOF_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(_tempDir);

        ExtractWithSecurityChecks(iaofPath, _tempDir);

        var metaPath = Path.Combine(_tempDir, "metadata.json");
        if (File.Exists(metaPath))
        {
            var json = File.ReadAllText(metaPath);
            Metadata = JsonSerializer.Deserialize<ProjectMetadata>(json);
        }

        DocumentPath = FindDocument(_tempDir);
        if (DocumentPath == null)
            System.Diagnostics.Debug.WriteLine("[ProjectArchiveAdapter] No document found in project archive");

        ChatHistoryPath = Path.Combine(_tempDir, "ai_chat_history.json");
    }

    /// <summary>
    /// Extracts a ZIP with path traversal and ZIP bomb protection.
    /// </summary>
    private static void ExtractWithSecurityChecks(string zipPath, string destinationDir)
    {
        var fullDestination = Path.GetFullPath(destinationDir);
        long totalSize = 0;

        using var archive = ZipFile.OpenRead(zipPath);

        if (archive.Entries.Count > MaxEntryCount)
            throw new InvalidDataException($"Project file contains too many entries ({archive.Entries.Count} > {MaxEntryCount})");

        foreach (var entry in archive.Entries)
        {
            var destPath = Path.GetFullPath(Path.Combine(fullDestination, entry.FullName));

            // Path traversal check: ensure extracted path is within destination
            if (!destPath.StartsWith(fullDestination + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                && !destPath.Equals(fullDestination, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException($"Path traversal detected in ZIP entry: {entry.FullName}");
            }

            // ZIP bomb check: accumulate total uncompressed size
            totalSize += entry.Length;
            if (totalSize > MaxExtractedSize)
                throw new InvalidDataException($"Project file exceeds maximum allowed size ({MaxExtractedSize / (1024 * 1024)} MB)");

            // Create directories for entry
            var entryDir = Path.GetDirectoryName(destPath);
            if (entryDir != null)
                Directory.CreateDirectory(entryDir);

            // Skip directory entries
            if (string.IsNullOrEmpty(entry.Name)) continue;

            entry.ExtractToFile(destPath, overwrite: true);
        }
    }

    /// <summary>Creates a new .iaof project from an existing document.</summary>
    public static string CreateFromDocument(string documentPath, string outputPath, string? author = null)
    {
        var ext = Path.GetExtension(documentPath).ToLowerInvariant();
        var innerName = "document" + ext;
        var docType = ext switch
        {
            ".docx" or ".doc" => "word",
            ".xlsx" or ".xls" or ".csv" => "excel",
            ".pptx" or ".ppt" => "pptx",
            _ => "unknown"
        };

        var tempDir = Path.Combine(Path.GetTempPath(), "IAOF_new_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(tempDir);

        try
        {
            // Copy document
            File.Copy(documentPath, Path.Combine(tempDir, innerName));

            // Create metadata
            var metadata = new ProjectMetadata
            {
                Version = "1.0",
                ProductCode = "IAOF",
                DocumentType = docType,
                OriginalFileName = Path.GetFileName(documentPath),
                InnerDocumentName = innerName,
                Author = author ?? Environment.UserName,
                CreatedAt = DateTime.UtcNow.ToString("o"),
                LastModifiedAt = DateTime.UtcNow.ToString("o"),
            };

            var metaJson = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(Path.Combine(tempDir, "metadata.json"), metaJson);

            // Empty chat history
            File.WriteAllText(Path.Combine(tempDir, "ai_chat_history.json"), "[]");

            // References directory
            Directory.CreateDirectory(Path.Combine(tempDir, "references"));

            // Create ZIP
            if (File.Exists(outputPath)) File.Delete(outputPath);
            ZipFile.CreateFromDirectory(tempDir, outputPath);

            return outputPath;
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); }
            catch (IOException) { /* cleanup best-effort — file may be locked */ }
        }
    }

    /// <summary>Saves modified content back to a .iaof file.</summary>
    public void Save(string outputPath, string? updatedDocumentPath = null)
    {
        if (_tempDir == null)
            throw new InvalidOperationException("No project is open");

        // Update document if provided
        if (updatedDocumentPath != null && DocumentPath != null)
        {
            File.Copy(updatedDocumentPath, DocumentPath, overwrite: true);
        }

        // Update metadata timestamp
        if (Metadata != null)
        {
            Metadata.LastModifiedAt = DateTime.UtcNow.ToString("o");
            var metaJson = JsonSerializer.Serialize(Metadata, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(Path.Combine(_tempDir, "metadata.json"), metaJson);
        }

        // Recreate ZIP (atomic: write to temp, then move)
        var tempZip = outputPath + ".tmp";
        if (File.Exists(tempZip)) File.Delete(tempZip);
        ZipFile.CreateFromDirectory(_tempDir, tempZip);

        if (File.Exists(outputPath)) File.Delete(outputPath);
        File.Move(tempZip, outputPath);
    }

    private static string? FindDocument(string dir)
    {
        foreach (var ext in new[] { ".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt", ".csv" })
        {
            var candidates = Directory.GetFiles(dir, $"document{ext}");
            if (candidates.Length > 0) return candidates[0];
        }
        // Fallback: any Office file
        foreach (var ext in new[] { "*.docx", "*.xlsx", "*.pptx" })
        {
            var candidates = Directory.GetFiles(dir, ext);
            if (candidates.Length > 0) return candidates[0];
        }
        return null;
    }

    public void Dispose()
    {
        if (_tempDir != null)
        {
            try { Directory.Delete(_tempDir, recursive: true); }
            catch (IOException) { /* best-effort — temp file may be locked */ }
            _tempDir = null;
        }
    }
}

public class ProjectMetadata
{
    public string Version { get; set; } = "1.0";
    public string ProductCode { get; set; } = "IAOF";
    public string DocumentType { get; set; } = "";
    public string OriginalFileName { get; set; } = "";
    public string InnerDocumentName { get; set; } = "";
    public string Author { get; set; } = "";
    public string CreatedAt { get; set; } = "";
    public string LastModifiedAt { get; set; } = "";
}
