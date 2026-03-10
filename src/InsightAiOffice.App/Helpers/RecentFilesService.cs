using System.IO;
using System.Text.Json;

namespace InsightAiOffice.App.Helpers;

public class RecentFileEntry
{
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public DateTime LastOpened { get; set; }
}

public class RecentFilesService
{
    private const int MaxEntries = 10;
    private readonly string _filePath;
    private List<RecentFileEntry> _entries = [];

    public RecentFilesService(string productCode)
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HarmonicInsight", productCode);
        Directory.CreateDirectory(dir);
        _filePath = Path.Combine(dir, "recent-files.json");
        Load();
    }

    public IReadOnlyList<RecentFileEntry> Entries => _entries;

    public void Add(string filePath)
    {
        var existing = _entries.FindIndex(e =>
            string.Equals(e.FilePath, filePath, StringComparison.OrdinalIgnoreCase));
        if (existing >= 0)
            _entries.RemoveAt(existing);

        _entries.Insert(0, new RecentFileEntry
        {
            FilePath = filePath,
            FileName = Path.GetFileName(filePath),
            LastOpened = DateTime.Now,
        });

        if (_entries.Count > MaxEntries)
            _entries.RemoveRange(MaxEntries, _entries.Count - MaxEntries);

        Save();
    }

    private void Load()
    {
        try
        {
            if (File.Exists(_filePath))
            {
                var json = File.ReadAllText(_filePath);
                _entries = JsonSerializer.Deserialize<List<RecentFileEntry>>(json) ?? [];
            }
        }
        catch
        {
            _entries = [];
        }
    }

    private void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(_entries, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_filePath, json);
        }
        catch
        {
            // ignore write errors
        }
    }
}
