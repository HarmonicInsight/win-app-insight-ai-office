using System.IO;
using System.IO.Compression;
using InsightAiOffice.Data.Repositories;
using Xunit;

namespace InsightAiOffice.Core.Tests;

public class ProjectArchiveAdapterTests
{
    [Fact]
    public void CreateFromDocument_CreatesValidZip()
    {
        var tempDoc = Path.GetTempFileName();
        var tempProject = Path.ChangeExtension(Path.GetTempFileName(), ".iaof");
        try
        {
            File.WriteAllText(tempDoc, "test content");

            var result = ProjectArchiveAdapter.CreateFromDocument(tempDoc, tempProject);

            Assert.Equal(tempProject, result);
            Assert.True(File.Exists(tempProject));
        }
        finally
        {
            if (File.Exists(tempDoc)) File.Delete(tempDoc);
            if (File.Exists(tempProject)) File.Delete(tempProject);
        }
    }

    [Fact]
    public void Open_LoadsExistingProject()
    {
        var tempDoc = Path.ChangeExtension(Path.GetTempFileName(), ".docx");
        var tempProject = Path.ChangeExtension(Path.GetTempFileName(), ".iaof");
        try
        {
            File.WriteAllText(tempDoc, "test content");
            ProjectArchiveAdapter.CreateFromDocument(tempDoc, tempProject);

            using var adapter = new ProjectArchiveAdapter();
            adapter.Open(tempProject);

            Assert.NotNull(adapter.Metadata);
            Assert.Equal("IAOF", adapter.Metadata!.ProductCode);
            Assert.Equal("word", adapter.Metadata.DocumentType);
            Assert.NotNull(adapter.DocumentPath);
        }
        finally
        {
            if (File.Exists(tempDoc)) File.Delete(tempDoc);
            if (File.Exists(tempProject)) File.Delete(tempProject);
        }
    }

    [Fact]
    public void CreateFromDocument_SetsCorrectDocumentType()
    {
        var tempDoc = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
        var tempProject = Path.ChangeExtension(Path.GetTempFileName(), ".iaof");
        try
        {
            File.WriteAllText(tempDoc, "fake excel");
            ProjectArchiveAdapter.CreateFromDocument(tempDoc, tempProject);

            using var adapter = new ProjectArchiveAdapter();
            adapter.Open(tempProject);

            Assert.Equal("excel", adapter.Metadata!.DocumentType);
        }
        finally
        {
            if (File.Exists(tempDoc)) File.Delete(tempDoc);
            if (File.Exists(tempProject)) File.Delete(tempProject);
        }
    }

    [Fact]
    public void Open_RejectsPathTraversal()
    {
        var maliciousZip = Path.ChangeExtension(Path.GetTempFileName(), ".iaof");
        try
        {
            // Create a ZIP with a path traversal entry
            using (var stream = File.Create(maliciousZip))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                var entry = archive.CreateEntry("../../../etc/evil.txt");
                using var writer = new StreamWriter(entry.Open());
                writer.Write("malicious");
            }

            using var adapter = new ProjectArchiveAdapter();
            Assert.Throws<InvalidDataException>(() => adapter.Open(maliciousZip));
        }
        finally
        {
            if (File.Exists(maliciousZip)) File.Delete(maliciousZip);
        }
    }

    [Fact]
    public void Open_RejectsTooManyEntries()
    {
        var bombZip = Path.ChangeExtension(Path.GetTempFileName(), ".iaof");
        try
        {
            using (var stream = File.Create(bombZip))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                for (int i = 0; i < 501; i++)
                {
                    var entry = archive.CreateEntry($"file_{i}.txt");
                    using var writer = new StreamWriter(entry.Open());
                    writer.Write("x");
                }
            }

            using var adapter = new ProjectArchiveAdapter();
            Assert.Throws<InvalidDataException>(() => adapter.Open(bombZip));
        }
        finally
        {
            if (File.Exists(bombZip)) File.Delete(bombZip);
        }
    }

    [Fact]
    public void Open_HandlesNullDocument()
    {
        var emptyZip = Path.ChangeExtension(Path.GetTempFileName(), ".iaof");
        try
        {
            // Create a ZIP with only metadata (no document)
            using (var stream = File.Create(emptyZip))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                var entry = archive.CreateEntry("metadata.json");
                using var writer = new StreamWriter(entry.Open());
                writer.Write("{\"Version\":\"1.0\",\"ProductCode\":\"IAOF\",\"DocumentType\":\"word\"}");
            }

            using var adapter = new ProjectArchiveAdapter();
            adapter.Open(emptyZip);
            Assert.Null(adapter.DocumentPath);
            Assert.NotNull(adapter.Metadata);
        }
        finally
        {
            if (File.Exists(emptyZip)) File.Delete(emptyZip);
        }
    }

    [Fact]
    public void Save_UpdatesLastModified()
    {
        var tempDoc = Path.GetTempFileName();
        var tempProject = Path.ChangeExtension(Path.GetTempFileName(), ".iaof");
        try
        {
            File.WriteAllText(tempDoc, "test content");
            ProjectArchiveAdapter.CreateFromDocument(tempDoc, tempProject);

            using var adapter = new ProjectArchiveAdapter();
            adapter.Open(tempProject);
            var originalModified = adapter.Metadata!.LastModifiedAt;

            // Wait briefly to ensure timestamp difference
            Thread.Sleep(10);
            adapter.Save(tempProject);

            // Reopen and verify
            using var adapter2 = new ProjectArchiveAdapter();
            adapter2.Open(tempProject);
            Assert.NotEqual(originalModified, adapter2.Metadata!.LastModifiedAt);
        }
        finally
        {
            if (File.Exists(tempDoc)) File.Delete(tempDoc);
            if (File.Exists(tempProject)) File.Delete(tempProject);
        }
    }
}
