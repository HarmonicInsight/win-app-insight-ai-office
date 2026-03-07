using InsightAiOffice.Core.Services;
using Xunit;

namespace InsightAiOffice.Core.Tests;

public class PromptEntryTests
{
    [Fact]
    public void PromptEntry_Record_PropertiesAreCorrect()
    {
        var entry = new PromptEntry("id1", "Title", "Content", "General", DateTime.UtcNow);

        Assert.Equal("id1", entry.Id);
        Assert.Equal("Title", entry.Title);
        Assert.Equal("Content", entry.Content);
        Assert.Equal("General", entry.Category);
    }

    [Fact]
    public void PromptEntry_Record_SupportsEquality()
    {
        var dt = DateTime.UtcNow;
        var a = new PromptEntry("id1", "Title", "Content", "General", dt);
        var b = new PromptEntry("id1", "Title", "Content", "General", dt);

        Assert.Equal(a, b);
    }
}
