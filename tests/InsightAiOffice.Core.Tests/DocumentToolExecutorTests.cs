using InsightAiOffice.App.Services;
using Xunit;

namespace InsightAiOffice.Core.Tests;

public class DocumentToolExecutorTests
{
    private static DocumentToolExecutor CreateTestExecutor(List<string>? insertedTexts = null)
    {
        insertedTexts ??= [];
        return new DocumentToolExecutor(
            insertText: text => insertedTexts.Add(text),
            getSelectedText: () => "",
            setStatus: _ => { });
    }

    [Fact]
    public void ParseAndExecute_ReturnsPlainText_WhenNoToolCall()
    {
        var executor = CreateTestExecutor();

        var (cleaned, count) = executor.ParseAndExecute("This is a normal response.");

        Assert.Equal("This is a normal response.", cleaned);
        Assert.Equal(0, count);
    }

    [Fact]
    public void ParseAndExecute_DetectsInsertTextTool()
    {
        var insertedTexts = new List<string>();
        var executor = CreateTestExecutor(insertedTexts);

        var response = """
            Here is my suggestion:
            {"tool":"insert_text","args":{"text":"Hello World"}}
            """;

        var (_, count) = executor.ParseAndExecute(response);

        Assert.Equal(1, count);
        Assert.Single(insertedTexts);
        Assert.Equal("Hello World", insertedTexts[0]);
    }

    [Fact]
    public void ParseAndExecute_ExecutesMultipleToolCalls()
    {
        var insertedTexts = new List<string>();
        var executor = CreateTestExecutor(insertedTexts);

        var response = """
            {"tool":"insert_text","args":{"text":"First"}}
            {"tool":"insert_text","args":{"text":"Second"}}
            """;

        var (_, count) = executor.ParseAndExecute(response);

        Assert.Equal(2, count);
        Assert.Equal(2, insertedTexts.Count);
    }

    [Fact]
    public void ParseAndExecute_HandlesReplaceSelection()
    {
        var insertedTexts = new List<string>();
        var executor = CreateTestExecutor(insertedTexts);

        var response = """{"tool":"replace_selection","args":{"text":"Replaced"}}""";

        var (_, count) = executor.ParseAndExecute(response);

        Assert.Equal(1, count);
        Assert.Equal("Replaced", insertedTexts[0]);
    }
}
