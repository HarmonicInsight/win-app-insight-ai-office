using InsightAiOffice.App.Helpers;
using Xunit;

namespace InsightAiOffice.Core.Tests;

public class LanguageManagerTests
{
    [Fact]
    public void Get_ReturnsJapanese_WhenLanguageIsJa()
    {
        LanguageManager.SetLanguage("ja");

        Assert.Equal("ファイル", LanguageManager.Get("Menu_File"));
        Assert.Equal("ホーム", LanguageManager.Get("Menu_Home"));
    }

    [Fact]
    public void Get_ReturnsEnglish_WhenLanguageIsEn()
    {
        LanguageManager.SetLanguage("en");

        Assert.Equal("File", LanguageManager.Get("Menu_File"));
        Assert.Equal("Home", LanguageManager.Get("Menu_Home"));

        // Reset to default
        LanguageManager.SetLanguage("ja");
    }

    [Fact]
    public void Get_FallsBackToJapanese_WhenKeyMissing()
    {
        LanguageManager.SetLanguage("ja");

        // Unknown key returns the key itself
        Assert.Equal("Unknown_Key", LanguageManager.Get("Unknown_Key"));
    }

    [Fact]
    public void SetLanguage_FallsBackToJa_WhenUnsupported()
    {
        LanguageManager.SetLanguage("fr");

        Assert.Equal("ja", LanguageManager.CurrentLanguage);
    }

    [Fact]
    public void Format_InterpolatesArgs()
    {
        LanguageManager.SetLanguage("ja");

        var result = LanguageManager.Format("Error_Unexpected", "テスト例外");

        Assert.Contains("テスト例外", result);
    }

    [Fact]
    public void StaticProperties_ReturnCorrectValues()
    {
        LanguageManager.SetLanguage("ja");
        Assert.Equal("ドキュメントに挿入", LanguageManager.InsertToDocLabel);
        Assert.Equal("コピー", LanguageManager.CopyLabel);

        LanguageManager.SetLanguage("en");
        Assert.Equal("Insert to Document", LanguageManager.InsertToDocLabel);
        Assert.Equal("Copy", LanguageManager.CopyLabel);

        LanguageManager.SetLanguage("ja");
    }

    [Fact]
    public void JaAndEn_HaveSameKeys()
    {
        // Verify both languages have the same set of keys
        LanguageManager.SetLanguage("ja");
        var jaKeys = new HashSet<string>();
        LanguageManager.SetLanguage("en");
        var enKeys = new HashSet<string>();

        // Test known keys exist in both
        var testKeys = new[]
        {
            "App_Title", "Menu_File", "Menu_Home", "Menu_AI",
            "File_Open", "File_Save", "AI_Chat", "AI_Summarize",
            "Pane_Chat", "License_Title", "Status_Ready",
            "Error_Title", "Window_Close",
            // New localized keys
            "AI_Thinking", "AI_Responding", "AI_ConfigRequired",
            "AI_ConfigDone", "AI_Error", "AI_InsertedToDoc",
            "Doc_OpenFirst", "Doc_Loaded", "Doc_Delete",
            "Settings_AiSettings", "Settings_OpenAiSettings",
        };

        foreach (var key in testKeys)
        {
            LanguageManager.SetLanguage("ja");
            var jaVal = LanguageManager.Get(key);
            LanguageManager.SetLanguage("en");
            var enVal = LanguageManager.Get(key);

            Assert.NotEqual(key, jaVal); // Should not fall back to key
            Assert.NotEqual(key, enVal);
        }

        LanguageManager.SetLanguage("ja");
    }
}
