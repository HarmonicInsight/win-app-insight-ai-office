// C# Script (.csx) to extract presets from sister apps and merge into IAOF BuiltInPresets.cs
// Run with: dotnet script tools/merge-presets.csx

using System.Text;
using System.Text.RegularExpressions;

var slideFile = @"C:\dev\win-app-insight-slide\src\InsightOfficeSlide\Services\PromptPresetService.cs";
var sheetFile = @"C:\dev\win-app-insight-sheet\src\HarmonicSheet.App\Helpers\PromptPresetService.cs";
var docFile   = @"C:\dev\win-app-insight-doc\src\HarmonicDoc.Core\Services\PromptPresetService.cs";

string ExtractPresetBlock(string filePath)
{
    var content = File.ReadAllText(filePath);
    // Find GetBuiltInPresets method body - the return [...] block
    var match = Regex.Match(content, @"private static List<UserPromptPreset> GetBuiltInPresets\(\)\s*\{.*?return\s*\[", RegexOptions.Singleline);
    if (!match.Success) { Console.WriteLine($"ERROR: Could not find GetBuiltInPresets in {filePath}"); return ""; }

    int startIdx = match.Index + match.Length;
    // Find matching closing bracket, counting nesting
    int depth = 1;
    int i = startIdx;
    while (i < content.Length && depth > 0)
    {
        if (content[i] == '[') depth++;
        else if (content[i] == ']') depth--;
        i++;
    }
    return content.Substring(startIdx, i - startIdx - 1).Trim();
}

string TransformPresets(string block, string idPrefix, string categoryPrefix, string emoji)
{
    // Replace builtin_ with prefix_
    var result = block.Replace("Id = \"builtin_", $"Id = \"{idPrefix}");

    // Replace Category = "xxx" with Category = "emoji prefix: xxx"
    result = Regex.Replace(result, @"Category\s*=\s*""([^""]+)""", m =>
        $"Category = \"{emoji} {categoryPrefix}: {m.Groups[1].Value}\"");

    // Remove IsDefault = true (only IAOF's first preset should be default)
    result = Regex.Replace(result, @",?\s*IsDefault\s*=\s*true\s*,?", m => {
        var s = m.Value;
        // Clean up comma handling
        if (s.StartsWith(",") && s.EndsWith(",")) return ",";
        return "";
    });

    // Remove "var now = new DateTime(2025, 1, 1);" lines if present
    result = Regex.Replace(result, @"\s*var now = new DateTime\([^)]+\);\s*", "");

    return result;
}

Console.WriteLine("Reading source files...");

var slideBlock = ExtractPresetBlock(slideFile);
var sheetBlock = ExtractPresetBlock(sheetFile);
var docBlock   = ExtractPresetBlock(docFile);

Console.WriteLine($"SLIDE block: {slideBlock.Length} chars");
Console.WriteLine($"SHEET block: {sheetBlock.Length} chars");
Console.WriteLine($"DOC block: {docBlock.Length} chars");

var slidePresets = TransformPresets(slideBlock, "slide_", "Slide", "📊");
var sheetPresets = TransformPresets(sheetBlock, "sheet_", "Sheet", "📗");
var docPresets   = TransformPresets(docBlock, "doc_", "Doc", "📄");

// Count presets (approximate by counting "new()")
int Count(string s) => Regex.Matches(s, @"\bnew\(\)").Count;
Console.WriteLine($"SLIDE presets: {Count(slidePresets)}");
Console.WriteLine($"SHEET presets: {Count(sheetPresets)}");
Console.WriteLine($"DOC presets: {Count(docPresets)}");

// Write the merged output
var outputPath = @"C:\dev\win-app-insight-ai-office\tools\merged-presets-fragment.cs";
var sb = new StringBuilder();
sb.AppendLine("        // ════════════════════════════════════════════════════════════");
sb.AppendLine("        // 📊 InsightSlide からインポート");
sb.AppendLine("        // ════════════════════════════════════════════════════════════");
sb.AppendLine(slidePresets);
sb.AppendLine();
sb.AppendLine("        // ════════════════════════════════════════════════════════════");
sb.AppendLine("        // 📗 InsightSheet からインポート");
sb.AppendLine("        // ════════════════════════════════════════════════════════════");
sb.AppendLine(sheetPresets);
sb.AppendLine();
sb.AppendLine("        // ════════════════════════════════════════════════════════════");
sb.AppendLine("        // 📄 InsightDoc からインポート");
sb.AppendLine("        // ════════════════════════════════════════════════════════════");
sb.AppendLine(docPresets);

File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
Console.WriteLine($"\nMerged fragment written to: {outputPath}");
Console.WriteLine($"Total size: {sb.Length} chars");
