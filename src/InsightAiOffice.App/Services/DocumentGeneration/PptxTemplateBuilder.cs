using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// Ivory &amp; Gold ビルトインテンプレート PPTX を生成・管理する。
/// ゼロベース作成時もこのテンプレートをベースにすることで、
/// プロフェッショナルなスライドマスター・レイアウトが適用される。
/// </summary>
public static class PptxTemplateBuilder
{
    private const long SW = 12192000;
    private const long SH = 6858000;
    private const string Font = "Yu Gothic UI";

    // Ivory & Gold
    private const string Gold = "B8942F";
    private const string GoldDark = "8A6F23";
    private const string GoldLight = "D4B94A";
    private const string Ivory = "FAF8F5";
    private const string White = "FFFFFF";
    private const string TextDark = "1C1917";
    private const string TextMid = "57534E";

    /// <summary>
    /// テンプレート PPTX のパスを返す。なければ生成する。
    /// </summary>
    public static string EnsureTemplate()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "HarmonicInsight", "InsightAiOffice", "templates");
        Directory.CreateDirectory(dir);

        var path = Path.Combine(dir, "ivory-gold-template.pptx");

        // 既に存在すればそのまま返す
        if (File.Exists(path)) return path;

        Build(path);
        return path;
    }

    /// <summary>
    /// テンプレートを再生成（バージョンアップ時等）
    /// </summary>
    public static string Rebuild()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "HarmonicInsight", "InsightAiOffice", "templates");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "ivory-gold-template.pptx");
        if (File.Exists(path)) File.Delete(path);
        Build(path);
        return path;
    }

    private static void Build(string outputPath)
    {
        using var pres = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation);
        var presPart = pres.AddPresentationPart();
        presPart.Presentation = new Presentation
        {
            SlideSize = new SlideSize { Cx = (int)SW, Cy = (int)SH, Type = SlideSizeValues.Custom },
            NotesSize = new NotesSize { Cx = 6858000, Cy = 9144000 },
        };

        // テーマ + スライドマスター
        var masterPart = presPart.AddNewPart<SlideMasterPart>("rIdMaster");
        var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rIdLayout");

        layoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()));
        layoutPart.SlideLayout.Save();

        // スライドマスターに背景 + ゴールドバー + ロゴテキスト
        var masterTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        // 背景: Ivory
        AddRect(masterTree, 2, "bg", 0, 0, SW, SH, Ivory);

        // 上部ゴールドバー（細）
        AddRect(masterTree, 3, "top-bar", 0, 0, SW, 50000, Gold);

        // 下部バー
        AddRect(masterTree, 4, "bottom-bar", 0, SH - 250000, SW, 250000, "F5F0E8");

        // 下部ゴールドライン
        AddRect(masterTree, 5, "bottom-line", 0, SH - 250000, SW, 18000, Gold);

        // フッター: "HARMONIC insight" テキスト
        AddText(masterTree, 6, "footer", 300000, SH - 200000, 3000000, 180000,
            "HARMONIC insight", 900, false, TextMid, A.TextAlignmentTypeValues.Left);

        // ページ番号エリア
        AddText(masterTree, 7, "slide-num", SW - 1000000, SH - 200000, 700000, 180000,
            "", 900, false, TextMid, A.TextAlignmentTypeValues.Right);

        masterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(masterTree),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
            },
            new SlideLayoutIdList(new SlideLayoutId { Id = 2147483649, RelationshipId = "rIdLayout" }));

        // テーマ
        var themePart = masterPart.AddNewPart<ThemePart>("rIdTheme");
        themePart.Theme = CreatePremiumTheme();
        themePart.Theme.Save();
        masterPart.SlideMaster.Save();

        presPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
            new SlideMasterId { Id = 2147483648, RelationshipId = "rIdMaster" });

        // ダミースライド1枚（テンプレートには最低1枚必要）
        var slideIdList = new SlideIdList();
        presPart.Presentation.SlideIdList = slideIdList;
        var slidePart = presPart.AddNewPart<SlidePart>("rId256");

        var slideTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        slidePart.Slide = new Slide(new CommonSlideData(slideTree));
        slidePart.Slide.Save();
        slideIdList.Append(new SlideId { Id = 256, RelationshipId = "rId256" });

        presPart.Presentation.Save();
    }

    private static A.Theme CreatePremiumTheme()
    {
        return new A.Theme(
            new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.RgbColorModelHex { Val = TextDark }),
                    new A.Light1Color(new A.RgbColorModelHex { Val = White }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = Gold }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = Ivory }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = Gold }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = GoldLight }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "16A34A" }),   // Green
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "2563EB" }),   // Blue
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "DC2626" }),   // Red
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "57534E" }),   // Slate
                    new A.Hyperlink(new A.RgbColorModelHex { Val = Gold }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = GoldDark })
                ) { Name = "Ivory & Gold" },
                new A.FontScheme(
                    new A.MajorFont(
                        new A.LatinFont { Typeface = Font },
                        new A.EastAsianFont { Typeface = Font },
                        new A.ComplexScriptFont { Typeface = Font }),
                    new A.MinorFont(
                        new A.LatinFont { Typeface = Font },
                        new A.EastAsianFont { Typeface = Font },
                        new A.ComplexScriptFont { Typeface = Font })
                ) { Name = "Ivory & Gold" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                    new A.LineStyleList(
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 9525 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 19050 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 25400 }),
                    new A.EffectStyleList(
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList())),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }))
                ) { Name = "Ivory & Gold" }
            )) { Name = "Ivory & Gold" };
    }

    // ── ヘルパー ──

    private static void AddRect(ShapeTree tree, uint id, string name,
        long x, long y, long cx, long cy, string fill)
    {
        tree.Append(new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.SolidFill(new A.RgbColorModelHex { Val = fill }),
                new A.Outline(new A.NoFill())),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph())));
    }

    private static void AddText(ShapeTree tree, uint id, string name,
        long x, long y, long cx, long cy,
        string text, int fontSize100, bool bold, string color,
        A.TextAlignmentTypeValues align)
    {
        var run = new A.Run(
            new A.RunProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = color }),
                new A.LatinFont { Typeface = Font },
                new A.EastAsianFont { Typeface = Font })
            { Language = "ja-JP", FontSize = fontSize100, Bold = bold },
            new A.Text(text));

        tree.Append(new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties { Wrap = A.TextWrappingValues.Square },
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties { Alignment = align },
                    run))));
    }
}
