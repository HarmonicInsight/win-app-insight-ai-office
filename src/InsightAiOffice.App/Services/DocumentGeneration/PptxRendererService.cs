using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using InsightAiOffice.App.Services.DocumentGeneration;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// SlideSpecItem リスト → .pptx レンダリング（Open XML SDK）
///
/// Ivory & Gold テーマ（Primary: #B8942F）
/// </summary>
public static class PptxRendererService
{
    private const long SlideWidth = 12192000;
    private const long SlideHeight = 6858000;
    private const long MarginLeft = 457200;
    private const long MarginTop = 914400;
    private const long MarginRight = 457200;
    private const long ContentWidth = SlideWidth - MarginLeft - MarginRight;

    private const string FontName = "Yu Gothic";

    // Ivory & Gold カラーパレット
    private const string ColorPrimary = "B8942F";    // Gold
    private const string ColorAccent = "8B6F1F";     // Dark Gold
    private const string ColorText = "333333";
    private const string ColorSubText = "666666";
    private const string ColorWhite = "FFFFFF";
    private const string ColorBgIvory = "FAF8F5";

    /// <summary>
    /// 新規 PPTX を生成
    /// </summary>
    public static Task<string> RenderAsync(
        List<SlideSpecItem> slides, string outputPath, CancellationToken ct = default)
    {
        return RenderAsync(slides, outputPath, templatePath: null, ct);
    }

    /// <summary>
    /// PPTX を生成（テンプレート指定時は既存ファイルにスライド追加）
    /// </summary>
    public static Task<string> RenderAsync(
        List<SlideSpecItem> slides, string outputPath, string? templatePath, CancellationToken ct = default)
    {
        return Task.Run(() =>
        {
            // テンプレートモード: 既存 PPTX をコピーしてスライド追加
            if (!string.IsNullOrEmpty(templatePath) && File.Exists(templatePath))
            {
                return RenderFromTemplate(slides, outputPath, templatePath, ct);
            }

            // ビルトインテンプレートベースで新規作成（プレミアム品質）
            var builtinTemplate = PptxTemplateBuilder.EnsureTemplate();
            if (File.Exists(builtinTemplate))
            {
                return RenderFromTemplate(slides, outputPath, builtinTemplate, ct);
            }

            // フォールバック: テンプレートなしで新規作成
            using var presentation = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation);
            var presentationPart = presentation.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            presentationPart.Presentation.SlideSize = new SlideSize
            {
                Cx = (int)SlideWidth, Cy = (int)SlideHeight, Type = SlideSizeValues.Custom
            };
            presentationPart.Presentation.NotesSize = new NotesSize { Cx = 6858000, Cy = 9144000 };

            var slideIdList = new SlideIdList();
            presentationPart.Presentation.SlideIdList = slideIdList;

            CreateSlideMaster(presentationPart);

            uint slideId = 256;
            foreach (var spec in slides.OrderBy(s => s.Order))
            {
                ct.ThrowIfCancellationRequested();
                var slidePart = presentationPart.AddNewPart<SlidePart>($"rId{slideId}");
                RenderSlide(slidePart, spec);
                slideIdList.Append(new SlideId { Id = slideId, RelationshipId = $"rId{slideId}" });
                slideId++;
            }

            presentationPart.Presentation.Save();
            return outputPath;
        }, ct);
    }

    /// <summary>
    /// 既存 PPTX テンプレートにスライドを追加
    /// </summary>
    private static string RenderFromTemplate(
        List<SlideSpecItem> slides, string outputPath, string templatePath, CancellationToken ct)
    {
        // テンプレートをコピー
        File.Copy(templatePath, outputPath, overwrite: true);

        using var presentation = PresentationDocument.Open(outputPath, true);
        var presentationPart = presentation.PresentationPart
            ?? throw new InvalidOperationException("Invalid PPTX template");

        var slideIdList = presentationPart.Presentation.SlideIdList
            ?? (presentationPart.Presentation.SlideIdList = new SlideIdList());

        // 既存スライドの最大 ID を取得
        uint maxId = 256;
        int maxRId = 0;
        foreach (var existingSlide in slideIdList.Elements<SlideId>())
        {
            if (existingSlide.Id != null && existingSlide.Id > maxId)
                maxId = existingSlide.Id;
            if (existingSlide.RelationshipId?.Value is string rId &&
                rId.StartsWith("rId") &&
                int.TryParse(rId[3..], out var rIdNum) &&
                rIdNum > maxRId)
            {
                maxRId = rIdNum;
            }
        }

        uint slideId = maxId + 1;
        int relId = maxRId + 1;

        foreach (var spec in slides.OrderBy(s => s.Order))
        {
            ct.ThrowIfCancellationRequested();
            var relationshipId = $"rId{relId}";
            var slidePart = presentationPart.AddNewPart<SlidePart>(relationshipId);
            RenderSlide(slidePart, spec);
            slideIdList.Append(new SlideId { Id = slideId, RelationshipId = relationshipId });
            slideId++;
            relId++;
        }

        presentationPart.Presentation.Save();
        return outputPath;
    }

    private static void CreateSlideMaster(PresentationPart presentationPart)
    {
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rIdMaster");
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rIdLayout");

        slideLayoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))),
            new ColorMapOverride(new A.MasterColorMapping()));
        slideLayoutPart.SlideLayout.Save();

        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))),
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

        var themePart = slideMasterPart.AddNewPart<ThemePart>("rIdTheme");
        themePart.Theme = CreateTheme();
        themePart.Theme.Save();
        slideMasterPart.SlideMaster.Save();

        presentationPart.Presentation!.SlideMasterIdList = new SlideMasterIdList(
            new SlideMasterId { Id = 2147483648, RelationshipId = "rIdMaster" });
    }

    private static A.Theme CreateTheme()
    {
        return new A.Theme(
            new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = ColorText }),
                    new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = ColorWhite }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = ColorPrimary }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = ColorBgIvory }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = ColorPrimary }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "D4A843" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "70AD47" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "FFC000" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "44546A" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" })
                ) { Name = "Ivory & Gold" },
                new A.FontScheme(
                    new A.MajorFont(new A.LatinFont { Typeface = FontName },
                        new A.EastAsianFont { Typeface = FontName },
                        new A.ComplexScriptFont { Typeface = FontName }),
                    new A.MinorFont(new A.LatinFont { Typeface = FontName },
                        new A.EastAsianFont { Typeface = FontName },
                        new A.ComplexScriptFont { Typeface = FontName })
                ) { Name = "Ivory & Gold" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                    new A.LineStyleList(
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 9525 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 9525 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 9525 }),
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

    private static void RenderSlide(SlidePart slidePart, SlideSpecItem spec)
    {
        var shapeTree = new ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        switch (spec.SlideType)
        {
            case SlideType.Title:
                RenderTitleSlide(shapeTree, spec);
                break;
            case SlideType.Agenda:
                RenderAgendaSlide(shapeTree, spec);
                break;
            case SlideType.Data:
                RenderDataSlide(shapeTree, spec);
                break;
            default:
                RenderContentSlide(shapeTree, spec);
                break;
        }

        slidePart.Slide = new Slide(new CommonSlideData(shapeTree));

        if (!string.IsNullOrEmpty(spec.SpeakerNotes))
            AddSpeakerNotes(slidePart, spec.SpeakerNotes);

        slidePart.Slide.Save();
    }

    private static void RenderTitleSlide(ShapeTree tree, SlideSpecItem spec)
    {
        AddShape(tree, 2, "bg", 0, 0, SlideWidth, SlideHeight, ColorPrimary);

        AddTextBox(tree, 3, "title", MarginLeft, SlideHeight / 3,
            ContentWidth, 1200000, spec.Title, 4000, true, ColorWhite,
            A.TextAlignmentTypeValues.Center);

        if (!string.IsNullOrEmpty(spec.KeyMessage))
        {
            AddTextBox(tree, 4, "keymsg", MarginLeft, SlideHeight / 3 + 1400000,
                ContentWidth, 800000, spec.KeyMessage, 2000, false, ColorBgIvory,
                A.TextAlignmentTypeValues.Center);
        }
    }

    private static void RenderAgendaSlide(ShapeTree tree, SlideSpecItem spec)
    {
        AddTextBox(tree, 2, "title", MarginLeft, 274320,
            ContentWidth, 640000, spec.Title, 2800, true, ColorPrimary);

        long yPos = MarginTop + 200000;
        for (int i = 0; i < spec.Bullets.Count; i++)
        {
            AddTextBox(tree, (uint)(3 + i), $"item{i}", MarginLeft + 200000, yPos,
                ContentWidth - 400000, 450000, $"{i + 1}. {spec.Bullets[i]}",
                2000, false, ColorText);
            yPos += 500000;
        }
    }

    private static void RenderDataSlide(ShapeTree tree, SlideSpecItem spec)
    {
        AddTextBox(tree, 2, "title", MarginLeft, 274320,
            ContentWidth, 640000, spec.Title, 2800, true, ColorPrimary);

        if (!string.IsNullOrEmpty(spec.KeyMessage))
        {
            AddTextBox(tree, 3, "keymsg", MarginLeft, MarginTop,
                ContentWidth, 400000, spec.KeyMessage, 1600, false, ColorAccent);
        }

        long yPos = MarginTop + 600000;
        for (int i = 0; i < spec.Bullets.Count && i < 5; i++)
        {
            AddTextBox(tree, (uint)(4 + i), $"data{i}", MarginLeft + 100000, yPos,
                ContentWidth - 200000, 350000, $"  {spec.Bullets[i]}", 1600, false, ColorText);
            yPos += 400000;
        }
    }

    private static void RenderContentSlide(ShapeTree tree, SlideSpecItem spec)
    {
        // Gold accent bar
        AddShape(tree, 2, "accent-bar", 0, 0, SlideWidth, 60000, ColorPrimary);

        AddTextBox(tree, 3, "title", MarginLeft, 150000,
            ContentWidth, 640000, spec.Title, 2800, true, ColorAccent);

        long contentTop = MarginTop;
        if (!string.IsNullOrEmpty(spec.KeyMessage))
        {
            AddTextBox(tree, 4, "keymsg", MarginLeft, contentTop,
                ContentWidth, 400000, spec.KeyMessage, 1600, true, ColorPrimary);
            contentTop += 500000;
        }

        int fontSize = spec.Bullets.Count > 7 ? 1400 : spec.Bullets.Count > 5 ? 1600 : 1800;
        long yPos = contentTop;
        int maxBullets = Math.Min(spec.Bullets.Count, 7);
        for (int i = 0; i < maxBullets; i++)
        {
            AddTextBox(tree, (uint)(5 + i), $"bullet{i}", MarginLeft + 200000, yPos,
                ContentWidth - 400000, 380000, $"  {spec.Bullets[i]}",
                fontSize, false, ColorText);
            yPos += 420000;
        }

        if (spec.Bullets.Count > 7)
        {
            var overflow = string.Join("\n", spec.Bullets.Skip(7).Select(b => $"  {b}"));
            spec.SpeakerNotes = $"[追加項目]\n{overflow}\n\n{spec.SpeakerNotes}";
        }
    }

    private static void AddTextBox(ShapeTree tree, uint id, string name,
        long x, long y, long cx, long cy,
        string text, int fontSizeHundredths, bool bold, string colorHex,
        A.TextAlignmentTypeValues? align = null)
    {
        var run = new A.Run(
            new A.RunProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = colorHex }),
                new A.LatinFont { Typeface = FontName },
                new A.EastAsianFont { Typeface = FontName })
            { Language = "ja-JP", FontSize = fontSizeHundredths, Bold = bold },
            new A.Text(text));

        var paragraph = new A.Paragraph(
            new A.ParagraphProperties { Alignment = align ?? A.TextAlignmentTypeValues.Left },
            run);

        var shape = new P.Shape(
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
                new A.ListStyle(), paragraph));

        tree.Append(shape);
    }

    private static void AddShape(ShapeTree tree, uint id, string name,
        long x, long y, long cx, long cy, string solidFill)
    {
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.SolidFill(new A.RgbColorModelHex { Val = solidFill })),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));

        tree.Append(shape);
    }

    private static void AddSpeakerNotes(SlidePart slidePart, string notesText)
    {
        var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
        var run = new A.Run(
            new A.RunProperties { Language = "ja-JP", FontSize = 1200 },
            new A.Text(notesText));

        notesSlidePart.NotesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 2, Name = "Notes" },
                            new P.NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1 })),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new A.BodyProperties(), new A.ListStyle(),
                            new A.Paragraph(run))))));
        notesSlidePart.NotesSlide.Save();
    }
}
