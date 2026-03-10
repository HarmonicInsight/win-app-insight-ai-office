using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace InsightAiOffice.App.Services;

/// <summary>
/// PPTX サービス。
/// OpenXML でスライド数・テキスト抽出を行い、
/// Syncfusion が利用可能ならレンダリング・スライド操作も提供。
/// </summary>
public static class PptxService
{
    private static Type? _presentationType;
    private static MethodInfo? _openMethod;
    private static object? _imageTypeBitmap;
    private static Type? _slideLayoutType;
    private static bool _resolved;
    private static bool _canRender;

    // ── レンダリング ──

    public static List<(BitmapSource Full, BitmapSource Thumbnail)> RenderAllSlides(
        string pptxPath, int thumbnailWidth = 280)
    {
        var results = new List<(BitmapSource, BitmapSource)>();
        if (!File.Exists(pptxPath)) return results;

        // Syncfusion レンダリングを試行
        if (EnsureSyncfusionResolved() && _canRender)
        {
            try
            {
                return RenderAllSlidesSyncfusion(pptxPath, thumbnailWidth);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PptxService] Syncfusion render failed, falling back to OpenXML: {ex.Message}");
            }
        }

        // フォールバック: OpenXML でスライド数を取得し、テキスト付きプレースホルダーを生成
        return RenderAllSlidesOpenXml(pptxPath, thumbnailWidth);
    }

    private static List<(BitmapSource Full, BitmapSource Thumbnail)> RenderAllSlidesSyncfusion(
        string pptxPath, int thumbnailWidth)
    {
        var results = new List<(BitmapSource, BitmapSource)>();

        using var stream = new FileStream(pptxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        dynamic presentation = _openMethod!.Invoke(null, [stream])!;

        try
        {
            int index = 0;
            foreach (var slide in presentation.Slides)
            {
                index++;
                try
                {
                    System.Drawing.Image image = slide.ConvertToImage(_imageTypeBitmap);
                    using (image)
                    {
                        var full = ConvertToBitmapSource(image, 0) ?? CreatePlaceholder(960, 540, index, null);
                        var thumb = ConvertToBitmapSource(image, thumbnailWidth) ?? CreatePlaceholder(280, 158, index, null);
                        results.Add((full, thumb));
                    }
                }
                catch
                {
                    results.Add((CreatePlaceholder(960, 540, index, null), CreatePlaceholder(280, 158, index, null)));
                }
            }
        }
        finally { ((IDisposable)presentation).Dispose(); }

        return results;
    }

    private static List<(BitmapSource Full, BitmapSource Thumbnail)> RenderAllSlidesOpenXml(
        string pptxPath, int thumbnailWidth)
    {
        var results = new List<(BitmapSource, BitmapSource)>();

        try
        {
            using var doc = PresentationDocument.Open(pptxPath, false);
            var presPart = doc.PresentationPart;
            if (presPart?.Presentation.SlideIdList == null) return results;

            int index = 0;
            foreach (var slideId in presPart.Presentation.SlideIdList.ChildElements.OfType<SlideId>())
            {
                index++;
                string? slideText = null;
                try
                {
                    var slidePart = (SlidePart)presPart.GetPartById(slideId.RelationshipId!);
                    slideText = ExtractSlideText(slidePart, maxChars: 120);
                }
                catch { /* ignore */ }

                var full = CreatePlaceholder(960, 540, index, slideText);
                var thumb = CreatePlaceholder(280, 158, index, slideText != null ? Truncate(slideText, 40) : null);
                results.Add((full, thumb));
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PptxService] OpenXML fallback error: {ex.Message}");
        }

        return results;
    }

    private static string? ExtractSlideText(SlidePart slidePart, int maxChars)
    {
        var texts = new List<string>();
        foreach (var textBody in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>())
        {
            foreach (var para in textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                var text = string.Concat(
                    para.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text));
                if (!string.IsNullOrWhiteSpace(text))
                    texts.Add(text.Trim());
            }
        }

        if (texts.Count == 0) return null;
        var joined = string.Join("\n", texts);
        return joined.Length > maxChars ? joined[..maxChars] + "…" : joined;
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s[..max] + "…";

    public static int GetSlideCount(string pptxPath)
    {
        if (!File.Exists(pptxPath)) return 0;

        // OpenXML でカウント（常に動作する）
        try
        {
            using var doc = PresentationDocument.Open(pptxPath, false);
            return doc.PresentationPart?.Presentation.SlideIdList?.ChildElements
                .OfType<SlideId>().Count() ?? 0;
        }
        catch { return 0; }
    }

    // ── PDF エクスポート ──

    public static void ConvertToPdf(string pptxPath, string outputPath)
    {
        if (!EnsureSyncfusionResolved())
            throw new InvalidOperationException("Syncfusion Presentation not available");

        using var stream = new FileStream(pptxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        dynamic presentation = _openMethod!.Invoke(null, [stream])!;

        try
        {
            var converterType = FindType("Syncfusion.PresentationToPdfConverter.PresentationToPdfConverter");
            if (converterType == null) throw new InvalidOperationException("PresentationToPdfConverter not found");

            var convertMethod = converterType.GetMethod("Convert", [_presentationType!]);
            if (convertMethod == null) throw new InvalidOperationException("Convert method not found");

            dynamic pdfDocument = convertMethod.Invoke(null, [presentation])!;
            try
            {
                using var outStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                pdfDocument.Save(outStream);
            }
            finally { ((IDisposable)pdfDocument).Dispose(); }
        }
        finally { ((IDisposable)presentation).Dispose(); }
    }

    // ── スライド操作 ──

    public static void AddSlide(string pptxPath, int afterSlideIndex)
    {
        WithPresentation(pptxPath, (presentation, stream) =>
        {
            var blankLayout = Enum.Parse(_slideLayoutType!, "Blank");
            presentation.Slides.Add(blankLayout);

            int count = presentation.Slides.Count;
            if (afterSlideIndex + 1 < count - 1)
                presentation.Slides.MoveTo(count - 1, afterSlideIndex + 1);

            SavePresentation(presentation, stream);
        });
    }

    public static void DuplicateSlide(string pptxPath, int slideIndex)
    {
        WithPresentation(pptxPath, (presentation, stream) =>
        {
            if (slideIndex < 0 || slideIndex >= presentation.Slides.Count) return;

            var slide = presentation.Slides[slideIndex].Clone();
            presentation.Slides.Insert(slideIndex + 1, slide);

            SavePresentation(presentation, stream);
        });
    }

    public static void DeleteSlide(string pptxPath, int slideIndex)
    {
        WithPresentation(pptxPath, (presentation, stream) =>
        {
            if (presentation.Slides.Count <= 1) return;
            if (slideIndex < 0 || slideIndex >= presentation.Slides.Count) return;

            presentation.Slides.RemoveAt(slideIndex);

            SavePresentation(presentation, stream);
        });
    }

    public static void MoveSlide(string pptxPath, int fromIndex, int toIndex)
    {
        WithPresentation(pptxPath, (presentation, stream) =>
        {
            if (fromIndex < 0 || fromIndex >= presentation.Slides.Count) return;
            if (toIndex < 0 || toIndex >= presentation.Slides.Count) return;

            presentation.Slides.MoveTo(fromIndex, toIndex);

            SavePresentation(presentation, stream);
        });
    }

    // ── ヘルパー ──

    private static void WithPresentation(string pptxPath, Action<dynamic, FileStream> action)
    {
        if (!EnsureSyncfusionResolved()) return;

        using var stream = new FileStream(pptxPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
        dynamic presentation = _openMethod!.Invoke(null, [stream])!;
        try
        {
            action(presentation, stream);
        }
        finally { ((IDisposable)presentation).Dispose(); }
    }

    private static void SavePresentation(dynamic presentation, FileStream stream)
    {
        stream.Position = 0;
        stream.SetLength(0);
        presentation.Save(stream);
    }

    private static BitmapSource? ConvertToBitmapSource(System.Drawing.Image? image, int decodeWidth)
    {
        if (image == null) return null;
        try
        {
            using var ms = new MemoryStream();
            image.Save(ms, ImageFormat.Png);
            ms.Position = 0;

            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.StreamSource = ms;
            if (decodeWidth > 0) bitmap.DecodePixelWidth = decodeWidth;
            bitmap.EndInit();
            bitmap.Freeze();
            return bitmap;
        }
        catch { return null; }
    }

    private static BitmapSource CreatePlaceholder(int width, int height, int slideNumber, string? previewText)
    {
        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen())
        {
            // 背景
            dc.DrawRectangle(
                new SolidColorBrush(Color.FromRgb(245, 245, 245)), null,
                new Rect(0, 0, width, height));

            // 枠線
            dc.DrawRectangle(null,
                new Pen(new SolidColorBrush(Color.FromRgb(220, 220, 220)), 1),
                new Rect(0, 0, width, height));

            // スライド番号
            var titleSize = width > 400 ? 24.0 : 14.0;
            dc.DrawText(
                new FormattedText($"Slide {slideNumber}",
                    System.Globalization.CultureInfo.InvariantCulture,
                    FlowDirection.LeftToRight,
                    new Typeface("Segoe UI"), titleSize,
                    new SolidColorBrush(Color.FromRgb(100, 100, 100)), 96),
                new Point(width > 400 ? 24 : 8, width > 400 ? 16 : 6));

            // テキストプレビュー
            if (!string.IsNullOrEmpty(previewText))
            {
                var textSize = width > 400 ? 14.0 : 9.0;
                var maxWidth = width - (width > 400 ? 48 : 16);
                var ft = new FormattedText(previewText,
                    System.Globalization.CultureInfo.InvariantCulture,
                    FlowDirection.LeftToRight,
                    new Typeface("Segoe UI"), textSize,
                    new SolidColorBrush(Color.FromRgb(80, 80, 80)), 96)
                {
                    MaxTextWidth = maxWidth,
                    MaxTextHeight = height - (width > 400 ? 60 : 24),
                    Trimming = TextTrimming.CharacterEllipsis
                };
                dc.DrawText(ft, new Point(width > 400 ? 24 : 8, width > 400 ? 52 : 22));
            }
        }
        var rt = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        rt.Render(visual);
        rt.Freeze();
        return rt;
    }

    // ── Syncfusion リフレクション解決 ──

    private static bool EnsureSyncfusionResolved()
    {
        if (_resolved) return _presentationType != null && _openMethod != null;
        _resolved = true;

        try
        {
            // アセンブリがまだロードされていない場合、明示的にロードを試みる
            TryLoadAssembly("Syncfusion.Presentation.Base");
            TryLoadAssembly("Syncfusion.Presentation.WPF");

            // Presentation 型を探す
            _presentationType = FindType("Syncfusion.Presentation.Presentation", "Syncfusion.Presentation.Base");
            _presentationType ??= FindType("Syncfusion.Presentation.Presentation", "Syncfusion.Presentation.WPF");
            _presentationType ??= FindType("Syncfusion.Presentation.Presentation");
            if (_presentationType == null)
            {
                System.Diagnostics.Debug.WriteLine("[PptxService] Presentation type not found. Loaded assemblies:");
                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    if (asm.GetName().Name?.Contains("Syncfusion") == true)
                        System.Diagnostics.Debug.WriteLine($"  - {asm.GetName().Name}");
                }
                return false;
            }

            _openMethod = _presentationType.GetMethod("Open", [typeof(Stream)]);
            if (_openMethod == null)
            {
                System.Diagnostics.Debug.WriteLine("[PptxService] Open(Stream) method not found on Presentation type");
                return false;
            }

            // ImageType.Bitmap（レンダリング用、なくても開くことは可能）
            var imageType = FindType("Syncfusion.Drawing.ImageType", "Syncfusion.Presentation.Base");
            imageType ??= FindType("Syncfusion.Drawing.ImageType", "Syncfusion.Presentation.WPF");
            imageType ??= FindType("Syncfusion.Drawing.ImageType");
            imageType ??= FindType("Syncfusion.Presentation.ImageType");
            if (imageType != null && imageType.IsEnum)
            {
                _imageTypeBitmap = Enum.Parse(imageType, "Bitmap");
                _canRender = true;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[PptxService] ImageType enum not found — rendering disabled, using placeholders");
                _canRender = false;
            }

            // SlideLayoutType
            _slideLayoutType = FindType("Syncfusion.Presentation.SlideLayoutType", "Syncfusion.Presentation.Base");
            _slideLayoutType ??= FindType("Syncfusion.Presentation.SlideLayoutType", "Syncfusion.Presentation.WPF");
            _slideLayoutType ??= FindType("Syncfusion.Presentation.SlideLayoutType");

            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PptxService] Resolve error: {ex.Message}");
            return false;
        }
    }

    private static void TryLoadAssembly(string assemblyName)
    {
        try
        {
            // 既にロード済みか確認
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (asm.GetName().Name == assemblyName) return;
            }
            Assembly.Load(assemblyName);
        }
        catch
        {
            // ロード失敗は無視（FindType のフォールバックに任せる）
        }
    }

    private static Type? FindType(string typeName, string? preferredAssembly = null)
    {
        if (preferredAssembly != null)
        {
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (asm.GetName().Name == preferredAssembly)
                {
                    var t = asm.GetType(typeName);
                    if (t != null) return t;
                }
            }
        }

        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
        {
            var t = asm.GetType(typeName);
            if (t != null) return t;
        }

        return null;
    }
}
