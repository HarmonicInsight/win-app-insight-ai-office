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
/// Syncfusion をリフレクション経由で利用（Base/Portable 型競合を回避）。
/// OpenXML はテキスト抽出とフォールバックに使用。
/// </summary>
public static class PptxService
{
    private static Type? _presentationType;
    private static MethodInfo? _openMethod;
    private static MethodInfo? _convertToImageMethod;
    private static object? _imageTypeBitmap;
    private static object? _slideLayoutBlank;
    private static bool _resolved;
    private static bool _canRender;

    // ── レンダリング ──

    public static List<(BitmapSource Full, BitmapSource Thumbnail)> RenderAllSlides(
        string pptxPath, int thumbnailWidth = 280)
    {
        var results = new List<(BitmapSource, BitmapSource)>();
        if (!File.Exists(pptxPath)) return results;

        if (EnsureResolved() && _canRender)
        {
            try
            {
                return RenderAllSlidesSyncfusion(pptxPath, thumbnailWidth);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PptxService] Syncfusion render failed: {ex.Message}");
            }
        }

        return RenderAllSlidesOpenXml(pptxPath, thumbnailWidth);
    }

    private static List<(BitmapSource Full, BitmapSource Thumbnail)> RenderAllSlidesSyncfusion(
        string pptxPath, int thumbnailWidth)
    {
        var results = new List<(BitmapSource, BitmapSource)>();

        using var stream = new FileStream(pptxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var presentation = _openMethod!.Invoke(null, [stream])!;

        try
        {
            // Get Slides collection via reflection (avoids dynamic dispatch issues)
            var slidesObj = presentation.GetType().GetProperty("Slides")?.GetValue(presentation);
            if (slidesObj == null) return results;

            var slidesEnumerable = (System.Collections.IEnumerable)slidesObj;
            int index = 0;

            foreach (var slide in slidesEnumerable)
            {
                index++;
                try
                {
                    // Use cached ConvertToImage MethodInfo — avoids dynamic dispatch
                    // which fails with extension methods and type-ambiguous enum parameters
                    var convertMethod = _convertToImageMethod;
                    if (convertMethod == null)
                    {
                        // Find ConvertToImage on the concrete slide type
                        var slideType = slide.GetType();
                        convertMethod = slideType.GetMethod("ConvertToImage");
                        if (convertMethod == null)
                        {
                            // Try interface types
                            foreach (var iface in slideType.GetInterfaces())
                            {
                                convertMethod = iface.GetMethod("ConvertToImage");
                                if (convertMethod != null) break;
                            }
                        }
                        _convertToImageMethod = convertMethod;
                    }

                    if (convertMethod != null)
                    {
                        var image = (System.Drawing.Image)convertMethod.Invoke(slide, [_imageTypeBitmap])!;
                        using (image)
                        {
                            var full = ConvertToBitmapSource(image, 0) ?? CreatePlaceholder(960, 540, index, null);
                            var thumb = ConvertToBitmapSource(image, thumbnailWidth) ?? CreatePlaceholder(280, 158, index, null);
                            results.Add((full, thumb));
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[PptxService] ConvertToImage method not found on {slide.GetType().FullName}");
                        results.Add((CreatePlaceholder(960, 540, index, null), CreatePlaceholder(280, 158, index, null)));
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[PptxService] Slide {index} render error: {ex.Message}");
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
        if (!EnsureResolved())
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
            presentation.Slides.Add(_slideLayoutBlank);

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
        if (!EnsureResolved()) return;

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
            dc.DrawRectangle(
                new SolidColorBrush(Color.FromRgb(245, 245, 245)), null,
                new Rect(0, 0, width, height));
            dc.DrawRectangle(null,
                new Pen(new SolidColorBrush(Color.FromRgb(220, 220, 220)), 1),
                new Rect(0, 0, width, height));

            var titleSize = width > 400 ? 24.0 : 14.0;
            dc.DrawText(
                new FormattedText($"Slide {slideNumber}",
                    System.Globalization.CultureInfo.InvariantCulture,
                    FlowDirection.LeftToRight,
                    new Typeface("Segoe UI"), titleSize,
                    new SolidColorBrush(Color.FromRgb(100, 100, 100)), 96),
                new Point(width > 400 ? 24 : 8, width > 400 ? 16 : 6));

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

    private static bool EnsureResolved()
    {
        if (_resolved) return _presentationType != null && _openMethod != null;
        _resolved = true;

        try
        {
            TryLoadAssembly("Syncfusion.Presentation.Base");

            _presentationType = FindType("Syncfusion.Presentation.Presentation", "Syncfusion.Presentation.Base");
            _presentationType ??= FindType("Syncfusion.Presentation.Presentation");
            if (_presentationType == null) return false;

            _openMethod = _presentationType.GetMethod("Open", [typeof(Stream)]);
            if (_openMethod == null) return false;

            // ImageType.Bitmap — try multiple possible namespaces
            var imageType = FindType("Syncfusion.Drawing.ImageType", "Syncfusion.Presentation.Base");
            imageType ??= FindType("Syncfusion.Drawing.ImageType");
            imageType ??= FindType("Syncfusion.Presentation.ImageType");
            if (imageType != null && imageType.IsEnum)
            {
                _imageTypeBitmap = Enum.Parse(imageType, "Bitmap");

                // Pre-resolve ConvertToImage from ISlide interface
                var slideType = FindType("Syncfusion.Presentation.ISlide", "Syncfusion.Presentation.Base");
                slideType ??= FindType("Syncfusion.Presentation.ISlide");
                if (slideType != null)
                {
                    _convertToImageMethod = slideType.GetMethod("ConvertToImage", [imageType]);
                    _convertToImageMethod ??= slideType.GetMethod("ConvertToImage");
                }

                _canRender = true;
            }

            // SlideLayoutType.Blank
            var slideLayoutType = FindType("Syncfusion.Presentation.SlideLayoutType", "Syncfusion.Presentation.Base");
            slideLayoutType ??= FindType("Syncfusion.Presentation.SlideLayoutType");
            if (slideLayoutType != null && slideLayoutType.IsEnum)
                _slideLayoutBlank = Enum.Parse(slideLayoutType, "Blank");

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void TryLoadAssembly(string assemblyName)
    {
        try
        {
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                if (asm.GetName().Name == assemblyName) return;
            Assembly.Load(assemblyName);
        }
        catch { }
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
