using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;

// Псевдонимы пространств имён OpenXml: P — для PresentationML, D — для DrawingML.
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace DocxToPdfConverter;

// Конвертер презентаций PPTX в PDF.
// В отличие от DOCX, в pptx каждая фигура имеет абсолютные координаты (в EMU)
// на полотне слайда. Один слайд — одна страница PDF в альбомной ориентации.
//
// Что поддерживается в v1:
//   - Текстовые блоки (с шрифтами, размером, жирным/курсивом, цветом, выравниванием)
//   - Растровые картинки (PNG, JPEG, BMP)
//   - Простые фигуры с заливкой (прямоугольник, эллипс)
//   - Фон слайда (если задан явной заливкой)
//   - Группы фигур (рекурсивно)
//
// Что НЕ поддерживается (просто пропускается):
//   - Графики, диаграммы (charts)
//   - SmartArt
//   - Таблицы (GraphicFrame с таблицей внутри)
//   - Векторные форматы картинок (WMF/EMF)
//   - Сложные эффекты, тени, градиенты
//   - Анимации, переходы (для PDF неактуально)
public class PptxConverter
{
    private const double EmuPerPoint = 12700.0;
    private const string DefaultFontName = "Calibri";
    private const double DefaultFontSize = 18;

    private static bool _fontResolverRegistered;

    private PdfDocument _pdf = null!;
    private PresentationPart _presPart = null!;
    private double _slideWidthPt;
    private double _slideHeightPt;
    private readonly List<IDisposable> _imageResources = new();

    public void Convert(string pptxPath, string pdfPath, IProgress<int>? progress = null)
    {
        if (!_fontResolverRegistered)
        {
            GlobalFontSettings.FontResolver = new WindowsFontResolver();
            _fontResolverRegistered = true;
        }

        _pdf = new PdfDocument();

        try
        {
            using var pres = PresentationDocument.Open(pptxPath, false);
            _presPart = pres.PresentationPart
                ?? throw new InvalidDataException("В презентации отсутствует основная часть.");

            var presentation = _presPart.Presentation
                ?? throw new InvalidDataException("Корневой элемент презентации отсутствует.");

            // Размер слайда: <p:sldSz cx="..." cy="..."/> в EMU.
            // Дефолт — 4:3 формат (10x7.5 дюйма).
            var slideSize = presentation.SlideSize;
            long cx = slideSize?.Cx?.Value ?? 9144000;
            long cy = slideSize?.Cy?.Value ?? 6858000;
            _slideWidthPt = cx / EmuPerPoint;
            _slideHeightPt = cy / EmuPerPoint;

            // Слайды в порядке, заданном в <p:sldIdLst>.
            var slideIds = presentation.SlideIdList?.Elements<P.SlideId>().ToList() ?? new();
            int total = slideIds.Count;
            int done = 0;

            foreach (var slideId in slideIds)
            {
                var relId = slideId.RelationshipId?.Value;
                if (relId == null) continue;
                if (_presPart.GetPartById(relId) is not SlidePart slidePart) continue;

                RenderSlide(slidePart);

                done++;
                progress?.Report(total == 0 ? 95 : done * 95 / total);
            }

            _pdf.Save(pdfPath);
            _pdf.Dispose();
            progress?.Report(100);
        }
        finally
        {
            foreach (var d in _imageResources)
            {
                try { d.Dispose(); } catch { }
            }
            _imageResources.Clear();
        }
    }

    // ---------- Рендер слайда ----------

    private void RenderSlide(SlidePart slidePart)
    {
        var page = _pdf.AddPage();
        page.Width = XUnit.FromPoint(_slideWidthPt);
        page.Height = XUnit.FromPoint(_slideHeightPt);

        using var gfx = XGraphics.FromPdfPage(page);

        // Фон слайда. Если явной заливки нет — рисуем белый.
        var bgFill = GetSlideBackgroundColor(slidePart);
        var bgBrush = new XSolidBrush(bgFill ?? XColors.White);
        gfx.DrawRectangle(bgBrush, 0, 0, _slideWidthPt, _slideHeightPt);

        // Дерево фигур (<p:spTree>) — все объекты на слайде.
        var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        if (shapeTree == null) return;

        foreach (var element in shapeTree.ChildElements)
        {
            RenderElement(element, gfx, slidePart);
        }
    }

    private void RenderElement(OpenXmlElement element, XGraphics gfx, SlidePart slidePart)
    {
        switch (element)
        {
            case P.Shape shape:
                RenderShape(shape, gfx);
                break;
            case P.Picture pic:
                RenderPicture(pic, gfx, slidePart);
                break;
            case P.GroupShape group:
                // Группа — обходим вложенные элементы рекурсивно.
                foreach (var child in group.ChildElements)
                    RenderElement(child, gfx, slidePart);
                break;
            // ConnectionShape, GraphicFrame (charts, tables) — пропускаем для v1.
        }
    }

    // ---------- Рендер фигур ----------

    private void RenderShape(P.Shape shape, XGraphics gfx)
    {
        var bounds = GetShapeBounds(shape.ShapeProperties);
        if (bounds == null) return;
        var rect = bounds.Value;

        // Заливка фигуры (если задана явным RGB).
        var fillColor = GetSolidFillColor(shape.ShapeProperties?.GetFirstChild<D.SolidFill>());
        if (fillColor.HasValue)
        {
            var brush = new XSolidBrush(fillColor.Value);
            var preset = shape.ShapeProperties?.GetFirstChild<D.PresetGeometry>()?.Preset?.Value;
            if (preset == D.ShapeTypeValues.Ellipse)
                gfx.DrawEllipse(brush, rect.X, rect.Y, rect.Width, rect.Height);
            else
                gfx.DrawRectangle(brush, rect.X, rect.Y, rect.Width, rect.Height);
        }

        // Текст внутри фигуры.
        if (shape.TextBody != null)
            RenderTextBody(shape.TextBody, rect, gfx);
    }

    private void RenderPicture(P.Picture pic, XGraphics gfx, SlidePart slidePart)
    {
        var bounds = GetShapeBounds(pic.ShapeProperties);
        if (bounds == null) return;

        var blip = pic.Descendants<D.Blip>().FirstOrDefault();
        var embedId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId)) return;

        if (slidePart.GetPartById(embedId) is not ImagePart imagePart) return;

        try
        {
            var ms = new MemoryStream();
            using (var src = imagePart.GetStream())
                src.CopyTo(ms);
            ms.Position = 0;

            XImage ximg;
            try { ximg = XImage.FromStream(ms); }
            catch { ms.Dispose(); return; } // неподдерживаемый формат — пропускаем

            _imageResources.Add(ximg);
            _imageResources.Add(ms);

            var rect = bounds.Value;
            gfx.DrawImage(ximg, rect.X, rect.Y, rect.Width, rect.Height);
        }
        catch { }
    }

    // Вытаскивает координаты и размеры фигуры из <a:xfrm>.
    private static (double X, double Y, double Width, double Height)? GetShapeBounds(P.ShapeProperties? props)
    {
        var xfrm = props?.Transform2D;
        if (xfrm?.Offset == null || xfrm.Extents == null) return null;

        long x = xfrm.Offset.X?.Value ?? 0;
        long y = xfrm.Offset.Y?.Value ?? 0;
        long cx = xfrm.Extents.Cx?.Value ?? 0;
        long cy = xfrm.Extents.Cy?.Value ?? 0;
        if (cx <= 0 || cy <= 0) return null;

        return (x / EmuPerPoint, y / EmuPerPoint, cx / EmuPerPoint, cy / EmuPerPoint);
    }

    // ---------- Цвета ----------

    private static XColor? GetSlideBackgroundColor(SlidePart slidePart)
    {
        var bg = slidePart.Slide?.CommonSlideData?.Background;
        if (bg == null) return null;
        var solidFill = bg.Descendants<D.SolidFill>().FirstOrDefault();
        return GetSolidFillColor(solidFill);
    }

    private static XColor? GetSolidFillColor(D.SolidFill? fill)
    {
        if (fill == null) return null;
        var rgb = fill.GetFirstChild<D.RgbColorModelHex>();
        if (rgb?.Val?.Value is not string hex || hex.Length != 6) return null;
        try
        {
            byte r = System.Convert.ToByte(hex.Substring(0, 2), 16);
            byte g = System.Convert.ToByte(hex.Substring(2, 2), 16);
            byte b = System.Convert.ToByte(hex.Substring(4, 2), 16);
            return XColor.FromArgb(r, g, b);
        }
        catch { return null; }
    }

    // ---------- Рендер текста ----------

    private void RenderTextBody(OpenXmlElement body, (double X, double Y, double Width, double Height) rect, XGraphics gfx)
    {
        // Простейшая логика: параграфы один под другим, начиная с верха фигуры.
        double y = rect.Y + 4;

        foreach (var paragraph in body.Elements<D.Paragraph>())
        {
            y = RenderParagraph(paragraph, rect, y, gfx);
            if (y > rect.Y + rect.Height) break; // не вылезаем за границы фигуры
        }
    }

    private double RenderParagraph(D.Paragraph paragraph, (double X, double Y, double Width, double Height) rect, double y, XGraphics gfx)
    {
        var alignment = ParseAlignment(paragraph.ParagraphProperties?.Alignment?.Value);

        var tokens = new List<TextToken>();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case D.Run run:
                    CollectRunTokens(run, tokens, gfx);
                    break;
                case D.Break:
                    tokens.Add(new TextToken("\n", null!, null!, 0));
                    break;
            }
        }

        // Пустой параграф — пропуск строки высотой со стандартный шрифт.
        if (tokens.Count == 0)
            return y + DefaultFontSize * 1.2;

        return LayoutLines(tokens, rect, y, alignment, gfx);
    }

    private void CollectRunTokens(D.Run run, List<TextToken> tokens, XGraphics gfx)
    {
        var props = run.RunProperties;
        var text = string.Join("", run.Elements<D.Text>().Select(t => t.Text));
        if (string.IsNullOrEmpty(text)) return;

        // Размер шрифта в pptx хранится в сотых долях пункта (sz="1800" = 18pt).
        double size = (props?.FontSize?.Value ?? 1800) / 100.0;
        bool bold = props?.Bold?.Value ?? false;
        bool italic = props?.Italic?.Value ?? false;
        bool underline = props?.Underline != null && props.Underline.Value != D.TextUnderlineValues.None;

        var fontName = props?.GetFirstChild<D.LatinFont>()?.Typeface?.Value ?? DefaultFontName;
        var color = GetSolidFillColor(props?.GetFirstChild<D.SolidFill>()) ?? XColors.Black;

        var style = XFontStyleEx.Regular;
        if (bold) style |= XFontStyleEx.Bold;
        if (italic) style |= XFontStyleEx.Italic;
        if (underline) style |= XFontStyleEx.Underline;

        XFont font;
        try { font = new XFont(fontName, size, style); }
        catch { font = new XFont(DefaultFontName, size, style); }

        var brush = new XSolidBrush(color);

        foreach (var word in SplitWords(text))
        {
            double width = word == "\n" ? 0 : gfx.MeasureString(word, font).Width;
            tokens.Add(new TextToken(word, font, brush, width));
        }
    }

    private static IEnumerable<string> SplitWords(string text)
    {
        if (string.IsNullOrEmpty(text)) yield break;

        int start = 0;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == ' ' || c == '\n')
            {
                if (i > start) yield return text.Substring(start, i - start);
                yield return c.ToString();
                start = i + 1;
            }
        }
        if (start < text.Length) yield return text.Substring(start);
    }

    private static double LayoutLines(List<TextToken> tokens, (double X, double Y, double Width, double Height) rect, double startY, XStringAlignment alignment, XGraphics gfx)
    {
        double maxWidth = rect.Width - 8; // небольшой внутренний отступ
        double y = startY;

        var line = new List<TextToken>();
        double lineWidth = 0;

        void Flush()
        {
            if (line.Count > 0)
            {
                while (line.Count > 0 && line[^1].Text == " ")
                {
                    lineWidth -= line[^1].Width;
                    line.RemoveAt(line.Count - 1);
                }
                if (line.Count > 0)
                {
                    DrawTextLine(line, lineWidth, rect.X + 4, maxWidth, y, alignment, gfx);
                    double lineHeight = line.Where(t => t.Font != null).Select(t => t.Font.GetHeight()).DefaultIfEmpty(DefaultFontSize).Max() * 1.2;
                    y += lineHeight;
                }
                else
                {
                    y += DefaultFontSize * 1.2; // пустая строка после очистки пробелов
                }
            }
            line.Clear();
            lineWidth = 0;
        }

        foreach (var tok in tokens)
        {
            if (tok.Text == "\n")
            {
                Flush();
                continue;
            }
            if (tok.Text == " " && line.Count == 0) continue;

            if (lineWidth + tok.Width > maxWidth && line.Count > 0)
            {
                Flush();
                if (tok.Text == " ") continue;
            }
            line.Add(tok);
            lineWidth += tok.Width;
        }
        Flush();

        return y;
    }

    private static void DrawTextLine(List<TextToken> line, double lineWidth, double x0, double maxWidth, double y, XStringAlignment alignment, XGraphics gfx)
    {
        double x = alignment switch
        {
            XStringAlignment.Center => x0 + (maxWidth - lineWidth) / 2,
            XStringAlignment.Far => x0 + (maxWidth - lineWidth),
            _ => x0
        };

        foreach (var tok in line)
        {
            gfx.DrawString(tok.Text, tok.Font, tok.Brush, x, y, XStringFormats.TopLeft);
            x += tok.Width;
        }
    }

    private static XStringAlignment ParseAlignment(D.TextAlignmentTypeValues? align)
    {
        if (align == null) return XStringAlignment.Near;
        if (align == D.TextAlignmentTypeValues.Center) return XStringAlignment.Center;
        if (align == D.TextAlignmentTypeValues.Right) return XStringAlignment.Far;
        return XStringAlignment.Near;
    }

    private readonly record struct TextToken(string Text, XFont Font, XBrush Brush, double Width);
}
