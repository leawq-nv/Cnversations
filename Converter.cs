using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;

// Псевдонимы для пространств имён OpenXml, связанных с рисованием —
// чтобы не плодить длинные DocumentFormat.OpenXml.Drawing.Wordprocessing.* в коде.
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxToPdfConverter;

// Логика конвертации docx -> pdf.
// Читаем docx через OpenXml, рисуем pdf через PdfSharp.
public class Converter
{
    // Отступы от краёв страницы (в пунктах PDF; 1 пункт = 1/72 дюйма).
    private const double Margin = 50;

    // Шрифт и размер по умолчанию, если в документе не указаны явно.
    private const string DefaultFontName = "Times New Roman";
    private const double DefaultFontSize = 11;

    // Множитель межстрочного интервала.
    private const double LineSpacing = 1.2;

    // Отступ после параграфа.
    private const double ParagraphSpacing = 4;

    // 1 пункт PDF = 12700 EMU (English Metric Units — единица измерения в OOXML).
    private const double EmuPerPoint = 12700.0;

    private static bool _fontResolverRegistered;

    private PdfDocument _pdf = null!;
    private PdfPage _page = null!;
    private XGraphics _gfx = null!;
    private double _y; // текущая Y-координата (растёт сверху вниз)

    // Главная часть документа — нужна для разрешения ссылок на встроенные картинки.
    private MainDocumentPart _mainPart = null!;

    // Ресурсы картинок (XImage и потоки), которые нужно держать живыми до Save и потом утилизировать.
    private readonly List<IDisposable> _imageResources = new();

    public void Convert(string docxPath, string pdfPath, IProgress<int>? progress = null)
    {
        // Регистрируем резолвер шрифтов один раз за всё время жизни приложения.
        if (!_fontResolverRegistered)
        {
            GlobalFontSettings.FontResolver = new WindowsFontResolver();
            _fontResolverRegistered = true;
        }

        _pdf = new PdfDocument();
        AddPage();

        try
        {
            using (var docx = WordprocessingDocument.Open(docxPath, false))
            {
                _mainPart = docx.MainDocumentPart
                    ?? throw new InvalidDataException("Главная часть документа отсутствует.");

                var body = _mainPart.Document?.Body
                    ?? throw new InvalidDataException("Документ не содержит тела (возможно, файл повреждён).");

                var paragraphs = body.Elements<Paragraph>().ToList();
                int total = paragraphs.Count;
                int done = 0;

                foreach (var paragraph in paragraphs)
                {
                    RenderParagraph(paragraph);
                    done++;
                    progress?.Report(total == 0 ? 95 : done * 95 / total);
                }
            }

            _gfx.Dispose();
            _pdf.Save(pdfPath);
            _pdf.Dispose();

            progress?.Report(100);
        }
        finally
        {
            // Утилизируем ресурсы картинок только после Save — XImage держит ссылки на потоки,
            // PdfSharp дочитывает их при сохранении.
            foreach (var d in _imageResources)
            {
                try { d.Dispose(); } catch { }
            }
            _imageResources.Clear();
        }
    }

    private void AddPage()
    {
        _gfx?.Dispose();
        _page = _pdf.AddPage();
        _page.Size = PdfSharp.PageSize.A4;
        _gfx = XGraphics.FromPdfPage(_page);
        _y = Margin;
    }

    // ---------- Рендер параграфа ----------

    private void RenderParagraph(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var headingLevel = ParseHeadingLevel(styleId);

        var alignment = ParseAlignment(paragraph.ParagraphProperties?.Justification?.Val);

        var tokens = new List<Token>();
        foreach (var run in paragraph.Elements<Run>())
        {
            CollectTokensFromRun(run, headingLevel, tokens);
        }

        if (tokens.Count == 0)
        {
            _y += DefaultFontSize * LineSpacing;
            EnsureSpace(0);
            return;
        }

        LayoutAndDraw(tokens, alignment);

        _y += ParagraphSpacing;
    }

    private void CollectTokensFromRun(Run run, int headingLevel, List<Token> tokens)
    {
        var props = run.RunProperties;

        bool bold = props?.Bold != null;
        bool italic = props?.Italic != null;
        bool underline = props?.Underline != null;

        double size;
        if (headingLevel > 0)
        {
            bold = true;
            size = headingLevel switch
            {
                1 => 20,
                2 => 17,
                3 => 14,
                4 => 12.5,
                _ => 11.5
            };
        }
        else
        {
            size = ParseFontSize(props) ?? DefaultFontSize;
        }

        var fontName = props?.RunFonts?.Ascii?.Value ?? DefaultFontName;
        var color = ParseColor(props?.Color?.Val?.Value) ?? XColors.Black;

        var style = XFontStyleEx.Regular;
        if (bold) style |= XFontStyleEx.Bold;
        if (italic) style |= XFontStyleEx.Italic;
        if (underline) style |= XFontStyleEx.Underline;

        XFont font;
        try { font = new XFont(fontName, size, style); }
        catch { font = new XFont(DefaultFontName, size, style); }

        var brush = new XSolidBrush(color);

        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case Text t:
                    foreach (var word in SplitToWords(t.Text))
                        tokens.Add(new Token(word, font, brush, _gfx.MeasureString(word, font).Width));
                    break;
                case Break br:
                    bool isPageBreak = br.Type != null && br.Type.Value == BreakValues.Page;
                    tokens.Add(new Token("", font, brush, 0, isPageBreak ? TokenKind.PageBreak : TokenKind.LineBreak));
                    break;
                case TabChar:
                    var tabWidth = _gfx.MeasureString("    ", font).Width;
                    tokens.Add(new Token("    ", font, brush, tabWidth));
                    break;
                case Drawing drawing:
                    var img = TryLoadImage(drawing);
                    if (img != null)
                    {
                        tokens.Add(new Token(
                            "", font, brush,
                            Width: img.Value.WidthPt,
                            Kind: TokenKind.Image,
                            Image: img.Value.Image,
                            Height: img.Value.HeightPt));
                    }
                    break;
            }
        }
    }

    private static IEnumerable<string> SplitToWords(string text)
    {
        if (string.IsNullOrEmpty(text)) yield break;

        int start = 0;
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] == ' ')
            {
                if (i > start)
                    yield return text.Substring(start, i - start);
                yield return " ";
                start = i + 1;
            }
        }
        if (start < text.Length)
            yield return text.Substring(start);
    }

    // ---------- Загрузка картинки из docx ----------

    // Извлекает встроенную (inline) картинку из элемента <w:drawing>.
    // Возвращает XImage и размеры в пунктах PDF, либо null, если что-то пошло не так.
    private (XImage Image, double WidthPt, double HeightPt)? TryLoadImage(Drawing drawing)
    {
        try
        {
            // Размер картинки из <wp:extent cx="..." cy="..."/> в EMU.
            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            if (extent?.Cx?.Value is not long cx || extent.Cy?.Value is not long cy) return null;

            double widthPt = cx / EmuPerPoint;
            double heightPt = cy / EmuPerPoint;
            if (widthPt <= 0 || heightPt <= 0) return null;

            // Ссылка на ресурс картинки через relationship ID в <a:blip r:embed="..."/>.
            var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
            var embedId = blip?.Embed?.Value;
            if (string.IsNullOrEmpty(embedId)) return null;

            if (_mainPart.GetPartById(embedId) is not ImagePart imagePart) return null;

            // Копируем байты картинки в MemoryStream — он переживёт чтение docx
            // и будет жить, пока PdfSharp дочитывает картинку при Save.
            var ms = new MemoryStream();
            using (var src = imagePart.GetStream())
            {
                src.CopyTo(ms);
            }
            ms.Position = 0;

            XImage ximg;
            try
            {
                ximg = XImage.FromStream(ms);
            }
            catch
            {
                ms.Dispose();
                return null; // неподдерживаемый формат (например, WMF/EMF) — пропускаем
            }

            _imageResources.Add(ximg);
            _imageResources.Add(ms);

            // Если картинка шире доступной области — масштабируем по ширине.
            double maxWidth = _page.Width.Point - 2 * Margin;
            if (widthPt > maxWidth)
            {
                double scale = maxWidth / widthPt;
                widthPt = maxWidth;
                heightPt *= scale;
            }

            // Если выше доступной области — масштабируем по высоте (с сохранением пропорций).
            double maxHeight = _page.Height.Point - 2 * Margin;
            if (heightPt > maxHeight)
            {
                double scale = maxHeight / heightPt;
                heightPt = maxHeight;
                widthPt *= scale;
            }

            return (ximg, widthPt, heightPt);
        }
        catch
        {
            return null;
        }
    }

    // ---------- Раскладка строк ----------

    private void LayoutAndDraw(List<Token> tokens, XStringAlignment alignment)
    {
        double maxWidth = _page.Width.Point - 2 * Margin;

        var line = new List<Token>();
        double lineWidth = 0;

        void Flush()
        {
            if (line.Count > 0)
            {
                while (line.Count > 0 && line[^1].Text == " " && line[^1].Kind == TokenKind.Text)
                {
                    lineWidth -= line[^1].Width;
                    line.RemoveAt(line.Count - 1);
                }
                if (line.Count > 0)
                    DrawLine(line, lineWidth, maxWidth, alignment);
            }
            line = new List<Token>();
            lineWidth = 0;
        }

        foreach (var tok in tokens)
        {
            if (tok.Kind == TokenKind.PageBreak)
            {
                Flush();
                AddPage();
                continue;
            }
            if (tok.Kind == TokenKind.LineBreak)
            {
                Flush();
                continue;
            }

            if (tok.Text == " " && tok.Kind == TokenKind.Text && line.Count == 0)
                continue;

            if (lineWidth + tok.Width > maxWidth && line.Count > 0)
            {
                Flush();
                if (tok.Text == " " && tok.Kind == TokenKind.Text) continue;
            }

            line.Add(tok);
            lineWidth += tok.Width;
        }

        Flush();
    }

    private void DrawLine(List<Token> line, double lineWidth, double maxWidth, XStringAlignment alignment)
    {
        // Высота строки: для текста — высота шрифта × межстрочный интервал, для картинок — их собственная высота.
        double lineHeight = line.Max(t => t.Kind == TokenKind.Image
            ? t.Height
            : t.Font.GetHeight() * LineSpacing);

        EnsureSpace(lineHeight);

        double x = alignment switch
        {
            XStringAlignment.Center => Margin + (maxWidth - lineWidth) / 2,
            XStringAlignment.Far => Margin + (maxWidth - lineWidth),
            _ => Margin
        };

        foreach (var tok in line)
        {
            if (tok.Kind == TokenKind.Image && tok.Image != null)
            {
                _gfx.DrawImage(tok.Image, x, _y, tok.Width, tok.Height);
            }
            else
            {
                _gfx.DrawString(tok.Text, tok.Font, tok.Brush, x, _y, XStringFormats.TopLeft);
            }
            x += tok.Width;
        }

        _y += lineHeight;
    }

    private void EnsureSpace(double needed)
    {
        if (_y + needed > _page.Height.Point - Margin)
        {
            AddPage();
        }
    }

    // ---------- Парсинг свойств docx ----------

    private static double? ParseFontSize(RunProperties? props)
    {
        var v = props?.FontSize?.Val?.Value;
        if (v == null) return null;
        if (double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out var halfPt))
            return halfPt / 2.0;
        return null;
    }

    private static XColor? ParseColor(string? hex)
    {
        if (string.IsNullOrEmpty(hex) || hex == "auto") return null;
        if (hex.Length != 6) return null;
        try
        {
            byte r = System.Convert.ToByte(hex.Substring(0, 2), 16);
            byte g = System.Convert.ToByte(hex.Substring(2, 2), 16);
            byte b = System.Convert.ToByte(hex.Substring(4, 2), 16);
            return XColor.FromArgb(r, g, b);
        }
        catch { return null; }
    }

    private static XStringAlignment ParseAlignment(DocumentFormat.OpenXml.EnumValue<JustificationValues>? value)
    {
        if (value?.Value == null) return XStringAlignment.Near;
        var v = value.Value;
        if (v == JustificationValues.Center) return XStringAlignment.Center;
        if (v == JustificationValues.Right || v == JustificationValues.End) return XStringAlignment.Far;
        return XStringAlignment.Near;
    }

    private static int ParseHeadingLevel(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return 0;
        var s = styleId.ToLowerInvariant();
        for (int level = 1; level <= 6; level++)
        {
            if (s == $"heading{level}" || s == $"заголовок{level}")
                return level;
        }
        return 0;
    }

    // ---------- Внутренние типы ----------

    private enum TokenKind { Text, LineBreak, PageBreak, Image }

    private readonly record struct Token(
        string Text,
        XFont Font,
        XBrush Brush,
        double Width,
        TokenKind Kind = TokenKind.Text,
        XImage? Image = null,
        double Height = 0);
}
