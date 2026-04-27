using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;

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

    private static bool _fontResolverRegistered;

    private PdfDocument _pdf = null!;
    private PdfPage _page = null!;
    private XGraphics _gfx = null!;
    private double _y; // текущая Y-координата (растёт сверху вниз)

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

        using (var docx = WordprocessingDocument.Open(docxPath, false))
        {
            var body = docx.MainDocumentPart?.Document?.Body
                ?? throw new InvalidDataException("Документ не содержит тела (возможно, файл повреждён).");

            // Считаем параграфы, чтобы знать общую долю работы для прогресс-бара.
            var paragraphs = body.Elements<Paragraph>().ToList();
            int total = paragraphs.Count;
            int done = 0;

            foreach (var paragraph in paragraphs)
            {
                RenderParagraph(paragraph);
                done++;

                // Сообщаем процент выполнения. Оставляем последние 5% на сохранение файла.
                progress?.Report(total == 0 ? 95 : done * 95 / total);
            }
        }

        _gfx.Dispose();
        _pdf.Save(pdfPath);
        _pdf.Dispose();

        progress?.Report(100);
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
        // Стиль (например, Heading1) — для подсказок размера.
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var headingLevel = ParseHeadingLevel(styleId);

        // Выравнивание абзаца.
        var alignment = ParseAlignment(paragraph.ParagraphProperties?.Justification?.Val);

        // Собираем токены (слово + форматирование) из всех Run-ов.
        var tokens = new List<Token>();
        foreach (var run in paragraph.Elements<Run>())
        {
            CollectTokensFromRun(run, headingLevel, tokens);
        }

        // Пустой параграф — просто пропуск строки.
        if (tokens.Count == 0)
        {
            _y += DefaultFontSize * LineSpacing;
            EnsureSpace(0);
            return;
        }

        // Раскладываем токены по строкам с учётом ширины страницы.
        LayoutAndDraw(tokens, alignment);

        _y += ParagraphSpacing;
    }

    private void CollectTokensFromRun(Run run, int headingLevel, List<Token> tokens)
    {
        var props = run.RunProperties;

        bool bold = props?.Bold != null;
        bool italic = props?.Italic != null;
        bool underline = props?.Underline != null;

        // Заголовки: жирный + увеличенный шрифт по уровню.
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

        // Проходим по всем дочерним элементам в порядке появления:
        // Text — обычный текст, Break — разрыв строки или страницы, TabChar — табуляция.
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
            }
        }
    }

    // Разбивает текст на токены: каждое слово и каждый пробел — отдельный токен.
    // Это нужно, чтобы корректно переносить строки по пробелам.
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
                // Убираем висящие пробелы в конце строки (они не должны влиять на выравнивание).
                while (line.Count > 0 && line[^1].Text == " ")
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

            // Пробел в начале строки игнорируем.
            if (tok.Text == " " && line.Count == 0)
                continue;

            // Если слово не помещается — переносим на новую строку.
            if (lineWidth + tok.Width > maxWidth && line.Count > 0)
            {
                Flush();
                if (tok.Text == " ") continue; // пробел в начале новой строки тоже не нужен
            }

            line.Add(tok);
            lineWidth += tok.Width;
        }

        Flush();
    }

    private void DrawLine(List<Token> line, double lineWidth, double maxWidth, XStringAlignment alignment)
    {
        // Высота строки = максимальная высота шрифта в строке * межстрочный интервал.
        double lineHeight = line.Max(t => t.Font.GetHeight()) * LineSpacing;

        EnsureSpace(lineHeight);

        // Стартовый X в зависимости от выравнивания.
        double x = alignment switch
        {
            XStringAlignment.Center => Margin + (maxWidth - lineWidth) / 2,
            XStringAlignment.Far => Margin + (maxWidth - lineWidth),
            _ => Margin
        };

        // Базовая линия — Y + высота шрифта без межстрочного бонуса.
        // PdfSharp рисует текст по верхнему левому углу при XStringFormats.TopLeft.
        foreach (var tok in line)
        {
            _gfx.DrawString(tok.Text, tok.Font, tok.Brush, x, _y, XStringFormats.TopLeft);
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
        // В docx размер шрифта хранится в полупунктах: значение 24 = 12pt.
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
        // Justify (по ширине) пока не реализован — рисуем как обычное левое.
        return XStringAlignment.Near;
    }

    private static int ParseHeadingLevel(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return 0;
        // Стандартные ID: "Heading1", "Heading2", ... или "Заголовок1" в локализованных версиях.
        var s = styleId.ToLowerInvariant();
        for (int level = 1; level <= 6; level++)
        {
            if (s == $"heading{level}" || s == $"заголовок{level}")
                return level;
        }
        return 0;
    }

    // ---------- Внутренние типы ----------

    private enum TokenKind { Text, LineBreak, PageBreak }

    private readonly record struct Token(
        string Text,
        XFont Font,
        XBrush Brush,
        double Width,
        TokenKind Kind = TokenKind.Text);
}
