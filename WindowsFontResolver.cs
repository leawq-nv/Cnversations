using PdfSharp.Fonts;

namespace DocxToPdfConverter;

// Резолвер шрифтов: говорит PdfSharp, где брать .ttf файлы для нужного шрифта.
// PdfSharp 6+ требует, чтобы мы сами указали источник шрифтов.
// Берём шрифты прямо из системной папки C:\Windows\Fonts.
public class WindowsFontResolver : IFontResolver
{
    public byte[]? GetFont(string faceName)
    {
        var fontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        var path = Path.Combine(fontsDir, faceName);
        return File.Exists(path) ? File.ReadAllBytes(path) : null;
    }

    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var key = familyName.Trim().ToLowerInvariant();

        string file = key switch
        {
            "arial" => PickArial(isBold, isItalic),
            "calibri" => PickCalibri(isBold, isItalic),
            "courier" or "courier new" => PickCourier(isBold, isItalic),
            "verdana" => PickVerdana(isBold, isItalic),
            "tahoma" => PickTahoma(isBold, isItalic),
            _ => PickTimes(isBold, isItalic) // по умолчанию — Times New Roman
        };

        return new FontResolverInfo(file);
    }

    private static string PickTimes(bool b, bool i) => (b, i) switch
    {
        (true, true) => "timesbi.ttf",
        (true, false) => "timesbd.ttf",
        (false, true) => "timesi.ttf",
        _ => "times.ttf"
    };

    private static string PickArial(bool b, bool i) => (b, i) switch
    {
        (true, true) => "arialbi.ttf",
        (true, false) => "arialbd.ttf",
        (false, true) => "ariali.ttf",
        _ => "arial.ttf"
    };

    private static string PickCalibri(bool b, bool i) => (b, i) switch
    {
        (true, true) => "calibriz.ttf",
        (true, false) => "calibrib.ttf",
        (false, true) => "calibrii.ttf",
        _ => "calibri.ttf"
    };

    private static string PickCourier(bool b, bool i) => (b, i) switch
    {
        (true, true) => "courbi.ttf",
        (true, false) => "courbd.ttf",
        (false, true) => "couri.ttf",
        _ => "cour.ttf"
    };

    private static string PickVerdana(bool b, bool i) => (b, i) switch
    {
        (true, true) => "verdanaz.ttf",
        (true, false) => "verdanab.ttf",
        (false, true) => "verdanai.ttf",
        _ => "verdana.ttf"
    };

    private static string PickTahoma(bool b, bool i) => (b, i) switch
    {
        (true, _) => "tahomabd.ttf",
        _ => "tahoma.ttf"
    };
}
