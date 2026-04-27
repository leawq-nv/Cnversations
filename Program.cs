namespace DocxToPdfConverter;

static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        // CLI-режим. Использование:
        //   DocxToPdfConverter.exe <input> <output.pdf>
        // Конвертер выбирается по расширению входного файла (.docx или .pptx).
        if (args.Length >= 2)
        {
            try
            {
                var ext = Path.GetExtension(args[0]).ToLowerInvariant();
                switch (ext)
                {
                    case ".docx":
                        new Converter().Convert(args[0], args[1]);
                        break;
                    case ".pptx":
                        new PptxConverter().Convert(args[0], args[1]);
                        break;
                    default:
                        Console.Error.WriteLine($"Неподдерживаемое расширение: {ext}. Поддерживаются .docx и .pptx.");
                        return 1;
                }
                Console.WriteLine($"OK: {args[1]}");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Ошибка: {ex.Message}");
                return 1;
            }
        }

        ApplicationConfiguration.Initialize();
        Application.Run(new Form1());
        return 0;
    }
}
