namespace DocxToPdfConverter;

static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        // Если переданы аргументы командной строки — запускаемся в CLI-режиме без GUI.
        // Использование: DocxToPdfConverter.exe <input.docx> <output.pdf>
        if (args.Length >= 2)
        {
            try
            {
                var converter = new Converter();
                converter.Convert(args[0], args[1]);
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
