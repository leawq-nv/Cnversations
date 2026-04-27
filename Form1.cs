namespace DocxToPdfConverter;

public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
    }

    // Открывает диалог выбора .docx файла.
    private void OnInputBrowseClick(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Выберите DOCX файл",
            Filter = "Word документы (*.docx)|*.docx|Все файлы (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _inputPath.Text = dialog.FileName;

            // Если выходной путь ещё не задан — подставляем тот же файл с расширением .pdf.
            if (string.IsNullOrWhiteSpace(_outputPath.Text))
            {
                _outputPath.Text = Path.ChangeExtension(dialog.FileName, ".pdf");
            }
        }
    }

    // Открывает диалог выбора места сохранения .pdf файла.
    private void OnOutputBrowseClick(object? sender, EventArgs e)
    {
        using var dialog = new SaveFileDialog
        {
            Title = "Сохранить PDF как",
            Filter = "PDF документы (*.pdf)|*.pdf|Все файлы (*.*)|*.*",
            DefaultExt = "pdf",
            AddExtension = true,
            FileName = string.IsNullOrWhiteSpace(_outputPath.Text)
                ? "output.pdf"
                : Path.GetFileName(_outputPath.Text)
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _outputPath.Text = dialog.FileName;
        }
    }

    // Запускает конвертацию (асинхронно, чтобы не подвисало окно).
    private async void OnConvertClick(object? sender, EventArgs e)
    {
        var input = _inputPath.Text.Trim();
        var output = _outputPath.Text.Trim();

        // Простая валидация ввода.
        if (string.IsNullOrEmpty(input))
        {
            ShowWarning("Не выбран входной DOCX файл.");
            return;
        }
        if (!File.Exists(input))
        {
            ShowWarning("Входной файл не найден.");
            return;
        }
        if (string.IsNullOrEmpty(output))
        {
            ShowWarning("Не указан путь для выходного PDF файла.");
            return;
        }

        SetBusy(true, "Идёт конвертация...");

        // Объект Progress<int> ловит отчёты прогресса из фонового потока
        // и сам перебрасывает их в UI-поток для безопасного обновления контролов.
        var progress = new Progress<int>(percent =>
        {
            _progressBar.Value = Math.Clamp(percent, 0, 100);
            _statusLabel.Text = $"Идёт конвертация... {percent}%";
        });

        try
        {
            // Конвертация выполняется на фоновом потоке, чтобы не замораживать GUI.
            await Task.Run(() =>
            {
                var converter = new Converter();
                converter.Convert(input, output, progress);
            });

            SetBusy(false, $"Готово. Файл сохранён: {output}");
            MessageBox.Show(this,
                $"Конвертация завершена.\n\n{output}",
                "Успех",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            SetBusy(false, "Ошибка при конвертации.");
            MessageBox.Show(this,
                $"Не удалось сконвертировать файл.\n\n{ex.Message}",
                "Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    // Переключает форму между состояниями "занят" и "готов".
    private void SetBusy(bool busy, string status)
    {
        _convertButton.Enabled = !busy;
        _inputBrowse.Enabled = !busy;
        _outputBrowse.Enabled = !busy;

        // Сбрасываем прогресс-бар в начало при старте новой задачи или возвращении к простою.
        if (busy)
            _progressBar.Value = 0;

        _statusLabel.Text = status;
    }

    private void ShowWarning(string text)
    {
        MessageBox.Show(this, text, "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    // Срабатывает, когда пользователь тащит файл в окно.
    // Здесь мы решаем — принимать перетаскивание или нет.
    private void OnFormDragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data != null && TryGetDocxPath(e.Data, out _))
        {
            e.Effect = DragDropEffects.Copy; // курсор покажет «плюсик»
        }
        else
        {
            e.Effect = DragDropEffects.None;
        }
    }

    // Срабатывает, когда пользователь отпустил файл над окном.
    private void OnFormDragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data == null || !TryGetDocxPath(e.Data, out var path))
            return;

        _inputPath.Text = path;

        // Если выходной путь ещё не задан — подставим .pdf рядом с docx.
        if (string.IsNullOrWhiteSpace(_outputPath.Text))
        {
            _outputPath.Text = Path.ChangeExtension(path, ".pdf");
        }

        _statusLabel.Text = $"Файл загружен: {Path.GetFileName(path)}";
    }

    // Извлекает путь к .docx файлу из объекта перетаскивания.
    // Если перетаскивают несколько файлов — берём первый docx.
    private static bool TryGetDocxPath(IDataObject data, out string path)
    {
        path = string.Empty;
        if (!data.GetDataPresent(DataFormats.FileDrop)) return false;

        if (data.GetData(DataFormats.FileDrop) is not string[] files) return false;

        var docx = files.FirstOrDefault(f =>
            !string.IsNullOrEmpty(f) &&
            File.Exists(f) &&
            string.Equals(Path.GetExtension(f), ".docx", StringComparison.OrdinalIgnoreCase));

        if (docx == null) return false;

        path = docx;
        return true;
    }
}
