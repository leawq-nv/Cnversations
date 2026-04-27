namespace DocxToPdfConverter;

// Универсальная панель одной вкладки-конвертера.
// Параметризуется расширением входного файла, описанием типа и функцией конвертации.
// Один и тот же класс переиспользуется для DOCX и PPTX вкладок.
public class ConverterPanel : Panel
{
    private readonly string _inputExt;       // например, ".docx" или ".pptx"
    private readonly string _inputTypeName;  // для диалога: "Word документы" / "PowerPoint презентации"
    private readonly string _dropZoneText;   // текст внутри зоны перетаскивания
    private readonly Action<string, string, IProgress<int>?> _converter; // сама конвертация

    private DropZone _dropZone = null!;
    private Label _dropZoneLabel = null!;
    private Label _inputLabel = null!;
    private TextBox _inputPath = null!;
    private Button _inputBrowse = null!;
    private Label _outputLabel = null!;
    private TextBox _outputPath = null!;
    private Button _outputBrowse = null!;
    private Button _convertButton = null!;
    private ProgressBar _progressBar = null!;
    private Label _statusLabel = null!;

    public ConverterPanel(
        string inputExt,
        string inputTypeName,
        string dropZoneText,
        Action<string, string, IProgress<int>?> converter)
    {
        _inputExt = inputExt;
        _inputTypeName = inputTypeName;
        _dropZoneText = dropZoneText;
        _converter = converter;

        // ВАЖНО: задать размер панели ДО добавления детей.
        // Иначе якоря (Anchor=Top|Left|Right) у дочерних контролов запомнят
        // расстояния от краёв относительно дефолтного крошечного размера панели,
        // и при разворачивании Dock=Fill контролы вылезут за пределы видимой области.
        ClientSize = new Size(620, 380);

        AllowDrop = true;
        DragEnter += OnPanelDragEnter;
        DragDrop += OnPanelDragDrop;

        BuildControls();

        // Dock устанавливаем последним — он применится, когда панель попадёт в родителя.
        Dock = DockStyle.Fill;
    }

    private void BuildControls()
    {
        // Зона перетаскивания.
        _dropZone = new DropZone
        {
            Location = new Point(20, 20),
            Size = new Size(580, 90),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            AllowDrop = true
        };
        _dropZone.DragEnter += OnDropZoneDragEnter;
        _dropZone.DragLeave += OnDropZoneDragLeave;
        _dropZone.DragDrop += OnDropZoneDragDrop;

        _dropZoneLabel = new Label
        {
            Text = _dropZoneText,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 11F, FontStyle.Regular),
            ForeColor = Color.FromArgb(90, 90, 90),
            BackColor = Color.Transparent,
            AllowDrop = true
        };
        _dropZoneLabel.DragEnter += OnDropZoneDragEnter;
        _dropZoneLabel.DragLeave += OnDropZoneDragLeave;
        _dropZoneLabel.DragDrop += OnDropZoneDragDrop;
        _dropZone.Controls.Add(_dropZoneLabel);

        _inputLabel = new Label
        {
            Text = "Или выберите путь вручную:",
            Location = new Point(20, 125),
            AutoSize = true
        };

        _inputPath = new TextBox
        {
            Location = new Point(20, 150),
            Size = new Size(480, 25),
            ReadOnly = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        _inputBrowse = new Button
        {
            Text = "Выбрать...",
            Location = new Point(510, 149),
            Size = new Size(90, 27),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _inputBrowse.Click += OnInputBrowseClick;

        _outputLabel = new Label
        {
            Text = "Выходной файл (PDF):",
            Location = new Point(20, 185),
            AutoSize = true
        };

        _outputPath = new TextBox
        {
            Location = new Point(20, 210),
            Size = new Size(480, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        _outputBrowse = new Button
        {
            Text = "Сохранить как...",
            Location = new Point(510, 209),
            Size = new Size(90, 27),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _outputBrowse.Click += OnOutputBrowseClick;

        _convertButton = new Button
        {
            Text = "Конвертировать",
            Location = new Point(20, 255),
            Size = new Size(580, 36),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold)
        };
        _convertButton.Click += OnConvertClick;

        _progressBar = new ProgressBar
        {
            Location = new Point(20, 305),
            Size = new Size(580, 18),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Style = ProgressBarStyle.Continuous,
            Minimum = 0,
            Maximum = 100,
            Value = 0
        };

        _statusLabel = new Label
        {
            Text = "Готов к работе.",
            Location = new Point(20, 333),
            AutoSize = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Left
        };

        Controls.Add(_dropZone);
        Controls.Add(_inputLabel);
        Controls.Add(_inputPath);
        Controls.Add(_inputBrowse);
        Controls.Add(_outputLabel);
        Controls.Add(_outputPath);
        Controls.Add(_outputBrowse);
        Controls.Add(_convertButton);
        Controls.Add(_progressBar);
        Controls.Add(_statusLabel);
    }

    // ---------- Кнопки выбора файла ----------

    private void OnInputBrowseClick(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = $"Выберите {_inputExt.TrimStart('.').ToUpperInvariant()} файл",
            Filter = $"{_inputTypeName} (*{_inputExt})|*{_inputExt}|Все файлы (*.*)|*.*",
            CheckFileExists = true
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _inputPath.Text = dialog.FileName;
            if (string.IsNullOrWhiteSpace(_outputPath.Text))
                _outputPath.Text = Path.ChangeExtension(dialog.FileName, ".pdf");
        }
    }

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
            _outputPath.Text = dialog.FileName;
    }

    // ---------- Запуск конвертации ----------

    private async void OnConvertClick(object? sender, EventArgs e)
    {
        var input = _inputPath.Text.Trim();
        var output = _outputPath.Text.Trim();

        if (string.IsNullOrEmpty(input)) { ShowWarning($"Не выбран входной {_inputExt.ToUpperInvariant()} файл."); return; }
        if (!File.Exists(input)) { ShowWarning("Входной файл не найден."); return; }
        if (string.IsNullOrEmpty(output)) { ShowWarning("Не указан путь для выходного PDF файла."); return; }

        SetBusy(true, "Идёт конвертация...");

        var progress = new Progress<int>(percent =>
        {
            _progressBar.Value = Math.Clamp(percent, 0, 100);
            _statusLabel.Text = $"Идёт конвертация... {percent}%";
        });

        try
        {
            await Task.Run(() => _converter(input, output, progress));

            SetBusy(false, $"Готово. Файл сохранён: {output}");
            MessageBox.Show(this,
                $"Конвертация завершена.\n\n{output}",
                "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            SetBusy(false, "Ошибка при конвертации.");
            MessageBox.Show(this,
                $"Не удалось сконвертировать файл.\n\n{ex.Message}",
                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SetBusy(bool busy, string status)
    {
        _convertButton.Enabled = !busy;
        _inputBrowse.Enabled = !busy;
        _outputBrowse.Enabled = !busy;
        if (busy) _progressBar.Value = 0;
        _statusLabel.Text = status;
    }

    private void ShowWarning(string text)
    {
        MessageBox.Show(this, text, "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    // ---------- Обработка перетаскивания ----------

    private void OnPanelDragEnter(object? sender, DragEventArgs e) => SetDragEffect(e);
    private void OnPanelDragDrop(object? sender, DragEventArgs e) => HandleDrop(e);

    private void OnDropZoneDragEnter(object? sender, DragEventArgs e)
    {
        SetDragEffect(e);
        _dropZone.IsDragHover = e.Effect == DragDropEffects.Copy;
    }

    private void OnDropZoneDragLeave(object? sender, EventArgs e) => _dropZone.IsDragHover = false;

    private void OnDropZoneDragDrop(object? sender, DragEventArgs e)
    {
        _dropZone.IsDragHover = false;
        HandleDrop(e);
    }

    private void SetDragEffect(DragEventArgs e)
    {
        if (e.Data != null && TryGetSourcePath(e.Data, _inputExt, out _))
            e.Effect = DragDropEffects.Copy;
        else
            e.Effect = DragDropEffects.None;
    }

    private void HandleDrop(DragEventArgs e)
    {
        if (e.Data == null || !TryGetSourcePath(e.Data, _inputExt, out var path)) return;

        _inputPath.Text = path;
        if (string.IsNullOrWhiteSpace(_outputPath.Text))
            _outputPath.Text = Path.ChangeExtension(path, ".pdf");

        _statusLabel.Text = $"Файл загружен: {Path.GetFileName(path)}";
    }

    // Ищет в перетаскиваемых файлах первый, у которого нужное расширение.
    private static bool TryGetSourcePath(IDataObject data, string ext, out string path)
    {
        path = string.Empty;
        if (!data.GetDataPresent(DataFormats.FileDrop)) return false;
        if (data.GetData(DataFormats.FileDrop) is not string[] files) return false;

        var match = files.FirstOrDefault(f =>
            !string.IsNullOrEmpty(f) &&
            File.Exists(f) &&
            string.Equals(Path.GetExtension(f), ext, StringComparison.OrdinalIgnoreCase));

        if (match == null) return false;
        path = match;
        return true;
    }
}
