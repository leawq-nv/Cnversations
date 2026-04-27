namespace DocxToPdfConverter;

partial class Form1
{
    private System.ComponentModel.IContainer components = null!;

    // Поля для контролов формы.
    private Label _inputLabel = null!;
    private TextBox _inputPath = null!;
    private Button _inputBrowse = null!;

    private Label _outputLabel = null!;
    private TextBox _outputPath = null!;
    private Button _outputBrowse = null!;

    private Button _convertButton = null!;
    private ProgressBar _progressBar = null!;
    private Label _statusLabel = null!;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(620, 260);
        Text = "Конвертер DOCX → PDF";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(500, 260);
        Font = new Font("Segoe UI", 9F);
        AllowDrop = true;
        DragEnter += OnFormDragEnter;
        DragDrop += OnFormDragDrop;

        // Метка "Входной файл".
        _inputLabel = new Label
        {
            Text = "Входной файл (DOCX):",
            Location = new Point(20, 20),
            AutoSize = true
        };

        // Поле пути входного файла.
        _inputPath = new TextBox
        {
            Location = new Point(20, 45),
            Size = new Size(480, 25),
            ReadOnly = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        // Кнопка выбора входного файла.
        _inputBrowse = new Button
        {
            Text = "Выбрать...",
            Location = new Point(510, 44),
            Size = new Size(90, 27),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _inputBrowse.Click += OnInputBrowseClick;

        // Метка "Выходной файл".
        _outputLabel = new Label
        {
            Text = "Выходной файл (PDF):",
            Location = new Point(20, 80),
            AutoSize = true
        };

        // Поле пути выходного файла.
        _outputPath = new TextBox
        {
            Location = new Point(20, 105),
            Size = new Size(480, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        // Кнопка выбора выходного файла.
        _outputBrowse = new Button
        {
            Text = "Сохранить как...",
            Location = new Point(510, 104),
            Size = new Size(90, 27),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _outputBrowse.Click += OnOutputBrowseClick;

        // Главная кнопка "Конвертировать".
        _convertButton = new Button
        {
            Text = "Конвертировать",
            Location = new Point(20, 150),
            Size = new Size(580, 36),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Font = new Font("Segoe UI", 10F, FontStyle.Bold)
        };
        _convertButton.Click += OnConvertClick;

        // Полоса прогресса (процентная — заполняется по мере конвертации).
        _progressBar = new ProgressBar
        {
            Location = new Point(20, 200),
            Size = new Size(580, 18),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Style = ProgressBarStyle.Continuous,
            Minimum = 0,
            Maximum = 100,
            Value = 0
        };

        // Статусная строка.
        _statusLabel = new Label
        {
            Text = "Готов к работе. Можно перетащить .docx файл прямо в окно.",
            Location = new Point(20, 225),
            AutoSize = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Left
        };

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
}
