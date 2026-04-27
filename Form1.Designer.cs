namespace DocxToPdfConverter;

partial class Form1
{
    private System.ComponentModel.IContainer components = null!;

    private TabControl _tabs = null!;
    private TabPage _docxTab = null!;
    private TabPage _pptxTab = null!;
    private ConverterPanel _docxPanel = null!;
    private ConverterPanel _pptxPanel = null!;

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
        ClientSize = new Size(640, 400);
        Text = "Конвертер OOXML → PDF";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(540, 400);
        Font = new Font("Segoe UI", 9F);

        // Корневой элемент — TabControl с двумя вкладками.
        _tabs = new TabControl
        {
            Dock = DockStyle.Fill,
            Padding = new Point(12, 4)
        };

        // Вкладка DOCX → PDF.
        _docxTab = new TabPage("DOCX → PDF") { Padding = new Padding(0) };
        _docxPanel = new ConverterPanel(
            inputExt: ".docx",
            inputTypeName: "Word документы",
            dropZoneText: "Перетащите DOCX файл сюда",
            converter: (input, output, progress) => new Converter().Convert(input, output, progress));
        _docxTab.Controls.Add(_docxPanel);

        // Вкладка PPTX → PDF.
        _pptxTab = new TabPage("PPTX → PDF") { Padding = new Padding(0) };
        _pptxPanel = new ConverterPanel(
            inputExt: ".pptx",
            inputTypeName: "PowerPoint презентации",
            dropZoneText: "Перетащите PPTX файл сюда",
            converter: (input, output, progress) => new PptxConverter().Convert(input, output, progress));
        _pptxTab.Controls.Add(_pptxPanel);

        _tabs.TabPages.Add(_docxTab);
        _tabs.TabPages.Add(_pptxTab);

        Controls.Add(_tabs);
    }
}
