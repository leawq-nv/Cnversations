using System.ComponentModel;
using System.Drawing.Drawing2D;

namespace DocxToPdfConverter;

// Панель-приёмник для перетаскивания файлов.
// Рисует пунктирную рамку и подсвечивает себя, когда над ней тащат подходящий файл.
public class DropZone : Panel
{
    private bool _isDragHover;

    // [Browsable(false)] + [DesignerSerializationVisibility(Hidden)] говорят дизайнеру:
    // не сохранять это свойство в Designer.cs — это runtime-состояние, а не настройка.
    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool IsDragHover
    {
        get => _isDragHover;
        set
        {
            if (_isDragHover == value) return;
            _isDragHover = value;
            Invalidate(); // перерисовать панель при изменении состояния
        }
    }

    public DropZone()
    {
        // Двойная буферизация — рамка не мерцает при перерисовке.
        SetStyle(
            ControlStyles.UserPaint |
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw,
            true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);

        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        // Цвета зависят от состояния: обычное / при наведении файла.
        var borderColor = _isDragHover ? Color.SteelBlue : Color.FromArgb(160, 160, 160);
        var fillColor = _isDragHover ? Color.FromArgb(225, 238, 250) : Color.FromArgb(248, 248, 248);

        using (var brush = new SolidBrush(fillColor))
        {
            g.FillRectangle(brush, ClientRectangle);
        }

        using var pen = new Pen(borderColor, 2)
        {
            DashStyle = DashStyle.Dash,
            DashPattern = new float[] { 4, 4 }
        };

        // Рисуем рамку чуть внутри границ панели, чтобы линия не обрезалась.
        var rect = new Rectangle(1, 1, Width - 3, Height - 3);
        g.DrawRectangle(pen, rect);
    }
}
