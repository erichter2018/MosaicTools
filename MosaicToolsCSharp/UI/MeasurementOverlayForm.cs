using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;
using MosaicTools.Services;
using static MosaicTools.Services.OcrService;

namespace MosaicTools.UI;

/// <summary>
/// Transparent overlay showing bidimensional measurement lines and text.
/// Auto-dismisses after 5 seconds. Click to dismiss early.
/// </summary>
public class MeasurementOverlayForm : Form
{
    private Bitmap? _bitmap;
    private System.Windows.Forms.Timer? _dismissTimer;

    private static MeasurementOverlayForm? _instance;

    public static void ShowMeasurement(MeasurementResult result)
    {
        CloseIfOpen();
        _instance = new MeasurementOverlayForm(result);
        _instance.Show();
    }

    public static void CloseIfOpen()
    {
        if (_instance != null && !_instance.IsDisposed)
        {
            _instance.Close();
            _instance = null;
        }
    }

    private MeasurementOverlayForm(MeasurementResult result)
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;

        // Bounding box of all measurement points + padding for text
        int pad = 80;
        int minX = Math.Min(Math.Min(result.MajorStart.X, result.MajorEnd.X),
            Math.Min(result.MinorStart.X, result.MinorEnd.X));
        int maxX = Math.Max(Math.Max(result.MajorStart.X, result.MajorEnd.X),
            Math.Max(result.MinorStart.X, result.MinorEnd.X));
        int minY = Math.Min(Math.Min(result.MajorStart.Y, result.MajorEnd.Y),
            Math.Min(result.MinorStart.Y, result.MinorEnd.Y));
        int maxY = Math.Max(Math.Max(result.MajorStart.Y, result.MajorEnd.Y),
            Math.Max(result.MinorStart.Y, result.MinorEnd.Y));

        int formX = minX - pad;
        int formY = minY - pad;
        int formW = Math.Max(200, maxX - minX + pad * 2);
        int formH = Math.Max(100, maxY - minY + pad * 2);

        Location = new Point(formX, formY);
        Size = new Size(formW, formH);

        // Offset for converting screen coords to local
        int ox = -formX, oy = -formY;

        _bitmap = new Bitmap(formW, formH, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(_bitmap))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

            // Measurement lines
            using var majorPen = new Pen(Color.FromArgb(220, 0, 255, 200), 2.5f);
            using var minorPen = new Pen(Color.FromArgb(220, 255, 200, 0), 2.5f);

            g.DrawLine(majorPen,
                result.MajorStart.X + ox, result.MajorStart.Y + oy,
                result.MajorEnd.X + ox, result.MajorEnd.Y + oy);
            g.DrawLine(minorPen,
                result.MinorStart.X + ox, result.MinorStart.Y + oy,
                result.MinorEnd.X + ox, result.MinorEnd.Y + oy);

            // Endpoint dots
            int dotR = 4;
            using var dotBrushMaj = new SolidBrush(Color.FromArgb(230, 0, 255, 200));
            using var dotBrushMin = new SolidBrush(Color.FromArgb(230, 255, 200, 0));
            g.FillEllipse(dotBrushMaj, result.MajorStart.X + ox - dotR, result.MajorStart.Y + oy - dotR, dotR * 2, dotR * 2);
            g.FillEllipse(dotBrushMaj, result.MajorEnd.X + ox - dotR, result.MajorEnd.Y + oy - dotR, dotR * 2, dotR * 2);
            g.FillEllipse(dotBrushMin, result.MinorStart.X + ox - dotR, result.MinorStart.Y + oy - dotR, dotR * 2, dotR * 2);
            g.FillEllipse(dotBrushMin, result.MinorEnd.X + ox - dotR, result.MinorEnd.Y + oy - dotR, dotR * 2, dotR * 2);

            // Text label
            string text = $"{result.MajorAxisCm:F1} \u00d7 {result.MinorAxisCm:F1} cm";
            using var font = new Font("Segoe UI", 12, FontStyle.Bold);
            var textSize = g.MeasureString(text, font);

            // Position label above the lines with clear separation
            int linesTopY = Math.Min(Math.Min(result.MajorStart.Y, result.MajorEnd.Y),
                Math.Min(result.MinorStart.Y, result.MinorEnd.Y)) + oy;
            int linesBottomY = Math.Max(Math.Max(result.MajorStart.Y, result.MajorEnd.Y),
                Math.Max(result.MinorStart.Y, result.MinorEnd.Y)) + oy;
            int textX = result.ScreenCenter.X + ox - (int)(textSize.Width / 2);
            int textY = linesTopY - (int)textSize.Height - 14;
            // If no room above, place below the lines
            if (textY < 4) textY = linesBottomY + 14;
            textX = Math.Max(4, Math.Min(formW - (int)textSize.Width - 4, textX));
            textY = Math.Max(4, Math.Min(formH - (int)textSize.Height - 4, textY));

            // Semi-transparent background
            using var bgBrush = new SolidBrush(Color.FromArgb(180, 0, 0, 0));
            g.FillRectangle(bgBrush, textX - 6, textY - 3,
                textSize.Width + 12, textSize.Height + 6);

            using var textBrush = new SolidBrush(Color.FromArgb(255, 0, 255, 200));
            g.DrawString(text, font, textBrush, textX, textY);
        }

        LayeredWindowHelper.PremultiplyBitmapAlpha(_bitmap);

        // Click to dismiss
        MouseClick += (_, _) => Close();

        // Auto-dismiss after 5 seconds
        _dismissTimer = new System.Windows.Forms.Timer { Interval = 5000 };
        _dismissTimer.Tick += (_, _) => { _dismissTimer.Stop(); Close(); };
        _dismissTimer.Start();
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        if (_bitmap != null)
            LayeredWindowHelper.SetBitmap(this, _bitmap);
    }

    protected override void OnPaint(PaintEventArgs e) { }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _dismissTimer?.Dispose();
        _bitmap?.Dispose();
        _bitmap = null;
        base.OnFormClosed(e);
        if (_instance == this) _instance = null;
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= 0x80;    // WS_EX_TOOLWINDOW — hide from Alt+Tab
            cp.ExStyle |= 0x80000; // WS_EX_LAYERED — per-pixel alpha
            return cp;
        }
    }
}
