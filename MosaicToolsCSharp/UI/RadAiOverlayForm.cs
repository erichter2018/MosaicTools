using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// [RadAI] RadAI impression overlay — positioned below the report popup.
/// Matches ReportPopupForm visual style exactly (transparent or opaque mode).
/// Used when ShowReportAfterProcess is enabled; otherwise the standalone popup is used.
/// </summary>
public class RadAiOverlayForm : Form
{
    #region P/Invoke for Layered Window

    [StructLayout(LayoutKind.Sequential)]
    private struct W32Point { public int X, Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct W32Size { public int Width, Height; }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    private struct BLENDFUNCTION
    {
        public byte BlendOp;
        public byte BlendFlags;
        public byte SourceConstantAlpha;
        public byte AlphaFormat;
    }

    private const int WS_EX_LAYERED = 0x80000;
    private const byte AC_SRC_OVER = 0x00;
    private const byte AC_SRC_ALPHA = 0x01;
    private const uint ULW_ALPHA = 0x02;

    [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
    private static extern bool UpdateLayeredWindow(
        IntPtr hwnd, IntPtr hdcDst, ref W32Point pptDst, ref W32Size psize,
        IntPtr hdcSrc, ref W32Point pptSrc, uint crKey,
        ref BLENDFUNCTION pblend, uint dwFlags);

    [DllImport("user32.dll", ExactSpelling = true)]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll", ExactSpelling = true)]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("gdi32.dll", ExactSpelling = true)]
    private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll", ExactSpelling = true)]
    private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

    [DllImport("gdi32.dll", ExactSpelling = true)]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("gdi32.dll", ExactSpelling = true)]
    private static extern bool DeleteDC(IntPtr hdc);

    [DllImport("user32.dll")]
    private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();

    private const int WM_NCLBUTTONDOWN = 0xA1;
    private const int HT_CAPTION = 0x2;

    #endregion

    private readonly bool _useLayeredWindow;
    private readonly string _formattedText;
    private readonly int _backgroundAlpha;
    private readonly Font _normalFont;
    private readonly Font _headerFont;
    private const string HeaderText = "RadAI Impression";

    // Transparent mode
    private int _scrollOffset;
    private int _totalContentHeight;

    // Opaque mode
    private RichTextBox? _richTextBox;
    private bool _isResizing;

    // Linked form tracking
    private Form? _linkedForm;

    // Drag tracking (transparent mode)
    private Point _formPosOnMouseDown;

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            if (_useLayeredWindow)
                cp.ExStyle |= WS_EX_LAYERED;
            return cp;
        }
    }

    public RadAiOverlayForm(Configuration config, string[] impressionItems)
    {
        _useLayeredWindow = config.ReportPopupTransparent;
        _backgroundAlpha = Math.Clamp((int)Math.Round(config.ReportPopupTransparency / 100.0 * 255), 30, 255);

        _normalFont = new Font(config.ReportPopupFontFamily, config.ReportPopupFontSize);
        _headerFont = new Font(config.ReportPopupFontFamily, config.ReportPopupFontSize + 2, FontStyle.Bold);

        // Format text: header + numbered items
        var sb = new StringBuilder();
        sb.AppendLine(HeaderText);
        for (int i = 0; i < impressionItems.Length; i++)
        {
            var item = ReportPopupForm.SanitizeText(impressionItems[i]);
            if (impressionItems.Length >= 2)
                sb.AppendLine($"{i + 1}. {item}");
            else
                sb.AppendLine(item);
        }
        _formattedText = sb.ToString().TrimEnd().Replace("\r\n", "\n");

        FormBorderStyle = FormBorderStyle.None;
        Text = "";
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;

        int width = Math.Max(config.ReportPopupWidth, 600);

        if (_useLayeredWindow)
            SetupTransparentMode(width);
        else
            SetupOpaqueMode(width);
    }

    #region Setup

    private void SetupTransparentMode(int width)
    {
        AutoSize = false;
        AutoScroll = false;

        int contentHeight = MeasureContentHeight(width);
        _totalContentHeight = contentHeight;

        var screen = Screen.PrimaryScreen ?? Screen.AllScreens[0];
        int maxHeight = screen.WorkingArea.Height - 50;
        int formHeight = Math.Clamp(contentHeight, 60, maxHeight);

        this.Size = new Size(width, formHeight);

        this.MouseDown += OnLayeredMouseDown;
        this.MouseWheel += OnLayeredMouseWheel;

        this.Load += (s, e) =>
        {
            this.Activate();
            RenderAndUpdate();
        };
    }

    private void SetupOpaqueMode(int width)
    {
        DoubleBuffered = true;
        BackColor = Color.FromArgb(30, 30, 30);
        AutoSize = false;
        AutoScroll = false;

        int padding = 20;

        _richTextBox = new RichTextBox
        {
            Width = width - (padding * 2),
            Location = new Point(padding, padding),
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.Gainsboro,
            Font = _normalFont,
            Cursor = Cursors.Arrow,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.None,
            DetectUrls = false,
            ShortcutsEnabled = false
        };

        _richTextBox.ContentsResized += (s, e) =>
        {
            if (_isResizing || !IsHandleCreated) return;
            BeginInvoke(new Action(() =>
            {
                if (_isResizing) return;
                _isResizing = true;
                try { PerformResize(); }
                finally { _isResizing = false; }
            }));
        };

        _richTextBox.Text = _formattedText;
        Controls.Add(_richTextBox);

        this.ClientSize = new Size(width, 100);

        SetupOpaqueInteractions(this);
        SetupOpaqueInteractions(_richTextBox);

        this.Load += (s, e) =>
        {
            this.Activate();
            ActiveControl = null;
            FormatHeader();
            PerformResize();
        };
    }

    #endregion

    #region Linked Positioning

    /// <summary>
    /// Link this overlay to follow a target form (always position below it).
    /// </summary>
    public void LinkToForm(Form target)
    {
        _linkedForm = target;
        target.LocationChanged += OnLinkedFormMoved;
        target.SizeChanged += OnLinkedFormMoved;
        target.FormClosed += OnLinkedFormClosed;
        PositionBelowLinkedForm();
    }

    private void OnLinkedFormMoved(object? sender, EventArgs e)
    {
        PositionBelowLinkedForm();
    }

    private void OnLinkedFormClosed(object? sender, EventArgs e)
    {
        if (_linkedForm != null)
        {
            _linkedForm.LocationChanged -= OnLinkedFormMoved;
            _linkedForm.SizeChanged -= OnLinkedFormMoved;
            _linkedForm.FormClosed -= OnLinkedFormClosed;
            _linkedForm = null;
        }
    }

    private void PositionBelowLinkedForm()
    {
        if (_linkedForm == null || _linkedForm.IsDisposed) return;

        int x = _linkedForm.Left;
        int y = _linkedForm.Bottom + 5;

        // Clamp to screen bounds
        var screen = Screen.FromControl(_linkedForm);
        if (y + Height > screen.WorkingArea.Bottom)
            y = screen.WorkingArea.Bottom - Height;
        if (x + Width > screen.WorkingArea.Right)
            x = screen.WorkingArea.Right - Width;

        var newLoc = new Point(x, y);
        if (this.Location != newLoc)
        {
            this.Location = newLoc;
            if (_useLayeredWindow && IsHandleCreated && !IsDisposed)
                RenderAndUpdate();
        }
    }

    #endregion

    #region Transparent Mode - Rendering

    private void RenderAndUpdate()
    {
        if (!IsHandleCreated || IsDisposed) return;

        int w = Width, h = Height;

        // Layer 1: semi-transparent background
        using var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.FromArgb(_backgroundAlpha, 20, 20, 20));
        }

        // Layer 2: ClearType text on opaque dark background
        using var textLayer = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(textLayer))
        {
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.Clear(Color.FromArgb(255, 20, 20, 20));
            DrawContent(g);
        }

        MergeTextLayer(bmp, textLayer);
        PremultiplyBitmapAlpha(bmp);
        SetBitmap(bmp);
    }

    private void DrawContent(Graphics g)
    {
        int padding = 20;
        int maxWidth = Width - padding * 2;
        float y = padding - _scrollOffset;

        using var sf = new StringFormat(StringFormat.GenericTypographic);
        sf.Trimming = StringTrimming.Word;
        sf.FormatFlags = 0;

        var lines = _formattedText.Split('\n');

        foreach (var line in lines)
        {
            if (string.IsNullOrEmpty(line))
            {
                y += _normalFont.Height;
                continue;
            }

            bool isHeader = (line == HeaderText);
            var font = isHeader ? _headerFont : _normalFont;

            var layoutRect = new RectangleF(padding, y, maxWidth, 10000);
            var measured = g.MeasureString(line, font, maxWidth, sf);
            float lineHeight = Math.Max(measured.Height, font.Height);

            if (y + lineHeight > 0 && y < Height)
            {
                using var brush = new SolidBrush(Color.White);
                g.DrawString(line, font, brush, layoutRect, sf);
            }

            y += lineHeight;
        }

        _totalContentHeight = (int)(y + _scrollOffset + padding);
    }

    private int MeasureContentHeight(int width)
    {
        int padding = 20;
        int maxWidth = width - padding * 2;
        float y = padding;

        using var bmp = new Bitmap(1, 1, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

        using var sf = new StringFormat(StringFormat.GenericTypographic);
        sf.Trimming = StringTrimming.Word;

        var lines = _formattedText.Split('\n');
        foreach (var line in lines)
        {
            if (string.IsNullOrEmpty(line))
            {
                y += _normalFont.Height;
                continue;
            }

            bool isHeader = (line == HeaderText);
            var font = isHeader ? _headerFont : _normalFont;

            var measured = g.MeasureString(line, font, maxWidth, sf);
            y += Math.Max(measured.Height, font.Height);
        }

        return (int)y + padding;
    }

    /// <summary>
    /// Stamp ClearType text pixels from textLayer onto dst bitmap.
    /// Same technique as ReportPopupForm for crisp text on transparent background.
    /// </summary>
    private static void MergeTextLayer(Bitmap dst, Bitmap textLayer)
    {
        var rect = new Rectangle(0, 0, dst.Width, dst.Height);
        var dstData = dst.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        var srcData = textLayer.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

        int byteCount = Math.Abs(dstData.Stride) * dstData.Height;
        byte[] dstPx = new byte[byteCount];
        byte[] srcPx = new byte[byteCount];
        Marshal.Copy(dstData.Scan0, dstPx, 0, byteCount);
        Marshal.Copy(srcData.Scan0, srcPx, 0, byteCount);

        const byte bgB = 20, bgG = 20, bgR = 20;
        const int threshold = 8;

        for (int i = 0; i < byteCount; i += 4)
        {
            int diff = Math.Abs(srcPx[i] - bgB) +
                       Math.Abs(srcPx[i + 1] - bgG) +
                       Math.Abs(srcPx[i + 2] - bgR);

            if (diff > threshold)
            {
                dstPx[i]     = srcPx[i];
                dstPx[i + 1] = srcPx[i + 1];
                dstPx[i + 2] = srcPx[i + 2];
                dstPx[i + 3] = 255;
            }
        }

        Marshal.Copy(dstPx, 0, dstData.Scan0, byteCount);
        dst.UnlockBits(dstData);
        textLayer.UnlockBits(srcData);
    }

    private static void PremultiplyBitmapAlpha(Bitmap bmp)
    {
        var rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
        var data = bmp.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        int byteCount = Math.Abs(data.Stride) * data.Height;
        byte[] pixels = new byte[byteCount];
        Marshal.Copy(data.Scan0, pixels, 0, byteCount);

        for (int i = 0; i < byteCount; i += 4)
        {
            byte a = pixels[i + 3];
            if (a == 255) continue;
            if (a == 0)
            {
                pixels[i] = pixels[i + 1] = pixels[i + 2] = 0;
                continue;
            }
            pixels[i]     = (byte)(pixels[i]     * a / 255);
            pixels[i + 1] = (byte)(pixels[i + 1] * a / 255);
            pixels[i + 2] = (byte)(pixels[i + 2] * a / 255);
        }

        Marshal.Copy(pixels, 0, data.Scan0, byteCount);
        bmp.UnlockBits(data);
    }

    private void SetBitmap(Bitmap bitmap)
    {
        if (!IsHandleCreated || IsDisposed) return;

        IntPtr screenDc = IntPtr.Zero, memDc = IntPtr.Zero, hBitmap = IntPtr.Zero, oldBitmap = IntPtr.Zero;

        try
        {
            screenDc = GetDC(IntPtr.Zero);
            memDc = CreateCompatibleDC(screenDc);
            hBitmap = bitmap.GetHbitmap(Color.FromArgb(0));
            oldBitmap = SelectObject(memDc, hBitmap);
            var blend = new BLENDFUNCTION
            {
                BlendOp = AC_SRC_OVER,
                BlendFlags = 0,
                SourceConstantAlpha = 255,
                AlphaFormat = AC_SRC_ALPHA
            };

            var size = new W32Size { Width = bitmap.Width, Height = bitmap.Height };
            var source = new W32Point { X = 0, Y = 0 };
            var topPos = new W32Point { X = Left, Y = Top };

            UpdateLayeredWindow(Handle, screenDc, ref topPos, ref size,
                memDc, ref source, 0, ref blend, ULW_ALPHA);
        }
        finally
        {
            if (oldBitmap != IntPtr.Zero) SelectObject(memDc, oldBitmap);
            if (hBitmap != IntPtr.Zero) DeleteObject(hBitmap);
            if (memDc != IntPtr.Zero) DeleteDC(memDc);
            if (screenDc != IntPtr.Zero) ReleaseDC(IntPtr.Zero, screenDc);
        }
    }

    #endregion

    #region Transparent Mode - Mouse Handling

    private void OnLayeredMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right)
        {
            Close();
            return;
        }

        // Only allow drag when not linked to another form
        if (e.Button == MouseButtons.Left && _linkedForm == null)
        {
            _formPosOnMouseDown = this.Location;
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);

            if (this.Location != _formPosOnMouseDown)
                RenderAndUpdate();
        }
    }

    private void OnLayeredMouseWheel(object? sender, MouseEventArgs e)
    {
        if (_totalContentHeight <= Height) return;

        int scrollAmount = _normalFont.Height * 3;
        _scrollOffset -= e.Delta > 0 ? scrollAmount : -scrollAmount;

        int maxScroll = Math.Max(0, _totalContentHeight - Height);
        _scrollOffset = Math.Clamp(_scrollOffset, 0, maxScroll);

        RenderAndUpdate();
    }

    #endregion

    #region Opaque Mode

    private void FormatHeader()
    {
        if (_richTextBox == null) return;

        int headerIdx = _richTextBox.Text.IndexOf(HeaderText, StringComparison.Ordinal);
        if (headerIdx >= 0)
        {
            _richTextBox.Select(headerIdx, HeaderText.Length);
            _richTextBox.SelectionFont = _headerFont;
            _richTextBox.SelectionColor = Color.White;
        }
        _richTextBox.Select(0, 0);
    }

    private void PerformResize()
    {
        if (_richTextBox == null) return;

        int contentHeight = 0;
        int textLength = _richTextBox.TextLength;

        if (textLength > 0)
        {
            Point pt = _richTextBox.GetPositionFromCharIndex(textLength - 1);
            contentHeight = pt.Y + _richTextBox.Font.Height + 10;
        }
        else
        {
            contentHeight = _richTextBox.Font.Height;
        }

        int padding = 20;
        int requiredHeight = contentHeight + (padding * 2);
        if (requiredHeight < 60) requiredHeight = 60;

        var screen = Screen.FromControl(this);
        int maxHeight = screen.WorkingArea.Height - 50;

        int finalHeight;
        bool needsScroll;

        if (requiredHeight > maxHeight)
        {
            finalHeight = maxHeight;
            needsScroll = true;
        }
        else
        {
            finalHeight = requiredHeight;
            needsScroll = false;
        }

        if (needsScroll && _richTextBox.ScrollBars != RichTextBoxScrollBars.Vertical)
            _richTextBox.ScrollBars = RichTextBoxScrollBars.Vertical;
        else if (!needsScroll && _richTextBox.ScrollBars != RichTextBoxScrollBars.None)
            _richTextBox.ScrollBars = RichTextBoxScrollBars.None;

        if (this.ClientSize.Height != finalHeight)
            this.ClientSize = new Size(this.ClientSize.Width, finalHeight);

        int rtbHeight = finalHeight - (padding * 2);
        if (_richTextBox.Height != rtbHeight)
            _richTextBox.Height = rtbHeight;
    }

    private void SetupOpaqueInteractions(Control control)
    {
        Point dragStart = Point.Empty;
        bool dragging = false;

        control.MouseDown += (s, e) =>
        {
            if (e.Button == MouseButtons.Right)
            {
                Close();
                return;
            }

            if (e.Button == MouseButtons.Left && _linkedForm == null)
            {
                if (control != this)
                {
                    ReleaseCapture();
                    SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                    dragging = false;
                }
                else
                {
                    dragStart = new Point(e.X, e.Y);
                    dragging = true;
                }
            }
        };

        control.MouseMove += (s, e) =>
        {
            if (dragging)
            {
                Point currentScreenPos = Cursor.Position;
                Location = new Point(
                    currentScreenPos.X - dragStart.X,
                    currentScreenPos.Y - dragStart.Y);
            }
        };

        control.MouseUp += (s, e) =>
        {
            dragging = false;
        };
    }

    #endregion

    #region Disposal

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (_linkedForm != null)
            {
                _linkedForm.LocationChanged -= OnLinkedFormMoved;
                _linkedForm.SizeChanged -= OnLinkedFormMoved;
                _linkedForm.FormClosed -= OnLinkedFormClosed;
                _linkedForm = null;
            }
            _normalFont?.Dispose();
            _headerFont?.Dispose();
        }
        base.Dispose(disposing);
    }

    #endregion
}
