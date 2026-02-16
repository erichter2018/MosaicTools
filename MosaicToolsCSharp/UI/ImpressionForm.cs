using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Floating window displaying the Impression from the report.
/// Shown after ProcessReport, hidden on SignReport.
/// Supports transparent (layered window) and opaque (standard controls) rendering modes.
/// </summary>
public class ImpressionForm : Form
{
    private readonly Configuration _config;
    private readonly Label _contentLabel;
    private readonly bool _useLayeredWindow;
    private readonly int _backgroundAlpha;
    private readonly Font _contentFont;

    // Transparent mode
    private ContextMenuStrip? _dragBarMenu;
    private Point _formPosOnMouseDown;

    // Layout constants
    private const int BorderWidth = 2;
    private const int DragBarHeight = 16;
    private const int ContentMarginLeft = 10;
    private const int ContentMarginRight = 10;
    private const int ContentMarginTop = 5;
    private const int ContentMarginBottom = 15;
    private const int MaxContentWidth = 500;

    // Drag state
    private Point _dragStart;
    private bool _dragging;

    // For center-based positioning
    private bool _initialPositionSet = false;

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            if (_useLayeredWindow)
                cp.ExStyle |= LayeredWindowHelper.WS_EX_LAYERED;
            return cp;
        }
    }

    public ImpressionForm(Configuration config)
    {
        _config = config;
        _useLayeredWindow = config.ReportPopupTransparent;
        _backgroundAlpha = Math.Clamp((int)Math.Round(config.ReportPopupTransparency / 100.0 * 255), 30, 255);
        _contentFont = new Font("Segoe UI", 11, FontStyle.Bold);

        // Form properties - frameless, topmost
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(60, 60, 60); // Border color
        StartPosition = FormStartPosition.Manual;

        // Content label - always created (data holder in transparent mode, visible in opaque)
        _contentLabel = new Label
        {
            Text = "(Waiting for impression...)",
            Font = _contentFont,
            ForeColor = Color.FromArgb(128, 128, 128),
            BackColor = Color.Black,
            AutoSize = true,
            Margin = new Padding(ContentMarginLeft, ContentMarginTop, ContentMarginRight, ContentMarginBottom),
            MaximumSize = new Size(MaxContentWidth, 0)
        };

        if (_useLayeredWindow)
            SetupTransparentMode();
        else
            SetupOpaqueMode();
    }

    #region Setup

    private void SetupTransparentMode()
    {
        AutoSize = false;
        Padding = Padding.Empty;

        _dragBarMenu = new ContextMenuStrip();
        _dragBarMenu.Items.Add("Close", null, (_, _) => BeginInvoke(() => Close()));

        MouseDown += OnTransparentMouseDown;

        var (fw, fh) = MeasureFormSize(_contentLabel.Text);
        Size = new Size(fw, fh);

        Shown += (_, _) => PositionFromCenter();
        SizeChanged += (_, _) => { if (_initialPositionSet) RepositionToCenter(); };
        Load += (_, _) => RenderTransparent();
    }

    private void SetupOpaqueMode()
    {
        Padding = new Padding(1);

        // Use TableLayoutPanel for proper auto-sizing
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.Black,
            ColumnCount = 1,
            RowCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(0),
            Margin = new Padding(0)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, DragBarHeight));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(layout);

        // Drag bar
        var dragBar = new Label
        {
            Text = "⋯",
            Font = new Font("Segoe UI", 8),
            ForeColor = Color.FromArgb(102, 102, 102),
            BackColor = Color.Black,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Cursor = Cursors.SizeAll,
            Margin = new Padding(0)
        };
        dragBar.MouseDown += OnDragStart;
        dragBar.MouseMove += OnDragMove;
        dragBar.MouseUp += OnDragEnd;
        layout.Controls.Add(dragBar, 0, 0);

        // Content label
        layout.Controls.Add(_contentLabel, 0, 1);

        // Make form auto-size to content
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;

        // Context menu for drag bar only (Close option)
        var menu = new ContextMenuStrip();
        menu.Items.Add("Close", null, (_, _) => BeginInvoke(() => Close()));
        dragBar.ContextMenuStrip = menu;

        // Left click to dismiss, right click to copy debug
        _contentLabel.MouseDown += OnContentLabelMouseDown;

        // Position from center after form sizes itself
        Shown += (_, _) => PositionFromCenter();
        SizeChanged += (_, _) => { if (_initialPositionSet) RepositionToCenter(); };
    }

    #endregion

    #region Positioning

    private void PositionFromCenter()
    {
        Location = new Point(
            _config.ImpressionX - Width / 2,
            _config.ImpressionY - Height / 2
        );
        _initialPositionSet = true;
    }

    private void RepositionToCenter()
    {
        if (_dragging) return;
        Location = new Point(
            _config.ImpressionX - Width / 2,
            _config.ImpressionY - Height / 2
        );
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Update the displayed impression text.
    /// </summary>
    public void SetImpression(string? text)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetImpression(text));
            return;
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            _contentLabel.Text = "(Waiting for impression...)";
            _contentLabel.ForeColor = Color.FromArgb(128, 128, 128);
        }
        else
        {
            _contentLabel.Text = text;
            _contentLabel.ForeColor = Color.FromArgb(200, 200, 200);
        }

        if (_useLayeredWindow && IsHandleCreated && !IsDisposed)
        {
            var (fw, fh) = MeasureFormSize(_contentLabel.Text);
            if (fw != Width || fh != Height)
                Size = new Size(fw, fh);
            RenderTransparent();
        }
    }

    /// <summary>
    /// Extract impression from raw report text.
    /// </summary>
    public static string? ExtractImpression(string? rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText))
            return null;

        if (!rawText.Contains("IMPRESSION", StringComparison.OrdinalIgnoreCase))
            return null;

        var sectionHeaders = @"TECHNIQUE|FINDINGS|CLINICAL HISTORY|COMPARISON|EXAM|PROCEDURE|INDICATION|CONCLUSION|RECOMMENDATION|SIGNATURE|ELECTRONICALLY SIGNED";

        var match = Regex.Match(rawText,
            $@"IMPRESSION[:\s]*\n?(.+?)(?=\n\s*({sectionHeaders})\s*[:\n]|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!match.Success || match.Groups.Count < 2)
            return null;

        var content = match.Groups[1].Value.Trim();
        if (string.IsNullOrWhiteSpace(content))
            return null;

        content = Regex.Replace(content, @"[\r\n\t ]+", " ").Trim();
        content = CleanText(content);
        content = FormatNumberedItems(content);

        return content;
    }

    public void EnsureOnTop()
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            BeginInvoke(EnsureOnTop);
            return;
        }
        if (!IsDisposed && IsHandleCreated)
            NativeWindows.ForceTopMost(this.Handle);
    }

    #endregion

    #region Transparent Mode - Rendering

    private void RenderTransparent()
    {
        if (!IsHandleCreated || IsDisposed) return;
        int w = Width, h = Height;
        if (w <= 0 || h <= 0) return;

        string text = _contentLabel.Text;
        Color textColor = _contentLabel.ForeColor;

        // Layer 1: semi-transparent background
        using var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.FromArgb(0, 0, 0, 0));
            int bw = BorderWidth;
            // Inner area - semi-transparent (draw first, border drawn on top)
            using var innerBrush = new SolidBrush(Color.FromArgb(_backgroundAlpha, 0, 0, 0));
            g.FillRectangle(innerBrush, bw, bw, w - bw * 2, h - bw * 2);
            // Border strips - fully opaque so it's always visible
            using var borderBrush = new SolidBrush(Color.FromArgb(255, 60, 60, 60));
            g.FillRectangle(borderBrush, 0, 0, w, bw);           // top
            g.FillRectangle(borderBrush, 0, h - bw, w, bw);      // bottom
            g.FillRectangle(borderBrush, 0, bw, bw, h - bw);     // left
            g.FillRectangle(borderBrush, w - bw, bw, bw, h - bw); // right
        }

        // Layer 2: ClearType text on opaque black
        using var textLayer = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(textLayer))
        {
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.Clear(Color.FromArgb(255, 0, 0, 0));

            // Drag bar "⋯"
            using var dragFont = new Font("Segoe UI", 8);
            using var dragBrush = new SolidBrush(Color.FromArgb(102, 102, 102));
            using var dragSf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString("⋯", dragFont, dragBrush,
                new RectangleF(BorderWidth, BorderWidth, w - BorderWidth * 2, DragBarHeight), dragSf);

            // Content text
            float textX = BorderWidth + ContentMarginLeft;
            float textY = BorderWidth + DragBarHeight + ContentMarginTop;
            using var sf = new StringFormat(StringFormat.GenericTypographic) { Trimming = StringTrimming.Word };
            using var brush = new SolidBrush(textColor);
            g.DrawString(text, _contentFont, brush,
                new RectangleF(textX, textY, MaxContentWidth, h - textY), sf);
        }

        LayeredWindowHelper.MergeTextLayer(bmp, textLayer);
        LayeredWindowHelper.PremultiplyBitmapAlpha(bmp);
        LayeredWindowHelper.SetBitmap(this, bmp);
    }

    private (int width, int height) MeasureFormSize(string text)
    {
        using var bmp = new Bitmap(1, 1);
        using var g = Graphics.FromImage(bmp);
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        using var sf = new StringFormat(StringFormat.GenericTypographic) { Trimming = StringTrimming.Word };

        var measured = g.MeasureString(text, _contentFont, MaxContentWidth, sf);

        int contentWidth = Math.Max((int)Math.Ceiling(measured.Width), 50);
        int contentHeight = Math.Max((int)Math.Ceiling(measured.Height), _contentFont.Height);

        int formWidth = BorderWidth * 2 + ContentMarginLeft + contentWidth + ContentMarginRight;
        int formHeight = BorderWidth * 2 + DragBarHeight + ContentMarginTop + contentHeight + ContentMarginBottom;

        return (Math.Max(formWidth, 100), Math.Max(formHeight, 50));
    }

    #endregion

    #region Transparent Mode - Mouse Handling

    private void OnTransparentMouseDown(object? sender, MouseEventArgs e)
    {
        bool inDragBar = e.Y < BorderWidth + DragBarHeight;

        if (e.Button == MouseButtons.Right)
        {
            if (inDragBar)
                _dragBarMenu?.Show(this, e.Location);
            else
                CopyDebugInfoToClipboard();
            return;
        }

        if (e.Button == MouseButtons.Left)
        {
            if (inDragBar)
            {
                _formPosOnMouseDown = Location;
                LayeredWindowHelper.ReleaseCapture();
                LayeredWindowHelper.SendMessage(Handle, LayeredWindowHelper.WM_NCLBUTTONDOWN, LayeredWindowHelper.HT_CAPTION, 0);
                // SendMessage returns when drag ends
                if (Location != _formPosOnMouseDown)
                {
                    _config.ImpressionX = Location.X + Width / 2;
                    _config.ImpressionY = Location.Y + Height / 2;
                    _config.Save();
                }
            }
            else
            {
                // Content click - dismiss
                Close();
            }
        }
    }

    #endregion

    #region Opaque Mode - Mouse Handling

    private void OnContentLabelMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            Close();
        }
        else if (e.Button == MouseButtons.Right)
        {
            CopyDebugInfoToClipboard();
        }
    }

    #endregion

    #region Drag Logic (Opaque Mode)

    private void OnDragStart(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            _dragging = true;
            _dragStart = e.Location;
        }
    }

    private void OnDragMove(object? sender, MouseEventArgs e)
    {
        if (_dragging)
        {
            Location = new Point(
                Location.X + e.X - _dragStart.X,
                Location.Y + e.Y - _dragStart.Y
            );
        }
    }

    private void OnDragEnd(object? sender, MouseEventArgs e)
    {
        _dragging = false;
        _config.ImpressionX = Location.X + Width / 2;
        _config.ImpressionY = Location.Y + Height / 2;
        _config.Save();
    }

    #endregion

    #region Text Utilities

    private static string FormatNumberedItems(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        if (!Regex.IsMatch(text, @"^\s*1\."))
            return text;

        var result = new System.Text.StringBuilder();
        int expectedNumber = 1;
        int i = 0;

        while (i < text.Length)
        {
            string nextPattern = $"{expectedNumber}.";

            if (i + nextPattern.Length <= text.Length)
            {
                string substring = text.Substring(i, nextPattern.Length);

                if (substring == nextPattern)
                {
                    bool isValidNumberedItem = (i + nextPattern.Length >= text.Length) ||
                        char.IsWhiteSpace(text[i + nextPattern.Length]) ||
                        char.IsLetter(text[i + nextPattern.Length]);

                    if (isValidNumberedItem)
                    {
                        if (expectedNumber > 1)
                        {
                            result.Append('\n');
                        }
                        result.Append(nextPattern);
                        i += nextPattern.Length;
                        expectedNumber++;
                        continue;
                    }
                }
            }

            result.Append(text[i]);
            i++;
        }

        return result.ToString();
    }

    private static string CleanText(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var sb = new System.Text.StringBuilder();
        foreach (char c in text)
        {
            if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || c == ' ' || c == '-' || c == '/' || c == '\'')
            {
                sb.Append(c);
            }
            else if (char.IsWhiteSpace(c) || char.IsControl(c))
            {
                sb.Append(' ');
            }
        }

        return Regex.Replace(sb.ToString(), @" +", " ").Trim();
    }

    private void CopyDebugInfoToClipboard()
    {
        var debugInfo = $"=== Impression Debug ===\r\n" +
                        $"Content: {_contentLabel.Text}";

        try
        {
            Clipboard.SetText(debugInfo);
            Services.Logger.Trace("Impression debug info copied to clipboard");
            ShowCopiedToast();
        }
        catch (Exception ex)
        {
            Services.Logger.Trace($"Failed to copy debug info: {ex.Message}");
        }
    }

    private void ShowCopiedToast()
    {
        var toast = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            ShowInTaskbar = false,
            TopMost = true,
            BackColor = Color.FromArgb(51, 51, 51),
            Size = new Size(120, 30),
            StartPosition = FormStartPosition.Manual,
            Location = new Point(Location.X + Width / 2 - 60, Location.Y - 35)
        };

        var label = new Label
        {
            Text = "Debug copied!",
            ForeColor = Color.White,
            BackColor = Color.FromArgb(51, 51, 51),
            Font = new Font("Segoe UI", 9),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter
        };
        toast.Controls.Add(label);
        toast.Show();

        var timer = new System.Windows.Forms.Timer { Interval = 1500 };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            if (!toast.IsDisposed) toast.Close();
        };
        timer.Start();
    }

    #endregion

    #region Disposal

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            // Detach ALL context menus before disposal to prevent ObjectDisposedException
            // in WinForms ModalMenuFilter.ProcessActivationChange on next focus change
            ContextMenuStrip = null;
            foreach (Control c in Controls)
                c.ContextMenuStrip = null;

            _contentFont?.Dispose();
            _dragBarMenu?.Dispose();
        }
        base.Dispose(disposing);
    }

    #endregion
}
