// [CustomSTT] Full transcript comparison popup — line-by-line grouped view with color-coded prefixes
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Scrollable, resizable popup showing transcripts grouped line-by-line:
/// each line from ensemble shown with corresponding line from each provider underneath.
/// Monospace font, color-coded prefixes. Remembers position and size.
/// </summary>
public class SttTranscriptComparisonForm : Form
{
    private readonly Configuration _config;
    private Point _dragOffset;
    private bool _dragging;

    private static readonly Color BgColor = Color.FromArgb(18, 18, 24);
    private static readonly Color DgColor = Color.FromArgb(80, 160, 255);     // Deepgram blue
    private static readonly Color S1Color = Color.FromArgb(80, 200, 180);     // Soniox teal
    private static readonly Color S2Color = Color.FromArgb(180, 130, 255);    // Speechmatics purple
    private static readonly Color EnsColor = Color.FromArgb(100, 220, 100);   // Ensemble green
    private static readonly Color TextColor = Color.FromArgb(200, 200, 210);
    private static readonly Color DimColor = Color.FromArgb(100, 100, 110);

    private readonly RichTextBox _rtb;
    private readonly Panel _titlePanel;

    private const int DefaultW = 700;
    private const int DefaultH = 450;
    private const int MinW = 300;
    private const int MinH = 150;

    // Resize grip
    private bool _resizing;
    private Point _resizeStart;
    private Size _resizeStartSize;

    public SttTranscriptComparisonForm(Configuration config)
    {
        _config = config;
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;
        BackColor = BgColor;
        Opacity = 0.95;
        AutoScaleMode = AutoScaleMode.None;
        MinimumSize = new Size(MinW, MinH);

        // Restore size
        int w = config.SttTranscriptComparisonW > 0 ? config.SttTranscriptComparisonW : DefaultW;
        int h = config.SttTranscriptComparisonH > 0 ? config.SttTranscriptComparisonH : DefaultH;
        Size = new Size(w, h);

        // Restore position
        if (config.SttTranscriptComparisonX != int.MinValue && config.SttTranscriptComparisonY != int.MinValue)
            Location = new Point(config.SttTranscriptComparisonX, config.SttTranscriptComparisonY);
        else
        {
            var screen = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
            Location = new Point(screen.Right - Width - 20, screen.Top + 60);
        }

        // Title bar
        _titlePanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 24,
            BackColor = Color.FromArgb(30, 30, 38)
        };
        var titleLabel = new Label
        {
            Text = "STT TRANSCRIPT COMPARISON",
            Location = new Point(8, 4),
            AutoSize = true,
            Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(140, 160, 200)
        };
        _titlePanel.Controls.Add(titleLabel);

        var closeBtn = new Label
        {
            Text = "\u00d7",
            Location = new Point(Width - 22, 2),
            AutoSize = true,
            Font = new Font("Segoe UI", 10f),
            ForeColor = Color.FromArgb(80, 80, 90),
            Cursor = Cursors.Hand
        };
        closeBtn.Click += (_, _) => Hide();
        closeBtn.MouseEnter += (_, _) => closeBtn.ForeColor = Color.IndianRed;
        closeBtn.MouseLeave += (_, _) => closeBtn.ForeColor = Color.FromArgb(80, 80, 90);
        _titlePanel.Controls.Add(closeBtn);

        // Dragging on title
        _titlePanel.MouseDown += OnDragMouseDown;
        _titlePanel.MouseMove += OnDragMouseMove;
        _titlePanel.MouseUp += OnDragMouseUp;
        titleLabel.MouseDown += OnDragMouseDown;
        titleLabel.MouseMove += OnDragMouseMove;
        titleLabel.MouseUp += OnDragMouseUp;

        // RichTextBox for content
        _rtb = new RichTextBox
        {
            Dock = DockStyle.Fill,
            BackColor = BgColor,
            ForeColor = TextColor,
            Font = new Font("Consolas", 9f),
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            WordWrap = true
        };
        _rtb.Click += (_, _) => Hide();

        // Add RTB first, then title — WinForms docks last-added first,
        // so title (Top) must be added after RTB (Fill) to sit above it.
        Controls.Add(_rtb);
        Controls.Add(_titlePanel);

        // Keep close button anchored right on resize
        Resize += (_, _) => closeBtn.Location = new Point(Width - 22, 2);
    }

    public void ShowTranscripts(string dgText, string s1Text, string s2Text, string ensText,
        string s1Name = "Soniox", string s2Name = "Speechmatics")
    {
        _rtb.Clear();

        var s1Tag = TagFor(s1Name);
        var s2Tag = TagFor(s2Name);
        const string ensTag = "ENS";
        const string dgTag = "DG";

        // Split each transcript into lines
        var ensLines = SplitLines(NormalizeText(ensText));
        var dgLines = SplitLines(NormalizeText(dgText));
        var s1Lines = SplitLines(NormalizeText(s1Text));
        var s2Lines = SplitLines(NormalizeText(s2Text));

        int maxLines = Math.Max(ensLines.Count,
            Math.Max(dgLines.Count, Math.Max(s1Lines.Count, s2Lines.Count)));

        var providers = new[]
        {
            (ensTag, EnsColor, ensLines),
            (dgTag, DgColor, dgLines),
            (s1Tag, S1Color, s1Lines),
            (s2Tag, S2Color, s2Lines)
        };

        int maxTag = 0;
        foreach (var (tag, _, _) in providers)
            if (tag.Length > maxTag) maxTag = tag.Length;

        for (int ln = 0; ln < maxLines; ln++)
        {
            if (ln > 0) _rtb.AppendText("\n");

            for (int p = 0; p < providers.Length; p++)
            {
                var (tag, color, lines) = providers[p];
                var paddedTag = tag.PadRight(maxTag);
                var line = ln < lines.Count ? lines[ln] : "";
                var hasText = !string.IsNullOrEmpty(line);

                AppendColored(paddedTag, color);
                AppendColored(" : ", DimColor);
                AppendColored(hasText ? line : "-", hasText ? TextColor : DimColor);

                if (p < providers.Length - 1 || ln < maxLines - 1)
                    _rtb.AppendText("\n");
            }
        }

        if (maxLines == 0)
            AppendColored("(no transcript data)", DimColor);

        _rtb.SelectionStart = 0;
        _rtb.ScrollToCaret();

        if (!Visible) Show();
        BringToFront();
    }

    private static List<string> SplitLines(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return new List<string>();
        var result = new List<string>();
        foreach (var line in text.Split('\n'))
        {
            var trimmed = line.Trim();
            if (trimmed.Length > 0) result.Add(trimmed);
        }
        return result;
    }

    private static string TagFor(string providerName)
    {
        if (providerName.StartsWith("Soniox", StringComparison.OrdinalIgnoreCase)) return "SX";
        if (providerName.StartsWith("Speechmatics", StringComparison.OrdinalIgnoreCase)) return "SM";
        if (providerName.StartsWith("Deepgram", StringComparison.OrdinalIgnoreCase)) return "DG";
        return providerName.Length >= 2
            ? providerName[..2].ToUpperInvariant()
            : providerName.ToUpperInvariant();
    }

    /// <summary>Collapse multiple spaces/tabs into single space, trim.</summary>
    private static string NormalizeText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        return Regex.Replace(text.Trim(), @"[ \t]+", " ");
    }

    private void AppendColored(string text, Color color)
    {
        int start = _rtb.TextLength;
        _rtb.AppendText(text);
        _rtb.Select(start, text.Length);
        _rtb.SelectionColor = color;
        _rtb.SelectionLength = 0;
    }

    private void SaveBounds()
    {
        _config.SttTranscriptComparisonX = Location.X;
        _config.SttTranscriptComparisonY = Location.Y;
        _config.SttTranscriptComparisonW = Width;
        _config.SttTranscriptComparisonH = Height;
        _config.Save();
    }

    // ── Dragging ──
    private void OnDragMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            _dragging = true;
            _dragOffset = e.Location;
            if (sender is Control c && c != this)
                _dragOffset = new Point(e.X + c.Left, e.Y + c.Top);
        }
    }

    private void OnDragMouseMove(object? sender, MouseEventArgs e)
    {
        if (_dragging)
        {
            var screenPos = sender is Control c && c != this
                ? c.PointToScreen(e.Location)
                : PointToScreen(e.Location);
            Location = new Point(screenPos.X - _dragOffset.X, screenPos.Y - _dragOffset.Y);
        }
    }

    private void OnDragMouseUp(object? sender, MouseEventArgs e)
    {
        if (_dragging)
        {
            _dragging = false;
            SaveBounds();
        }
    }

    // ── Resize via edges/corners ──
    private const int GripSize = 6;

    protected override void OnMouseDown(MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left && GetResizeCursor(e.Location) != Cursors.Default)
        {
            _resizing = true;
            _resizeStart = PointToScreen(e.Location);
            _resizeStartSize = Size;
            Capture = true;
        }
        base.OnMouseDown(e);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        if (_resizing)
        {
            var screen = PointToScreen(e.Location);
            int dx = screen.X - _resizeStart.X;
            int dy = screen.Y - _resizeStart.Y;
            int newW = Math.Max(MinW, _resizeStartSize.Width + dx);
            int newH = Math.Max(MinH, _resizeStartSize.Height + dy);
            Size = new Size(newW, newH);
        }
        else
        {
            Cursor = GetResizeCursor(e.Location);
        }
        base.OnMouseMove(e);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        if (_resizing)
        {
            _resizing = false;
            Capture = false;
            SaveBounds();
        }
        base.OnMouseUp(e);
    }

    private Cursor GetResizeCursor(Point p)
    {
        bool right = p.X >= Width - GripSize;
        bool bottom = p.Y >= Height - GripSize;
        if (right && bottom) return Cursors.SizeNWSE;
        if (right) return Cursors.SizeWE;
        if (bottom) return Cursors.SizeNS;
        return Cursors.Default;
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= 0x80; // WS_EX_TOOLWINDOW
            return cp;
        }
    }
}
