using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Display mode for the report popup click cycle.
/// </summary>
public enum ReportDisplayMode
{
    Changes,
    Rainbow,
    OrphanFindings
}

/// <summary>
/// Report popup window.
/// Supports two rendering modes:
/// - Transparent: Layered window with per-pixel alpha (image shows through background)
/// - Opaque: Traditional RichTextBox rendering (solid dark background)
/// Supports diff highlighting and Rainbow Mode correlation in both modes.
/// </summary>
public class ReportPopupForm : Form
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

    // Mode
    private readonly bool _useLayeredWindow;

    // Config and state
    private readonly Configuration _config;
    private string? _baselineReport;
    private string _currentReportText;
    private string _formattedText;
    private ReportDisplayMode _displayMode;
    private readonly bool _changesEnabled;
    private readonly bool _correlationEnabled;
    private bool _baselineIsSectionOnly; // True when baseline is from template DB (diff only FINDINGS+IMPRESSION)
    private CorrelationResult? _correlationResult;
    private CorrelationResult? _previousCorrelation; // Retained to prevent regression on noisy scrapes
    private bool _showingStaleContent; // True when showing cached report while live report is being updated
    private bool _radAiImpressionActive; // True when RadAI impression was successfully inserted
    private System.Windows.Forms.Timer? _stalePulseTimer;
    private float _stalePulsePhase; // 0..2π for smooth sine wave
    private int? _accessionSeed; // Hash of accession for consistent rainbow colors per study

    // Transparent mode rendering
    private int _scrollOffset;
    private int _totalContentHeight;
    private readonly int _backgroundAlpha;
    private readonly Font _normalFont;
    private readonly Font _headerFont;
    private readonly Font _modeLabelFont;
    private readonly Font _staleIndicatorFont;
    private readonly Font _staleLabelFont; // Larger font for opaque mode stale label

    private readonly record struct HighlightEntry(string Text, Color BackColor);
    private List<HighlightEntry> _highlights = new();

    // Opaque mode controls
    private RichTextBox? _richTextBox;
    private Label? _modeLabel;
    private bool _isResizing;

    // Deletable impression points (opaque mode only)
    private readonly bool _deletableEnabled;
    private List<Label> _trashIcons = new();
    private List<(int charIndex, string itemText)>? _impressionItems;
    private List<Label> _impressionFixerLabels = new();
    private List<ImpressionFixerEntry> _impressionFixerEntries = new();
    private System.Windows.Forms.Timer? _debounceTimer;
    private bool _deletePending;

    /// <summary>
    /// Fired after debounce when user deletes impression points.
    /// The string is the new impression text (lines without numbers, joined by \r\n).
    /// </summary>
    public event Action<string>? ImpressionDeleteRequested;

    // Drag
    private Point _formPosOnMouseDown;
    private Point _dragStart;
    private bool _dragging;

    // Horizontal resize
    private const int ResizeGripWidth = 6;
    private const int MinFormWidth = 400;
    private bool _resizingLeft;
    private bool _resizingRight;
    private int _resizeStartX;
    private int _resizeStartWidth;
    private int _resizeStartLeft;

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

    public ReportPopupForm(Configuration config, string reportText, string? baselineReport = null,
        bool changesEnabled = false, bool correlationEnabled = false, bool baselineIsSectionOnly = false,
        string? accession = null)
    {
        _config = config;
        _baselineReport = baselineReport;
        _currentReportText = reportText;
        _changesEnabled = changesEnabled;
        _correlationEnabled = correlationEnabled;
        _baselineIsSectionOnly = baselineIsSectionOnly;
        _useLayeredWindow = config.ReportPopupTransparent;
        _deletableEnabled = !_useLayeredWindow && config.ImpressionDeletablePoints;
        _accessionSeed = accession?.GetHashCode();

        // Compute background alpha from transparency setting
        _backgroundAlpha = (int)Math.Round(config.ReportPopupTransparency / 100.0 * 255);
        _backgroundAlpha = Math.Clamp(_backgroundAlpha, 30, 255);

        // Display mode
        if (_changesEnabled)
            _displayMode = ReportDisplayMode.Changes;
        else if (_correlationEnabled)
            _displayMode = ReportDisplayMode.Rainbow;
        else
            _displayMode = ReportDisplayMode.Changes;

        // Fonts
        _normalFont = new Font(config.ReportPopupFontFamily, config.ReportPopupFontSize);
        _headerFont = new Font(config.ReportPopupFontFamily, config.ReportPopupFontSize + 2, FontStyle.Bold);
        _modeLabelFont = new Font("Segoe UI", 8f);
        _staleIndicatorFont = new Font("Segoe UI", 13f, FontStyle.Bold);
        _staleLabelFont = new Font("Segoe UI", 13f, FontStyle.Bold);

        // Form properties
        FormBorderStyle = FormBorderStyle.None;
        Text = "";
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;

        int width = Math.Max(config.ReportPopupWidth, 600);

        // Format text
        _formattedText = SanitizeText(FormatReportText(reportText)).Replace("\r\n", "\n");

        if (_useLayeredWindow)
        {
            SetupTransparentMode(width);
        }
        else
        {
            SetupOpaqueMode(width);
        }

        // Common setup
        this.Location = ScreenHelper.EnsureOnScreen(config.ReportPopupX, config.ReportPopupY, Size.Width, Size.Height);

        LocationChanged += (_, _) =>
        {
            _config.ReportPopupX = Location.X;
            _config.ReportPopupY = Location.Y;
        };

        FormClosed += (_, _) => _config.Save();
    }

    #region Setup Methods

    private void SetupTransparentMode(int width)
    {
        AutoSize = false;
        AutoScroll = false;

        ComputeHighlights();

        int contentHeight = MeasureContentHeight(width);
        _totalContentHeight = contentHeight;

        var screen = Screen.FromPoint(new Point(_config.ReportPopupX, _config.ReportPopupY));
        int maxHeight = screen.WorkingArea.Bottom - _config.ReportPopupY - 10;
        int formHeight = Math.Clamp(contentHeight, 100, maxHeight);

        this.Size = new Size(width, formHeight);

        this.MouseDown += OnLayeredMouseDown;
        this.MouseMove += OnLayeredMouseMove;
        this.MouseUp += OnLayeredMouseUp;
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
            ForeColor = Color.White,
            Font = _normalFont,
            Cursor = Cursors.Hand,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.None,
            DetectUrls = false,
            ShortcutsEnabled = false
        };

        _richTextBox.ContentsResized += RichTextBox_ContentsResized;
        _richTextBox.ForeColor = Color.Gainsboro;
        if (_deletableEnabled)
            _richTextBox.RightMargin = _richTextBox.Width - 26;
        _richTextBox.Text = _formattedText;

        Controls.Add(_richTextBox);

        _modeLabel = new Label
        {
            AutoSize = true,
            ForeColor = Color.FromArgb(140, 140, 140),
            BackColor = Color.FromArgb(30, 30, 30),
            Font = _modeLabelFont,
            Text = "",
            Cursor = Cursors.Hand
        };
        Controls.Add(_modeLabel);
        _modeLabel.BringToFront();
        UpdateModeLabel();

        this.ClientSize = new Size(width, 200);

        SetupOpaqueInteractions(this);
        SetupOpaqueInteractions(_richTextBox);

        if (_deletableEnabled)
        {
            _debounceTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            _debounceTimer.Tick += OnDebounceTimerTick;
        }

        this.Load += (s, e) =>
        {
            this.Activate();
            ActiveControl = null;
            ApplyCurrentModeFormatting();
            PerformResize();
            if (_deletableEnabled) PositionTrashIcons();
        };

        this.Resize += (s, e) =>
        {
            PositionModeLabel();
            if (_deletableEnabled) PositionTrashIcons();
        };

        _richTextBox.VScroll += (s, e) =>
        {
            if (_deletableEnabled) PositionTrashIcons();
        };
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Handle click cycle: Changes -> Rainbow -> Close.
    /// Returns true if handled (don't close), false if should close.
    /// </summary>
    public bool HandleClickCycle()
    {
        Logger.Trace($"HandleClickCycle: mode={_displayMode}, changesEnabled={_changesEnabled}, correlationEnabled={_correlationEnabled}, orphanEnabled={_config.OrphanFindingsEnabled}");

        if (_displayMode == ReportDisplayMode.Changes && _correlationEnabled)
        {
            _displayMode = ReportDisplayMode.Rainbow;
            Logger.Trace("Switching to Rainbow mode");

            if (_useLayeredWindow)
            {
                ComputeHighlights();
                RenderAndUpdate();
            }
            else
            {
                UpdateModeLabel();
                ResetAndReapplyFormatting();
            }
            return true;
        }

        if (_displayMode == ReportDisplayMode.Rainbow && _config.OrphanFindingsEnabled)
        {
            _displayMode = ReportDisplayMode.OrphanFindings;
            Logger.Trace("Switching to Unmatched mode");

            if (_useLayeredWindow)
            {
                ComputeHighlights();
                RenderAndUpdate();
            }
            else
            {
                UpdateModeLabel();
                ResetAndReapplyFormatting();
            }
            return true;
        }

        Logger.Trace("Click cycle: closing");
        return false;
    }

    /// <summary>
    /// Update the report text and re-apply highlighting.
    /// Called when Process Report is pressed while popup is open.
    /// </summary>
    public void UpdateReport(string newReportText, string? baseline = null, bool baselineIsSectionOnly = false)
    {
        // Don't overwrite while user is actively deleting impression points
        if (_deletePending) return;

        _showingStaleContent = false; // Clear stale flag when we get new content
        if (baseline != null)
        {
            _baselineReport = baseline;
            _baselineIsSectionOnly = baselineIsSectionOnly;
        }

        // Don't reset display mode - preserve user's current selection (Changes vs Rainbow)
        // Reset correlation for recalculation; only keep previous for anti-regression if content is similar
        // (anti-regression is for scrape noise, not genuine edits)
        bool contentChanged = newReportText != _currentReportText;
        _currentReportText = newReportText;

        if (contentChanged)
        {
            _previousCorrelation = null;
            _correlationResult = null;
        }
        else
        {
            // Same content (e.g., baseline update) - allow anti-regression
            _previousCorrelation = _correlationResult;
            _correlationResult = null;
        }
        _formattedText = SanitizeText(FormatReportText(newReportText)).Replace("\r\n", "\n");
        if (_radAiImpressionActive)
            _formattedText = _formattedText.Replace("IMPRESSION:", "RadAI IMPRESSION:");

        if (_useLayeredWindow)
        {
            _scrollOffset = 0;
            ComputeHighlights();

            int contentHeight = MeasureContentHeight(Width);
            _totalContentHeight = contentHeight;

            var screen = Screen.FromControl(this);
            int maxHeight = screen.WorkingArea.Bottom - Top - 10;
            int formHeight = Math.Clamp(contentHeight, 100, maxHeight);

            this.Size = new Size(Width, formHeight);
            RenderAndUpdate();
        }
        else
        {
            _richTextBox!.Text = _formattedText;
            _richTextBox.SelectAll();
            _richTextBox.SelectionFont = _richTextBox.Font;
            _richTextBox.SelectionColor = Color.Gainsboro;
            _richTextBox.SelectionBackColor = _richTextBox.BackColor;
            _richTextBox.Select(0, 0);

            UpdateModeLabel();
            ApplyCurrentModeFormatting();
            PerformResize();

            // Deferred reposition: RTB layout from ContentsResized is async,
            // so trash icons may need repositioning after the form finishes resizing
            if (_deletableEnabled)
                BeginInvoke(new Action(PositionTrashIcons));
        }

        Logger.Trace($"ReportPopup updated: {newReportText.Length} chars, baseline={_baselineReport?.Length ?? 0} chars, mode={_displayMode}");
    }

    /// <summary>
    /// Mark the impression section as RadAI-generated. Changes "IMPRESSION:" header to "RadAI IMPRESSION:".
    /// </summary>
    public void SetRadAiImpressionActive(bool active)
    {
        if (_radAiImpressionActive == active) return;
        _radAiImpressionActive = active;

        // Refresh the formatted text to apply/remove the label
        _formattedText = SanitizeText(FormatReportText(_currentReportText)).Replace("\r\n", "\n");
        if (_radAiImpressionActive)
            _formattedText = _formattedText.Replace("IMPRESSION:", "RadAI IMPRESSION:");

        if (_useLayeredWindow)
        {
            RenderAndUpdate();
        }
        else if (_richTextBox != null)
        {
            _richTextBox.Text = _formattedText;
            _richTextBox.SelectAll();
            _richTextBox.SelectionFont = _richTextBox.Font;
            _richTextBox.SelectionColor = Color.Gainsboro;
            _richTextBox.SelectionBackColor = _richTextBox.BackColor;
            _richTextBox.Select(0, 0);
            ApplyCurrentModeFormatting();
        }
    }

    /// <summary>
    /// Mark the popup as showing stale/cached content while the live report is being updated.
    /// </summary>
    public void SetStaleState(bool isStale)
    {
        if (_showingStaleContent != isStale)
        {
            _showingStaleContent = isStale;

            if (isStale)
            {
                // Start pulsing border animation
                _stalePulsePhase = 0;
                if (_stalePulseTimer == null)
                {
                    _stalePulseTimer = new System.Windows.Forms.Timer { Interval = 40 }; // ~25fps
                    _stalePulseTimer.Tick += (_, _) =>
                    {
                        _stalePulsePhase += 0.12f; // ~2.4s full cycle
                        if (_stalePulsePhase > (float)(2 * Math.PI)) _stalePulsePhase -= (float)(2 * Math.PI);

                        if (_useLayeredWindow)
                            RenderAndUpdate();
                        else
                            Invalidate(); // triggers OnPaint for border
                    };
                }
                _stalePulseTimer.Start();
            }
            else
            {
                // Stop pulsing
                _stalePulseTimer?.Stop();
            }

            if (_useLayeredWindow)
            {
                RenderAndUpdate();
            }
            else
            {
                UpdateModeLabel();
                Invalidate();
            }
        }
    }

    /// <summary>
    /// Set the impression fixer entries to show on the IMPRESSION header line.
    /// Called by ActionController with entries filtered by study description.
    /// </summary>
    public void SetImpressionFixers(List<ImpressionFixerEntry> entries)
    {
        if (_useLayeredWindow) return; // Only opaque mode supports interactive labels
        if (!_deletableEnabled) return;

        _impressionFixerEntries = entries;

        // Remove old labels
        foreach (var lbl in _impressionFixerLabels)
        {
            Controls.Remove(lbl);
            lbl.Dispose();
        }
        _impressionFixerLabels.Clear();

        var bg = Color.FromArgb(30, 30, 30);
        var font = new Font("Segoe UI", 9.5f);

        foreach (var entry in entries)
        {
            var prefix = entry.ReplaceMode ? "=" : "+";
            var normalColor = entry.ReplaceMode
                ? Color.FromArgb(255, 180, 0)   // amber for replace
                : Color.FromArgb(0, 200, 200);  // cyan for insert
            var hoverColor = entry.ReplaceMode
                ? Color.FromArgb(255, 220, 80)
                : Color.FromArgb(100, 255, 255);

            var lbl = new Label
            {
                Text = $"{prefix}{entry.Blurb}",
                AutoSize = true,
                Font = font,
                ForeColor = normalColor,
                BackColor = bg,
                Cursor = Cursors.Hand,
                Tag = entry,
                Visible = false
            };

            var normalCopy = normalColor;
            var hoverCopy = hoverColor;
            lbl.MouseEnter += (_, _) => lbl.ForeColor = hoverCopy;
            lbl.MouseLeave += (_, _) => lbl.ForeColor = normalCopy;
            lbl.Click += OnFixerLabelClick;

            Controls.Add(lbl);
            lbl.BringToFront();
            _impressionFixerLabels.Add(lbl);
        }

        PositionTrashIcons();
    }

    #endregion

    #region Transparent Mode - Rendering

    /// <summary>
    /// Two-layer rendering for ClearType text on transparent background:
    /// 1. Background layer: semi-transparent dark fill + highlight rectangles + mode label
    /// 2. Text layer: ClearType text rendered onto opaque dark surface
    /// 3. Merge: text pixels stamped onto background layer with full opacity
    /// Result: crisp ClearType text over a see-through background.
    /// </summary>
    private void RenderAndUpdate()
    {
        if (!IsHandleCreated || IsDisposed) return;

        int w = Width, h = Height;

        // Layer 1: semi-transparent background + mode label
        using var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            g.Clear(Color.FromArgb(_backgroundAlpha, 20, 20, 20));
            DrawModeLabel(g);
            DrawStalePulseBorder(g);
        }

        // Layer 2: highlights + ClearType text on opaque dark background
        using var textLayer = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(textLayer))
        {
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.Clear(Color.FromArgb(255, 20, 20, 20));
            DrawContentParts(g, highlights: true, text: true);
        }

        // Merge text pixels onto background bitmap, then premultiply
        MergeTextLayer(bmp, textLayer);
        PremultiplyBitmapAlpha(bmp);
        SetBitmap(bmp);
    }

    /// <summary>
    /// Draw content with optional highlights and/or text.
    /// Both passes share the same layout so positions align exactly.
    /// </summary>
    private void DrawContentParts(Graphics g, bool highlights, bool text)
    {
        int padding = 20;
        int maxWidth = Width - padding * 2;
        float y = padding - _scrollOffset;

        using var sf = new StringFormat(StringFormat.GenericTypographic);
        sf.Trimming = StringTrimming.Word;
        sf.FormatFlags = 0;

        var textColor = Color.White;
        var lines = _formattedText.Split('\n');

        foreach (var line in lines)
        {
            if (string.IsNullOrEmpty(line))
            {
                y += _normalFont.Height;
                continue;
            }

            bool isHeader = (line == "FINDINGS:" || line == "IMPRESSION:" || line == "RadAI IMPRESSION:");
            var font = isHeader ? _headerFont : _normalFont;

            var layoutRect = new RectangleF(padding, y, maxWidth, 10000);
            var measured = g.MeasureString(line, font, maxWidth, sf);
            float lineHeight = Math.Max(measured.Height, font.Height);

            if (y + lineHeight > 0 && y < Height)
            {
                if (highlights)
                    DrawLineHighlights(g, line, font, layoutRect, sf);

                if (text)
                {
                    using var brush = new SolidBrush(textColor);
                    g.DrawString(line, font, brush, layoutRect, sf);
                }
            }

            y += lineHeight;
        }

        if (text)
            _totalContentHeight = (int)(y + _scrollOffset + padding);
    }

    private void DrawLineHighlights(Graphics g, string line, Font font, RectangleF layoutRect, StringFormat parentSf)
    {
        // Collect all character ranges per color, then merge overlapping ranges
        // to prevent semi-transparent rectangles from stacking.
        var rangesByColor = new Dictionary<Color, List<(int Start, int End)>>();

        foreach (var hl in _highlights)
        {
            int idx = line.IndexOf(hl.Text, StringComparison.Ordinal);
            if (idx < 0) continue;

            if (!rangesByColor.TryGetValue(hl.BackColor, out var ranges))
            {
                ranges = new List<(int, int)>();
                rangesByColor[hl.BackColor] = ranges;
            }
            ranges.Add((idx, idx + hl.Text.Length));
        }

        if (rangesByColor.Count == 0) return;

        foreach (var (color, ranges) in rangesByColor)
        {
            // Merge overlapping/adjacent character ranges
            ranges.Sort((a, b) => a.Start.CompareTo(b.Start));
            var merged = new List<(int Start, int End)> { ranges[0] };
            for (int i = 1; i < ranges.Count; i++)
            {
                var last = merged[^1];
                if (ranges[i].Start <= last.End)
                    merged[^1] = (last.Start, Math.Max(last.End, ranges[i].End));
                else
                    merged.Add(ranges[i]);
            }

            // Draw each merged range using the SAME StringFormat as text layout
            foreach (var (start, end) in merged)
            {
                try
                {
                    using var sf = new StringFormat(parentSf);
                    sf.SetMeasurableCharacterRanges(new[] { new CharacterRange(start, end - start) });

                    var regions = g.MeasureCharacterRanges(line, font, layoutRect, sf);
                    using var identity = new System.Drawing.Drawing2D.Matrix();
                    foreach (var region in regions)
                    {
                        // Use GetRegionScans instead of GetBounds to get separate
                        // tight rectangles per visual line when text wraps.
                        // GetBounds returns one big bounding box that fills the
                        // entire width, making wrapped highlights look like blocks.
                        var scans = region.GetRegionScans(identity);
                        foreach (var scanRect in scans)
                        {
                            if (scanRect.Width > 0 && scanRect.Height > 0)
                            {
                                    using var hlBrush = new SolidBrush(color);
                                g.FillRectangle(hlBrush, scanRect);
                            }
                        }
                        region.Dispose();
                    }
                }
                catch { }
            }
        }
    }


    private void DrawModeLabel(Graphics g)
    {
        // Show stale content indicator (always visible when flag is set)
        if (_showingStaleContent)
        {
            string staleLabel = "Updating...";
            var staleSize = g.MeasureString(staleLabel, _staleIndicatorFont);
            float staleX = (Width - staleSize.Width) / 2; // Center horizontally
            float staleY = 8;

            // Semi-transparent yellow background banner
            using (var bgBrush = new SolidBrush(Color.FromArgb(180, 200, 180, 0)))
            {
                g.FillRectangle(bgBrush, 0, 0, Width, staleSize.Height + 16);
            }

            // Black text on yellow background for contrast
            using (var textBrush = new SolidBrush(Color.FromArgb(255, 0, 0, 0)))
            {
                g.DrawString(staleLabel, _staleIndicatorFont, textBrush, staleX, staleY);
            }

            return; // Don't show mode label when stale indicator is showing
        }

        // Show mode label when more than one mode is available in the cycle
        bool hasMultipleModes = (_changesEnabled && _correlationEnabled) ||
                                (_correlationEnabled && _config.OrphanFindingsEnabled);
        if (!hasMultipleModes) return;

        string label = _displayMode switch
        {
            ReportDisplayMode.Changes => "Changes",
            ReportDisplayMode.Rainbow => "Rainbow",
            ReportDisplayMode.OrphanFindings => "Unmatched",
            _ => "Changes"
        };
        var labelSize = g.MeasureString(label, _modeLabelFont);
        float x = Width - labelSize.Width - 8;
        float y = 4;

        using var brush = new SolidBrush(Color.FromArgb(200, 140, 140, 140));
        g.DrawString(label, _modeLabelFont, brush, x, y);
    }

    /// <summary>
    /// Draw a pulsing yellow border when showing stale content.
    /// Opacity oscillates smoothly via sine wave.
    /// </summary>
    private void DrawStalePulseBorder(Graphics g)
    {
        if (!_showingStaleContent) return;

        // Sine wave: oscillate alpha between 40 and 160
        float t = (float)((Math.Sin(_stalePulsePhase) + 1.0) / 2.0); // 0..1
        int alpha = (int)(40 + t * 120);
        int borderWidth = 2;

        using var pen = new Pen(Color.FromArgb(alpha, 220, 180, 0), borderWidth);
        g.DrawRectangle(pen, borderWidth / 2, borderWidth / 2,
            Width - borderWidth, Height - borderWidth);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (!_useLayeredWindow)
            DrawStalePulseBorder(e.Graphics);
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

            bool isHeader = (line == "FINDINGS:" || line == "IMPRESSION:" || line == "RadAI IMPRESSION:");
            var font = isHeader ? _headerFont : _normalFont;

            var measured = g.MeasureString(line, font, maxWidth, sf);
            y += Math.Max(measured.Height, font.Height);
        }

        return (int)y + padding;
    }

    /// <summary>
    /// Stamp ClearType text pixels from textLayer onto dst bitmap.
    /// Any pixel on textLayer that differs from the opaque background (20,20,20)
    /// is a text pixel — copy its RGB and set alpha to 255 (fully opaque).
    /// Background pixels are left untouched (semi-transparent).
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
                // Text pixel: copy ClearType RGB, make fully opaque
                dstPx[i]     = srcPx[i];     // B
                dstPx[i + 1] = srcPx[i + 1]; // G
                dstPx[i + 2] = srcPx[i + 2]; // R
                dstPx[i + 3] = 255;          // A
            }
        }

        Marshal.Copy(dstPx, 0, dstData.Scan0, byteCount);
        dst.UnlockBits(dstData);
        textLayer.UnlockBits(srcData);
    }

    /// <summary>
    /// Premultiply alpha channel for correct UpdateLayeredWindow rendering.
    /// GDI+ stores straight alpha, but UpdateLayeredWindow expects premultiplied.
    /// </summary>
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
            pixels[i]     = (byte)(pixels[i]     * a / 255); // B
            pixels[i + 1] = (byte)(pixels[i + 1] * a / 255); // G
            pixels[i + 2] = (byte)(pixels[i + 2] * a / 255); // R
        }

        Marshal.Copy(pixels, 0, data.Scan0, byteCount);
        bmp.UnlockBits(data);
    }

    private void SetBitmap(Bitmap bitmap)
    {
        if (!IsHandleCreated || IsDisposed) return;

        IntPtr screenDc = IntPtr.Zero;
        IntPtr memDc = IntPtr.Zero;
        IntPtr hBitmap = IntPtr.Zero;
        IntPtr oldBitmap = IntPtr.Zero;

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

    #region Transparent Mode - Highlights

    private void ComputeHighlights()
    {
        _highlights.Clear();

        if (_displayMode == ReportDisplayMode.Changes && !string.IsNullOrEmpty(_baselineReport))
        {
            ComputeDiffHighlights();
        }
        else if (_displayMode == ReportDisplayMode.Rainbow)
        {
            ComputeRainbowHighlights();
        }
        else if (_displayMode == ReportDisplayMode.OrphanFindings)
        {
            ComputeOrphanHighlights();
        }
    }

    private void ComputeDiffHighlights()
    {
        try
        {
            Color highlightColor;
            try
            {
                var baseColor = ColorTranslator.FromHtml(_config.ReportChangesColor);
                float alpha = _config.ReportChangesAlpha / 100f;
                int a = (int)(alpha * 200);
                highlightColor = Color.FromArgb(Math.Clamp(a, 30, 200), baseColor.R, baseColor.G, baseColor.B);
            }
            catch
            {
                highlightColor = Color.FromArgb(80, 144, 238, 144);
            }

            // When baseline is from template DB, diff only FINDINGS+IMPRESSION sections
            string baselineText = _baselineReport!;
            string currentText = _currentReportText;
            if (_baselineIsSectionOnly)
            {
                var (curFindings, curImpression) = CorrelationService.ExtractSections(_currentReportText);
                if (!string.IsNullOrWhiteSpace(curFindings) && !string.IsNullOrWhiteSpace(curImpression))
                {
                    currentText = $"FINDINGS:\n{curFindings}\nIMPRESSION:\n{curImpression}";
                }
            }

            var baselineSentences = SplitIntoSentences(baselineText);
            var currentSentences = SplitIntoSentences(currentText);

            var baselineSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in baselineSentences)
            {
                var normalized = NormalizeSentence(s);
                if (!string.IsNullOrWhiteSpace(normalized))
                    baselineSet.Add(normalized);
                var stripped = StripSectionHeading(s);
                var strippedNorm = NormalizeSentence(stripped);
                if (!string.IsNullOrWhiteSpace(strippedNorm))
                    baselineSet.Add(strippedNorm);
            }

            int count = 0;
            foreach (var sentence in currentSentences)
            {
                var content = StripSectionHeading(sentence).Trim();
                var normalized = NormalizeSentence(content);

                if (string.IsNullOrWhiteSpace(normalized)) continue;
                if (content.Length < 3) continue;
                if (content.EndsWith(":")) continue;
                if (IsSectionHeading(content)) continue;
                if (baselineSet.Contains(normalized)) continue;

                _highlights.Add(new HighlightEntry(content, highlightColor));
                count++;
            }

            Logger.Trace($"Diff highlights computed: {count} entries");
        }
        catch (Exception ex)
        {
            Logger.Trace($"Diff highlight computation error: {ex.Message}");
        }
    }

    /// <summary>
    /// Extract the set of dictated (non-template) sentences by diffing against baseline.
    /// Returns null if no baseline is available.
    /// </summary>
    private HashSet<string>? GetDictatedSentences()
    {
        if (string.IsNullOrEmpty(_baselineReport)) return null;

        try
        {
            var baselineSentences = SplitIntoSentences(_baselineReport);
            var currentSentences = SplitIntoSentences(_currentReportText);

            var baselineSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in baselineSentences)
            {
                var normalized = NormalizeSentence(s);
                if (!string.IsNullOrWhiteSpace(normalized))
                    baselineSet.Add(normalized);
                var stripped = StripSectionHeading(s);
                var strippedNorm = NormalizeSentence(stripped);
                if (!string.IsNullOrWhiteSpace(strippedNorm))
                    baselineSet.Add(strippedNorm);
            }

            var dictated = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sentence in currentSentences)
            {
                var content = StripSectionHeading(sentence).Trim();
                var normalized = NormalizeSentence(content);

                if (string.IsNullOrWhiteSpace(normalized)) continue;
                if (content.Length < 3) continue;
                if (content.EndsWith(":")) continue;
                if (IsSectionHeading(content)) continue;
                if (baselineSet.Contains(normalized)) continue;

                dictated.Add(normalized);
            }

            Logger.Trace($"GetDictatedSentences: {dictated.Count} dictated sentences from {currentSentences.Count} total");
            // Return the set even when empty — an empty set means "baseline exists,
            // nothing was dictated" which produces zero filtered findings (no highlights).
            // Returning null would mean "no baseline available" and fall back to
            // unfiltered mode, incorrectly highlighting template text.
            return dictated;
        }
        catch (Exception ex)
        {
            Logger.Trace($"GetDictatedSentences error: {ex.Message}");
            return null;
        }
    }

    private void ComputeRainbowHighlights()
    {
        try
        {
            var dictated = GetDictatedSentences();
            _correlationResult ??= CorrelationService.CorrelateReversed(_formattedText, dictated, _accessionSeed);

            // Don't regress: if previous correlation had more matched groups, keep it
            // (scrape noise / transitional states during Process Report can cause transient correlation failures)
            if (_previousCorrelation != null)
            {
                // Validate previous correlation still applies to current text.
                // Use NormalizeSentence for comparison — exact Ordinal matching fails because
                // SanitizeText/FormatReportText can produce slightly different whitespace or
                // invisible-char stripping across scrapes.
                var normalizedText = NormalizeSentence(_formattedText);
                int prevMatched = 0;
                foreach (var item in _previousCorrelation.Items)
                {
                    if (string.IsNullOrEmpty(item.ImpressionText)) continue;
                    if (normalizedText.Contains(NormalizeSentence(item.ImpressionText), StringComparison.Ordinal))
                        prevMatched++;
                }
                int curMatched = _correlationResult.Items.Count(i => !string.IsNullOrEmpty(i.ImpressionText));
                if (prevMatched > curMatched)
                {
                    Logger.Trace($"Rainbow mode: Keeping previous correlation ({prevMatched} valid matches > {curMatched} current)");
                    _correlationResult = _previousCorrelation;
                }
                _previousCorrelation = null;
            }

            if (_correlationResult.Items.Count == 0)
            {
                Logger.Trace("Rainbow mode: No correlations found");
                return;
            }

            // Find section boundaries for highlight placement
            int findingsSectionIdx = _formattedText.IndexOf("FINDINGS:", StringComparison.OrdinalIgnoreCase);
            int impressionSectionIdx = _formattedText.IndexOf("IMPRESSION:", StringComparison.OrdinalIgnoreCase);
            if (impressionSectionIdx < 0)
                impressionSectionIdx = _formattedText.IndexOf("RadAI IMPRESSION:", StringComparison.OrdinalIgnoreCase);

            var bgColor = Color.FromArgb(20, 20, 20);
            int count = 0;

            foreach (var item in _correlationResult.Items)
            {
                var paletteColor = item.HighlightColor ?? CorrelationService.Palette[item.ColorIndex];
                var blended = CorrelationService.BlendWithBackground(paletteColor, bgColor);
                var hlColor = Color.FromArgb(160, blended.R, blended.G, blended.B);

                // Only highlight impression for matched groups (not orphans)
                // Search only within IMPRESSION section
                if (!string.IsNullOrWhiteSpace(item.ImpressionText))
                {
                    int searchFrom = impressionSectionIdx >= 0 ? impressionSectionIdx : 0;
                    int idx = _formattedText.IndexOf(item.ImpressionText, searchFrom, StringComparison.Ordinal);
                    if (idx >= 0)
                    {
                        _highlights.Add(new HighlightEntry(item.ImpressionText, hlColor));
                        count++;
                    }
                }

                // Search only within FINDINGS section (before IMPRESSION)
                foreach (var finding in item.MatchedFindings)
                {
                    if (IsSectionHeading(finding)) continue;
                    if (!string.IsNullOrWhiteSpace(finding))
                    {
                        int searchFrom = findingsSectionIdx >= 0 ? findingsSectionIdx : 0;
                        int searchEnd = impressionSectionIdx >= 0 ? impressionSectionIdx : _formattedText.Length;
                        int idx = _formattedText.IndexOf(finding, searchFrom, searchEnd - searchFrom, StringComparison.Ordinal);
                        if (idx >= 0)
                        {
                            _highlights.Add(new HighlightEntry(finding, hlColor));
                            count++;
                        }
                    }
                }
            }

            int orphanCount = _correlationResult.Items.Count(i => string.IsNullOrEmpty(i.ImpressionText));
            Logger.Trace($"Rainbow highlights computed: {count} entries from {_correlationResult.Items.Count} correlations ({orphanCount} orphans)");
        }
        catch (Exception ex)
        {
            Logger.Trace($"Rainbow highlight computation error: {ex.Message}");
        }
    }

    private void ComputeOrphanHighlights()
    {
        try
        {
            var dictated = GetDictatedSentences();
            _correlationResult ??= CorrelationService.CorrelateReversed(_formattedText, dictated, _accessionSeed);

            if (_correlationResult.Items.Count == 0)
            {
                Logger.Trace("Unmatched mode: No correlations found");
                return;
            }

            var bgColor = Color.FromArgb(20, 20, 20);
            int count = 0;

            foreach (var item in _correlationResult.Items)
            {
                // Only show orphan findings (items with no impression match)
                if (!string.IsNullOrEmpty(item.ImpressionText)) continue;

                var paletteColor = item.HighlightColor ?? CorrelationService.Palette[item.ColorIndex];
                var blended = CorrelationService.BlendWithBackground(paletteColor, bgColor);
                var hlColor = Color.FromArgb(160, blended.R, blended.G, blended.B);

                foreach (var finding in item.MatchedFindings)
                {
                    if (IsSectionHeading(finding)) continue;
                    if (!string.IsNullOrWhiteSpace(finding))
                    {
                        _highlights.Add(new HighlightEntry(finding, hlColor));
                        count++;
                    }
                }
            }

            int orphanCount = _correlationResult.Items.Count(i => string.IsNullOrEmpty(i.ImpressionText));
            Logger.Trace($"Unmatched highlights computed: {count} entries from {orphanCount} orphan findings");
        }
        catch (Exception ex)
        {
            Logger.Trace($"Unmatched highlight computation error: {ex.Message}");
        }
    }

    #endregion

    #region Transparent Mode - Mouse Handling

    private void OnLayeredMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            // Check if on resize edge
            if (e.X <= ResizeGripWidth)
            {
                _resizingLeft = true;
                _resizeStartX = Cursor.Position.X;
                _resizeStartWidth = Width;
                _resizeStartLeft = Left;
                return;
            }
            if (e.X >= Width - ResizeGripWidth)
            {
                _resizingRight = true;
                _resizeStartX = Cursor.Position.X;
                _resizeStartWidth = Width;
                return;
            }

            _formPosOnMouseDown = this.Location;

            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);

            if (this.Location == _formPosOnMouseDown)
            {
                if (!HandleClickCycle())
                    Close();
            }
            else
            {
                RenderAndUpdate();
            }
        }
        else if (e.Button == MouseButtons.Right)
        {
            Close();
        }
    }

    private void OnLayeredMouseMove(object? sender, MouseEventArgs e)
    {
        if (_resizingLeft || _resizingRight)
        {
            int delta = Cursor.Position.X - _resizeStartX;
            int newWidth;

            if (_resizingLeft)
            {
                newWidth = Math.Max(MinFormWidth, _resizeStartWidth - delta);
                int actualDelta = _resizeStartWidth - newWidth;
                Left = _resizeStartLeft + actualDelta;
            }
            else
            {
                newWidth = Math.Max(MinFormWidth, _resizeStartWidth + delta);
            }

            Width = newWidth;
            _totalContentHeight = MeasureContentHeight(Width);
            var screen = Screen.FromControl(this);
            int maxHeight = screen.WorkingArea.Bottom - Top - 10;
            Height = Math.Clamp(_totalContentHeight, 100, maxHeight);
            RenderAndUpdate();
            return;
        }

        // Update cursor for edge hover
        if (e.X <= ResizeGripWidth || e.X >= Width - ResizeGripWidth)
            Cursor = Cursors.SizeWE;
        else
            Cursor = Cursors.Default;
    }

    private void OnLayeredMouseUp(object? sender, MouseEventArgs e)
    {
        if ((_resizingLeft || _resizingRight) && e.Button == MouseButtons.Left)
        {
            _config.ReportPopupWidth = Width;
            _resizingLeft = false;
            _resizingRight = false;
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

    #region Opaque Mode - Rendering

    private void UpdateModeLabel()
    {
        if (_modeLabel == null) return;

        // Show stale indicator when report is being updated
        if (_showingStaleContent)
        {
            _modeLabel.Text = "  Updating...  ";
            _modeLabel.Font = _staleLabelFont;
            _modeLabel.ForeColor = Color.FromArgb(200, 180, 0); // Yellow
            _modeLabel.BackColor = Color.FromArgb(80, 60, 0); // Dark yellow background
            _modeLabel.Padding = new Padding(4, 2, 4, 2);
            _modeLabel.Visible = true;
        }
        else if ((_changesEnabled && _correlationEnabled) ||
                 (_correlationEnabled && _config.OrphanFindingsEnabled))
        {
            _modeLabel.Text = _displayMode switch
            {
                ReportDisplayMode.Changes => "Changes",
                ReportDisplayMode.Rainbow => "Rainbow",
                ReportDisplayMode.OrphanFindings => "Unmatched",
                _ => "Changes"
            };
            _modeLabel.Font = _modeLabelFont; // Restore normal font
            _modeLabel.ForeColor = Color.FromArgb(140, 140, 140); // Gray (default)
            _modeLabel.BackColor = Color.FromArgb(30, 30, 30); // Dark background (default)
            _modeLabel.Padding = Padding.Empty;
            _modeLabel.Visible = true;
        }
        else
        {
            _modeLabel.Visible = false;
        }
        PositionModeLabel();
    }

    private void PositionModeLabel()
    {
        if (_modeLabel != null && _modeLabel.Visible)
        {
            if (_showingStaleContent)
            {
                // Center horizontally for stale indicator
                _modeLabel.Location = new Point((this.ClientSize.Width - _modeLabel.Width) / 2, 4);
            }
            else
            {
                _modeLabel.Location = new Point(this.ClientSize.Width - _modeLabel.Width - 8, 4);
            }
        }
    }

    private void ResetAndReapplyFormatting()
    {
        if (_richTextBox == null) return;

        _richTextBox.Text = _formattedText;
        _richTextBox.SelectAll();
        _richTextBox.SelectionFont = _richTextBox.Font;
        _richTextBox.SelectionColor = Color.Gainsboro;
        _richTextBox.SelectionBackColor = _richTextBox.BackColor;
        _richTextBox.Select(0, 0);

        ApplyCurrentModeFormatting();
        PerformResize();
    }

    private void ApplyCurrentModeFormatting()
    {
        if (_richTextBox == null) return;

        FormatKeywords(_richTextBox, new[] { "IMPRESSION:", "RadAI IMPRESSION:", "FINDINGS:" });

        if (_displayMode == ReportDisplayMode.Changes)
        {
            if (!string.IsNullOrEmpty(_baselineReport))
            {
                string baselineText = _baselineReport;
                string currentText = _currentReportText;
                if (_baselineIsSectionOnly)
                {
                    var (curFindings, curImpression) = CorrelationService.ExtractSections(_currentReportText);
                    if (!string.IsNullOrWhiteSpace(curFindings) && !string.IsNullOrWhiteSpace(curImpression))
                    {
                        currentText = $"FINDINGS:\n{curFindings}\nIMPRESSION:\n{curImpression}";
                    }
                }
                ApplyDiffHighlighting(_richTextBox, baselineText, currentText);
            }
        }
        else if (_displayMode == ReportDisplayMode.Rainbow)
        {
            ApplyCorrelationHighlighting();
        }
        else if (_displayMode == ReportDisplayMode.OrphanFindings)
        {
            ApplyOrphanHighlighting();
        }
    }

    private void ApplyCorrelationHighlighting()
    {
        if (_richTextBox == null) return;

        try
        {
            var dictated = GetDictatedSentences();
            _correlationResult ??= CorrelationService.CorrelateReversed(_richTextBox.Text, dictated, _accessionSeed);

            // Don't regress: if previous correlation had more matched groups, keep it
            // (scrape noise / transitional states during Process Report can cause transient correlation failures)
            if (_previousCorrelation != null)
            {
                // Use NormalizeSentence for comparison — exact Ordinal matching fails because
                // SanitizeText/FormatReportText can produce slightly different whitespace or
                // invisible-char stripping across scrapes.
                var normalizedRtb = NormalizeSentence(_richTextBox.Text);
                int prevMatched = 0;
                foreach (var item in _previousCorrelation.Items)
                {
                    if (string.IsNullOrEmpty(item.ImpressionText)) continue;
                    if (normalizedRtb.Contains(NormalizeSentence(item.ImpressionText), StringComparison.Ordinal))
                        prevMatched++;
                }
                int curMatched = _correlationResult.Items.Count(i => !string.IsNullOrEmpty(i.ImpressionText));
                if (prevMatched > curMatched)
                {
                    Logger.Trace($"Rainbow mode: Keeping previous correlation ({prevMatched} valid matches > {curMatched} current)");
                    _correlationResult = _previousCorrelation;
                }
                _previousCorrelation = null;
            }

            if (_correlationResult.Items.Count == 0)
            {
                Logger.Trace("Rainbow mode: No correlations found");
                return;
            }

            var rtbText = _richTextBox.Text;
            var bg = BackColor;
            int highlightCount = 0;

            // Find section boundaries so highlights stay in their respective sections
            int findingsSectionStart = rtbText.IndexOf("FINDINGS:", StringComparison.OrdinalIgnoreCase);
            int impressionSectionStart = rtbText.IndexOf("IMPRESSION:", StringComparison.OrdinalIgnoreCase);
            if (impressionSectionStart < 0)
                impressionSectionStart = rtbText.IndexOf("RadAI IMPRESSION:", StringComparison.OrdinalIgnoreCase);

            foreach (var item in _correlationResult.Items)
            {
                var paletteColor = item.HighlightColor ?? CorrelationService.Palette[item.ColorIndex];
                var blended = CorrelationService.BlendWithBackground(paletteColor, bg);

                // Only highlight impression for matched groups (not orphans)
                // Search only in IMPRESSION section to avoid matching identical text in FINDINGS
                if (!string.IsNullOrEmpty(item.ImpressionText))
                {
                    int impressionIdx = impressionSectionStart >= 0
                        ? rtbText.IndexOf(item.ImpressionText, impressionSectionStart, StringComparison.Ordinal)
                        : rtbText.IndexOf(item.ImpressionText, StringComparison.Ordinal);
                    if (impressionIdx >= 0)
                    {
                        _richTextBox.Select(impressionIdx, item.ImpressionText.Length);
                        _richTextBox.SelectionBackColor = blended;
                        highlightCount++;
                    }
                }

                // Search only within FINDINGS section (before IMPRESSION) to avoid
                // highlighting identical text that appears in clinical history or impression
                foreach (var finding in item.MatchedFindings)
                {
                    var content = finding;
                    if (IsSectionHeading(content)) continue;

                    int searchFrom = findingsSectionStart >= 0 ? findingsSectionStart : 0;
                    int searchLen = impressionSectionStart >= 0
                        ? impressionSectionStart - searchFrom
                        : rtbText.Length - searchFrom;
                    if (searchLen <= 0) searchLen = rtbText.Length - searchFrom;

                    int findingIdx = rtbText.IndexOf(content, searchFrom, searchLen, StringComparison.Ordinal);
                    if (findingIdx >= 0)
                    {
                        _richTextBox.Select(findingIdx, content.Length);
                        _richTextBox.SelectionBackColor = blended;
                        highlightCount++;
                    }
                }
            }

            _richTextBox.Select(0, 0);
            int orphanCount = _correlationResult.Items.Count(i => string.IsNullOrEmpty(i.ImpressionText));
            Logger.Trace($"Rainbow mode: {_correlationResult.Items.Count} correlations ({orphanCount} orphans), {highlightCount} regions highlighted");
        }
        catch (Exception ex)
        {
            Logger.Trace($"Rainbow mode error: {ex.Message}");
        }
    }

    private void ApplyOrphanHighlighting()
    {
        if (_richTextBox == null) return;

        try
        {
            var dictated = GetDictatedSentences();
            _correlationResult ??= CorrelationService.CorrelateReversed(_richTextBox.Text, dictated, _accessionSeed);

            if (_correlationResult.Items.Count == 0)
            {
                Logger.Trace("Unmatched mode: No correlations found");
                return;
            }

            var rtbText = _richTextBox.Text;
            var bg = BackColor;
            int highlightCount = 0;

            // Section boundaries for finding highlights
            int findingsSectionStart = rtbText.IndexOf("FINDINGS:", StringComparison.OrdinalIgnoreCase);
            int impressionSectionStart = rtbText.IndexOf("IMPRESSION:", StringComparison.OrdinalIgnoreCase);
            if (impressionSectionStart < 0)
                impressionSectionStart = rtbText.IndexOf("RadAI IMPRESSION:", StringComparison.OrdinalIgnoreCase);

            foreach (var item in _correlationResult.Items)
            {
                // Only show orphan findings (items with no impression match)
                if (!string.IsNullOrEmpty(item.ImpressionText)) continue;

                var paletteColor = item.HighlightColor ?? CorrelationService.Palette[item.ColorIndex];
                var blended = CorrelationService.BlendWithBackground(paletteColor, bg);

                foreach (var finding in item.MatchedFindings)
                {
                    if (IsSectionHeading(finding)) continue;

                    int searchFrom = findingsSectionStart >= 0 ? findingsSectionStart : 0;
                    int searchLen = impressionSectionStart >= 0
                        ? impressionSectionStart - searchFrom
                        : rtbText.Length - searchFrom;
                    if (searchLen <= 0) searchLen = rtbText.Length - searchFrom;

                    int findingIdx = rtbText.IndexOf(finding, searchFrom, searchLen, StringComparison.Ordinal);
                    if (findingIdx >= 0)
                    {
                        _richTextBox.Select(findingIdx, finding.Length);
                        _richTextBox.SelectionBackColor = blended;
                        highlightCount++;
                    }
                }
            }

            _richTextBox.Select(0, 0);
            int orphanCount = _correlationResult.Items.Count(i => string.IsNullOrEmpty(i.ImpressionText));
            Logger.Trace($"Unmatched mode: {orphanCount} orphan findings, {highlightCount} regions highlighted");
        }
        catch (Exception ex)
        {
            Logger.Trace($"Unmatched mode error: {ex.Message}");
        }
    }

    private void FormatKeywords(RichTextBox rtb, string[] keywords)
    {
        float baseSize = rtb.Font.Size;
        using var highlightFont = new Font(rtb.Font.FontFamily, baseSize + 2, FontStyle.Bold);
        Color highlightColor = Color.White;

        foreach (var word in keywords)
        {
            int startIndex = 0;
            while (startIndex < rtb.TextLength)
            {
                int foundIndex = rtb.Find(word, startIndex, RichTextBoxFinds.WholeWord | RichTextBoxFinds.MatchCase);
                if (foundIndex == -1) break;

                rtb.Select(foundIndex, word.Length);
                rtb.SelectionColor = highlightColor;
                rtb.SelectionFont = highlightFont;

                startIndex = foundIndex + word.Length;
            }
        }

        rtb.Select(0, 0);
    }

    private void ApplyDiffHighlighting(RichTextBox rtb, string baseline, string current)
    {
        try
        {
            Color highlightColor;
            try
            {
                var baseColor = ColorTranslator.FromHtml(_config.ReportChangesColor);
                var bg = BackColor;
                float alpha = _config.ReportChangesAlpha / 100f;
                highlightColor = Color.FromArgb(
                    (int)(bg.R + (baseColor.R - bg.R) * alpha),
                    (int)(bg.G + (baseColor.G - bg.G) * alpha),
                    (int)(bg.B + (baseColor.B - bg.B) * alpha)
                );
            }
            catch
            {
                highlightColor = Color.FromArgb(58, 82, 58);
            }

            var baselineSentences = SplitIntoSentences(baseline);
            var currentSentences = SplitIntoSentences(current);

            var baselineSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in baselineSentences)
            {
                var normalized = NormalizeSentence(s);
                if (!string.IsNullOrWhiteSpace(normalized))
                    baselineSet.Add(normalized);
            }

            var newSentences = new List<(string original, string content)>();
            foreach (var sentence in currentSentences)
            {
                var content = StripSectionHeading(sentence);
                var normalized = NormalizeSentence(content);
                if (!string.IsNullOrWhiteSpace(normalized) && !baselineSet.Contains(normalized))
                {
                    newSentences.Add((sentence.Trim(), content.Trim()));
                }
            }

            foreach (var s in baselineSentences)
            {
                var content = StripSectionHeading(s);
                var normalized = NormalizeSentence(content);
                if (!string.IsNullOrWhiteSpace(normalized))
                    baselineSet.Add(normalized);
            }

            newSentences = newSentences.Where(x =>
                !baselineSet.Contains(NormalizeSentence(x.content))).ToList();

            string rtbText = rtb.Text;
            int highlightCount = 0;

            foreach (var (original, content) in newSentences)
            {
                if (string.IsNullOrWhiteSpace(content)) continue;
                if (content.Length < 3) continue;
                if (content.EndsWith(":")) continue;
                if (IsSectionHeading(content)) continue;

                int idx = rtbText.IndexOf(content, StringComparison.Ordinal);
                if (idx >= 0)
                {
                    rtb.Select(idx, content.Length);
                    rtb.SelectionBackColor = highlightColor;
                    highlightCount++;
                }
            }

            rtb.Select(0, 0);
            Logger.Trace($"Applied diff highlighting: {highlightCount} new sentences highlighted (from {newSentences.Count} detected)");
        }
        catch (Exception ex)
        {
            Logger.Trace($"Diff highlighting error: {ex.Message}");
        }
    }

    #endregion

    #region Opaque Mode - Resize

    private void RichTextBox_ContentsResized(object? sender, ContentsResizedEventArgs e)
    {
        if (_isResizing) return;
        if (!this.IsHandleCreated) return;

        this.BeginInvoke(new Action(() =>
        {
            if (_isResizing) return;
            _isResizing = true;
            try
            {
                PerformResize();
            }
            finally
            {
                _isResizing = false;
            }
        }));
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
        int requiredTotalHeight = contentHeight + (padding * 2);

        if (requiredTotalHeight < 100) requiredTotalHeight = 100;

        Rectangle workArea = Screen.FromControl(this).WorkingArea;
        int maxHeight = workArea.Bottom - Top - 10;

        int finalFormHeight;
        bool needsScroll;

        if (requiredTotalHeight > maxHeight)
        {
            finalFormHeight = maxHeight;
            needsScroll = true;
        }
        else
        {
            finalFormHeight = requiredTotalHeight;
            needsScroll = false;
        }

        if (needsScroll && _richTextBox.ScrollBars != RichTextBoxScrollBars.Vertical)
        {
            _richTextBox.ScrollBars = RichTextBoxScrollBars.Vertical;
        }
        else if (!needsScroll && _richTextBox.ScrollBars != RichTextBoxScrollBars.None)
        {
            _richTextBox.ScrollBars = RichTextBoxScrollBars.None;
        }

        if (this.ClientSize.Height != finalFormHeight)
        {
            this.ClientSize = new Size(this.ClientSize.Width, finalFormHeight);
        }

        int rtbHeight = finalFormHeight - (padding * 2);
        if (_richTextBox.Height != rtbHeight)
        {
            _richTextBox.Height = rtbHeight;
        }

        PositionModeLabel();
    }

    #endregion

    #region Opaque Mode - Interactions

    private void SetupOpaqueInteractions(Control control)
    {
        Point formPosOnMouseDown = Point.Empty;

        control.MouseDown += (s, e) =>
        {
            if (e.Button == MouseButtons.Left)
            {
                // Check resize edges (form-level coordinates)
                int formX = (control == this) ? e.X : e.X + control.Left;
                if (formX <= ResizeGripWidth)
                {
                    _resizingLeft = true;
                    _resizeStartX = Cursor.Position.X;
                    _resizeStartWidth = Width;
                    _resizeStartLeft = Left;
                    return;
                }
                if (formX >= Width - ResizeGripWidth)
                {
                    _resizingRight = true;
                    _resizeStartX = Cursor.Position.X;
                    _resizeStartWidth = Width;
                    return;
                }

                formPosOnMouseDown = this.Location;

                _dragStart = new Point(e.X, e.Y);

                if (control != this)
                {
                    ReleaseCapture();
                    SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);

                    if (this.Location == formPosOnMouseDown)
                    {
                        if (!HandleClickCycle())
                            Close();
                    }
                    _dragging = false;
                }
                else
                {
                    _dragging = true;
                }
            }
            if (e.Button == MouseButtons.Right) Close();
        };

        control.MouseMove += (s, e) =>
        {
            if (_resizingLeft || _resizingRight)
            {
                int delta = Cursor.Position.X - _resizeStartX;
                int newWidth;

                if (_resizingLeft)
                {
                    newWidth = Math.Max(MinFormWidth, _resizeStartWidth - delta);
                    int actualDelta = _resizeStartWidth - newWidth;
                    Left = _resizeStartLeft + actualDelta;
                }
                else
                {
                    newWidth = Math.Max(MinFormWidth, _resizeStartWidth + delta);
                }

                if (newWidth != Width)
                {
                    Width = newWidth;
                    if (_richTextBox != null)
                    {
                        int padding = 20;
                        _richTextBox.Width = Width - (padding * 2);
                        if (_deletableEnabled)
                            _richTextBox.RightMargin = _richTextBox.Width - 26;
                    }
                    PerformResize();
                }
                return;
            }

            if (_dragging)
            {
                Point currentScreenPos = Cursor.Position;
                Location = new Point(
                    currentScreenPos.X - _dragStart.X,
                    currentScreenPos.Y - _dragStart.Y
                );
            }
            else
            {
                // Update cursor for edge hover
                int formX = (control == this) ? e.X : e.X + control.Left;
                if (formX <= ResizeGripWidth || formX >= Width - ResizeGripWidth)
                    control.Cursor = Cursors.SizeWE;
                else
                    control.Cursor = Cursors.Hand;
            }
        };

        control.MouseUp += (s, e) =>
        {
            if ((_resizingLeft || _resizingRight) && e.Button == MouseButtons.Left)
            {
                _config.ReportPopupWidth = Width;
                _resizingLeft = false;
                _resizingRight = false;
                return;
            }
            if (_dragging && e.Button == MouseButtons.Left)
            {
                if (this.Location == formPosOnMouseDown)
                {
                    if (!HandleClickCycle())
                        Close();
                }
            }
            _dragging = false;
        };
    }

    #endregion

    #region Deletable Impression Points

    /// <summary>
    /// Parse impression items from the RichTextBox text and create/position trash icons.
    /// </summary>
    private void PositionTrashIcons()
    {
        if (_richTextBox == null || !_deletableEnabled) return;

        // Parse impression items from RichTextBox text
        var rtbText = _richTextBox.Text;
        int impressionStart = rtbText.IndexOf("IMPRESSION:", StringComparison.OrdinalIgnoreCase);
        if (impressionStart < 0)
        {
            ClearTrashIcons();
            foreach (var lbl in _impressionFixerLabels) lbl.Visible = false;
            return;
        }

        // Position impression fixer labels on the IMPRESSION header line, right-aligned
        if (_impressionFixerLabels.Count > 0)
        {
            var headerPos = _richTextBox.GetPositionFromCharIndex(impressionStart);
            int labelY = _richTextBox.Top + headerPos.Y;
            bool inView = (labelY >= _richTextBox.Top && labelY + 10 < _richTextBox.Bottom);

            // Check if COMPARISON has an actual date and parse it
            bool hasComparison = false;
            DateTime? comparisonDate = null;
            int compIdx = rtbText.IndexOf("COMPARISON:", StringComparison.OrdinalIgnoreCase);
            if (compIdx >= 0)
            {
                int compEnd = rtbText.Length;
                var nextHeader = Regex.Match(rtbText.Substring(compIdx + 11), @"^[A-Z ]{2,}:", RegexOptions.Multiline);
                if (nextHeader.Success) compEnd = compIdx + 11 + nextHeader.Index;
                var compSection = rtbText.Substring(compIdx, Math.Min(compEnd - compIdx, 200));
                var dateMatch = Regex.Match(compSection, @"\d{1,2}/\d{1,2}/\d{2,4}");
                hasComparison = !compSection.Contains("None available", StringComparison.OrdinalIgnoreCase)
                    && dateMatch.Success;
                if (hasComparison && DateTime.TryParse(dateMatch.Value, out var parsed))
                    comparisonDate = parsed;
            }

            // Layout right-to-left
            int x = _richTextBox.Right - 2;
            for (int i = _impressionFixerLabels.Count - 1; i >= 0; i--)
            {
                var lbl = _impressionFixerLabels[i];
                var entry = _impressionFixerEntries[i];
                bool entryVisible = inView;
                if (entry.RequireComparison)
                {
                    if (!hasComparison)
                        entryVisible = false;
                    else if (entry.MaxComparisonWeeks > 0 && comparisonDate.HasValue)
                        entryVisible = entryVisible && (DateTime.Today - comparisonDate.Value.Date).TotalDays <= entry.MaxComparisonWeeks * 7;
                }
                if (entryVisible)
                {
                    x -= lbl.Width + 4;
                    lbl.Location = new Point(x, labelY);
                }
                lbl.Visible = entryVisible;
            }
        }

        // Find numbered items after IMPRESSION:
        int searchFrom = impressionStart + "IMPRESSION:".Length;
        var items = new List<(int charIndex, string itemText)>();
        var regex = new Regex(@"^\d+\.\s*(.*)$", RegexOptions.Multiline);

        foreach (Match m in regex.Matches(rtbText, searchFrom))
        {
            items.Add((m.Index, m.Groups[1].Value.Trim()));
        }

        if (items.Count == 0)
        {
            ClearTrashIcons();
            return;
        }

        _impressionItems = items;

        // Ensure we have the right number of trash icon labels
        while (_trashIcons.Count < items.Count)
        {
            int idx = _trashIcons.Count;
            var trashLabel = new Label
            {
                Text = "\u2715", // ✕
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 100, 100),
                BackColor = Color.FromArgb(30, 30, 30),
                Size = new Size(22, 22),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                Visible = false,
                Tag = idx
            };
            trashLabel.MouseEnter += (_, _) => trashLabel.ForeColor = Color.FromArgb(220, 60, 60);
            trashLabel.MouseLeave += (_, _) => trashLabel.ForeColor = Color.FromArgb(100, 100, 100);
            trashLabel.Click += OnTrashIconClick;
            Controls.Add(trashLabel);
            trashLabel.BringToFront();
            _trashIcons.Add(trashLabel);
        }

        // Hide excess icons
        for (int i = items.Count; i < _trashIcons.Count; i++)
            _trashIcons[i].Visible = false;

        // Position each trash icon next to its impression line
        for (int i = 0; i < items.Count; i++)
        {
            var (charIdx, _) = items[i];
            var pos = _richTextBox.GetPositionFromCharIndex(charIdx);

            // Convert RichTextBox-local position to form-local position
            int formX = _richTextBox.Right - 24;
            int formY = _richTextBox.Top + pos.Y;

            _trashIcons[i].Tag = i;
            _trashIcons[i].Location = new Point(formX, formY);
            _trashIcons[i].Visible = (formY >= _richTextBox.Top && formY + 10 < _richTextBox.Bottom);
        }
    }

    private void ClearTrashIcons()
    {
        foreach (var icon in _trashIcons)
            icon.Visible = false;
        _impressionItems = null;
    }

    private void OnTrashIconClick(object? sender, EventArgs e)
    {
        if (sender is not Label label || _impressionItems == null || _richTextBox == null) return;
        int index = (int)label.Tag!;
        if (index < 0 || index >= _impressionItems.Count) return;

        // Remove the line from the RichTextBox
        var (charIdx, _) = _impressionItems[index];
        int lineStart = charIdx;
        int lineEnd = _richTextBox.Text.IndexOf('\n', charIdx);
        if (lineEnd < 0) lineEnd = _richTextBox.Text.Length;
        else lineEnd++; // Include the newline

        // Also remove leading blank line if present
        if (lineStart > 0 && _richTextBox.Text[lineStart - 1] == '\n')
        {
            // Don't remove if this would eat into non-impression content
            // Only remove the preceding newline if the previous char is also a newline (blank line)
        }

        _richTextBox.Select(lineStart, lineEnd - lineStart);
        _richTextBox.ReadOnly = false;
        _richTextBox.SelectedText = "";
        _richTextBox.ReadOnly = true;

        // Renumber remaining impression items
        RenumberImpressionItems();

        // Re-apply formatting since we changed the text
        _formattedText = _richTextBox.Text;
        ApplyCurrentModeFormatting();

        _deletePending = true;

        // Reposition trash icons for updated content
        PositionTrashIcons();

        // Reset debounce timer
        _debounceTimer?.Stop();
        _debounceTimer?.Start();
    }

    private void OnFixerLabelClick(object? sender, EventArgs e)
    {
        if (_richTextBox == null || sender is not Label lbl || lbl.Tag is not ImpressionFixerEntry entry) return;

        var rtbText = _richTextBox.Text;
        int impressionStart = rtbText.IndexOf("IMPRESSION:", StringComparison.OrdinalIgnoreCase);
        if (impressionStart < 0) return;

        if (entry.ReplaceMode)
        {
            // Replace all text after IMPRESSION: header
            int lineEnd = rtbText.IndexOf('\n', impressionStart);
            int replaceFrom = lineEnd >= 0 ? lineEnd + 1 : rtbText.Length;
            int replaceLen = rtbText.Length - replaceFrom;

            _richTextBox.ReadOnly = false;
            _richTextBox.Select(replaceFrom, replaceLen);
            _richTextBox.SelectedText = $"1. {entry.Text}\n";
            _richTextBox.ReadOnly = true;
        }
        else
        {
            // Insert: append at end with next number
            int searchFrom = impressionStart + "IMPRESSION:".Length;
            var regex = new Regex(@"^\d+\.\s", RegexOptions.Multiline);
            int count = regex.Matches(rtbText, searchFrom).Count;
            int nextNum = count + 1;

            int insertPos = rtbText.Length;
            var prefix = insertPos > 0 && rtbText[insertPos - 1] != '\n' ? "\n" : "";
            var newLine = $"{prefix}{nextNum}. {entry.Text}\n";
            _richTextBox.ReadOnly = false;
            _richTextBox.Select(insertPos, 0);
            _richTextBox.SelectedText = newLine;
            _richTextBox.ReadOnly = true;
        }

        _formattedText = _richTextBox.Text;
        ApplyCurrentModeFormatting();
        _deletePending = true;
        PositionTrashIcons();

        _debounceTimer?.Stop();
        _debounceTimer?.Start();
    }

    private void RenumberImpressionItems()
    {
        if (_richTextBox == null) return;

        var rtbText = _richTextBox.Text;
        int impressionStart = rtbText.IndexOf("IMPRESSION:", StringComparison.OrdinalIgnoreCase);
        if (impressionStart < 0) return;

        int searchFrom = impressionStart + "IMPRESSION:".Length;
        var regex = new Regex(@"^(\d+)\.\s", RegexOptions.Multiline);
        var matches = regex.Matches(rtbText, searchFrom);

        _richTextBox.ReadOnly = false;
        int expectedNum = 1;
        int offset = 0; // Track text length changes from renumbering

        foreach (Match m in matches)
        {
            int adjustedIndex = m.Groups[1].Index + offset;
            string currentNum = m.Groups[1].Value;
            string expectedStr = expectedNum.ToString();

            if (currentNum != expectedStr)
            {
                _richTextBox.Select(adjustedIndex, currentNum.Length);
                _richTextBox.SelectedText = expectedStr;
                offset += expectedStr.Length - currentNum.Length;
            }
            expectedNum++;
        }
        _richTextBox.ReadOnly = true;
    }

    private void OnDebounceTimerTick(object? sender, EventArgs e)
    {
        _debounceTimer?.Stop();

        if (_richTextBox == null) return;

        // Extract current impression items from the (modified) text
        var rtbText = _richTextBox.Text;
        int impressionStart = rtbText.IndexOf("IMPRESSION:", StringComparison.OrdinalIgnoreCase);

        if (impressionStart < 0)
        {
            ImpressionDeleteRequested?.Invoke("");
            return;
        }

        int searchFrom = impressionStart + "IMPRESSION:".Length;
        var regex = new Regex(@"^\d+\.\s*(.*)$", RegexOptions.Multiline);
        var matches = regex.Matches(rtbText, searchFrom);

        var itemTexts = new List<string>();
        foreach (Match m in matches)
        {
            var text = m.Groups[1].Value.Trim();
            if (!string.IsNullOrEmpty(text))
                itemTexts.Add(text);
        }

        // Join without numbers — Mosaic auto-numbers impression lines
        var newText = string.Join("\r\n", itemTexts);
        ImpressionDeleteRequested?.Invoke(newText);
    }

    /// <summary>
    /// Clear the delete-pending flag so scrape updates resume.
    /// Called by ActionController after paste completes.
    /// </summary>
    public void ClearDeletePending()
    {
        if (InvokeRequired)
        {
            Invoke(ClearDeletePending);
            return;
        }
        _deletePending = false;
    }

    #endregion

    #region Disposal

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _normalFont?.Dispose();
            _headerFont?.Dispose();
            _modeLabelFont?.Dispose();
            _staleIndicatorFont?.Dispose();
            _staleLabelFont?.Dispose();
            if (_debounceTimer != null)
            {
                _debounceTimer.Stop();
                _debounceTimer.Dispose();
            }
            if (_stalePulseTimer != null)
            {
                _stalePulseTimer.Stop();
                _stalePulseTimer.Dispose();
            }
        }
        base.Dispose(disposing);
    }

    #endregion

    #region Text Formatting (shared)

    /// <summary>
    /// Format scraped report text for display.
    /// - Single blank line between major sections
    /// - Content in most sections joined into paragraphs
    /// - FINDINGS: subsections on own lines with blank lines between, sentences joined
    /// - IMPRESSION: numbered items on own lines
    /// </summary>
    internal static string FormatReportText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return text;

        text = text.Replace("\r\n", "\n").Replace("\r", "\n");

        var lines = text.Split('\n');
        var outputLines = new List<string>();

        var majorSections = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "EXAM:", "COMPARISON:", "CLINICAL HISTORY:", "FINDINGS:", "IMPRESSION:",
            "TECHNIQUE:", "INDICATION:", "PROCEDURE:", "CONCLUSION:", "RECOMMENDATION:"
        };

        string currentSection = "";
        var pendingContent = new List<string>();

        void FlushContent(bool asJoined = true)
        {
            if (pendingContent.Count > 0)
            {
                if (asJoined)
                    outputLines.Add(string.Join(" ", pendingContent));
                else
                    outputLines.AddRange(pendingContent);
                pendingContent.Clear();
            }
        }

        foreach (var rawLine in lines)
        {
            var line = Regex.Replace(rawLine.Trim(), @"\s+", " ");

            if (string.IsNullOrWhiteSpace(line))
                continue;

            bool isMajorSection = majorSections.Contains(line);
            bool isSubsectionHeader = !isMajorSection && IsSubsectionHeader(line);
            bool isNumberedItem = Regex.IsMatch(line, @"^\d+\.");

            if (isMajorSection)
            {
                FlushContent();

                if (outputLines.Count > 0)
                    outputLines.Add("");

                outputLines.Add(line);
                currentSection = line.TrimEnd(':').ToUpperInvariant();
            }
            else if (currentSection == "FINDINGS")
            {
                if (isSubsectionHeader)
                {
                    FlushContent();

                    if (outputLines.Count > 0 && outputLines[outputLines.Count - 1] != "FINDINGS:")
                        outputLines.Add("");

                    pendingContent.Add(line);
                }
                else
                {
                    pendingContent.Add(line);
                }
            }
            else if (currentSection == "IMPRESSION")
            {
                if (isNumberedItem)
                {
                    FlushContent();
                    outputLines.Add(line);
                }
                else if (outputLines.Count > 0 && Regex.IsMatch(outputLines[^1], @"^\d+\."))
                {
                    // Continuation of previous numbered item — join so the impression
                    // stays on one logical line for correlation highlight matching
                    outputLines[^1] += " " + line;
                }
                else
                {
                    pendingContent.Add(line);
                }
            }
            else
            {
                pendingContent.Add(line);
            }
        }

        FlushContent();

        return string.Join(Environment.NewLine, outputLines);
    }

    private static bool IsSubsectionHeader(string line)
    {
        int colonIdx = line.IndexOf(':');
        if (colonIdx <= 0) return false;

        string prefix = line.Substring(0, colonIdx);

        if (prefix.Length < 2) return false;

        int upperCount = 0;
        int letterCount = 0;
        foreach (char c in prefix)
        {
            if (char.IsLetter(c))
            {
                letterCount++;
                if (char.IsUpper(c)) upperCount++;
            }
        }

        return letterCount > 0 && (double)upperCount / letterCount > 0.8;
    }

    private static List<string> SplitIntoSentences(string text)
    {
        var sentences = new List<string>();
        if (string.IsNullOrEmpty(text)) return sentences;

        text = text.Replace("\r\n", "\n").Replace("\r", "\n");

        var current = new StringBuilder();

        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            current.Append(c);

            bool isSentenceEnd = (c == '.' || c == '!' || c == '?') &&
                                 (i == text.Length - 1 || char.IsWhiteSpace(text[i + 1]));

            bool isNewline = c == '\n';

            if (isSentenceEnd || isNewline)
            {
                var sentence = current.ToString().Trim();
                if (!string.IsNullOrWhiteSpace(sentence))
                {
                    sentences.Add(sentence);
                }
                current.Clear();
            }
        }

        var remaining = current.ToString().Trim();
        if (!string.IsNullOrWhiteSpace(remaining))
        {
            sentences.Add(remaining);
        }

        return sentences;
    }

    private static string NormalizeSentence(string sentence)
    {
        if (string.IsNullOrEmpty(sentence)) return "";

        // Strip invisible Unicode characters (zero-width spaces, format chars, etc.)
        // that Mosaic's ProseMirror editor inserts. Without this, baseline template
        // text with/without invisible chars won't match, causing false diff highlights.
        var sb = new StringBuilder(sentence.Length);
        foreach (char c in sentence)
        {
            var cat = char.GetUnicodeCategory(c);
            if (cat == System.Globalization.UnicodeCategory.Format) continue;
            if (char.IsControl(c) && c != '\n' && c != '\r' && c != '\t') continue;
            if (c == '\uFFFC' || c == '\uFFFD') continue;
            sb.Append(c == '\u00A0' ? ' ' : c); // Non-breaking space → space
        }

        return Regex.Replace(sb.ToString().ToLowerInvariant().Trim(), @"\s+", " ");
    }

    private static bool IsSectionHeading(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;

        var trimmed = text.Trim().TrimEnd(':');
        if (trimmed.Length < 2) return false;

        int upperCount = 0;
        int letterCount = 0;
        foreach (char c in trimmed)
        {
            if (char.IsLetter(c))
            {
                letterCount++;
                if (char.IsUpper(c)) upperCount++;
            }
        }

        return letterCount > 0 && (double)upperCount / letterCount > 0.8;
    }

    /// <summary>
    /// Remove invisible Unicode characters (zero-width spaces, BOM, format chars, etc.)
    /// that Mosaic's ProseMirror editor inserts. These render as visible boxes
    /// with GDI+ DrawString but are invisible in RichTextBox.
    /// </summary>
    internal static string SanitizeText(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var sb = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            // Keep standard whitespace
            if (c == '\r' || c == '\n' || c == '\t')
            {
                sb.Append(c);
                continue;
            }

            // Replace non-breaking space with regular space
            if (c == '\u00A0')
            {
                sb.Append(' ');
                continue;
            }

            var category = char.GetUnicodeCategory(c);

            // Skip: Format chars (Cf) - zero-width spaces, direction marks, joiners, BOM, etc.
            // Skip: Control chars (Cc) - except the ones we already kept above
            // Skip: Surrogates (Cs), Private Use (Co), Not Assigned (Cn)
            // Skip: Object/Replacement characters U+FFFC, U+FFFD
            if (category == System.Globalization.UnicodeCategory.Format ||
                category == System.Globalization.UnicodeCategory.Control ||
                category == System.Globalization.UnicodeCategory.Surrogate ||
                category == System.Globalization.UnicodeCategory.PrivateUse ||
                category == System.Globalization.UnicodeCategory.OtherNotAssigned ||
                c == '\uFFFC' || c == '\uFFFD')
            {
                continue;
            }

            sb.Append(c);
        }
        return sb.ToString();
    }

    private static string StripSectionHeading(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return text;

        var trimmed = text.Trim();
        int colonIdx = trimmed.IndexOf(':');
        if (colonIdx < 0 || colonIdx >= trimmed.Length - 1) return trimmed;

        var prefix = trimmed.Substring(0, colonIdx);
        if (IsSectionHeading(prefix + ":"))
        {
            return trimmed.Substring(colonIdx + 1).Trim();
        }

        return trimmed;
    }

    #endregion
}
