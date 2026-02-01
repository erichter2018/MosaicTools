using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Alert types for the notification box (alerts-only mode).
/// </summary>
public enum AlertType
{
    TemplateMismatch,
    GenderMismatch,
    StrokeDetected
}

/// <summary>
/// Floating window displaying clinical history from Clario scrape.
/// Can also operate in "alerts only" mode where it only appears when alerts trigger.
/// Supports transparent (layered window) and opaque (standard controls) rendering modes.
/// </summary>
public class ClinicalHistoryForm : Form
{
    private readonly Configuration _config;
    private readonly Label _contentLabel;
    private readonly bool _useLayeredWindow;
    private readonly int _backgroundAlpha;
    private readonly Font _contentFont;

    // Drag state
    private Point _dragStart;
    private bool _dragging;

    // For center-based positioning
    private bool _initialPositionSet = false;

    // Layout constants (TransparentBorderWidth only affects transparent mode rendering)
    private const int TransparentBorderWidth = 2;
    private const int DragBarHeight = 16;
    private const int ContentMarginLeft = 10;
    private const int ContentMarginRight = 10;
    private const int ContentMarginTop = 5;
    private const int ContentMarginBottom = 15;

    // Border colors - Priority order: Green (drafted) < Purple (stroke) < Flashing Red (gender) < Solid Red (template mismatch)
    private static readonly Color NormalBorderColor = Color.FromArgb(60, 60, 60);
    private static readonly Color DraftedBorderColor = Color.FromArgb(0, 180, 0); // Green - drafted status
    private static readonly Color StrokeBorderColor = Color.FromArgb(140, 80, 200); // Purple - stroke case
    private static readonly Color TemplateMismatchBorderColor = Color.FromArgb(220, 50, 50); // Red - template mismatch

    // Text colors
    private static readonly Color NormalTextColor = Color.FromArgb(200, 200, 200); // White-ish
    private static readonly Color FixedTextColor = Color.FromArgb(255, 255, 120); // Light yellow
    private static readonly Color EmptyTextColor = Color.FromArgb(128, 128, 128); // Gray

    // Template matching state for debug
    private string? _lastDescription;
    private string? _lastTemplateName;
    private bool _templateMismatch = false;

    // Auto-fix tracking - local fallback, but prefer session-wide callbacks if set
    private string? _lastAutoFixedAccession;
    private DateTime _lastAutoFixTime = DateTime.MinValue;
    private volatile bool _autoFixInProgress = false;

    // Session-wide clinical history fix tracking callbacks (set by MainForm)
    private Func<string?, bool>? _hasClinicalHistoryFixed;
    private Action<string?>? _markClinicalHistoryFixed;
    private Action? _onAutoFixComplete;
    private Func<bool>? _isAddendumOpen;

    // Track displayed text and whether it was "fixed" from original
    private string? _currentDisplayedText;
    private bool _currentTextWasFixed = false;

    // Gender check - terms that are impossible for the opposite gender
    private static readonly string[] FemaleOnlyTerms = {
        "uterus", "uterine", "ovary", "ovaries", "ovarian",
        "fallopian", "endometrium", "endometrial",
        "vagina", "vaginal", "vulva", "vulvar",
        "adnexa", "adnexal", "cervix",
        "pregnancy", "pregnant", "gravid", "gestational",
        "placenta", "placental", "fetus", "fetal",
        "hysterectomy", "myomectomy", "endometrial ablation",
        "oophorectomy", "salpingectomy", "salpingo-oophorectomy",
        "tubal ligation", "colposcopy", "conization", "cone biopsy", "LEEP",
        "vaginoplasty", "colporrhaphy", "cervical cerclage",
        "cesarean", "c-section", "D&C", "dilation and curettage",
        "vulvectomy", "labiaplasty"
    };

    private static readonly string[] MaleOnlyTerms = {
        "prostate", "prostatic", "seminal vesicle", "seminal vesicles",
        "testicle", "testis", "testes", "testicular",
        "scrotum", "scrotal", "epididymis", "epididymal",
        "spermatic cord", "penis", "penile", "vas deferens",
        "prostatectomy", "orchiectomy", "vasectomy", "TURP",
        "transurethral resection of prostate"
    };

    // Gender warning state
    private bool _genderWarningActive = false;
    private string? _genderWarningText;
    private string? _savedClinicalHistoryText;

    // Stroke detection state
    private bool _strokeDetected = false;
    private System.Windows.Forms.Timer? _blinkTimer;
    private bool _blinkState = false;
    private bool _noteCreated = false;
    private Label? _noteCreatedIndicator;

    // Clinical history fix indicator
    private bool _historyFixInserted = false;
    private Label? _historyFixedIndicator;

    // Alert-only mode state
    private bool _showingAlert = false;
    private AlertType? _currentAlertType = null;

    // Tooltip for border color explanation
    private readonly ToolTip _borderTooltip;

    // Callback for stroke note creation (set by MainForm)
    private Func<bool>? _onStrokeNoteClick;

    // Callback for critical note creation
    private Action? _onCriticalNoteClick;

    // Transparent mode mouse handling
    private Point _formPosOnMouseDown;
    private ContextMenuStrip? _transparentDragMenu;
    private ContextMenuStrip? _transparentContentMenu;

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

    public ClinicalHistoryForm(Configuration config)
    {
        _config = config;
        _useLayeredWindow = config.ReportPopupTransparent;
        _backgroundAlpha = Math.Clamp((int)Math.Round(config.ReportPopupTransparency / 100.0 * 255), 30, 255);
        _contentFont = new Font("Segoe UI", 11, FontStyle.Bold);

        // Form properties - frameless, topmost
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = NormalBorderColor;
        StartPosition = FormStartPosition.Manual;

        // Content label - always created (data holder in transparent mode, visible in opaque)
        _contentLabel = new Label
        {
            Text = "(No clinical history)",
            Font = _contentFont,
            ForeColor = Color.FromArgb(200, 200, 200),
            BackColor = Color.Black,
            AutoSize = true,
            Margin = new Padding(ContentMarginLeft, ContentMarginTop, ContentMarginRight, ContentMarginBottom),
            MaximumSize = new Size(0, 0)
        };

        if (_useLayeredWindow)
            SetupTransparentMode();
        else
            SetupOpaqueMode();

        // Tooltip for border color explanation (shared)
        _borderTooltip = new ToolTip
        {
            InitialDelay = 500,
            AutoPopDelay = 10000,
            ReshowDelay = 200
        };
        UpdateBorderTooltip();
    }

    #region Setup

    private void SetupTransparentMode()
    {
        AutoSize = false;
        Padding = Padding.Empty;

        // Context menus for transparent mode (shown programmatically)
        _transparentDragMenu = new ContextMenuStrip();
        _transparentDragMenu.Items.Add("Close", null, (_, _) => Close());

        _transparentContentMenu = new ContextMenuStrip();
        _transparentContentMenu.Items.Add("Create Critical Note", null, (_, _) => _onCriticalNoteClick?.Invoke());
        _transparentContentMenu.Items.Add(new ToolStripSeparator());
        _transparentContentMenu.Items.Add("Copy Debug Info", null, (_, _) => CopyDebugInfoToClipboard());

        MouseDown += OnTransparentMouseDown;

        var (fw, fh) = MeasureFormSize();
        Size = new Size(fw, fh);

        Shown += (_, _) => PositionFromCenter();
        SizeChanged += (_, _) => { if (_initialPositionSet) RepositionToCenter(); };
        Load += (_, _) => RenderTransparent();
    }

    private void SetupOpaqueMode()
    {
        Padding = new Padding(1);

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

        // Note created indicator (purple checkmark)
        _noteCreatedIndicator = new Label
        {
            Text = "✓",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = StrokeBorderColor,
            BackColor = Color.Black,
            Size = new Size(18, 16),
            Visible = false,
            Cursor = Cursors.Default,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _noteCreatedIndicator.Location = new Point(dragBar.Width - 20, 0);
        dragBar.Controls.Add(_noteCreatedIndicator);

        // History fixed indicator (yellow checkmark)
        _historyFixedIndicator = new Label
        {
            Text = "✓",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = FixedTextColor,
            BackColor = Color.Black,
            Size = new Size(18, 16),
            Visible = false,
            Cursor = Cursors.Default,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _historyFixedIndicator.Location = new Point(dragBar.Width - 38, 0);
        dragBar.Controls.Add(_historyFixedIndicator);

        // Make form auto-size to content
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;

        // Context menu for drag bar
        var menu = new ContextMenuStrip();
        menu.Items.Add("Close", null, (_, _) => Close());
        dragBar.ContextMenuStrip = menu;

        // Context menu for content label
        var contentMenu = new ContextMenuStrip();
        contentMenu.Items.Add("Create Critical Note", null, (_, _) => _onCriticalNoteClick?.Invoke());
        contentMenu.Items.Add(new ToolStripSeparator());
        contentMenu.Items.Add("Copy Debug Info", null, (_, _) => CopyDebugInfoToClipboard());
        _contentLabel.ContextMenuStrip = contentMenu;

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
            _config.ClinicalHistoryX - Width / 2,
            _config.ClinicalHistoryY - Height / 2
        );
        _initialPositionSet = true;
    }

    private void RepositionToCenter()
    {
        if (_dragging) return;
        Location = new Point(
            _config.ClinicalHistoryX - Width / 2,
            _config.ClinicalHistoryY - Height / 2
        );
    }

    #endregion

    #region Transparent Mode - Rendering

    private void RequestRender()
    {
        if (!_useLayeredWindow || !IsHandleCreated || IsDisposed) return;

        var (fw, fh) = MeasureFormSize();
        if (fw != Width || fh != Height)
            Size = new Size(fw, fh);
        RenderTransparent();
    }

    private void RenderTransparent()
    {
        if (!IsHandleCreated || IsDisposed) return;
        int w = Width, h = Height;
        if (w <= 0 || h <= 0) return;

        Color borderColor = BackColor;
        Color innerBg = _contentLabel.BackColor;
        string text = _contentLabel.Text;
        Color textColor = _contentLabel.ForeColor;

        // Layer 1: semi-transparent background
        using var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.FromArgb(0, 0, 0, 0));
            int bw = TransparentBorderWidth;
            // Inner area - semi-transparent (draw first, border drawn on top)
            using var innerBrush = new SolidBrush(Color.FromArgb(_backgroundAlpha, innerBg.R, innerBg.G, innerBg.B));
            g.FillRectangle(innerBrush, bw, bw, w - bw * 2, h - bw * 2);
            // Border strips - fully opaque so state colors (green/purple/red) are always visible
            using var borderBrush = new SolidBrush(Color.FromArgb(255, borderColor.R, borderColor.G, borderColor.B));
            g.FillRectangle(borderBrush, 0, 0, w, bw);           // top
            g.FillRectangle(borderBrush, 0, h - bw, w, bw);      // bottom
            g.FillRectangle(borderBrush, 0, 0, bw, h);           // left
            g.FillRectangle(borderBrush, w - bw, 0, bw, h);      // right
        }

        // Layer 2: ClearType text on opaque inner background
        using var textLayer = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(textLayer))
        {
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.Clear(Color.FromArgb(255, innerBg.R, innerBg.G, innerBg.B));

            // Drag bar "⋯"
            using var dragFont = new Font("Segoe UI", 8);
            using var dragBrush = new SolidBrush(Color.FromArgb(102, 102, 102));
            using var dragSf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString("⋯", dragFont, dragBrush,
                new RectangleF(TransparentBorderWidth, TransparentBorderWidth, w - TransparentBorderWidth * 2, DragBarHeight), dragSf);

            // Checkmark indicators on drag bar (right side)
            using var checkFont = new Font("Segoe UI", 10, FontStyle.Bold);

            if (_noteCreated)
            {
                using var purpleBrush = new SolidBrush(StrokeBorderColor);
                g.DrawString("✓", checkFont, purpleBrush, w - 22, TransparentBorderWidth);
            }

            if (_historyFixInserted)
            {
                using var yellowBrush = new SolidBrush(FixedTextColor);
                int xOffset = _noteCreated ? 40 : 22;
                g.DrawString("✓", checkFont, yellowBrush, w - xOffset, TransparentBorderWidth);
            }

            // Content text
            float textX = TransparentBorderWidth + ContentMarginLeft;
            float textY = TransparentBorderWidth + DragBarHeight + ContentMarginTop;
            float maxWidth = w - TransparentBorderWidth * 2 - ContentMarginLeft - ContentMarginRight;
            if (maxWidth < 10) maxWidth = 10;

            using var sf = new StringFormat(StringFormat.GenericTypographic) { Trimming = StringTrimming.Word };
            using var brush = new SolidBrush(textColor);
            g.DrawString(text, _contentFont, brush,
                new RectangleF(textX, textY, maxWidth, h - textY), sf);
        }

        LayeredWindowHelper.MergeTextLayer(bmp, textLayer, innerBg.R, innerBg.G, innerBg.B);
        LayeredWindowHelper.PremultiplyBitmapAlpha(bmp);
        LayeredWindowHelper.SetBitmap(this, bmp);
    }

    private (int width, int height) MeasureFormSize()
    {
        string text = _contentLabel.Text;
        using var bmp = new Bitmap(1, 1);
        using var g = Graphics.FromImage(bmp);
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        using var sf = new StringFormat(StringFormat.GenericTypographic) { Trimming = StringTrimming.Word };

        // First measure single-line
        var measured = g.MeasureString(text, _contentFont, int.MaxValue, sf);

        int contentWidth = (int)Math.Ceiling(measured.Width);
        int contentHeight = (int)Math.Ceiling(measured.Height);

        // Cap to screen width
        var screen = Screen.FromPoint(new Point(_config.ClinicalHistoryX, _config.ClinicalHistoryY));
        int maxFormWidth = screen.WorkingArea.Width - 40;
        int maxContentWidth = maxFormWidth - TransparentBorderWidth * 2 - ContentMarginLeft - ContentMarginRight;

        if (contentWidth > maxContentWidth)
        {
            contentWidth = maxContentWidth;
            measured = g.MeasureString(text, _contentFont, contentWidth, sf);
            contentHeight = (int)Math.Ceiling(measured.Height);
        }

        contentWidth = Math.Max(contentWidth, 50);
        contentHeight = Math.Max(contentHeight, _contentFont.Height);

        int formWidth = TransparentBorderWidth * 2 + ContentMarginLeft + contentWidth + ContentMarginRight;
        int formHeight = TransparentBorderWidth * 2 + DragBarHeight + ContentMarginTop + contentHeight + ContentMarginBottom;

        return (Math.Max(formWidth, 80), Math.Max(formHeight, 40));
    }

    #endregion

    #region Transparent Mode - Mouse Handling

    private void OnTransparentMouseDown(object? sender, MouseEventArgs e)
    {
        bool inDragBar = e.Y < TransparentBorderWidth + DragBarHeight;

        if (e.Button == MouseButtons.Right)
        {
            if (inDragBar)
                _transparentDragMenu?.Show(this, e.Location);
            else
                _transparentContentMenu?.Show(this, e.Location);
            return;
        }

        if (e.Button == MouseButtons.Left)
        {
            if (inDragBar)
            {
                _formPosOnMouseDown = Location;
                LayeredWindowHelper.ReleaseCapture();
                LayeredWindowHelper.SendMessage(Handle, LayeredWindowHelper.WM_NCLBUTTONDOWN, LayeredWindowHelper.HT_CAPTION, 0);
                if (Location != _formPosOnMouseDown)
                {
                    _config.ClinicalHistoryX = Location.X + Width / 2;
                    _config.ClinicalHistoryY = Location.Y + Height / 2;
                    _config.Save();
                }
            }
            else
            {
                // Content click - same logic as OnContentLabelMouseDown
                HandleContentClick();
            }
        }
    }

    private void HandleContentClick()
    {
        // Ctrl+Click = create critical note
        if (ModifierKeys.HasFlag(Keys.Control))
        {
            Services.Logger.Trace("ClinicalHistory: Ctrl+Click - invoking critical note callback");
            _onCriticalNoteClick?.Invoke();
            return;
        }

        Services.Logger.Trace($"ClinicalHistory click: strokeDetected={_strokeDetected}, clickToCreate={_config.StrokeClickToCreateNote}, callbackSet={_onStrokeNoteClick != null}");

        // If stroke detected and click-to-create enabled, create note
        if (_strokeDetected && _config.StrokeClickToCreateNote && _onStrokeNoteClick != null)
        {
            Services.Logger.Trace("ClinicalHistory: Invoking stroke note callback");
            bool created = _onStrokeNoteClick.Invoke();
            Services.Logger.Trace($"ClinicalHistory: Callback returned created={created}");
            if (!created)
                ShowAlreadyCreatedTooltip();
            return;
        }

        // Default: paste clinical history
        PasteClinicalHistoryToMosaic();
    }

    #endregion

    #region Public Methods - State Updates

    /// <summary>
    /// Update the displayed clinical history text.
    /// </summary>
    public void SetClinicalHistory(string? text, bool wasFixed = false)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetClinicalHistory(text, wasFixed));
            return;
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            if (text == null && !string.IsNullOrWhiteSpace(_currentDisplayedText))
                return;

            _contentLabel.Text = "(No clinical history)";
            _contentLabel.ForeColor = EmptyTextColor;
            _currentDisplayedText = null;
            _currentTextWasFixed = false;
        }
        else
        {
            _contentLabel.Text = text;
            _currentDisplayedText = text;
            _currentTextWasFixed = wasFixed;
            _contentLabel.ForeColor = wasFixed ? FixedTextColor : NormalTextColor;
        }

        RequestRender();
    }

    /// <summary>
    /// Set clinical history with auto-fix support.
    /// </summary>
    public void SetClinicalHistoryWithAutoFix(string? preCleaned, string? cleaned, string? accession = null)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetClinicalHistoryWithAutoFix(preCleaned, cleaned, accession));
            return;
        }

        bool wasFixed = !string.IsNullOrWhiteSpace(preCleaned) &&
                        !string.IsNullOrWhiteSpace(cleaned) &&
                        !string.Equals(preCleaned, cleaned, StringComparison.Ordinal);

        SetClinicalHistory(cleaned, wasFixed);

        if (!_config.AutoFixClinicalHistory)
        {
            Logger.Trace("Auto-fix: disabled in config");
            _onAutoFixComplete?.Invoke();
            return;
        }

        if (string.IsNullOrWhiteSpace(cleaned))
        {
            Logger.Trace("Auto-fix: cleaned is empty");
            _onAutoFixComplete?.Invoke();
            return;
        }

        Logger.Trace($"Auto-fix check: wasFixed={wasFixed}, accession='{accession}', preCleaned='{preCleaned?.Substring(0, Math.Min(50, preCleaned?.Length ?? 0))}...'");

        if (!wasFixed)
        {
            _onAutoFixComplete?.Invoke();
            return;
        }

        if (_hasClinicalHistoryFixed != null)
        {
            if (_hasClinicalHistoryFixed(accession))
            {
                Logger.Trace($"Auto-fix: already fixed accession {accession} (session tracking)");
                _onAutoFixComplete?.Invoke();
                return;
            }
        }
        else
        {
            if (!string.IsNullOrEmpty(accession) && string.Equals(_lastAutoFixedAccession, accession, StringComparison.Ordinal))
            {
                Logger.Trace($"Auto-fix: already fixed accession {accession} (local tracking)");
                _onAutoFixComplete?.Invoke();
                return;
            }
        }

        if (_autoFixInProgress)
        {
            Logger.Trace($"Auto-fix: already in progress for '{accession}', skipping");
            return;
        }
        _autoFixInProgress = true;

        Logger.Trace($"Auto-fix: malformed clinical history detected for '{accession}', waiting to recheck...");
        _lastAutoFixedAccession = accession;
        var capturedAccession = accession;

        ThreadPool.QueueUserWorkItem(_ =>
        {
            try
            {
                Thread.Sleep(5000);
                bool stillFixed = false;
                if (IsDisposed || !IsHandleCreated) { _autoFixInProgress = false; return; }
                Invoke(() => { stillFixed = _currentTextWasFixed; });

                if (!stillFixed)
                {
                    Logger.Trace($"Auto-fix: Mosaic self-corrected for '{capturedAccession}', skipping paste");
                    _autoFixInProgress = false;
                    _onAutoFixComplete?.Invoke();
                    return;
                }

                Logger.Trace($"Auto-fix: still malformed after recheck for '{capturedAccession}', pasting fix");
                _lastAutoFixTime = DateTime.Now;
                _markClinicalHistoryFixed?.Invoke(capturedAccession);
                PasteClinicalHistoryToMosaic(showYellowCheckmark: true);
                _autoFixInProgress = false;
            }
            catch (Exception ex)
            {
                Logger.Trace($"Auto-fix recheck error: {ex.Message}");
                _autoFixInProgress = false;
            }
        });
    }

    public void ResetAutoFixTracking()
    {
        _lastAutoFixedAccession = null;
        _autoFixInProgress = false;
    }

    public void OnStudyChanged(bool isNewStudy = true)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => OnStudyChanged(isNewStudy));
            return;
        }

        Logger.Trace($"ClinicalHistoryForm: Study changed (isNewStudy={isNewStudy}) - resetting state");

        _lastAutoFixedAccession = null;
        _lastAutoFixTime = DateTime.MinValue;
        _templateMismatch = false;
        _strokeDetected = false;
        _noteCreated = false;
        _historyFixInserted = false;
        if (_noteCreatedIndicator != null)
            _noteCreatedIndicator.Visible = false;
        if (_historyFixedIndicator != null)
            _historyFixedIndicator.Visible = false;
        _lastDescription = null;
        _lastTemplateName = null;
        BackColor = NormalBorderColor;
        UpdateBorderTooltip();

        _contentLabel.Text = isNewStudy ? "(Loading...)" : "(No clinical history)";
        _contentLabel.ForeColor = EmptyTextColor;
        _contentLabel.BackColor = Color.Black;
        _currentDisplayedText = null;
        _currentTextWasFixed = false;

        RequestRender();
    }

    public void UpdateTextColorFromFinalReport(string? finalReportText)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => UpdateTextColorFromFinalReport(finalReportText));
            return;
        }

        if (string.IsNullOrWhiteSpace(_currentDisplayedText))
            return;

        if (!_currentTextWasFixed)
        {
            _contentLabel.ForeColor = NormalTextColor;
            RequestRender();
            return;
        }

        var reportHistory = ExtractClinicalHistoryFromReport(finalReportText);
        bool matches = CompareHistoryText(_currentDisplayedText, reportHistory);

        if (matches)
        {
            _contentLabel.ForeColor = NormalTextColor;
            Logger.Trace("Clinical history matches final report - text color: white");
        }
        else
        {
            _contentLabel.ForeColor = FixedTextColor;
        }

        RequestRender();
    }

    public void SetDraftedState(bool isDrafted)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetDraftedState(isDrafted));
            return;
        }

        if (_templateMismatch)
            BackColor = TemplateMismatchBorderColor;
        else if (_strokeDetected)
            BackColor = StrokeBorderColor;
        else
            BackColor = isDrafted ? DraftedBorderColor : NormalBorderColor;

        UpdateBorderTooltip();
        RequestRender();
    }

    public void SetTemplateMismatchState(bool isMismatch, string? description, string? templateName)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetTemplateMismatchState(isMismatch, description, templateName));
            return;
        }

        _templateMismatch = isMismatch;
        _lastDescription = description;
        _lastTemplateName = templateName;

        if (isMismatch)
        {
            BackColor = TemplateMismatchBorderColor;
            UpdateBorderTooltip();
        }

        RequestRender();
    }

    public void SetStrokeState(bool isStroke)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetStrokeState(isStroke));
            return;
        }

        _strokeDetected = isStroke;

        if (isStroke && !_templateMismatch && !_genderWarningActive)
            BackColor = StrokeBorderColor;

        UpdateBorderTooltip();
        RequestRender();
    }

    public bool IsStrokeDetected => _strokeDetected;

    public void SetNoteCreated(bool created)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetNoteCreated(created));
            return;
        }

        _noteCreated = created;
        if (_noteCreatedIndicator != null)
            _noteCreatedIndicator.Visible = created;

        UpdateBorderTooltip();
        RequestRender();
    }

    public bool IsNoteCreated => _noteCreated;

    public void SetHistoryFixInserted(bool inserted)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => SetHistoryFixInserted(inserted));
            return;
        }

        _historyFixInserted = inserted;
        if (_historyFixedIndicator != null)
            _historyFixedIndicator.Visible = inserted;

        UpdateBorderTooltip();
        RequestRender();
    }

    public bool IsHistoryFixInserted => _historyFixInserted;

    #endregion

    #region Callback Setters

    public void SetStrokeNoteClickCallback(Func<bool> callback)
    {
        _onStrokeNoteClick = callback;
    }

    public void SetCriticalNoteClickCallback(Action callback)
    {
        _onCriticalNoteClick = callback;
    }

    public void SetAddendumCheckCallback(Func<bool> callback)
    {
        _isAddendumOpen = callback;
    }

    public void SetClinicalHistoryFixCallbacks(Func<string?, bool> hasFixed, Action<string?> markFixed)
    {
        _hasClinicalHistoryFixed = hasFixed;
        _markClinicalHistoryFixed = markFixed;
    }

    public void SetAutoFixCompleteCallback(Action callback)
    {
        _onAutoFixComplete = callback;
    }

    #endregion

    #region Alert Mode

    public bool IsShowingAlert => _showingAlert;

    public void ShowAlertOnly(AlertType type, string details)
    {
        if (InvokeRequired)
        {
            BeginInvoke(() => ShowAlertOnly(type, details));
            return;
        }

        _showingAlert = true;
        _currentAlertType = type;

        switch (type)
        {
            case AlertType.TemplateMismatch:
                BackColor = TemplateMismatchBorderColor;
                _contentLabel.BackColor = Color.Black;
                _contentLabel.ForeColor = Color.FromArgb(255, 150, 150);
                _contentLabel.Text = $"TEMPLATE MISMATCH\n{details}";
                break;

            case AlertType.GenderMismatch:
                break;

            case AlertType.StrokeDetected:
                BackColor = StrokeBorderColor;
                _contentLabel.BackColor = Color.Black;
                _contentLabel.ForeColor = Color.FromArgb(200, 150, 255);
                _contentLabel.Text = $"STROKE PROTOCOL\n{details}";
                break;
        }

        UpdateBorderTooltip();
        RequestRender();
    }

    public void ClearAlert()
    {
        if (InvokeRequired)
        {
            BeginInvoke(ClearAlert);
            return;
        }

        _showingAlert = false;
        _currentAlertType = null;

        if (!string.IsNullOrEmpty(_currentDisplayedText))
        {
            _contentLabel.Text = _currentDisplayedText;
            _contentLabel.ForeColor = _currentTextWasFixed ? FixedTextColor : NormalTextColor;
            _contentLabel.BackColor = Color.Black;
        }
        else
        {
            _contentLabel.Text = "(No clinical history)";
            _contentLabel.ForeColor = EmptyTextColor;
            _contentLabel.BackColor = Color.Black;
        }

        BackColor = NormalBorderColor;
        UpdateBorderTooltip();
        RequestRender();
    }

    public string GetAlertText(AlertType type)
    {
        switch (type)
        {
            case AlertType.TemplateMismatch:
                if (_lastDescription != null && _lastTemplateName != null)
                    return $"Study: {_lastDescription}\nTemplate: {_lastTemplateName}";
                return "Study/template mismatch detected";

            case AlertType.StrokeDetected:
                return "Study flagged as stroke protocol";

            case AlertType.GenderMismatch:
                return _genderWarningText ?? "Gender mismatch detected";

            default:
                return "";
        }
    }

    #endregion

    #region Tooltip

    private void UpdateBorderTooltip()
    {
        string tooltip;

        if (_genderWarningActive)
            tooltip = "Red (flashing): Gender mismatch - report contains terms that don't match patient gender";
        else if (_templateMismatch)
            tooltip = "Red: Template mismatch - study description doesn't match the report template";
        else if (_strokeDetected)
        {
            if (_noteCreated)
                tooltip = "Purple: Stroke protocol - Critical Communication Note created ✓";
            else if (_config.StrokeClickToCreateNote)
                tooltip = "Purple: Stroke protocol - click to create critical note";
            else
                tooltip = "Purple: Stroke protocol detected";
        }
        else if (BackColor == DraftedBorderColor)
            tooltip = "Green: Report is drafted";
        else
            tooltip = "Gray: Normal - no alerts";

        _borderTooltip.SetToolTip(this, tooltip);

        if (!_useLayeredWindow)
        {
            _borderTooltip.SetToolTip(_contentLabel, tooltip);
            if (_noteCreatedIndicator != null)
                _borderTooltip.SetToolTip(_noteCreatedIndicator, "Critical Communication Note created in Clario");
            if (_historyFixedIndicator != null)
                _borderTooltip.SetToolTip(_historyFixedIndicator, "Corrected clinical history inserted to Mosaic");
        }
    }

    #endregion

    #region Content Click & Paste (Opaque Mode)

    private void OnContentLabelMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            HandleContentClick();
        }
    }

    #endregion

    #region Clinical History Paste

    private void ShowAlreadyCreatedTooltip()
    {
        _borderTooltip.Show("Critical note already created for this study", this,
            Width / 2, Height / 2, 2000);
    }

    private void PasteClinicalHistoryToMosaic(bool showYellowCheckmark = false)
    {
        var text = _contentLabel.Text;
        if (string.IsNullOrWhiteSpace(text) || text == "(No clinical history)")
        {
            ShowToast("No clinical history to paste");
            return;
        }

        if (!showYellowCheckmark && _currentTextWasFixed)
            showYellowCheckmark = true;

        if (_isAddendumOpen?.Invoke() == true)
        {
            ShowToast("Cannot paste into addendum");
            return;
        }

        var trimmedText = text.TrimEnd(' ', '.');
        var formatted = $"\nClinical history: {trimmedText}.\n";

        System.Threading.Tasks.Task.Run(() =>
        {
            lock (ActionController.PasteLock)
            {
                try
                {
                    NativeWindows.SavePreviousFocus();

                    Logger.Trace($"Clinical history paste: setting clipboard to '{formatted.Substring(0, Math.Min(50, formatted.Length))}...'");
                    var textToPaste = (_config.SeparatePastedItems && !formatted.StartsWith("\n")) ? "\n" + formatted : formatted;
                    if (IsDisposed || !IsHandleCreated) return;
                    Invoke(() => Clipboard.SetText(textToPaste));
                    System.Threading.Thread.Sleep(50);

                    if (!NativeWindows.ActivateMosaicForcefully())
                    {
                        if (!IsDisposed && IsHandleCreated)
                            BeginInvoke(() => ShowToast("Mosaic not found"));
                        return;
                    }

                    System.Threading.Thread.Sleep(200);
                    NativeWindows.SendHotkey("ctrl+v");
                    System.Threading.Thread.Sleep(100);

                    ActionController.LastPasteTime = DateTime.Now;
                    NativeWindows.RestorePreviousFocus();

                    if (showYellowCheckmark)
                    {
                        if (!IsDisposed && IsHandleCreated)
                        {
                            BeginInvoke(() => SetHistoryFixInserted(true));
                            BeginInvoke(() => ShowToast("Corrected history pasted"));
                        }
                        _onAutoFixComplete?.Invoke();
                    }
                    else
                    {
                        if (!IsDisposed && IsHandleCreated)
                            BeginInvoke(() => ShowToast("Pasted to Mosaic"));
                    }
                }
                catch (Exception ex)
                {
                    Logger.Trace($"PasteClinicalHistoryToMosaic error: {ex.Message}");
                    try { if (!IsDisposed && IsHandleCreated) BeginInvoke(() => ShowToast("Paste failed")); } catch { }
                }
            }
        });
    }

    #endregion

    #region Debug & Toast

    private void CopyDebugInfoToClipboard()
    {
        var debugInfo = $"=== Template Matching Debug ===\r\n" +
                        $"Description: {_lastDescription ?? "(none)"}\r\n" +
                        $"Template: {_lastTemplateName ?? "(none)"}\r\n" +
                        $"Mismatch: {_templateMismatch}";

        try
        {
            Clipboard.SetText(debugInfo);
            Services.Logger.Trace("Template debug info copied to clipboard");
            ShowToast("Debug copied!");
        }
        catch (Exception ex)
        {
            Services.Logger.Trace($"Failed to copy debug info: {ex.Message}");
        }
    }

    private void ShowToast(string message)
    {
        var toast = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            ShowInTaskbar = false,
            TopMost = true,
            BackColor = Color.FromArgb(51, 51, 51),
            Size = new Size(140, 30),
            StartPosition = FormStartPosition.Manual,
            Location = new Point(Location.X + Width / 2 - 70, Location.Y - 35)
        };

        var label = new Label
        {
            Text = message,
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

    #region Text Cleaning & Extraction

    private static string CleanText(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var sb = new System.Text.StringBuilder();
        foreach (char c in text)
        {
            if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) || c == '-' || c == '/' || c == '\'')
                sb.Append(c);
            else if (char.IsControl(c))
                sb.Append(' ');
        }

        var cleaned = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
        cleaned = Regex.Replace(cleaned, @"Other\s*\(Please Specify\)\s*;?\s*", " ", RegexOptions.IgnoreCase);
        cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();

        var words = cleaned.Split(' ');
        var dedupedWords = new System.Collections.Generic.List<string>();
        for (int i = 0; i < words.Length; i++)
        {
            if (i == 0 || !string.Equals(words[i].TrimEnd('.', ',', ';'), words[i - 1].TrimEnd('.', ',', ';'), StringComparison.OrdinalIgnoreCase))
                dedupedWords.Add(words[i]);
        }
        cleaned = string.Join(" ", dedupedWords);

        return RemoveRepeatingPhrases(cleaned);
    }

    private static string RemoveRepeatingPhrases(string text)
    {
        if (string.IsNullOrEmpty(text) || text.Length < 6)
            return text;

        string result = text;

        for (int patternLen = 1; patternLen <= (result.Length + 1) / 2; patternLen++)
        {
            string pattern = result.Substring(0, patternLen);
            var expected = new System.Text.StringBuilder();
            while (expected.Length < result.Length)
                expected.Append(pattern);

            string expectedStr = expected.ToString().Substring(0, result.Length);
            if (string.Equals(result, expectedStr, StringComparison.OrdinalIgnoreCase))
                return pattern.Trim();
        }

        bool foundDuplicate = true;
        while (foundDuplicate)
        {
            foundDuplicate = false;
            int maxPhraseLen = result.Length / 2;

            for (int phraseLen = maxPhraseLen; phraseLen >= 5; phraseLen--)
            {
                for (int startPos = 0; startPos <= result.Length - phraseLen * 2; startPos++)
                {
                    string phrase = result.Substring(startPos, phraseLen).Trim();
                    if (phrase.Length < 5) continue;

                    int afterFirstPhrase = startPos + phraseLen;
                    while (afterFirstPhrase < result.Length &&
                           (char.IsWhiteSpace(result[afterFirstPhrase]) || result[afterFirstPhrase] == ';' || result[afterFirstPhrase] == ','))
                        afterFirstPhrase++;

                    if (afterFirstPhrase + phrase.Length <= result.Length)
                    {
                        string nextChunk = result.Substring(afterFirstPhrase, phrase.Length);
                        if (string.Equals(phrase, nextChunk.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            result = result.Substring(0, startPos + phraseLen) + result.Substring(afterFirstPhrase + phrase.Length);
                            result = Regex.Replace(result, @"\s+", " ").Trim();
                            foundDuplicate = true;
                            break;
                        }
                    }
                }
                if (foundDuplicate) break;
            }
        }

        return result;
    }

    public static string? ExtractClinicalHistory(string? rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText))
        {
            Services.Logger.Trace("ExtractClinicalHistory: rawText is null or empty");
            return null;
        }

        Services.Logger.Trace($"ExtractClinicalHistory: Input length={rawText.Length}");

        if (!rawText.Contains("CLINICAL HISTORY", StringComparison.OrdinalIgnoreCase))
        {
            Services.Logger.Trace("ExtractClinicalHistory: 'CLINICAL HISTORY' not found in text");
            return null;
        }

        var sectionHeaders = @"TECHNIQUE|FINDINGS|IMPRESSION|COMPARISON|PROCEDURE|CONCLUSION|RECOMMENDATION";
        var match = Regex.Match(rawText,
            $@"CLINICAL HISTORY[:\s]*(.*?)(?=\b({sectionHeaders})\b|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!match.Success || match.Groups.Count < 2)
        {
            Services.Logger.Trace("ExtractClinicalHistory: Regex did not match");
            return null;
        }

        Services.Logger.Trace($"ExtractClinicalHistory: Regex matched, group1 length={match.Groups[1].Value.Length}");
        var content = match.Groups[1].Value.Trim();

        if (string.IsNullOrWhiteSpace(content))
            return string.Empty;

        var lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var uniqueLines = new System.Collections.Generic.List<string>();

        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
                continue;

            bool isDuplicate = false;
            foreach (var existing in uniqueLines)
            {
                if (string.Equals(existing, trimmed, StringComparison.OrdinalIgnoreCase))
                {
                    isDuplicate = true;
                    break;
                }
            }

            if (!isDuplicate)
                uniqueLines.Add(trimmed);
        }

        if (uniqueLines.Count == 0)
            return null;

        return CleanText(string.Join(" ", uniqueLines));
    }

    public static (string? preCleaned, string? cleaned) ExtractClinicalHistoryWithFixInfo(string? rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText))
            return (null, null);

        if (!rawText.Contains("CLINICAL HISTORY", StringComparison.OrdinalIgnoreCase))
            return (null, null);

        var sectionHeaders = @"TECHNIQUE|FINDINGS|IMPRESSION|COMPARISON|PROCEDURE|CONCLUSION|RECOMMENDATION";
        var match = Regex.Match(rawText,
            $@"CLINICAL HISTORY[:\s]*(.*?)(?=\b({sectionHeaders})\b|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!match.Success || match.Groups.Count < 2)
            return (null, null);

        var content = match.Groups[1].Value.Trim();

        if (string.IsNullOrWhiteSpace(content))
            return (string.Empty, string.Empty);

        var allLines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var allTrimmedLines = new System.Collections.Generic.List<string>();
        foreach (var line in allLines)
        {
            var trimmed = line.Trim();
            if (!string.IsNullOrWhiteSpace(trimmed))
                allTrimmedLines.Add(trimmed);
        }

        if (allTrimmedLines.Count == 0)
            return (string.Empty, string.Empty);

        var preCleanedRaw = string.Join(" ", allTrimmedLines);
        var sb = new System.Text.StringBuilder();
        foreach (char c in preCleanedRaw)
        {
            if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) || c == '-' || c == '/' || c == '\'')
                sb.Append(c);
            else if (char.IsControl(c))
                sb.Append(' ');
        }
        var preCleaned = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();

        var uniqueLines = new System.Collections.Generic.List<string>();
        foreach (var line in allTrimmedLines)
        {
            bool isDuplicate = false;
            foreach (var existing in uniqueLines)
            {
                if (string.Equals(existing, line, StringComparison.OrdinalIgnoreCase))
                {
                    isDuplicate = true;
                    break;
                }
            }
            if (!isDuplicate)
                uniqueLines.Add(line);
        }

        var cleaned = CleanText(string.Join(" ", uniqueLines));
        return (preCleaned, cleaned);
    }

    private static string? ExtractClinicalHistoryFromReport(string? reportText)
    {
        if (string.IsNullOrWhiteSpace(reportText))
            return null;

        var match = Regex.Match(reportText,
            @"CLINICAL HISTORY[:\s]*(.*?)(?=\b(?:TECHNIQUE|FINDINGS|IMPRESSION|COMPARISON|PROCEDURE|CONCLUSION|RECOMMENDATION)\b|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!match.Success || match.Groups.Count < 2)
            return null;

        return match.Groups[1].Value.Trim();
    }

    private static bool CompareHistoryText(string? text1, string? text2)
    {
        if (string.IsNullOrWhiteSpace(text1) || string.IsNullOrWhiteSpace(text2))
            return false;

        var norm1 = NormalizeForComparison(text1);
        var norm2 = NormalizeForComparison(text2);

        return string.Equals(norm1, norm2, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeForComparison(string text)
    {
        var result = Regex.Replace(text, @"\s+", " ").Trim();
        result = result.TrimEnd('.', ',', ';', ' ');
        return result;
    }

    #endregion

    #region EnsureOnTop

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
        _config.ClinicalHistoryX = Location.X + Width / 2;
        _config.ClinicalHistoryY = Location.Y + Height / 2;
        _config.Save();
    }

    #endregion

    #region Gender Check

    public static List<string> CheckGenderMismatch(string? reportText, string? patientGender)
    {
        var mismatches = new List<string>();

        if (string.IsNullOrWhiteSpace(reportText) || string.IsNullOrWhiteSpace(patientGender))
            return mismatches;

        var reportLower = reportText.ToLowerInvariant();
        var genderUpper = patientGender.ToUpperInvariant();

        if (genderUpper == "MALE")
        {
            foreach (var term in FemaleOnlyTerms)
            {
                if (Regex.IsMatch(reportLower, $@"\b{Regex.Escape(term)}\b"))
                    mismatches.Add(term);
            }
        }
        else if (genderUpper == "FEMALE")
        {
            foreach (var term in MaleOnlyTerms)
            {
                if (Regex.IsMatch(reportLower, $@"\b{Regex.Escape(term)}\b"))
                    mismatches.Add(term);
            }
        }

        return mismatches;
    }

    public void SetGenderWarning(bool active, string? patientGender, List<string>? mismatchedTerms)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetGenderWarning(active, patientGender, mismatchedTerms));
            return;
        }

        if (active && mismatchedTerms != null && mismatchedTerms.Count > 0)
        {
            if (!_genderWarningActive)
                _savedClinicalHistoryText = _contentLabel.Text;

            _genderWarningActive = true;
            _genderWarningText = $"GENDER MISMATCH!\nPatient: {patientGender}\nTerms: {string.Join(", ", mismatchedTerms)}";

            StartBlinking();
            UpdateBorderTooltip();
            RequestRender();
        }
        else
        {
            if (_genderWarningActive)
            {
                _genderWarningActive = false;
                StopBlinking();

                _contentLabel.BackColor = Color.Black;

                if (!string.IsNullOrEmpty(_currentDisplayedText))
                {
                    _contentLabel.Text = _currentDisplayedText;
                    _contentLabel.ForeColor = _currentTextWasFixed ? FixedTextColor : NormalTextColor;
                }
                else if (!string.IsNullOrEmpty(_savedClinicalHistoryText))
                {
                    _contentLabel.Text = _savedClinicalHistoryText;
                    _contentLabel.ForeColor = NormalTextColor;
                }

                if (!_templateMismatch)
                    BackColor = NormalBorderColor;

                UpdateBorderTooltip();
                _savedClinicalHistoryText = null;
                RequestRender();
            }
        }
    }

    private void StartBlinking()
    {
        if (_blinkTimer != null) return;

        _blinkState = true;
        UpdateBlinkDisplay();

        _blinkTimer = new System.Windows.Forms.Timer { Interval = 500 };
        _blinkTimer.Tick += (_, _) =>
        {
            _blinkState = !_blinkState;
            UpdateBlinkDisplay();
        };
        _blinkTimer.Start();
    }

    private void StopBlinking()
    {
        if (_blinkTimer != null)
        {
            _blinkTimer.Stop();
            _blinkTimer.Dispose();
            _blinkTimer = null;
        }
    }

    private void UpdateBlinkDisplay()
    {
        if (!_genderWarningActive) return;

        if (_blinkState)
        {
            BackColor = Color.FromArgb(220, 0, 0);
            _contentLabel.BackColor = Color.FromArgb(180, 0, 0);
            _contentLabel.ForeColor = Color.White;
            _contentLabel.Text = _genderWarningText ?? "GENDER MISMATCH!";
        }
        else
        {
            BackColor = Color.FromArgb(120, 0, 0);
            _contentLabel.BackColor = Color.FromArgb(80, 0, 0);
            _contentLabel.ForeColor = Color.FromArgb(255, 200, 200);
            _contentLabel.Text = _genderWarningText ?? "GENDER MISMATCH!";
        }

        RequestRender();
    }

    public bool IsGenderWarningActive => _genderWarningActive;

    #endregion

    #region Disposal

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _contentFont?.Dispose();
            _blinkTimer?.Stop();
            _blinkTimer?.Dispose();
            _transparentDragMenu?.Dispose();
            _transparentContentMenu?.Dispose();
        }
        base.Dispose(disposing);
    }

    #endregion
}
