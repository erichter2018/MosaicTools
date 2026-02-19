// [CustomSTT] Minimal floating indicator showing live dictation text.
// Auto-clears after each paste, auto-hides when recording stops.
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Minimal dark-themed floating indicator for live STT dictation.
/// Shows interim/final text, clears after each paste, draggable, remembers position.
/// </summary>
public class TranscriptionForm : Form
{
    private readonly Configuration _config;
    private readonly RichTextBox _textBox;

    // Drag state
    private Point _dragStart;
    private bool _dragging;

    // Text tracking
    private int _insertionPoint;
    private string _interimText = "";

    // Auto-grow
    private int _maxHeight;
    private bool _isAutoGrowing;
    private const int InitialHeight = 40;
    private const int TextPadding = 6;

    // Position save debounce
    private System.Threading.Timer? _saveTimer;

    // Don't steal focus from Mosaic when the form is shown
    protected override bool ShowWithoutActivation => true;

    public TranscriptionForm(Configuration config)
    {
        _config = config;
        _maxHeight = Math.Max(80, config.TranscriptionFormHeight);

        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        DoubleBuffered = true;
        Text = "MosaicToolsTranscription";
        BackColor = Color.FromArgb(30, 30, 30);
        Opacity = 0.9;
        MinimumSize = new Size(200, InitialHeight);
        Size = new Size(Math.Max(200, config.TranscriptionFormWidth), InitialHeight);
        StartPosition = FormStartPosition.Manual;
        Cursor = Cursors.SizeAll;
        Padding = new Padding(6);

        Location = ScreenHelper.EnsureOnScreen(config.TranscriptionFormX, config.TranscriptionFormY);

        // Text display — read-only, no scrollbars, minimal
        _textBox = new RichTextBox
        {
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.FromArgb(200, 200, 200),
            Font = new Font("Segoe UI", 11),
            BorderStyle = BorderStyle.None,
            Dock = DockStyle.Fill,
            ReadOnly = true,
            ScrollBars = RichTextBoxScrollBars.None,
            DetectUrls = false,
            WordWrap = true,
            Cursor = Cursors.SizeAll,
            TabStop = false
        };

        // Forward mouse events from textbox to form for dragging
        _textBox.MouseDown += OnDragStart;
        _textBox.MouseMove += OnDragMove;
        _textBox.MouseUp += OnDragEnd;

        Controls.Add(_textBox);

        // Drag on form background too
        MouseDown += OnDragStart;
        MouseMove += OnDragMove;
        MouseUp += OnDragEnd;

        _saveTimer = new System.Threading.Timer(_ =>
        {
            try { BeginInvoke(SavePositionToConfig); } catch { }
        }, null, Timeout.Infinite, Timeout.Infinite);

        LocationChanged += (_, _) => _saveTimer?.Change(500, Timeout.Infinite);
        SizeChanged += (_, _) => _saveTimer?.Change(500, Timeout.Infinite);
    }

    private void SavePositionToConfig()
    {
        if (WindowState != FormWindowState.Normal) return;
        _config.TranscriptionFormX = Location.X;
        _config.TranscriptionFormY = Location.Y;
        _config.TranscriptionFormWidth = Width;
        _config.Save();
    }

    private void AutoGrowHeight()
    {
        if (_isAutoGrowing) return;
        _isAutoGrowing = true;
        try
        {
            if (_textBox.TextLength == 0)
            {
                Height = InitialHeight;
                return;
            }
            var lastCharPos = _textBox.GetPositionFromCharIndex(Math.Max(0, _textBox.TextLength - 1));
            int contentHeight = lastCharPos.Y + _textBox.Font.Height + TextPadding * 2 + (int)Padding.Vertical;
            int targetHeight = Math.Clamp(contentHeight, InitialHeight, _maxHeight);
            if (targetHeight != Height)
                Height = targetHeight;
        }
        catch { }
        finally { _isAutoGrowing = false; }
    }

    /// <summary>
    /// Append a transcription result. Shows interim text in gray, final in white.
    /// Clears previous text on each new final result (since it's already pasted to Mosaic).
    /// </summary>
    public void AppendResult(SttResult result)
    {
        if (InvokeRequired) { BeginInvoke(() => AppendResult(result)); return; }

        // Temporarily allow edits — ReadOnly RichTextBox produces system beeps on modification
        _textBox.ReadOnly = false;

        if (result.IsFinal)
        {
            // Clear and show the final text briefly before it disappears
            _textBox.Clear();
            _insertionPoint = 0;
            _interimText = "";

            _textBox.SelectionStart = 0;
            _textBox.SelectionColor = Color.White;
            _textBox.SelectionFont = new Font(_textBox.Font, FontStyle.Regular);
            _textBox.SelectedText = result.Transcript;
            _insertionPoint = _textBox.TextLength;

            _textBox.SelectionStart = _textBox.TextLength;
            AutoGrowHeight();
        }
        else
        {
            // Show interim text in gray italic — replaces previous interim
            RemoveInterimText();

            _interimText = result.Transcript;
            _textBox.SelectionStart = _insertionPoint;
            _textBox.SelectionLength = 0;
            _textBox.SelectionColor = Color.FromArgb(128, 128, 128);
            _textBox.SelectionFont = new Font(_textBox.Font, FontStyle.Italic);
            _textBox.SelectedText = _interimText;

            _textBox.SelectionStart = _textBox.TextLength;
            AutoGrowHeight();
        }

        _textBox.ReadOnly = true;
    }

    private void RemoveInterimText()
    {
        if (_interimText.Length > 0)
        {
            _textBox.ReadOnly = false;
            if (_insertionPoint + _interimText.Length <= _textBox.TextLength)
            {
                _textBox.SelectionStart = _insertionPoint;
                _textBox.SelectionLength = _interimText.Length;
                _textBox.SelectedText = "";
            }
            _interimText = "";
            _textBox.ReadOnly = true;
        }
    }

    public void ClearTranscript()
    {
        if (InvokeRequired) { BeginInvoke(ClearTranscript); return; }
        _textBox.ReadOnly = false;
        _textBox.Clear();
        _insertionPoint = 0;
        _interimText = "";
        _textBox.ReadOnly = true;
        AutoGrowHeight();
    }

    public void SetRecordingState(bool recording)
    {
        if (InvokeRequired) { BeginInvoke(() => SetRecordingState(recording)); return; }
        // No-op — recording state is handled by ActionController showing/hiding the form
    }

    // Keep as no-op for compatibility with MainForm calls
    public void UpdateStatus(string status) { }
    public string GetTranscriptText() => "";
    public void OnRecordingStarted() { }

    public void EnsureOnTop()
    {
        if (!IsDisposed && Visible)
            NativeWindows.ForceTopMost(Handle);
    }

    protected override void OnVisibleChanged(EventArgs e)
    {
        base.OnVisibleChanged(e);
        if (!Visible) SavePositionToConfig();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _config.TranscriptionFormHeight = _maxHeight;
        SavePositionToConfig();
        _saveTimer?.Dispose();
        base.OnFormClosing(e);
    }

    protected override void WndProc(ref Message m)
    {
        const int WM_NCHITTEST = 0x84;
        const int HTRIGHT = 11;
        const int HTLEFT = 10;

        // Allow horizontal resize from edges
        if (m.Msg == WM_NCHITTEST)
        {
            base.WndProc(ref m);
            var pt = PointToClient(new Point(m.LParam.ToInt32() & 0xFFFF, m.LParam.ToInt32() >> 16));
            if (pt.X <= 6) m.Result = (IntPtr)HTLEFT;
            else if (pt.X >= ClientSize.Width - 6) m.Result = (IntPtr)HTRIGHT;
            return;
        }

        const int WM_EXITSIZEMOVE = 0x232;
        if (m.Msg == WM_EXITSIZEMOVE && !_isAutoGrowing)
        {
            _maxHeight = Math.Max(_maxHeight, Height);
            _config.TranscriptionFormHeight = _maxHeight;
        }

        base.WndProc(ref m);
    }

    #region Drag Support

    private void OnDragStart(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left) { _dragging = true; _dragStart = e.Location; }
    }

    private void OnDragMove(object? sender, MouseEventArgs e)
    {
        if (_dragging)
        {
            var screen = (sender as Control)?.PointToScreen(e.Location) ?? Cursor.Position;
            Location = new Point(screen.X - _dragStart.X, screen.Y - _dragStart.Y);
        }
    }

    private void OnDragEnd(object? sender, MouseEventArgs e) => _dragging = false;

    #endregion
}
