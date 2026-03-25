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

    // Accumulated final results for full re-render (RichTextBox loses per-char colors on edit)
    private readonly List<FinalSegment> _finalSegments = new();
    private record FinalSegment(string[] Words, Color[] Colors);

    // Auto-grow
    private int _maxHeight;
    private bool _isAutoGrowing;
    private const int InitialHeight = 34;
    private const int TextPadding = 4;

    // Position save debounce
    private System.Threading.Timer? _saveTimer;

    // Delayed clear — keeps final text visible briefly after paste
    private System.Threading.Timer? _clearTimer;
    private const int ClearDelayMs = 500;

    // Delayed hide — keeps form visible briefly after recording stops so user can verify
    private System.Threading.Timer? _hideTimer;
    private const int LingerDelayMs = 2000;

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
            Font = new Font("Segoe UI", 12.5f),
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

        _clearTimer = new System.Threading.Timer(_ =>
        {
            try { BeginInvoke(DoClearTranscript); } catch { }
        }, null, Timeout.Infinite, Timeout.Infinite);

        _hideTimer = new System.Threading.Timer(_ =>
        {
            try { BeginInvoke(DoDelayedHide); } catch { }
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
    /// Append a transcription result. Shows interim text in gray italic, speech-final in white.
    /// Only clears on SpeechFinal (utterance end). IsFinal-only results are paste-only (no display).
    /// </summary>
    public void AppendResult(SttResult result)
    {
        if (InvokeRequired) { BeginInvoke(() => AppendResult(result)); return; }

        // Cancel any pending delayed clear/hide — new text takes priority
        _clearTimer?.Change(Timeout.Infinite, Timeout.Infinite);
        _hideTimer?.Change(Timeout.Infinite, Timeout.Infinite);

        // Temporarily allow edits — ReadOnly RichTextBox produces system beeps on modification
        _textBox.ReadOnly = false;

        if (result.SpeechFinal)
        {
            // Build word list and color list for this segment
            HashSet<int>? mediumSet = result.MediumConfWordIndices is { Length: > 0 }
                ? new HashSet<int>(result.MediumConfWordIndices) : null;
            HashSet<int>? lowSet = result.LowConfWordIndices is { Length: > 0 }
                ? new HashSet<int>(result.LowConfWordIndices) : null;

            if (result.Words is { Length: > 0 })
            {
                var transcriptWords = result.Transcript.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                bool useTranscript = transcriptWords.Length == result.Words.Length;
                var words = new string[result.Words.Length];
                var colors = new Color[result.Words.Length];

                for (int i = 0; i < result.Words.Length; i++)
                {
                    var w = result.Words[i];
                    words[i] = useTranscript ? transcriptWords[i] : (w.PunctuatedText ?? w.Text);

                    if (lowSet != null && lowSet.Contains(i))
                        colors[i] = ConfidenceColor(0.50f);
                    else if (mediumSet != null && mediumSet.Contains(i))
                        colors[i] = ConfidenceColor(0.75f);
                    else if (mediumSet != null || lowSet != null)
                        colors[i] = ConfidenceColor(0.98f);
                    else
                        colors[i] = ConfidenceColor(w.Confidence);
                }
                _finalSegments.Add(new FinalSegment(words, colors));
            }
            else if (!string.IsNullOrEmpty(result.Transcript))
            {
                var words = result.Transcript.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                var defaultColor = Color.FromArgb(220, 220, 220);
                var colors = new Color[words.Length];
                Array.Fill(colors, defaultColor);
                _finalSegments.Add(new FinalSegment(words, colors));
            }

            _interimText = "";
            RebuildDisplay(null);
        }
        else if (result.IsFinal)
        {
            // IsFinal but not SpeechFinal: paste-only result (e.g. Soniox token batches).
            // Skip display update — the interim handles the overlay text.
        }
        else
        {
            // Interim: full rebuild with trailing interim text in gray italic
            _interimText = result.Transcript;
            RebuildDisplay(_interimText);
        }

        _textBox.ReadOnly = true;
    }

    /// <summary>
    /// Build RTF string directly and assign to RichTextBox.Rtf.
    /// Bypasses SelectionColor API which loses per-character colors on rapid updates.
    /// </summary>
    private void RebuildDisplay(string? interimText)
    {
        var sb = new System.Text.StringBuilder(512);

        // RTF header with color table
        // Collect unique colors first
        var colorList = new List<Color>();
        var colorIndex = new Dictionary<int, int>(); // argb -> 1-based index in color table

        void EnsureColor(Color c)
        {
            int argb = c.ToArgb();
            if (!colorIndex.ContainsKey(argb))
            {
                colorList.Add(c);
                colorIndex[argb] = colorList.Count; // 1-based (0 = default)
            }
        }

        // Pre-register all used colors
        var defaultColor = Color.FromArgb(220, 220, 220);
        var grayColor = Color.FromArgb(128, 128, 128);
        EnsureColor(defaultColor);
        EnsureColor(grayColor);
        foreach (var seg in _finalSegments)
            foreach (var c in seg.Colors)
                EnsureColor(c);

        // Font size in half-points
        int fontSize = (int)(_textBox.Font.Size * 2);

        sb.Append(@"{\rtf1\ansi\deff0");
        // Font table
        sb.Append(@"{\fonttbl{\f0 ").Append(_textBox.Font.Name).Append(@";}}");
        // Color table (entry 0 is auto/default, then our colors)
        sb.Append(@"{\colortbl ;");
        foreach (var c in colorList)
            sb.Append($@"\red{c.R}\green{c.G}\blue{c.B};");
        sb.Append('}');

        // Background color
        sb.Append($@"\viewkind4\f0\fs{fontSize}\cb0 ");

        // Render final segments
        for (int s = 0; s < _finalSegments.Count; s++)
        {
            if (s > 0)
                sb.Append($@"\cf{colorIndex[defaultColor.ToArgb()]}  ");

            var seg = _finalSegments[s];
            for (int i = 0; i < seg.Words.Length; i++)
            {
                int ci = colorIndex[seg.Colors[i].ToArgb()];
                sb.Append($@"\cf{ci} ");
                AppendRtfEscaped(sb, seg.Words[i]);
                sb.Append(' ');
            }
        }

        // Track where final text ends (approximate char count for _insertionPoint)
        int finalCharCount = 0;
        foreach (var seg in _finalSegments)
        {
            if (finalCharCount > 0) finalCharCount++; // space separator
            foreach (var w in seg.Words)
                finalCharCount += w.Length + 1; // word + space
        }
        _insertionPoint = finalCharCount;

        // Interim text in gray italic
        if (!string.IsNullOrEmpty(interimText))
        {
            int gi = colorIndex[grayColor.ToArgb()];
            sb.Append($@"\cf{gi}\i ");
            AppendRtfEscaped(sb, interimText);
            sb.Append(@"\i0 ");
        }

        sb.Append('}');

        _textBox.Rtf = sb.ToString();
        _textBox.SelectionStart = _textBox.TextLength;
        AutoGrowHeight();
    }

    private static void AppendRtfEscaped(System.Text.StringBuilder sb, string text)
    {
        foreach (char c in text)
        {
            if (c == '\\') sb.Append(@"\\");
            else if (c == '{') sb.Append(@"\{");
            else if (c == '}') sb.Append(@"\}");
            else if (c > 127) sb.Append($@"\u{(int)c}?");
            else sb.Append(c);
        }
    }

    // Kept for compatibility but no longer used by AppendResult
    private void RemoveInterimText() { }

    public void ClearTranscript()
    {
        if (InvokeRequired) { BeginInvoke(ClearTranscript); return; }
        // Schedule clear after delay so final text stays visible briefly
        _clearTimer?.Change(ClearDelayMs, Timeout.Infinite);
    }

    private void DoClearTranscript()
    {
        _textBox.ReadOnly = false;
        _textBox.Clear();
        _insertionPoint = 0;
        _interimText = "";
        _finalSegments.Clear();
        _textBox.ReadOnly = true;
        AutoGrowHeight();
    }

    /// <summary>
    /// Hide the form after a delay, keeping final text visible so user can verify.
    /// Cancelled automatically if new text arrives or a new recording starts.
    /// </summary>
    public void DelayedHide()
    {
        if (InvokeRequired) { BeginInvoke(DelayedHide); return; }
        _hideTimer?.Change(LingerDelayMs, Timeout.Infinite);
    }

    /// <summary>Cancel any pending delayed hide (e.g. new recording starting).</summary>
    public void CancelDelayedHide()
    {
        _hideTimer?.Change(Timeout.Infinite, Timeout.Infinite);
    }

    private void DoDelayedHide()
    {
        Hide();
        DoClearTranscript();
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
        _hideTimer?.Dispose();
        _clearTimer?.Dispose();
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

    /// <summary>
    /// Maps word confidence to a subtle color gradient.
    /// High confidence = white, medium = soft amber, low = muted orange.
    /// </summary>
    private static Color ConfidenceColor(float confidence) => confidence switch
    {
        >= 0.95f => Color.FromArgb(220, 220, 220),     // white — high confidence
        >= 0.80f => Color.FromArgb(220, 200, 150),     // soft amber — medium
        >= 0.60f => Color.FromArgb(220, 170, 100),     // muted orange — low
        _        => Color.FromArgb(210, 130, 80)        // deeper orange — very low
    };

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
