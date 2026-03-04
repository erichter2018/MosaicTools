// [CustomSTT] Live ensemble metrics popup — draggable, always-on-top stats display
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Small draggable overlay showing live ensemble STT statistics.
/// Shows Deepgram confidence, per-provider corrections/confirms, last correction detail.
/// </summary>
public class EnsembleMetricsForm : Form
{
    private EnsembleStats? _lastStats;
    private bool _recording;
    private Point _dragOffset;
    private bool _dragging;

    // Top section
    private readonly Label _titleLabel;
    private readonly Panel _confidenceBar;
    private readonly Panel _confidenceBarBg;
    private readonly Label _confidenceLabel;
    private readonly Label _wordsLabel;
    private readonly Label _lowConfLabel;
    private readonly Label _correctedLabel;

    // Provider sections
    private readonly Label _aaiHeader;
    private readonly Label _aaiFinalsLabel;
    private readonly Label _aaiFixedLabel;
    private readonly Label _aaiConfirmedLabel;
    private readonly Panel _aaiBar;
    private readonly Panel _aaiBarBg;

    private readonly Label _smHeader;
    private readonly Label _smFinalsLabel;
    private readonly Label _smFixedLabel;
    private readonly Label _smConfirmedLabel;
    private readonly Panel _smBar;
    private readonly Panel _smBarBg;

    // Bottom
    private readonly Label _consensusLabel;
    private readonly Label _lastCorrLabel;

    private static readonly Color DimColor = Color.FromArgb(120, 120, 130);
    private static readonly Color BrightColor = Color.FromArgb(210, 210, 220);
    private static readonly Color AaiColor = Color.FromArgb(100, 180, 255);  // blue
    private static readonly Color SmColor = Color.FromArgb(180, 130, 255);   // purple
    private static readonly Color FixColor = Color.FromArgb(100, 220, 100);  // green
    private static readonly Color ConfirmColor = Color.FromArgb(200, 200, 100); // yellow-ish

    public EnsembleMetricsForm()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;
        Size = new Size(260, 270);
        BackColor = Color.FromArgb(22, 22, 28);
        Opacity = 0.93;

        // Position at bottom-left above taskbar
        var screen = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
        Location = new Point(screen.Left + 10, screen.Bottom - Height - 10);

        var bodyFont = new Font("Segoe UI", 8f);
        var smallFont = new Font("Segoe UI", 7.5f);
        var tinyFont = new Font("Segoe UI", 7f);

        int y = 6;
        int w = Width;
        int pad = 10;
        int contentW = w - pad * 2;

        // ── Title + close ──
        _titleLabel = new Label
        {
            Text = "ENSEMBLE",
            Location = new Point(pad, y),
            AutoSize = true,
            Font = new Font("Segoe UI", 7f, FontStyle.Bold),
            ForeColor = Color.FromArgb(100, 160, 255),
        };
        Controls.Add(_titleLabel);

        var closeBtn = new Label
        {
            Text = "\u00d7",
            Location = new Point(w - 22, y - 2),
            AutoSize = true,
            Font = new Font("Segoe UI", 10f),
            ForeColor = Color.FromArgb(80, 80, 90),
            Cursor = Cursors.Hand
        };
        closeBtn.Click += (_, _) => Hide();
        closeBtn.MouseEnter += (_, _) => closeBtn.ForeColor = Color.IndianRed;
        closeBtn.MouseLeave += (_, _) => closeBtn.ForeColor = Color.FromArgb(80, 80, 90);
        Controls.Add(closeBtn);
        y += 18;

        // ── Confidence bar ──
        _confidenceBarBg = new Panel
        {
            Location = new Point(pad, y),
            Size = new Size(contentW, 8),
            BackColor = Color.FromArgb(40, 40, 48)
        };
        Controls.Add(_confidenceBarBg);

        _confidenceBar = new Panel
        {
            Location = new Point(0, 0),
            Size = new Size(0, 8),
            BackColor = Color.FromArgb(80, 160, 255)
        };
        _confidenceBarBg.Controls.Add(_confidenceBar);
        y += 12;

        // ── Confidence + words row ──
        _confidenceLabel = new Label
        {
            Text = "Conf: --",
            Location = new Point(pad, y),
            AutoSize = true,
            Font = bodyFont,
            ForeColor = BrightColor
        };
        Controls.Add(_confidenceLabel);

        _wordsLabel = new Label
        {
            Text = "Words: 0",
            Location = new Point(pad + 120, y),
            AutoSize = true,
            Font = bodyFont,
            ForeColor = DimColor
        };
        Controls.Add(_wordsLabel);
        y += 16;

        // ── Low conf + corrected row ──
        _lowConfLabel = new Label
        {
            Text = "Low conf: 0",
            Location = new Point(pad, y),
            AutoSize = true,
            Font = bodyFont,
            ForeColor = DimColor
        };
        Controls.Add(_lowConfLabel);

        _correctedLabel = new Label
        {
            Text = "Corrected: 0",
            Location = new Point(pad + 120, y),
            AutoSize = true,
            Font = bodyFont,
            ForeColor = DimColor
        };
        Controls.Add(_correctedLabel);
        y += 20;

        // ── Divider ──
        var div1 = new Panel
        {
            Location = new Point(pad, y),
            Size = new Size(contentW, 1),
            BackColor = Color.FromArgb(45, 45, 55)
        };
        Controls.Add(div1);
        y += 5;

        // ── AssemblyAI section ──
        _aaiHeader = new Label
        {
            Text = "AAI",
            Location = new Point(pad, y),
            AutoSize = true,
            Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
            ForeColor = AaiColor
        };
        Controls.Add(_aaiHeader);

        _aaiFinalsLabel = new Label
        {
            Text = "0 finals",
            Location = new Point(pad + 35, y),
            AutoSize = true,
            Font = smallFont,
            ForeColor = DimColor
        };
        Controls.Add(_aaiFinalsLabel);
        y += 15;

        // AAI contribution bar background
        _aaiBarBg = new Panel
        {
            Location = new Point(pad, y),
            Size = new Size(contentW, 6),
            BackColor = Color.FromArgb(35, 35, 45)
        };
        Controls.Add(_aaiBarBg);

        _aaiBar = new Panel
        {
            Location = new Point(0, 0),
            Size = new Size(0, 6),
            BackColor = AaiColor
        };
        _aaiBarBg.Controls.Add(_aaiBar);
        y += 9;

        _aaiFixedLabel = new Label
        {
            Text = "0 fixed",
            Location = new Point(pad + 4, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = FixColor
        };
        Controls.Add(_aaiFixedLabel);

        _aaiConfirmedLabel = new Label
        {
            Text = "0 confirmed",
            Location = new Point(pad + 80, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = ConfirmColor
        };
        Controls.Add(_aaiConfirmedLabel);
        y += 17;

        // ── Speechmatics section ──
        _smHeader = new Label
        {
            Text = "SM",
            Location = new Point(pad, y),
            AutoSize = true,
            Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
            ForeColor = SmColor
        };
        Controls.Add(_smHeader);

        _smFinalsLabel = new Label
        {
            Text = "0 finals",
            Location = new Point(pad + 35, y),
            AutoSize = true,
            Font = smallFont,
            ForeColor = DimColor
        };
        Controls.Add(_smFinalsLabel);
        y += 15;

        // SM contribution bar background
        _smBarBg = new Panel
        {
            Location = new Point(pad, y),
            Size = new Size(contentW, 6),
            BackColor = Color.FromArgb(35, 35, 45)
        };
        Controls.Add(_smBarBg);

        _smBar = new Panel
        {
            Location = new Point(0, 0),
            Size = new Size(0, 6),
            BackColor = SmColor
        };
        _smBarBg.Controls.Add(_smBar);
        y += 9;

        _smFixedLabel = new Label
        {
            Text = "0 fixed",
            Location = new Point(pad + 4, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = FixColor
        };
        Controls.Add(_smFixedLabel);

        _smConfirmedLabel = new Label
        {
            Text = "0 confirmed",
            Location = new Point(pad + 80, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = ConfirmColor
        };
        Controls.Add(_smConfirmedLabel);
        y += 17;

        // ── Divider ──
        var div2 = new Panel
        {
            Location = new Point(pad, y),
            Size = new Size(contentW, 1),
            BackColor = Color.FromArgb(45, 45, 55)
        };
        Controls.Add(div2);
        y += 5;

        // ── Consensus + last correction ──
        _consensusLabel = new Label
        {
            Text = "",
            Location = new Point(pad, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = Color.FromArgb(100, 200, 180)
        };
        Controls.Add(_consensusLabel);
        y += 14;

        _lastCorrLabel = new Label
        {
            Text = "",
            Location = new Point(pad, y),
            Size = new Size(contentW, 28),
            Font = tinyFont,
            ForeColor = Color.FromArgb(200, 170, 80)
        };
        Controls.Add(_lastCorrLabel);

        // Make the form draggable — attach to all non-interactive controls
        MouseDown += OnFormMouseDown;
        MouseMove += OnFormMouseMove;
        MouseUp += OnFormMouseUp;
        foreach (Control c in Controls)
        {
            if (c is Label lbl && lbl != closeBtn)
            {
                lbl.MouseDown += OnFormMouseDown;
                lbl.MouseMove += OnFormMouseMove;
                lbl.MouseUp += OnFormMouseUp;
            }
            else if (c is Panel p)
            {
                p.MouseDown += OnFormMouseDown;
                p.MouseMove += OnFormMouseMove;
                p.MouseUp += OnFormMouseUp;
            }
        }
    }

    public void UpdateStats(EnsembleStats stats)
    {
        _lastStats = stats;
        if (InvokeRequired)
        {
            try { BeginInvoke(() => ApplyStats(stats)); } catch { }
            return;
        }
        ApplyStats(stats);
    }

    public void SetRecording(bool recording)
    {
        _recording = recording;
        if (InvokeRequired)
        {
            try { BeginInvoke(() => UpdateTitle()); } catch { }
            return;
        }
        UpdateTitle();
    }

    private void UpdateTitle()
    {
        _titleLabel.Text = _recording ? "ENSEMBLE \u25cf REC" : "ENSEMBLE";
        _titleLabel.ForeColor = _recording ? Color.FromArgb(255, 100, 100) : Color.FromArgb(100, 160, 255);
    }

    private void ApplyStats(EnsembleStats stats)
    {
        // ── Confidence ──
        var conf = stats.AverageConfidence;
        _confidenceLabel.Text = $"Conf: {conf:P1}";
        _confidenceLabel.ForeColor = conf >= 0.95 ? Color.FromArgb(100, 220, 100) :
            conf >= 0.85 ? Color.FromArgb(200, 200, 100) :
            conf >= 0.70 ? Color.FromArgb(255, 160, 60) :
            Color.IndianRed;

        var barWidth = (int)(_confidenceBarBg.Width * Math.Clamp(conf, 0, 1));
        _confidenceBar.Width = barWidth;
        _confidenceBar.BackColor = conf >= 0.95 ? Color.FromArgb(80, 200, 80) :
            conf >= 0.85 ? Color.FromArgb(200, 200, 60) :
            conf >= 0.70 ? Color.FromArgb(255, 160, 40) :
            Color.IndianRed;

        _wordsLabel.Text = $"Words: {stats.TotalWords:N0}";
        _lowConfLabel.Text = $"Low conf: {stats.LowConfidenceWords}";

        _correctedLabel.Text = $"Corrected: {stats.CorrectedWords}";
        _correctedLabel.ForeColor = stats.CorrectedWords > 0 ? FixColor : DimColor;

        // ── AssemblyAI ──
        _aaiFinalsLabel.Text = $"{stats.S1Arrivals} finals";
        _aaiFixedLabel.Text = $"{stats.S1Corrections} fixed";
        _aaiFixedLabel.ForeColor = stats.S1Corrections > 0 ? FixColor : DimColor;
        _aaiConfirmedLabel.Text = $"{stats.S1Confirms} confirmed";
        _aaiConfirmedLabel.ForeColor = stats.S1Confirms > 0 ? ConfirmColor : DimColor;

        // AAI contribution bar: proportion of total corrections from this provider
        var aaiRatio = stats.CorrectedWords > 0 ? (double)stats.S1Corrections / stats.CorrectedWords : 0;
        _aaiBar.Width = (int)(_aaiBarBg.Width * Math.Clamp(aaiRatio, 0, 1));

        // ── Speechmatics ──
        _smFinalsLabel.Text = $"{stats.S2Arrivals} finals";
        _smFixedLabel.Text = $"{stats.S2Corrections} fixed";
        _smFixedLabel.ForeColor = stats.S2Corrections > 0 ? FixColor : DimColor;
        _smConfirmedLabel.Text = $"{stats.S2Confirms} confirmed";
        _smConfirmedLabel.ForeColor = stats.S2Confirms > 0 ? ConfirmColor : DimColor;

        var smRatio = stats.CorrectedWords > 0 ? (double)stats.S2Corrections / stats.CorrectedWords : 0;
        _smBar.Width = (int)(_smBarBg.Width * Math.Clamp(smRatio, 0, 1));

        // ── Consensus ──
        if (stats.ConsensusCorrections > 0)
            _consensusLabel.Text = $"Both agreed on {stats.ConsensusCorrections} correction{(stats.ConsensusCorrections == 1 ? "" : "s")}";
        else
            _consensusLabel.Text = "";

        // ── Last correction detail ──
        if (!string.IsNullOrEmpty(stats.LastCorrectionDetail))
            _lastCorrLabel.Text = stats.LastCorrectionDetail;
    }

    // ── Dragging ──
    private void OnFormMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            _dragging = true;
            _dragOffset = e.Location;
            if (sender is Control c && c != this)
                _dragOffset = new Point(e.X + c.Left, e.Y + c.Top);
        }
    }

    private void OnFormMouseMove(object? sender, MouseEventArgs e)
    {
        if (_dragging)
        {
            var screenPos = PointToScreen(e.Location);
            if (sender is Control c && c != this)
                screenPos = c.PointToScreen(e.Location);
            Location = new Point(screenPos.X - _dragOffset.X, screenPos.Y - _dragOffset.Y);
        }
    }

    private void OnFormMouseUp(object? sender, MouseEventArgs e)
    {
        _dragging = false;
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= 0x80; // WS_EX_TOOLWINDOW — doesn't show in taskbar/alt-tab
            return cp;
        }
    }
}
