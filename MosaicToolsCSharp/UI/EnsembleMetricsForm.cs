// [CustomSTT] Live ensemble metrics popup — draggable, always-on-top stats display
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Small draggable overlay showing live ensemble STT statistics.
/// Shows Deepgram confidence, per-provider corrections/confirms, accuracy impact, last correction detail.
/// </summary>
public class EnsembleMetricsForm : Form
{
    private readonly Configuration _config;
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
    private readonly Label _snxHeader;
    private readonly Label _snxFinalsLabel;
    private readonly Label _snxFixedLabel;
    private readonly Label _snxConfirmedLabel;
    private readonly Panel _snxBar;
    private readonly Panel _snxBarBg;

    private readonly Label _smHeader;
    private readonly Label _smFinalsLabel;
    private readonly Label _smFixedLabel;
    private readonly Label _smConfirmedLabel;
    private readonly Panel _smBar;
    private readonly Panel _smBarBg;

    // Bottom
    private readonly Label _consensusLabel;
    private readonly Label _lastCorrLabel;

    // Accuracy impact
    private readonly Label _accuracyHeader;
    private readonly Label _sessionAccuracyLabel;
    private readonly Label _alltimeAccuracyLabel;
    private readonly Label _precisionLabel;

    // All-time
    private readonly Label _alltimeHeader;
    private readonly Label _alltimeWordsLabel;
    private readonly Label _alltimeCorrectedLabel;
    private readonly Label _alltimeConfLabel;

    private static readonly Color DimColor = Color.FromArgb(120, 120, 130);
    private static readonly Color BrightColor = Color.FromArgb(210, 210, 220);
    private static readonly Color FixColor = Color.FromArgb(100, 220, 100);  // green
    private static readonly Color ConfirmColor = Color.FromArgb(200, 200, 100); // yellow-ish

    private static Color ProviderColor(string name) => name switch
    {
        "soniox" => Color.FromArgb(80, 200, 180),    // teal
        "speechmatics" => Color.FromArgb(180, 130, 255), // purple
        "assemblyai" => Color.FromArgb(100, 180, 255),   // blue
        _ => Color.FromArgb(150, 150, 200)
    };

    private static string ShortName(string name) => name switch
    {
        "soniox" => "SNX", "speechmatics" => "SM", "assemblyai" => "AAI",
        _ => name.ToUpperInvariant()[..Math.Min(3, name.Length)]
    };

    public EnsembleMetricsForm(Configuration config, string s1Name = "soniox", string s2Name = "speechmatics")
    {
        _config = config;
        var s1Color = ProviderColor(s1Name);
        var s2Color = ProviderColor(s2Name);
        var s1Short = ShortName(s1Name);
        var s2Short = ShortName(s2Name);
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;
        Size = new Size(260, 376);
        BackColor = Color.FromArgb(22, 22, 28);
        Opacity = 0.93;

        // Restore saved position, or default to bottom-left above taskbar
        if (config.SttEnsembleMetricsX != int.MinValue && config.SttEnsembleMetricsY != int.MinValue)
        {
            Location = new Point(config.SttEnsembleMetricsX, config.SttEnsembleMetricsY);
        }
        else
        {
            var screen = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
            Location = new Point(screen.Left + 10, screen.Bottom - Height - 10);
        }

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
        Controls.Add(new Panel { Location = new Point(pad, y), Size = new Size(contentW, 1), BackColor = Color.FromArgb(45, 45, 55) });
        y += 5;

        // ── Secondary 1 section ──
        _snxHeader = new Label
        {
            Text = s1Short,
            Location = new Point(pad, y),
            AutoSize = true,
            Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
            ForeColor = s1Color
        };
        Controls.Add(_snxHeader);

        _snxFinalsLabel = new Label
        {
            Text = "0 finals",
            Location = new Point(pad + 35, y),
            AutoSize = true,
            Font = smallFont,
            ForeColor = DimColor
        };
        Controls.Add(_snxFinalsLabel);
        y += 15;

        // S1 contribution bar background
        _snxBarBg = new Panel
        {
            Location = new Point(pad, y),
            Size = new Size(contentW, 6),
            BackColor = Color.FromArgb(35, 35, 45)
        };
        Controls.Add(_snxBarBg);

        _snxBar = new Panel
        {
            Location = new Point(0, 0),
            Size = new Size(0, 6),
            BackColor = s1Color
        };
        _snxBarBg.Controls.Add(_snxBar);
        y += 9;

        _snxFixedLabel = new Label
        {
            Text = "0 fixed",
            Location = new Point(pad + 4, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = FixColor
        };
        Controls.Add(_snxFixedLabel);

        _snxConfirmedLabel = new Label
        {
            Text = "0 confirmed",
            Location = new Point(pad + 80, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = ConfirmColor
        };
        Controls.Add(_snxConfirmedLabel);
        y += 17;

        // ── Secondary 2 section ──
        _smHeader = new Label
        {
            Text = s2Short,
            Location = new Point(pad, y),
            AutoSize = true,
            Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
            ForeColor = s2Color
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

        // S2 contribution bar background
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
            BackColor = s2Color
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
        Controls.Add(new Panel { Location = new Point(pad, y), Size = new Size(contentW, 1), BackColor = Color.FromArgb(45, 45, 55) });
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
        y += 30;

        // ── Divider ──
        Controls.Add(new Panel { Location = new Point(pad, y), Size = new Size(contentW, 1), BackColor = Color.FromArgb(45, 45, 55) });
        y += 5;

        // ── Accuracy impact ──
        _accuracyHeader = new Label
        {
            Text = "ACCURACY vs DG SOLO",
            Location = new Point(pad, y),
            AutoSize = true,
            Font = new Font("Segoe UI", 7f, FontStyle.Bold),
            ForeColor = Color.FromArgb(90, 160, 120)
        };
        Controls.Add(_accuracyHeader);
        y += 14;

        _sessionAccuracyLabel = new Label
        {
            Text = "Session: +0.00%",
            Location = new Point(pad + 4, y),
            AutoSize = true,
            Font = bodyFont,
            ForeColor = DimColor
        };
        Controls.Add(_sessionAccuracyLabel);
        y += 16;

        _alltimeAccuracyLabel = new Label
        {
            Text = "All-time: +0.00%",
            Location = new Point(pad + 4, y),
            AutoSize = true,
            Font = bodyFont,
            ForeColor = DimColor
        };
        Controls.Add(_alltimeAccuracyLabel);
        y += 16;

        _precisionLabel = new Label
        {
            Text = "Precision: \u2014",
            Location = new Point(pad + 4, y),
            AutoSize = true,
            Font = bodyFont,
            ForeColor = DimColor
        };
        Controls.Add(_precisionLabel);
        y += 20;

        // ── Divider ──
        Controls.Add(new Panel { Location = new Point(pad, y), Size = new Size(contentW, 1), BackColor = Color.FromArgb(45, 45, 55) });
        y += 5;

        // ── All-time stats ──
        _alltimeHeader = new Label
        {
            Text = "ALL-TIME",
            Location = new Point(pad, y),
            AutoSize = true,
            Font = new Font("Segoe UI", 7f, FontStyle.Bold),
            ForeColor = Color.FromArgb(90, 90, 100)
        };
        Controls.Add(_alltimeHeader);
        y += 14;

        _alltimeWordsLabel = new Label
        {
            Text = "0 words",
            Location = new Point(pad + 4, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = DimColor
        };
        Controls.Add(_alltimeWordsLabel);

        _alltimeCorrectedLabel = new Label
        {
            Text = "0 corrected",
            Location = new Point(pad + 80, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = DimColor
        };
        Controls.Add(_alltimeCorrectedLabel);

        _alltimeConfLabel = new Label
        {
            Text = "Conf: --",
            Location = new Point(pad + 165, y),
            AutoSize = true,
            Font = tinyFont,
            ForeColor = DimColor
        };
        Controls.Add(_alltimeConfLabel);

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

        // ── Soniox ──
        _snxFinalsLabel.Text = $"{stats.S1Arrivals} finals";
        _snxFixedLabel.Text = $"{stats.S1Corrections} fixed";
        _snxFixedLabel.ForeColor = stats.S1Corrections > 0 ? FixColor : DimColor;
        _snxConfirmedLabel.Text = $"{stats.S1Confirms} confirmed";
        _snxConfirmedLabel.ForeColor = stats.S1Confirms > 0 ? ConfirmColor : DimColor;

        // Soniox contribution bar: proportion of total corrections from this provider
        var snxRatio = stats.CorrectedWords > 0 ? (double)stats.S1Corrections / stats.CorrectedWords : 0;
        _snxBar.Width = (int)(_snxBarBg.Width * Math.Clamp(snxRatio, 0, 1));

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

        // ── Accuracy impact ──
        // Correction rate = corrections / total words. Each correction is a word that DG
        // got wrong (low confidence) and secondaries fixed, so this is the accuracy delta.
        var sessionImpact = stats.TotalWords > 0 ? (double)stats.CorrectedWords / stats.TotalWords * 100 : 0;
        _sessionAccuracyLabel.Text = $"Session: +{sessionImpact:F2}% ({stats.CorrectedWords} of {stats.TotalWords:N0})";
        _sessionAccuracyLabel.ForeColor = stats.CorrectedWords > 0 ? FixColor : DimColor;

        var alltimeImpact = stats.AlltimeWords > 0 ? (double)stats.AlltimeCorrected / stats.AlltimeWords * 100 : 0;
        _alltimeAccuracyLabel.Text = $"All-time: +{alltimeImpact:F2}% ({stats.AlltimeCorrected:N0} of {stats.AlltimeWords:N0})";
        _alltimeAccuracyLabel.ForeColor = stats.AlltimeCorrected > 0 ? FixColor : DimColor;

        // ── Precision (validated corrections) ──
        var sessTotal = stats.SessionValidated + stats.SessionRejected;
        var atTotal = stats.AlltimeValidated + stats.AlltimeRejected;
        if (sessTotal > 0 || atTotal > 0)
        {
            var sessPart = sessTotal > 0
                ? $"{stats.SessionValidated}/{sessTotal} ({(double)stats.SessionValidated / sessTotal:P0})"
                : "\u2014";
            var atPart = atTotal > 0
                ? $"{stats.AlltimeValidated}/{atTotal} ({(double)stats.AlltimeValidated / atTotal:P0})"
                : "\u2014";
            _precisionLabel.Text = $"Precision: {sessPart} | All-time: {atPart}";
            _precisionLabel.ForeColor = BrightColor;
        }
        else
        {
            _precisionLabel.Text = "Precision: \u2014";
            _precisionLabel.ForeColor = DimColor;
        }

        // ── All-time ──
        _alltimeWordsLabel.Text = $"{stats.AlltimeWords:N0} words";
        _alltimeCorrectedLabel.Text = $"{stats.AlltimeCorrected:N0} corrected";
        if (stats.AlltimeWords > 0)
        {
            _alltimeConfLabel.Text = $"Conf: {stats.AlltimeAverageConfidence:P1}";
            _alltimeConfLabel.ForeColor = stats.AlltimeAverageConfidence >= 0.95 ? Color.FromArgb(100, 220, 100) :
                stats.AlltimeAverageConfidence >= 0.85 ? Color.FromArgb(200, 200, 100) : DimColor;
        }
        _alltimeCorrectedLabel.ForeColor = stats.AlltimeCorrected > 0 ? FixColor : DimColor;
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
        if (_dragging)
        {
            _dragging = false;
            _config.SttEnsembleMetricsX = Location.X;
            _config.SttEnsembleMetricsY = Location.Y;
            _config.Save();
        }
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
