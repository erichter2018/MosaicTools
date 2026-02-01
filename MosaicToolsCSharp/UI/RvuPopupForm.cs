using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace MosaicTools.UI;

/// <summary>
/// Small borderless topmost popup showing secondary RVU metrics.
/// Used for hover popup (hides on mouse leave) and vertical stack drawer (persistent).
/// Styled to match the main bar: dark gray outer border, black inner background.
/// Labels are right-aligned so colons stack vertically.
/// </summary>
public class RvuPopupForm : Form
{
    private readonly Panel _innerPanel;
    private readonly List<Control> _metricControls = new();

    /// <summary>
    /// When true, the popup won't auto-hide on mouse leave (used for drawer mode).
    /// </summary>
    public bool Persistent { get; set; }

    public RvuPopupForm()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(51, 51, 51); // Same outer color as main bar
        StartPosition = FormStartPosition.Manual;
        AutoSize = false;
        DoubleBuffered = true;
        Padding = new Padding(1); // 1px border like the main bar

        _innerPanel = new Panel
        {
            BackColor = Color.Black,
            Dock = DockStyle.Fill
        };
        typeof(Panel).GetProperty("DoubleBuffered",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .SetValue(_innerPanel, true);
        Controls.Add(_innerPanel);
    }

    /// <summary>
    /// Update the popup with metric lines. Each entry is (label, value, color).
    /// Labels are right-aligned so the colons line up vertically.
    /// </summary>
    public void SetMetrics(List<(string Label, string Value, Color ValueColor)> metrics)
    {
        SuspendLayout();
        _innerPanel.SuspendLayout();

        var labelFont = new Font("Segoe UI", 8.5f);
        var valueFont = new Font("Segoe UI", 8.5f, FontStyle.Bold);

        int neededControls = metrics.Count * 2; // label + value per metric
        bool rebuild = _metricControls.Count != neededControls;

        if (rebuild)
        {
            // Control count changed — recreate
            foreach (var ctrl in _metricControls)
                ctrl.Dispose();
            _metricControls.Clear();
            _innerPanel.Controls.Clear();

            for (int i = 0; i < metrics.Count; i++)
            {
                var lblLabel = new Label
                {
                    Font = labelFont,
                    ForeColor = Color.FromArgb(140, 140, 140),
                    BackColor = Color.Black,
                    AutoSize = true
                };
                lblLabel.MouseLeave += OnChildMouseLeave;
                _innerPanel.Controls.Add(lblLabel);
                _metricControls.Add(lblLabel);

                var lblValue = new Label
                {
                    Font = valueFont,
                    BackColor = Color.Black,
                    AutoSize = true
                };
                lblValue.MouseLeave += OnChildMouseLeave;
                _innerPanel.Controls.Add(lblValue);
                _metricControls.Add(lblValue);
            }
        }

        // First pass: measure all labels to find the widest one
        int maxLabelWidth = 0;
        foreach (var (label, _, _) in metrics)
        {
            var size = TextRenderer.MeasureText(label, labelFont);
            if (size.Width > maxLabelWidth)
                maxLabelWidth = size.Width;
        }

        // Column where values start (after right-aligned labels)
        int valueColumnX = maxLabelWidth + 6;

        // Second pass: update text, color, and position
        int y = 3;
        int maxRight = 0;

        for (int i = 0; i < metrics.Count; i++)
        {
            var (label, value, color) = metrics[i];
            var lblLabel = (Label)_metricControls[i * 2];
            var lblValue = (Label)_metricControls[i * 2 + 1];

            lblLabel.Text = label;
            lblValue.Text = value;
            lblValue.ForeColor = color;

            var labelSize = TextRenderer.MeasureText(label, labelFont);
            lblLabel.Location = new Point(valueColumnX - labelSize.Width, y);
            lblValue.Location = new Point(valueColumnX + 2, y);

            var valueSize = TextRenderer.MeasureText(value, valueFont);
            int right = valueColumnX + 2 + valueSize.Width;
            if (right > maxRight) maxRight = right;

            y += Math.Max(labelSize.Height, valueSize.Height) - 2;
        }

        // +2 for the 1px border on each side
        Size = new Size(maxRight + 8 + 2, y + 4 + 2);
        _innerPanel.ResumeLayout(false);
        ResumeLayout(false);
    }

    private void OnChildMouseLeave(object? sender, EventArgs e)
    {
        if (Persistent) return;
        CheckAndHide();
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        if (Persistent) return;
        CheckAndHide();
    }

    private void CheckAndHide()
    {
        // Check if mouse actually left the form bounds (not just moved to a child)
        var pos = PointToClient(Cursor.Position);
        if (!ClientRectangle.Contains(pos))
            Hide();
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= 0x80; // WS_EX_TOOLWINDOW - don't show in taskbar/alt-tab
            return cp;
        }
    }
}
