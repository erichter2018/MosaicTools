using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Report Display settings: Changes highlighting, Rainbow mode, transparency.
/// </summary>
public class ReportDisplaySection : SettingsSection
{
    public override string SectionId => "report";

    private readonly CheckBox _showReportChangesCheck;
    private readonly Panel _reportChangesColorPanel;
    private readonly TrackBar _reportChangesAlphaSlider;
    private readonly Label _reportChangesAlphaLabel;
    private readonly RichTextBox _reportChangesPreview;
    private readonly CheckBox _correlationEnabledCheck;
    private readonly CheckBox _reportTransparentCheck;
    private readonly TrackBar _reportTransparencySlider;
    private readonly Label _reportTransparencyLabel;
    private readonly CheckBox _showReportAfterProcessCheck;
    private readonly CheckBox _showImpressionCheck;

    private Color _reportChangesColor = Color.FromArgb(144, 238, 144);

    public ReportDisplaySection(ToolTip toolTip) : base("Report Display", toolTip)
    {
        // Show Report After Process
        _showReportAfterProcessCheck = AddCheckBox("Show Report overlay after Process Report", LeftMargin, _nextY,
            "Automatically open report popup after Process Report.\nShows Changes or Rainbow highlighting if enabled.");
        _nextY += RowHeight;

        _showImpressionCheck = AddCheckBox("Show Impression popup after Process Report", LeftMargin, _nextY,
            "Display impression text after Process Report.\nClicks to dismiss, auto-hides on sign.");
        _nextY += RowHeight + 5;

        // Changes Highlighting
        AddSectionDivider("Changes Highlighting");

        _showReportChangesCheck = AddCheckBox("Highlight report changes", LeftMargin, _nextY,
            "Highlight new text in report popup with color.");

        _reportChangesColorPanel = AddColorPanel(LeftMargin + 180, _nextY - 2, 35, 20, _reportChangesColor,
            OnReportChangesColorClick, "Click to pick highlight color.");

        _reportChangesAlphaSlider = AddTrackBar(LeftMargin + 225, _nextY, 100, 5, 100, 30,
            "Transparency of highlight color.");
        _reportChangesAlphaSlider.ValueChanged += (s, e) =>
        {
            _reportChangesAlphaLabel.Text = $"{_reportChangesAlphaSlider.Value}%";
            UpdateReportChangesPreview();
        };

        _reportChangesAlphaLabel = AddLabel("30%", LeftMargin + 330, _nextY + 3);
        _reportChangesAlphaLabel.ForeColor = Color.Gray;
        _nextY += RowHeight + 5;

        // Preview
        _reportChangesPreview = new RichTextBox
        {
            Location = new Point(LeftMargin, _nextY),
            Size = new Size(390, 36),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            ReadOnly = true,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Segoe UI", 9)
        };
        Controls.Add(_reportChangesPreview);
        _nextY += 42;

        // Rainbow Mode
        _correlationEnabledCheck = AddCheckBox("Rainbow Mode (findings-impression correlation)", LeftMargin, _nextY,
            "Color-codes matching concepts between Findings and Impression.\nClick report popup to cycle between Changes and Rainbow modes.");
        _nextY += RowHeight + 5;

        // Transparent Overlay
        AddSectionDivider("Overlay Transparency");

        _reportTransparentCheck = AddCheckBox("Transparent overlay (see image through report)", LeftMargin, _nextY,
            "When enabled, the report popup background is semi-transparent\nso the radiology image shows through.");
        _reportTransparentCheck.CheckedChanged += (s, e) =>
        {
            _reportTransparencySlider.Enabled = _reportTransparentCheck.Checked;
            _reportTransparencyLabel.ForeColor = _reportTransparentCheck.Checked ? Color.Gray : Color.FromArgb(80, 80, 80);
        };
        _nextY += RowHeight;

        AddLabel("Transparency:", LeftMargin + 25, _nextY + 3);
        _reportTransparencySlider = AddTrackBar(LeftMargin + 130, _nextY, 160, 10, 100, 55,
            "Controls how see-through the report overlay is.\nLower values let more of the image show through.");

        _reportTransparencyLabel = AddLabel("55%", LeftMargin + 300, _nextY + 3);
        _reportTransparencyLabel.ForeColor = Color.Gray;
        _reportTransparencySlider.ValueChanged += (s, e) =>
        {
            _reportTransparencyLabel.Text = $"{_reportTransparencySlider.Value}%";
        };
        _nextY += RowHeight;

        UpdateReportChangesPreview();
        UpdateHeight();
    }

    private void OnReportChangesColorClick(object? sender, EventArgs e)
    {
        using var colorDialog = new ColorDialog
        {
            Color = _reportChangesColor,
            FullOpen = true
        };

        if (colorDialog.ShowDialog() == DialogResult.OK)
        {
            _reportChangesColor = colorDialog.Color;
            _reportChangesColorPanel.BackColor = _reportChangesColor;
            UpdateReportChangesPreview();
        }
    }

    private void UpdateReportChangesPreview()
    {
        _reportChangesPreview.Clear();
        _reportChangesPreview.Text = "Normal text. New dictated text appears highlighted.";

        // Calculate highlight color with alpha
        var alpha = (int)(_reportChangesAlphaSlider.Value / 100.0 * 255);
        var highlightColor = Color.FromArgb(
            (_reportChangesColor.R * alpha + 50 * (255 - alpha)) / 255,
            (_reportChangesColor.G * alpha + 50 * (255 - alpha)) / 255,
            (_reportChangesColor.B * alpha + 50 * (255 - alpha)) / 255);

        // Highlight "New dictated text appears highlighted"
        _reportChangesPreview.Select(13, 35);
        _reportChangesPreview.SelectionBackColor = highlightColor;
        _reportChangesPreview.Select(0, 0);
    }

    public override void LoadSettings(Configuration config)
    {
        _showReportAfterProcessCheck.Checked = config.ShowReportAfterProcess;
        _showImpressionCheck.Checked = config.ShowImpression;
        _showReportChangesCheck.Checked = config.ShowReportChanges;

        // Parse hex color string (e.g., "#90EE90")
        try
        {
            _reportChangesColor = ColorTranslator.FromHtml(config.ReportChangesColor);
        }
        catch
        {
            _reportChangesColor = Color.FromArgb(144, 238, 144); // Default light green
        }
        _reportChangesColorPanel.BackColor = _reportChangesColor;
        _reportChangesAlphaSlider.Value = config.ReportChangesAlpha;
        _reportChangesAlphaLabel.Text = $"{config.ReportChangesAlpha}%";

        _correlationEnabledCheck.Checked = config.CorrelationEnabled;

        _reportTransparentCheck.Checked = config.ReportPopupTransparent;
        _reportTransparencySlider.Value = config.ReportPopupTransparency;
        _reportTransparencySlider.Enabled = config.ReportPopupTransparent;
        _reportTransparencyLabel.Text = $"{config.ReportPopupTransparency}%";
        _reportTransparencyLabel.ForeColor = config.ReportPopupTransparent ? Color.Gray : Color.FromArgb(80, 80, 80);

        UpdateReportChangesPreview();
    }

    public override void SaveSettings(Configuration config)
    {
        config.ShowReportAfterProcess = _showReportAfterProcessCheck.Checked;
        config.ShowImpression = _showImpressionCheck.Checked;
        config.ShowReportChanges = _showReportChangesCheck.Checked;

        // Save as hex color string
        config.ReportChangesColor = $"#{_reportChangesColor.R:X2}{_reportChangesColor.G:X2}{_reportChangesColor.B:X2}";
        config.ReportChangesAlpha = _reportChangesAlphaSlider.Value;

        config.CorrelationEnabled = _correlationEnabledCheck.Checked;

        config.ReportPopupTransparent = _reportTransparentCheck.Checked;
        config.ReportPopupTransparency = _reportTransparencySlider.Value;
    }
}
