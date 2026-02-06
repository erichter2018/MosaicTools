using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Behavior settings: Scraping, scrolling, focus, InteleViewer integration.
/// </summary>
public class BehaviorSection : SettingsSection
{
    public override string SectionId => "behavior";

    private readonly NumericUpDown _scrapeIntervalUpDown;
    private readonly CheckBox _scrollToBottomCheck;
    private readonly CheckBox _showLineCountToastCheck;
    private readonly NumericUpDown _scrollThreshold1;
    private readonly NumericUpDown _scrollThreshold2;
    private readonly NumericUpDown _scrollThreshold3;
    private readonly CheckBox _ignoreInpatientDraftedCheck;
    private readonly RadioButton _ignoreInpatientAllXrRadio;
    private readonly RadioButton _ignoreInpatientChestOnlyRadio;
    private readonly TextBox? _windowLevelKeysBox;

    public BehaviorSection(ToolTip toolTip, bool isHeadless) : base("Behavior", toolTip)
    {
        // Background Monitoring
        AddSectionDivider("Background Monitoring");

        AddLabel("Scrape Mosaic every", LeftMargin, _nextY + 3);
        _scrapeIntervalUpDown = AddNumericUpDown(LeftMargin + 150, _nextY, 50, 1, 30, 1);
        AddLabel("seconds", LeftMargin + 205, _nextY + 3);
        _nextY += SubRowHeight;
        AddHintLabel("Keep this at 1s unless you are having massive performance degradation.", LeftMargin + 25);
        _nextY += RowHeight + 5;

        // Report Processing
        AddSectionDivider("Report Processing");

        _scrollToBottomCheck = AddCheckBox("Scroll to bottom on Process Report", LeftMargin, _nextY,
            "Scroll report to bottom after Process Report to show IMPRESSION section.");
        _scrollToBottomCheck.CheckedChanged += (s, e) => UpdateScrollSubStates();
        _nextY += SubRowHeight;

        _showLineCountToastCheck = AddCheckBox("Show line count toast", LeftMargin + 25, _nextY,
            "Toast showing number of lines after Process Report.", true);
        _nextY += RowHeight;

        // Smart Scroll Thresholds
        AddLabel("Smart Scroll Thresholds (lines):", LeftMargin + 25, _nextY);
        _nextY += 20;

        AddLabel("1 PgDn >", LeftMargin + 25, _nextY + 2);
        _scrollThreshold1 = AddNumericUpDown(LeftMargin + 85, _nextY, 50, 1, 500, 20);

        AddLabel("2 PgDn >", LeftMargin + 150, _nextY + 2);
        _scrollThreshold2 = AddNumericUpDown(LeftMargin + 210, _nextY, 50, 1, 500, 40);

        AddLabel("3 PgDn >", LeftMargin + 275, _nextY + 2);
        _scrollThreshold3 = AddNumericUpDown(LeftMargin + 335, _nextY, 50, 1, 500, 60,
            "Lines at which to add Page Down presses.");
        _nextY += RowHeight + 5;

        // Threshold constraints
        _scrollThreshold1.ValueChanged += (s, e) =>
        {
            if (_scrollThreshold2.Value <= _scrollThreshold1.Value)
                _scrollThreshold2.Value = Math.Min(_scrollThreshold2.Maximum, _scrollThreshold1.Value + 1);
        };
        _scrollThreshold2.ValueChanged += (s, e) =>
        {
            if (_scrollThreshold1.Value >= _scrollThreshold2.Value)
                _scrollThreshold1.Value = Math.Max(_scrollThreshold1.Minimum, _scrollThreshold2.Value - 1);
            if (_scrollThreshold3.Value <= _scrollThreshold2.Value)
                _scrollThreshold3.Value = Math.Min(_scrollThreshold3.Maximum, _scrollThreshold2.Value + 1);
        };
        _scrollThreshold3.ValueChanged += (s, e) =>
        {
            if (_scrollThreshold2.Value >= _scrollThreshold3.Value)
                _scrollThreshold2.Value = Math.Max(_scrollThreshold2.Minimum, _scrollThreshold3.Value - 1);
        };

        // Inpatient XR Handling
        AddSectionDivider("Inpatient XR Handling");

        _ignoreInpatientDraftedCheck = AddCheckBox("Ignore Inpatient Drafted (select all after auto-insertions)", LeftMargin, _nextY,
            "Auto-select all text for inpatient XR studies after macro/clinical history insertions.");
        _ignoreInpatientDraftedCheck.CheckedChanged += (s, e) => UpdateInpatientSubStates();
        _nextY += SubRowHeight;

        _ignoreInpatientAllXrRadio = new RadioButton
        {
            Text = "All Inpatient XR",
            Location = new Point(LeftMargin + 25, _nextY),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 180, 180),
            Font = new Font("Segoe UI", 9),
            Checked = true,
            Enabled = false
        };
        Controls.Add(_ignoreInpatientAllXrRadio);

        _ignoreInpatientChestOnlyRadio = new RadioButton
        {
            Text = "Inpatient Chest XR only",
            Location = new Point(LeftMargin + 170, _nextY),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 180, 180),
            Font = new Font("Segoe UI", 9),
            Enabled = false
        };
        Controls.Add(_ignoreInpatientChestOnlyRadio);
        _toolTip.SetToolTip(_ignoreInpatientChestOnlyRadio, "Apply to all inpatient X-ray studies or only inpatient chest X-rays.");
        _nextY += RowHeight + 5;

        // InteleViewer Integration (non-headless only)
        if (!isHeadless)
        {
            AddSectionDivider("InteleViewer Integration");

            AddLabel("Window/Level cycle keys:", LeftMargin, _nextY + 3);
            _windowLevelKeysBox = AddTextBox(LeftMargin + 170, _nextY, 180,
                "Comma-separated keys sent to InteleViewer for window/level cycling.\nE.g., F4, F5, F7, F6");
            _nextY += SubRowHeight;
            AddHintLabel("Comma-separated keys (e.g., F4, F5, F7, F6)", LeftMargin);
        }

        UpdateHeight();
    }

    private void UpdateScrollSubStates()
    {
        bool enabled = _scrollToBottomCheck.Checked;
        _showLineCountToastCheck.Enabled = enabled;
        _scrollThreshold1.Enabled = enabled;
        _scrollThreshold2.Enabled = enabled;
        _scrollThreshold3.Enabled = enabled;

        var subColor = enabled ? Color.FromArgb(180, 180, 180) : Color.FromArgb(100, 100, 100);
        _showLineCountToastCheck.ForeColor = subColor;
    }

    private void UpdateInpatientSubStates()
    {
        bool enabled = _ignoreInpatientDraftedCheck.Checked;
        _ignoreInpatientAllXrRadio.Enabled = enabled;
        _ignoreInpatientChestOnlyRadio.Enabled = enabled;

        var subColor = enabled ? Color.FromArgb(180, 180, 180) : Color.FromArgb(100, 100, 100);
        _ignoreInpatientAllXrRadio.ForeColor = subColor;
        _ignoreInpatientChestOnlyRadio.ForeColor = subColor;
    }

    public override void LoadSettings(Configuration config)
    {
        _scrapeIntervalUpDown.Value = Math.Clamp(config.ScrapeIntervalSeconds, 1, 30);

        _scrollToBottomCheck.Checked = config.ScrollToBottomOnProcess;
        _showLineCountToastCheck.Checked = config.ShowLineCountToast;
        _scrollThreshold1.Value = config.ScrollThreshold1;
        _scrollThreshold2.Value = config.ScrollThreshold2;
        _scrollThreshold3.Value = config.ScrollThreshold3;

        _ignoreInpatientDraftedCheck.Checked = config.IgnoreInpatientDrafted;
        if (config.IgnoreInpatientDraftedMode == 1)
            _ignoreInpatientChestOnlyRadio.Checked = true;
        else
            _ignoreInpatientAllXrRadio.Checked = true;

        if (_windowLevelKeysBox != null)
            _windowLevelKeysBox.Text = string.Join(", ", config.WindowLevelKeys ?? new List<string>());

        UpdateScrollSubStates();
        UpdateInpatientSubStates();
    }

    public override void SaveSettings(Configuration config)
    {
        config.ScrapeIntervalSeconds = (int)_scrapeIntervalUpDown.Value;

        config.ScrollToBottomOnProcess = _scrollToBottomCheck.Checked;
        config.ShowLineCountToast = _showLineCountToastCheck.Checked;
        config.ScrollThreshold1 = (int)_scrollThreshold1.Value;
        config.ScrollThreshold2 = (int)_scrollThreshold2.Value;
        config.ScrollThreshold3 = (int)_scrollThreshold3.Value;

        config.IgnoreInpatientDrafted = _ignoreInpatientDraftedCheck.Checked;
        config.IgnoreInpatientDraftedMode = _ignoreInpatientChestOnlyRadio.Checked ? 1 : 0;

        if (_windowLevelKeysBox != null)
        {
            config.WindowLevelKeys = _windowLevelKeysBox.Text
                .Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }
    }
}
