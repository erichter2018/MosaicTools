using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// RVU & Metrics settings: RVUCounter integration, metrics selection, goal.
/// </summary>
public class RvuMetricsSection : SettingsSection
{
    public override string SectionId => "rvu";

    private readonly CheckBox _rvuMetricTotalCheck;
    private readonly CheckBox _rvuMetricPerHourCheck;
    private readonly CheckBox _rvuMetricCurrentHourCheck;
    private readonly CheckBox _rvuMetricPriorHourCheck;
    private readonly CheckBox _rvuMetricEstTotalCheck;
    private readonly CheckBox _rvuMetricRvuPerStudyCheck;
    private readonly CheckBox _rvuMetricAvgPerHourCheck;
    private readonly CheckBox _rvuMetricPaceDeltaCheck;
    private readonly ComboBox _rvuOverflowLayoutCombo;
    private readonly Label _rvuOverflowLayoutLabel;
    private readonly CheckBox _rvuGoalEnabledCheck;
    private readonly NumericUpDown _rvuGoalValueBox;
    private readonly CheckBox _paceCarEnabledCheck;
    private readonly NumericUpDown _paceCarSecondsBox;
    private readonly CheckBox _clarioRvuInjectCheck;
    private readonly TextBox _rvuCounterPathBox;
    private readonly Label _rvuCounterStatusLabel;

    public RvuMetricsSection(ToolTip toolTip) : base("TBWU & Metrics", toolTip)
    {
        // Metrics Selection
        AddSectionDivider("Display Metrics");

        _rvuMetricTotalCheck = AddCheckBox("Total", LeftMargin, _nextY);
        _rvuMetricPerHourCheck = AddCheckBox("TBWU/h", LeftMargin + 70, _nextY);
        _rvuMetricCurrentHourCheck = AddCheckBox("This Hour", LeftMargin + 140, _nextY);
        _rvuMetricPriorHourCheck = AddCheckBox("Prev Hour", LeftMargin + 230, _nextY);
        _rvuMetricEstTotalCheck = AddCheckBox("Est Total", LeftMargin + 325, _nextY);
        _nextY += RowHeight;

        _rvuMetricRvuPerStudyCheck = AddCheckBox("TBWU/Study", LeftMargin, _nextY);
        _rvuMetricAvgPerHourCheck = AddCheckBox("Avg/h", LeftMargin + 100, _nextY);
        _rvuMetricPaceDeltaCheck = AddCheckBox("Pace +/-", LeftMargin + 170, _nextY,
            "Shows only current pace delta in metrics (for example, +2.1 ahead).");

        // Wire up checkbox changes to enable/disable layout combo
        EventHandler updateLayoutState = (s, e) => UpdateOverflowLayoutState();
        _rvuMetricTotalCheck.CheckedChanged += updateLayoutState;
        _rvuMetricPerHourCheck.CheckedChanged += updateLayoutState;
        _rvuMetricCurrentHourCheck.CheckedChanged += updateLayoutState;
        _rvuMetricPriorHourCheck.CheckedChanged += updateLayoutState;
        _rvuMetricEstTotalCheck.CheckedChanged += updateLayoutState;
        _rvuMetricRvuPerStudyCheck.CheckedChanged += updateLayoutState;
        _rvuMetricAvgPerHourCheck.CheckedChanged += updateLayoutState;
        _rvuMetricPaceDeltaCheck.CheckedChanged += updateLayoutState;
        _nextY += RowHeight;

        // Layout combo (for 3+ metrics)
        _rvuOverflowLayoutLabel = AddLabel("3+ layout:", LeftMargin, _nextY + 2);
        _rvuOverflowLayoutLabel.ForeColor = Color.Gray;
        _rvuOverflowLayoutLabel.Font = new Font("Segoe UI", 8);

        _rvuOverflowLayoutCombo = AddComboBox(LeftMargin + 70, _nextY, 110,
            new[] { "Horizontal", "Vertical Stack", "Hover Popup", "Carousel" },
            "Layout when 3+ metrics selected:\nHorizontal = wide bar, Vertical = stacked rows,\nHover = popup on mouse-over, Carousel = cycles every 4s.");
        _rvuOverflowLayoutCombo.Enabled = false;
        _nextY += RowHeight + 5;

        // Goal
        AddSectionDivider("Goal");

        _rvuGoalEnabledCheck = AddCheckBox("Goal:", LeftMargin, _nextY);
        _rvuGoalValueBox = AddNumericUpDown(LeftMargin + 65, _nextY, 55, 1, 100, 10);
        _rvuGoalValueBox.DecimalPlaces = 1;
        _rvuGoalValueBox.Increment = 0.5m;
        AddLabel("/h (colors TBWU/h red when below)", LeftMargin + 125, _nextY + 3).ForeColor = Color.Gray;
        _nextY += RowHeight + 5;

        // Pace Car
        AddSectionDivider("Pace Car");

        _paceCarEnabledCheck = AddCheckBox("Show full pace car information", LeftMargin, _nextY,
            "Alternates between TBWU metrics and pace comparison display.");
        _nextY += RowHeight;

        AddLabel("Alternate every", LeftMargin + 20, _nextY + 3, isSubItem: true);
        _paceCarSecondsBox = AddNumericUpDown(LeftMargin + 120, _nextY, 50, 3, 30, 8,
            "Seconds between switching TBWU metrics and pace car display (3-30).");
        AddLabel("sec", LeftMargin + 175, _nextY + 3, isSubItem: true);
        _nextY += RowHeight + 5;

        // Clario Integration
        AddSectionDivider("Clario Integration");

        _clarioRvuInjectCheck = AddCheckBox("Inject TBWU values into Clario worklist", LeftMargin, _nextY,
            "Replaces Clario's broken TBWU display with live RVUCounter values via CDP.");
        _nextY += RowHeight + 5;

        // Database Path
        AddSectionDivider("Database");

        _rvuCounterPathBox = AddTextBox(LeftMargin, _nextY, 280);
        _rvuCounterPathBox.ReadOnly = true;
        _rvuCounterPathBox.ForeColor = Color.LightGray;

        var findBtn = AddButton("Find", LeftMargin + 290, _nextY - 2, 45, 22, OnFindRvuCounterClick,
            "Auto-detect RVUCounter database location.");

        var browseBtn = AddButton("...", LeftMargin + 340, _nextY - 2, 30, 22, OnBrowseDatabaseClick,
            "Browse for database file.");
        _nextY += RowHeight;

        _rvuCounterStatusLabel = AddLabel("", LeftMargin, _nextY);
        _rvuCounterStatusLabel.ForeColor = Color.Gray;
        _rvuCounterStatusLabel.AutoSize = true;
        _nextY += SubRowHeight;

        UpdateHeight();
    }

    private void UpdateOverflowLayoutState()
    {
        var count = new[] { _rvuMetricTotalCheck, _rvuMetricPerHourCheck, _rvuMetricCurrentHourCheck,
            _rvuMetricPriorHourCheck, _rvuMetricEstTotalCheck, _rvuMetricRvuPerStudyCheck,
            _rvuMetricAvgPerHourCheck, _rvuMetricPaceDeltaCheck }.Count(c => c.Checked);
        _rvuOverflowLayoutCombo.Enabled = count >= 3;
        _rvuOverflowLayoutLabel.ForeColor = count >= 3 ? Color.Gray : Color.FromArgb(80, 80, 80);
    }

    private void OnFindRvuCounterClick(object? sender, EventArgs e)
    {
        // Try common locations
        var paths = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RVUCounter", "rvu_data.db"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RVUCounter", "rvu_data.db"),
            @"C:\RVUCounter\rvu_data.db"
        };

        foreach (var path in paths)
        {
            if (File.Exists(path))
            {
                _rvuCounterPathBox.Text = path;
                _rvuCounterStatusLabel.Text = "Found!";
                _rvuCounterStatusLabel.ForeColor = Color.FromArgb(100, 180, 100);
                return;
            }
        }

        _rvuCounterStatusLabel.Text = "Not found. Use Browse to locate manually.";
        _rvuCounterStatusLabel.ForeColor = Color.FromArgb(220, 100, 100);
    }

    private void OnBrowseDatabaseClick(object? sender, EventArgs e)
    {
        using var ofd = new OpenFileDialog
        {
            Title = "Select RVUCounter Database",
            Filter = "Database files (*.db)|*.db|All files (*.*)|*.*",
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        };

        if (ofd.ShowDialog() == DialogResult.OK)
        {
            _rvuCounterPathBox.Text = ofd.FileName;
            _rvuCounterStatusLabel.Text = "Selected.";
            _rvuCounterStatusLabel.ForeColor = Color.FromArgb(100, 180, 100);
        }
    }

    public override void LoadSettings(Configuration config)
    {
        // Load metrics from flags enum
        _rvuMetricTotalCheck.Checked = config.RvuMetrics.HasFlag(RvuMetric.Total);
        _rvuMetricPerHourCheck.Checked = config.RvuMetrics.HasFlag(RvuMetric.PerHour);
        _rvuMetricCurrentHourCheck.Checked = config.RvuMetrics.HasFlag(RvuMetric.CurrentHour);
        _rvuMetricPriorHourCheck.Checked = config.RvuMetrics.HasFlag(RvuMetric.PriorHour);
        _rvuMetricEstTotalCheck.Checked = config.RvuMetrics.HasFlag(RvuMetric.EstimatedTotal);
        _rvuMetricRvuPerStudyCheck.Checked = config.RvuMetrics.HasFlag(RvuMetric.RvuPerStudy);
        _rvuMetricAvgPerHourCheck.Checked = config.RvuMetrics.HasFlag(RvuMetric.AvgPerHour);
        _rvuMetricPaceDeltaCheck.Checked = config.RvuMetrics.HasFlag(RvuMetric.PaceDelta);

        _rvuOverflowLayoutCombo.SelectedIndex = (int)config.RvuOverflowLayout;

        _rvuGoalEnabledCheck.Checked = config.RvuGoalEnabled;
        _rvuGoalValueBox.Value = (decimal)config.RvuGoalPerHour;

        _paceCarEnabledCheck.Checked = config.PaceCarEnabled;
        _paceCarSecondsBox.Value = Math.Max(3, Math.Min(30, config.PaceCarAlternateSeconds));

        _clarioRvuInjectCheck.Checked = config.ClarioRvuInjectEnabled;

        _rvuCounterPathBox.Text = config.RvuCounterPath ?? "";

        UpdateOverflowLayoutState();
    }

    public override void SaveSettings(Configuration config)
    {
        // Save metrics as flags enum
        var metrics = RvuMetric.None;
        if (_rvuMetricTotalCheck.Checked) metrics |= RvuMetric.Total;
        if (_rvuMetricPerHourCheck.Checked) metrics |= RvuMetric.PerHour;
        if (_rvuMetricCurrentHourCheck.Checked) metrics |= RvuMetric.CurrentHour;
        if (_rvuMetricPriorHourCheck.Checked) metrics |= RvuMetric.PriorHour;
        if (_rvuMetricEstTotalCheck.Checked) metrics |= RvuMetric.EstimatedTotal;
        if (_rvuMetricRvuPerStudyCheck.Checked) metrics |= RvuMetric.RvuPerStudy;
        if (_rvuMetricAvgPerHourCheck.Checked) metrics |= RvuMetric.AvgPerHour;
        if (_rvuMetricPaceDeltaCheck.Checked) metrics |= RvuMetric.PaceDelta;
        config.RvuMetrics = metrics;

        config.RvuOverflowLayout = (RvuOverflowLayout)_rvuOverflowLayoutCombo.SelectedIndex;

        config.RvuGoalEnabled = _rvuGoalEnabledCheck.Checked;
        config.RvuGoalPerHour = (double)_rvuGoalValueBox.Value;

        config.PaceCarEnabled = _paceCarEnabledCheck.Checked;
        config.PaceCarAlternateSeconds = (int)_paceCarSecondsBox.Value;

        config.ClarioRvuInjectEnabled = _clarioRvuInjectCheck.Checked;

        config.RvuCounterPath = _rvuCounterPathBox.Text;
    }
}
