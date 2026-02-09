using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Experimental settings: Network monitor.
/// </summary>
public class ExperimentalSection : SettingsSection
{
    public override string SectionId => "experimental";

    private readonly CheckBox _connectivityMonitorEnabledCheck;
    private readonly NumericUpDown _connectivityIntervalUpDown;
    private readonly NumericUpDown _connectivityTimeoutUpDown;

    public ExperimentalSection(ToolTip toolTip) : base("Experimental", toolTip)
    {

        var warningLabel = AddLabel("⚠ These features may change or be removed", LeftMargin, _nextY);
        warningLabel.ForeColor = Color.FromArgb(255, 180, 50);
        warningLabel.Font = new Font("Segoe UI", 8, FontStyle.Italic);
        _nextY += RowHeight;

        // Network Monitor
        AddSectionDivider("Network Monitor");

        _connectivityMonitorEnabledCheck = AddCheckBox("Show connectivity status dots", LeftMargin, _nextY,
            "Show connectivity dots in widget bar for network status monitoring.");
        _connectivityMonitorEnabledCheck.CheckedChanged += (s, e) => UpdateNetworkSettingsStates();
        _nextY += RowHeight;

        AddLabel("Check every:", LeftMargin + 25, _nextY + 3);
        _connectivityIntervalUpDown = AddNumericUpDown(LeftMargin + 110, _nextY, 50, 10, 120, 30);
        AddLabel("sec", LeftMargin + 165, _nextY + 3);

        AddLabel("Timeout:", LeftMargin + 210, _nextY + 3);
        _connectivityTimeoutUpDown = AddNumericUpDown(LeftMargin + 270, _nextY, 50, 1, 10, 5);
        AddLabel("sec", LeftMargin + 325, _nextY + 3);
        _nextY += SubRowHeight;

        AddHintLabel("Shows 4 status dots: Mirth, Mosaic, Clario, InteleViewer", LeftMargin + 25);

        UpdateHeight();
    }

    private void UpdateNetworkSettingsStates()
    {
        bool enabled = _connectivityMonitorEnabledCheck.Checked;
        _connectivityIntervalUpDown.Enabled = enabled;
        _connectivityTimeoutUpDown.Enabled = enabled;
    }

    public override void LoadSettings(Configuration config)
    {
        _connectivityMonitorEnabledCheck.Checked = config.ConnectivityMonitorEnabled;
        _connectivityIntervalUpDown.Value = config.ConnectivityCheckIntervalSeconds;
        // Config stores ms, UI shows seconds
        _connectivityTimeoutUpDown.Value = Math.Max(1, (config.ConnectivityTimeoutMs + 500) / 1000);

        UpdateNetworkSettingsStates();
    }

    public override void SaveSettings(Configuration config)
    {
        config.ConnectivityMonitorEnabled = _connectivityMonitorEnabledCheck.Checked;
        config.ConnectivityCheckIntervalSeconds = (int)_connectivityIntervalUpDown.Value;
        // UI shows seconds, config stores ms
        config.ConnectivityTimeoutMs = (int)_connectivityTimeoutUpDown.Value * 1000;
    }
}
