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
    private readonly CheckBox _useSendInputInsertCheck;
    private readonly CheckBox _cdpEnabledCheck;
    private readonly CheckBox _cdpScrollFixCheck;
    private readonly CheckBox _cdpAutoScrollCheck;
    private readonly CheckBox _cdpHideDragHandlesCheck;
    private readonly CheckBox _cdpFlashingAlertTextCheck;

    public ExperimentalSection(ToolTip toolTip) : base("Experimental", toolTip)
    {

        var warningLabel = AddLabel("⚠ These features may change or be removed", LeftMargin, _nextY);
        warningLabel.ForeColor = Color.FromArgb(255, 180, 50);
        warningLabel.Font = new Font("Segoe UI", 8, FontStyle.Italic);
        _nextY += RowHeight;

        AddSectionDivider("Insertion");

        _useSendInputInsertCheck = AddCheckBox("Use SendInput instead of Ctrl+V", LeftMargin, _nextY,
            "Experimental: insert text via SendInput (no Ctrl+V). Default remains clipboard + Ctrl+V.");
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

        // CDP (Direct DOM Access)
        AddSectionDivider("Mosaic CDP (Direct DOM Access)");

        _cdpEnabledCheck = AddCheckBox("Use CDP for Mosaic interaction", LeftMargin, _nextY,
            "Falls back to UI Automation if CDP unavailable. Requires Mosaic restart on first enable.");
        _cdpEnabledCheck.CheckedChanged += (s, e) => UpdateCdpSettingsStates();
        _nextY += RowHeight;

        AddHintLabel("Reads Mosaic DOM directly via Chrome DevTools Protocol — faster, no COM leaks", LeftMargin + 25);

        _cdpScrollFixCheck = AddCheckBox("Independent column scrolling", LeftMargin + 25, _nextY,
            "Makes Transcript, Report, and sidebar columns scroll independently instead of the whole page.");
        _nextY += RowHeight;

        _cdpAutoScrollCheck = AddCheckBox("Auto-scroll to cursor during dictation", LeftMargin + 25, _nextY,
            "Keeps the cursor visible by scrolling the editor area when text is inserted near the bottom.");
        _nextY += RowHeight;

        _cdpHideDragHandlesCheck = AddCheckBox("Hide report drag/delete handles", LeftMargin + 25, _nextY,
            "Hides the drag-to-reorder and delete icons on paragraphs in the report editor.");
        _nextY += RowHeight;

        _cdpFlashingAlertTextCheck = AddCheckBox("Flashing Alert Text", LeftMargin + 25, _nextY,
            "When gender mismatch or findings/impression mismatch is active, matching final report text flashes via CDP until corrected.");
        _nextY += RowHeight;

        AddHintLabel("Alert colors:", LeftMargin + 25);
        var genderColorLabel = AddLabel("Gender mismatch = red (#DC0000 / #780000)", LeftMargin + 40, _nextY);
        genderColorLabel.ForeColor = Color.FromArgb(255, 120, 120);
        genderColorLabel.Font = new Font("Segoe UI", 8);
        _nextY += 18;

        var fimColorLabel = AddLabel("FIM mismatch = orange (#FFA000 / #B46E00)", LeftMargin + 40, _nextY);
        fimColorLabel.ForeColor = Color.FromArgb(255, 190, 110);
        fimColorLabel.Font = new Font("Segoe UI", 8);
        _nextY += 18;

        UpdateHeight();
    }

    private void UpdateNetworkSettingsStates()
    {
        bool enabled = _connectivityMonitorEnabledCheck.Checked;
        _connectivityIntervalUpDown.Enabled = enabled;
        _connectivityTimeoutUpDown.Enabled = enabled;
    }

    private void UpdateCdpSettingsStates()
    {
        _cdpScrollFixCheck.Enabled = _cdpEnabledCheck.Checked;
        _cdpAutoScrollCheck.Enabled = _cdpEnabledCheck.Checked;
        _cdpHideDragHandlesCheck.Enabled = _cdpEnabledCheck.Checked;
        _cdpFlashingAlertTextCheck.Enabled = _cdpEnabledCheck.Checked;
    }

    public override void LoadSettings(Configuration config)
    {
        _connectivityMonitorEnabledCheck.Checked = config.ConnectivityMonitorEnabled;
        _connectivityIntervalUpDown.Value = config.ConnectivityCheckIntervalSeconds;
        // Config stores ms, UI shows seconds
        _connectivityTimeoutUpDown.Value = Math.Max(1, (config.ConnectivityTimeoutMs + 500) / 1000);
        _useSendInputInsertCheck.Checked = config.ExperimentalUseSendInputInsert;
        _cdpEnabledCheck.Checked = config.CdpEnabled;
        _cdpScrollFixCheck.Checked = config.CdpIndependentScrolling;
        _cdpAutoScrollCheck.Checked = config.CdpAutoScrollEnabled;
        _cdpHideDragHandlesCheck.Checked = config.CdpHideDragHandles;
        _cdpFlashingAlertTextCheck.Checked = config.CdpFlashingAlertText;

        UpdateNetworkSettingsStates();
        UpdateCdpSettingsStates();
    }

    public override void SaveSettings(Configuration config)
    {
        config.ConnectivityMonitorEnabled = _connectivityMonitorEnabledCheck.Checked;
        config.ConnectivityCheckIntervalSeconds = (int)_connectivityIntervalUpDown.Value;
        // UI shows seconds, config stores ms
        config.ConnectivityTimeoutMs = (int)_connectivityTimeoutUpDown.Value * 1000;
        config.ExperimentalUseSendInputInsert = _useSendInputInsertCheck.Checked;
        config.CdpEnabled = _cdpEnabledCheck.Checked;
        config.CdpIndependentScrolling = _cdpScrollFixCheck.Checked;
        config.CdpAutoScrollEnabled = _cdpAutoScrollCheck.Checked;
        config.CdpHideDragHandles = _cdpHideDragHandlesCheck.Checked;
        config.CdpFlashingAlertText = _cdpFlashingAlertTextCheck.Checked;
    }
}
