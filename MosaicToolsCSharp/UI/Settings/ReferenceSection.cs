using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Reference section: AHK integration docs, debug tips, element dump tool.
/// </summary>
public class ReferenceSection : SettingsSection
{
    public override string SectionId => "reference";

    private readonly TextBox _infoBox;
    private readonly ComboBox _targetAppCombo;
    private readonly ComboBox _methodCombo;
    private readonly Button _dumpButton;
    private readonly Label _statusLabel;

    public ReferenceSection(ToolTip toolTip) : base("Reference", toolTip)
    {
        _infoBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            WordWrap = false,
            ScrollBars = ScrollBars.Vertical,
            Location = new Point(LeftMargin, _nextY),
            Size = new Size(450, 380),
            BackColor = Color.FromArgb(35, 35, 38),
            ForeColor = Color.LightGray,
            Font = new Font("Consolas", 8.5f),
            BorderStyle = BorderStyle.FixedSingle,
            Text =
"=== AHK / External Integration ===\r\n" +
"\r\n" +
"Send Windows Messages to trigger actions from\r\n" +
"AutoHotkey or other programs:\r\n" +
"\r\n" +
"WM_TRIGGER_SCRAPE = 0x0401        # Critical Findings\r\n" +
"WM_TRIGGER_BEEP = 0x0403          # System Beep\r\n" +
"WM_TRIGGER_SHOW_REPORT = 0x0404   # Show Report\r\n" +
"WM_TRIGGER_CAPTURE_SERIES = 0x0405 # Capture Series\r\n" +
"WM_TRIGGER_GET_PRIOR = 0x0406     # Get Prior\r\n" +
"WM_TRIGGER_TOGGLE_RECORD = 0x0407 # Toggle Record\r\n" +
"WM_TRIGGER_PROCESS_REPORT = 0x0408 # Process Report\r\n" +
"WM_TRIGGER_SIGN_REPORT = 0x0409   # Sign Report\r\n" +
"WM_TRIGGER_OPEN_SETTINGS = 0x040A # Open Settings\r\n" +
"WM_TRIGGER_CREATE_IMPRESSION = 0x040B # Create Impression\r\n" +
"WM_TRIGGER_DISCARD_STUDY = 0x040C # Discard Study\r\n" +
"WM_TRIGGER_CHECK_UPDATES = 0x040D # Check for Updates\r\n" +
"WM_TRIGGER_SHOW_PICK_LISTS = 0x040E # Show Pick Lists\r\n" +
"WM_TRIGGER_CREATE_CRITICAL_NOTE = 0x040F # Create Critical Note\r\n" +
"\r\n" +
"Example AHK:\r\n" +
"DetectHiddenWindows, On\r\n" +
"PostMessage, 0x0401, 0, 0,, ahk_class WindowsForms\r\n" +
"\r\n" +
"=== Debug Tips ===\r\n" +
"\r\n" +
"- Hold Win key while triggering Critical Findings\r\n" +
"  to see raw data without pasting.\r\n" +
"\r\n" +
"=== Config File ===\r\n" +
"Settings: %LOCALAPPDATA%\\MosaicTools\\MosaicToolsSettings.json"
        };
        Controls.Add(_infoBox);
        _nextY += 390;

        // Element Dump section
        AddSectionDivider("Element Dump (Debug)");

        // Target app dropdown
        AddLabel("Target App:", LeftMargin, _nextY + 3);
        _targetAppCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 100, _nextY),
            Width = 120,
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _targetAppCombo.Items.AddRange(new[] { "Clario", "Mosaic", "InteleViewer" });
        _targetAppCombo.SelectedIndex = 0;
        Controls.Add(_targetAppCombo);
        _nextY += RowHeight;

        // Method dropdown
        AddLabel("Method:", LeftMargin, _nextY + 3);
        _methodCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 100, _nextY),
            Width = 120,
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _methodCombo.Items.AddRange(new[] { "Both", "Name", "LegacyValue" });
        _methodCombo.SelectedIndex = 0;
        Controls.Add(_methodCombo);
        _nextY += RowHeight;

        // Dump button
        _dumpButton = new Button
        {
            Text = "Dump Elements to Clipboard",
            Location = new Point(LeftMargin, _nextY),
            Size = new Size(180, 28),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        _dumpButton.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
        _dumpButton.Click += OnDumpClick;
        Controls.Add(_dumpButton);

        _statusLabel = new Label
        {
            Location = new Point(LeftMargin + 190, _nextY + 5),
            AutoSize = true,
            ForeColor = Color.FromArgb(120, 200, 120),
            Font = new Font("Segoe UI", 9)
        };
        Controls.Add(_statusLabel);
        _nextY += 35;

        _toolTip.SetToolTip(_targetAppCombo, "Select the application to scan for UI elements.");
        _toolTip.SetToolTip(_methodCombo, "Name = UIA Name property\nLegacyValue = LegacyIAccessible.Value\nBoth = Compare both methods");
        _toolTip.SetToolTip(_dumpButton, "Scan all UI elements and copy to clipboard.\nUse to compare Name vs LegacyIAccessible.Value.");

        UpdateHeight();
    }

    private void OnDumpClick(object? sender, EventArgs e)
    {
        _statusLabel.Text = "Scanning...";
        _statusLabel.ForeColor = Color.FromArgb(200, 200, 120);
        _statusLabel.Refresh();

        try
        {
            var targetApp = _targetAppCombo.SelectedItem?.ToString() ?? "Clario";
            var method = _methodCombo.SelectedItem?.ToString() ?? "Both";

            using var automation = new AutomationService();
            var result = automation.DumpElements(targetApp, method);

            Clipboard.SetText(result);

            if (result.StartsWith("ERROR:"))
            {
                _statusLabel.Text = result;
                _statusLabel.ForeColor = Color.FromArgb(255, 120, 120);
            }
            else
            {
                // Count elements from the result
                var lines = result.Split('\n');
                var countLine = Array.Find(lines, l => l.StartsWith("Total elements"));
                _statusLabel.Text = countLine ?? "Copied to clipboard!";
                _statusLabel.ForeColor = Color.FromArgb(120, 200, 120);
            }
        }
        catch (Exception ex)
        {
            _statusLabel.Text = $"Error: {ex.Message}";
            _statusLabel.ForeColor = Color.FromArgb(255, 120, 120);
        }
    }

    public override void LoadSettings(Configuration config)
    {
        // Nothing to load - read-only reference
    }

    public override void SaveSettings(Configuration config)
    {
        // Nothing to save - read-only reference
    }
}
