using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Reference section: AHK integration docs, debug tips.
/// </summary>
public class ReferenceSection : SettingsSection
{
    public override string SectionId => "reference";

    private readonly TextBox _infoBox;

    public ReferenceSection(ToolTip toolTip) : base("Reference", toolTip)
    {
        _infoBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            WordWrap = false,
            ScrollBars = ScrollBars.Vertical,
            Location = new Point(LeftMargin, _nextY),
            Size = new Size(450, 480),
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
"=== Critical Note Creation ===\r\n" +
"\r\n" +
"Create a Critical Communication Note in Clario:\r\n" +
"- Map 'Create Critical Note' to hotkey/mic button\r\n" +
"- Ctrl+Click on Clinical History window\r\n" +
"- Right-click Clinical History -> Create Critical Note\r\n" +
"- Windows message 0x040F from AHK\r\n" +
"\r\n" +
"=== Debug Tips ===\r\n" +
"\r\n" +
"- Hold Win key while triggering Critical Findings\r\n" +
"  to see raw data without pasting.\r\n" +
"\r\n" +
"- Right-click Clinical History or Impression\r\n" +
"  windows for context menu and debug options.\r\n" +
"\r\n" +
"=== Config File ===\r\n" +
"Settings: %LOCALAPPDATA%\\MosaicTools\\MosaicToolsSettings.json"
        };
        Controls.Add(_infoBox);
        _nextY += 490;

        UpdateHeight();
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
