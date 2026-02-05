using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Keys & Buttons section: Links to open Keys and IV Buttons dialogs.
/// </summary>
public class KeysButtonsSection : SettingsSection
{
    public override string SectionId => "keys";

    private readonly CheckBox _floatingToolbarCheck;
    private readonly ComboBox _microphoneCombo;
    private readonly Label _micStatusLabel;
    private readonly Configuration _config;
    private readonly ActionController _controller;

    public KeysButtonsSection(ToolTip toolTip, Configuration config, ActionController controller, bool isHeadless) : base("Keys & Buttons", toolTip)
    {
        _config = config;
        _controller = controller;

        // Microphone Selection
        AddSectionDivider("Dictation Microphone");

        AddLabel("Microphone:", LeftMargin, _nextY + 3);
        _microphoneCombo = AddComboBox(LeftMargin + 90, _nextY, 140,
            new[] { "Auto", "PowerMic", "SpeechMike" },
            "Select which dictation microphone to use.\nAuto will connect to any available device.");
        _nextY += SubRowHeight + 2;

        // Status label showing detected device
        _micStatusLabel = new Label
        {
            Location = new Point(LeftMargin, _nextY),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        Controls.Add(_micStatusLabel);
        UpdateMicrophoneStatus();
        _nextY += SubRowHeight + 5;

        // Keys Configuration
        AddSectionDivider("Action Mappings");

        AddLabel("Configure hotkeys and mic button mappings for actions.", LeftMargin, _nextY);
        _nextY += SubRowHeight;

        var openKeysBtn = AddButton("Open Keys Configuration...", LeftMargin, _nextY, 200, 28, OnOpenKeysClick,
            "Configure hotkeys and mic button mappings for all actions.");
        _nextY += RowHeight + 10;

        // IV Buttons (non-headless only)
        if (!isHeadless)
        {
            AddSectionDivider("InteleViewer Buttons");

            _floatingToolbarCheck = AddCheckBox("Show InteleViewer Buttons toolbar", LeftMargin, _nextY,
                "Shows configurable buttons for InteleViewer shortcuts\n(window/level presets, zoom, etc.).");
            _nextY += SubRowHeight;

            var openButtonsBtn = AddButton("Open Button Studio...", LeftMargin, _nextY, 200, 28, OnOpenButtonsClick,
                "Configure the InteleViewer floating button toolbar.");
            _nextY += RowHeight + 10;
        }
        else
        {
            _floatingToolbarCheck = new CheckBox { Visible = false };
        }

        // Hardcoded buttons info
        AddSectionDivider("Hardcoded Mic Buttons");

        // Column headers
        var powerMicHeader = AddLabel("PowerMic", LeftMargin, _nextY);
        powerMicHeader.Font = new Font("Segoe UI", 8, FontStyle.Bold);
        powerMicHeader.ForeColor = Color.LightGray;

        var speechMikeHeader = AddLabel("SpeechMike", LeftMargin + 180, _nextY);
        speechMikeHeader.Font = new Font("Segoe UI", 8, FontStyle.Bold);
        speechMikeHeader.ForeColor = Color.LightGray;
        _nextY += 18;

        // Function labels (center column)
        var functions = new[] { "Process Report", "Generate Impression", "Start/Stop Dictation", "Sign Report" };
        var powerMicButtons = new[] { "Skip Back", "Skip Forward", "Record Button", "Checkmark" };
        var speechMikeButtons = new[] { "Ins/Ovr", "-i-", "Record", "EoL" };

        for (int i = 0; i < functions.Length; i++)
        {
            var pmLabel = AddLabel($"{powerMicButtons[i]} →", LeftMargin, _nextY);
            pmLabel.ForeColor = Color.DarkGray;
            pmLabel.Font = new Font("Segoe UI", 8, FontStyle.Italic);
            pmLabel.Width = 90;

            var funcLabel = AddLabel(functions[i], LeftMargin + 90, _nextY);
            funcLabel.ForeColor = Color.Gray;
            funcLabel.Font = new Font("Segoe UI", 8);
            funcLabel.Width = 110;

            var smLabel = AddLabel($"← {speechMikeButtons[i]}", LeftMargin + 200, _nextY);
            smLabel.ForeColor = Color.DarkGray;
            smLabel.Font = new Font("Segoe UI", 8, FontStyle.Italic);

            _nextY += 16;
        }
        _nextY += 10;

        UpdateHeight();
    }

    private void UpdateMicrophoneStatus()
    {
        var connected = _controller.GetConnectedMicrophoneName();
        var available = HidService.GetAvailableDevices();

        if (!string.IsNullOrEmpty(connected))
        {
            _micStatusLabel.Text = $"Connected: {connected}";
            _micStatusLabel.ForeColor = Color.LightGreen;
        }
        else if (available.Count > 0)
        {
            var names = string.Join(", ", available.ConvertAll(d => d.Name));
            _micStatusLabel.Text = $"Available: {names}";
            _micStatusLabel.ForeColor = Color.Yellow;
        }
        else
        {
            _micStatusLabel.Text = "No dictation microphone detected";
            _micStatusLabel.ForeColor = Color.Gray;
        }
    }

    private void OnOpenKeysClick(object? sender, EventArgs e)
    {
        using var dialog = new KeyMappingsDialog(_config, _controller, App.IsHeadless);
        dialog.ShowDialog();
    }

    private void OnOpenButtonsClick(object? sender, EventArgs e)
    {
        using var dialog = new ButtonStudioDialog(_config);
        dialog.ShowDialog();
    }

    public override void LoadSettings(Configuration config)
    {
        _floatingToolbarCheck.Checked = config.FloatingToolbarEnabled;

        // Select microphone preference
        _microphoneCombo.SelectedIndex = config.PreferredMicrophone switch
        {
            HidService.DevicePowerMic => 1,
            HidService.DeviceSpeechMike => 2,
            _ => 0  // Auto
        };

        UpdateMicrophoneStatus();
    }

    public override void SaveSettings(Configuration config)
    {
        config.FloatingToolbarEnabled = _floatingToolbarCheck.Checked;

        config.PreferredMicrophone = _microphoneCombo.SelectedIndex switch
        {
            1 => HidService.DevicePowerMic,
            2 => HidService.DeviceSpeechMike,
            _ => HidService.DeviceAuto
        };
    }
}
