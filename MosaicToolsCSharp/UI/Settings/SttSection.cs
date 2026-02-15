// [CustomSTT] Settings section for custom speech-to-text
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Settings section for Custom STT Mode configuration.
/// </summary>
public class SttSection : SettingsSection
{
    public override string SectionId => "stt";

    private readonly CheckBox _enabledCheck;
    private readonly ComboBox _providerCombo;
    private readonly TextBox _apiKeyBox;
    private readonly Button _getKeyButton;
    private readonly ComboBox _audioDeviceCombo;
    private readonly CheckBox _autoPunctuateCheck;
    private readonly CheckBox _startBeepCheck;
    private readonly CheckBox _stopBeepCheck;
    private readonly TrackBar _startBeepVolume;
    private readonly TrackBar _stopBeepVolume;
    private readonly Label _startBeepVolLabel;
    private readonly Label _stopBeepVolLabel;
    private readonly CheckBox _showIndicatorCheck;

    // Hotkey capture boxes
    private readonly TextBox _dictateHotkeyBox;
    private readonly TextBox _processHotkeyBox;
    private readonly TextBox _signHotkeyBox;

    public SttSection(ToolTip toolTip) : base("Speech-to-Text", toolTip)
    {
        _searchTerms.AddRange(new[] { "stt", "speech", "transcription", "deepgram", "dictation", "custom stt", "beep", "punctuation" });

        _enabledCheck = AddCheckBox("Enable Custom STT Mode", LeftMargin, _nextY,
            "Use cloud STT instead of Mosaic's built-in speech recognition. Cancel Mosaic's WebHID prompt when enabled.");
        _enabledCheck.CheckedChanged += (_, _) => UpdateControlStates();
        _nextY += RowHeight;

        AddHintLabel("Cancel Mosaic's WebHID prompt when this is enabled", LeftMargin + 25);

        // Provider
        AddSectionDivider("Provider");

        AddLabel("Provider:", LeftMargin + 25, _nextY + 3);
        _providerCombo = AddComboBox(LeftMargin + 110, _nextY, 200,
            new[] { "Deepgram Nova-3 Medical", "Deepgram Nova-3" });
        _nextY += RowHeight;

        // API Key
        AddLabel("API Key:", LeftMargin + 25, _nextY + 3);
        _apiKeyBox = AddTextBox(LeftMargin + 110, _nextY, 200);
        _apiKeyBox.UseSystemPasswordChar = true;

        _getKeyButton = AddButton("Get API Key", LeftMargin + 320, _nextY - 1, 90, 24, OnGetKeyClick,
            "Opens Deepgram signup page to create a free account and get an API key.");

        // Help button (?) with full instructions
        var helpButton = AddButton("?", LeftMargin + 415, _nextY - 1, 24, 24, OnApiKeyHelpClick,
            "How to get a Deepgram API key");
        helpButton.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        _nextY += SubRowHeight;

        AddHintLabel("Free tier: $200 credit. Nova-3 Medical: $0.0077/min streaming", LeftMargin + 25);

        // Audio Device
        AddSectionDivider("Audio");

        AddLabel("Audio Device:", LeftMargin + 25, _nextY + 3);
        _audioDeviceCombo = AddComboBox(LeftMargin + 120, _nextY, 250, new[] { "(Auto-detect PowerMic)" });
        _nextY += RowHeight;

        // Populate audio devices
        try
        {
            var devices = SttService.GetAudioDevices();
            foreach (var device in devices)
            {
                _audioDeviceCombo.Items.Add(device);
            }
        }
        catch { /* NAudio not available in design mode */ }

        // Keyboard Shortcuts
        AddSectionDivider("Keyboard Shortcuts");

        AddLabel("Dictate:", LeftMargin + 25, _nextY + 3);
        _dictateHotkeyBox = CreateHotkeyBox(LeftMargin + 140, _nextY, 140,
            "Keyboard shortcut to start/stop dictation. Press Escape to clear.");
        _nextY += RowHeight;

        AddLabel("Process Report:", LeftMargin + 25, _nextY + 3);
        _processHotkeyBox = CreateHotkeyBox(LeftMargin + 140, _nextY, 140,
            "Keyboard shortcut to process report. Press Escape to clear.");
        _nextY += RowHeight;

        AddLabel("Sign Report:", LeftMargin + 25, _nextY + 3);
        _signHotkeyBox = CreateHotkeyBox(LeftMargin + 140, _nextY, 140,
            "Keyboard shortcut to sign report. Press Escape to clear.");
        _nextY += SubRowHeight;

        AddHintLabel("Click box then press key combo. Escape to clear.", LeftMargin + 25);

        // Punctuation
        AddSectionDivider("Punctuation");

        _autoPunctuateCheck = AddCheckBox("Auto-punctuate", LeftMargin + 25, _nextY,
            "Automatically insert punctuation. When off, say \"period\", \"comma\", etc. to punctuate.");
        _nextY += SubRowHeight;

        AddHintLabel("Off = dictate punctuation (say \"period\", \"comma\"). On = auto-inserted.", LeftMargin + 25);

        // Beep Settings
        AddSectionDivider("Beeps");

        _startBeepCheck = AddCheckBox("Start beep", LeftMargin + 25, _nextY,
            "Play beep when recording starts.");
        _startBeepVolLabel = AddLabel("Vol:", LeftMargin + 160, _nextY + 3);
        _startBeepVolume = AddTrackBar(LeftMargin + 190, _nextY, 120, 0, 100, 8,
            "Start beep volume");
        _nextY += RowHeight;

        _stopBeepCheck = AddCheckBox("Stop beep", LeftMargin + 25, _nextY,
            "Play beep when recording stops.");
        _stopBeepVolLabel = AddLabel("Vol:", LeftMargin + 160, _nextY + 3);
        _stopBeepVolume = AddTrackBar(LeftMargin + 190, _nextY, 120, 0, 100, 8,
            "Stop beep volume");
        _nextY += RowHeight;

        // Display options
        AddSectionDivider("Display");

        _showIndicatorCheck = AddCheckBox("Show dictation indicator", LeftMargin + 25, _nextY,
            "Show a floating indicator with live transcription text while dictating.");
        _nextY += RowHeight;

        UpdateHeight();
    }

    /// <summary>
    /// Create a hotkey capture text box matching the KeyMappingsDialog pattern.
    /// </summary>
    private TextBox CreateHotkeyBox(int x, int y, int width, string? tooltip)
    {
        var box = new TextBox
        {
            Location = new Point(x, y),
            Width = width,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            ReadOnly = true,
            Cursor = Cursors.Hand,
            BorderStyle = BorderStyle.FixedSingle
        };

        // Let Alt combinations reach the KeyDown handler
        box.PreviewKeyDown += (_, e) => e.IsInputKey = true;

        box.KeyDown += (_, e) =>
        {
            e.SuppressKeyPress = true;
            e.Handled = true;

            // Escape clears
            if (e.KeyCode == Keys.Escape)
            {
                box.Text = "";
                return;
            }

            var parts = new List<string>();
            if (e.Control) parts.Add("Ctrl");
            if (e.Alt) parts.Add("Alt");
            if (e.Shift) parts.Add("Shift");

            if (e.KeyCode != Keys.ControlKey && e.KeyCode != Keys.Menu &&
                e.KeyCode != Keys.ShiftKey && e.KeyCode != Keys.None)
            {
                parts.Add(KeyCodeToDisplayName(e.KeyCode));
                box.Text = string.Join("+", parts);
            }
        };

        box.Click += (_, _) => box.SelectAll();

        Controls.Add(box);

        if (tooltip != null)
        {
            _toolTip.SetToolTip(box, tooltip);
        }

        return box;
    }

    /// <summary>
    /// Convert WinForms Keys enum to display name matching KeyboardService.VKToName format.
    /// </summary>
    private static string KeyCodeToDisplayName(Keys keyCode)
    {
        if (keyCode >= Keys.D0 && keyCode <= Keys.D9)
            return ((char)('0' + (keyCode - Keys.D0))).ToString();
        if (keyCode >= Keys.NumPad0 && keyCode <= Keys.NumPad9)
            return ((char)('0' + (keyCode - Keys.NumPad0))).ToString();
        return keyCode.ToString();
    }

    private void UpdateControlStates()
    {
        bool enabled = _enabledCheck.Checked;
        _providerCombo.Enabled = enabled;
        _apiKeyBox.Enabled = enabled;
        _getKeyButton.Enabled = enabled;
        _audioDeviceCombo.Enabled = enabled;
        _autoPunctuateCheck.Enabled = enabled;
        _dictateHotkeyBox.Enabled = enabled;
        _processHotkeyBox.Enabled = enabled;
        _signHotkeyBox.Enabled = enabled;
        _startBeepCheck.Enabled = enabled;
        _stopBeepCheck.Enabled = enabled;
        _startBeepVolume.Enabled = enabled;
        _stopBeepVolume.Enabled = enabled;
        _showIndicatorCheck.Enabled = enabled;
    }

    private void OnGetKeyClick(object? sender, EventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo("https://console.deepgram.com/signup") { UseShellExecute = true });
        }
        catch { }
    }

    private void OnApiKeyHelpClick(object? sender, EventArgs e)
    {
        MessageBox.Show(
            "How to get a Deepgram API key:\n\n" +
            "1. Click \"Get API Key\" to open the Deepgram signup page\n" +
            "2. Create a free account (email or Google/GitHub login)\n" +
            "3. After signing in, you'll land on the Deepgram Console\n" +
            "4. Click \"API Keys\" in the left sidebar (or go to Settings > API Keys)\n" +
            "5. Click \"Create a New API Key\"\n" +
            "6. Give it a name (e.g. \"MosaicTools\"), leave permissions as default\n" +
            "7. Click \"Create Key\" and copy the key that appears\n" +
            "8. Paste the key into the API Key field above\n\n" +
            "Free tier includes $200 in credit — enough for ~430 hours of\n" +
            "Nova-3 Medical dictation at $0.0077/min.\n\n" +
            "The key starts with \"dg_\" or a similar prefix.\n" +
            "Keep it private — do not share it.",
            "Deepgram API Key Setup",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    public override void LoadSettings(Configuration config)
    {
        _enabledCheck.Checked = config.CustomSttEnabled;

        _providerCombo.SelectedIndex = config.SttModel switch
        {
            "nova-3-medical" => 0,
            "nova-3" => 1,
            _ => 0
        };

        _apiKeyBox.Text = config.SttApiKey;

        // Select audio device
        if (string.IsNullOrEmpty(config.SttAudioDeviceName))
        {
            _audioDeviceCombo.SelectedIndex = 0; // Auto-detect
        }
        else
        {
            int idx = _audioDeviceCombo.FindStringExact(config.SttAudioDeviceName);
            _audioDeviceCombo.SelectedIndex = idx >= 0 ? idx : 0;
        }

        // Load hotkeys from action mappings
        _dictateHotkeyBox.Text = config.ActionMappings.GetValueOrDefault(Actions.ToggleRecord)?.Hotkey ?? "";
        _processHotkeyBox.Text = config.ActionMappings.GetValueOrDefault(Actions.ProcessReport)?.Hotkey ?? "";
        _signHotkeyBox.Text = config.ActionMappings.GetValueOrDefault(Actions.SignReport)?.Hotkey ?? "";

        _autoPunctuateCheck.Checked = config.SttAutoPunctuate;
        _startBeepCheck.Checked = config.SttStartBeepEnabled;
        _stopBeepCheck.Checked = config.SttStopBeepEnabled;
        _startBeepVolume.Value = Math.Clamp((int)(config.SttStartBeepVolume * 100), 0, 100);
        _stopBeepVolume.Value = Math.Clamp((int)(config.SttStopBeepVolume * 100), 0, 100);
        _showIndicatorCheck.Checked = config.SttShowIndicator;

        UpdateControlStates();
    }

    public override void SaveSettings(Configuration config)
    {
        config.CustomSttEnabled = _enabledCheck.Checked;
        config.SttProvider = "deepgram"; // Only provider for now

        config.SttModel = _providerCombo.SelectedIndex switch
        {
            0 => "nova-3-medical",
            1 => "nova-3",
            _ => "nova-3-medical"
        };

        config.SttApiKey = _apiKeyBox.Text.Trim();

        // Save audio device (empty = auto-detect)
        if (_audioDeviceCombo.SelectedIndex <= 0)
        {
            config.SttAudioDeviceName = "";
        }
        else
        {
            config.SttAudioDeviceName = _audioDeviceCombo.SelectedItem?.ToString() ?? "";
        }

        // Save hotkeys to action mappings
        SaveHotkeyToMapping(config, Actions.ToggleRecord, _dictateHotkeyBox.Text);
        SaveHotkeyToMapping(config, Actions.ProcessReport, _processHotkeyBox.Text);
        SaveHotkeyToMapping(config, Actions.SignReport, _signHotkeyBox.Text);

        config.SttAutoPunctuate = _autoPunctuateCheck.Checked;
        config.SttStartBeepEnabled = _startBeepCheck.Checked;
        config.SttStopBeepEnabled = _stopBeepCheck.Checked;
        config.SttStartBeepVolume = _startBeepVolume.Value / 100.0;
        config.SttStopBeepVolume = _stopBeepVolume.Value / 100.0;
        config.SttShowIndicator = _showIndicatorCheck.Checked;
    }

    /// <summary>
    /// Save a hotkey to the action mappings, preserving existing mic button mapping.
    /// </summary>
    private static void SaveHotkeyToMapping(Configuration config, string action, string hotkey)
    {
        var trimmed = string.IsNullOrWhiteSpace(hotkey) ? null : hotkey.Trim();
        var existing = config.ActionMappings.GetValueOrDefault(action);

        if (trimmed != null || existing?.MicButton != null)
        {
            config.ActionMappings[action] = new ActionMapping
            {
                Hotkey = trimmed ?? "",
                MicButton = existing?.MicButton ?? ""
            };
        }
    }
}
