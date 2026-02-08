using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Desktop mode settings: Hotkey, recording indicator, audio feedback.
/// Only shown in non-headless mode.
/// </summary>
public class DesktopSection : SettingsSection
{
    public override string SectionId => "desktop";

    private readonly TextBox _ivHotkeyBox;
    private readonly CheckBox _indicatorCheck;
    private readonly CheckBox _hideIndicatorWhenNoStudyCheck;
    private readonly CheckBox _autoStopCheck;
    private readonly CheckBox _deadManCheck;
    private readonly CheckBox _startBeepCheck;
    private readonly CheckBox _stopBeepCheck;
    private readonly TrackBar _startVolumeSlider;
    private readonly TrackBar _stopVolumeSlider;
    private readonly Label _startVolLabel;
    private readonly Label _stopVolLabel;
    private readonly NumericUpDown _dictationPauseNum;
    private readonly CheckBox _autoUpdateCheck;

    private readonly MainForm _mainForm;

    public DesktopSection(ToolTip toolTip, MainForm mainForm) : base("Desktop Mode", toolTip)
    {
        _mainForm = mainForm;

        // IV Report Hotkey
        AddLabel("IV Report Hotkey:", LeftMargin, _nextY + 3);
        _ivHotkeyBox = AddTextBox(LeftMargin + 130, _nextY, 120,
            "Hotkey to copy report from InteleViewer.\nClick to capture a new key combination.");
        _ivHotkeyBox.ReadOnly = true;
        _ivHotkeyBox.Cursor = Cursors.Hand;
        SetupHotkeyCapture(_ivHotkeyBox);
        _nextY += RowHeight + 5;

        // Recording Indicator
        _indicatorCheck = AddCheckBox("Show Recording Indicator", LeftMargin, _nextY,
            "Shows a small colored rectangle indicating dictation state\n(red = recording, gray = stopped).");
        _indicatorCheck.CheckedChanged += (s, e) => UpdateIndicatorSubState();
        _nextY += SubRowHeight;

        _hideIndicatorWhenNoStudyCheck = AddCheckBox("Hide when no study open", LeftMargin + 25, _nextY,
            "Hides the recording indicator when no study is open in Mosaic.", true);
        _hideIndicatorWhenNoStudyCheck.CheckedChanged += (s, e) =>
        {
            // Live update
            _mainForm.UpdateIndicatorVisibility();
        };
        _nextY += RowHeight;

        // Auto-Stop
        _autoStopCheck = AddCheckBox("Auto-Stop Dictation on Process Report", LeftMargin, _nextY,
            "Automatically stops dictation when you press Process Report.");
        _nextY += RowHeight;

        // Dead Man's Switch
        _deadManCheck = AddCheckBox("Push-to-Talk (Dead Man's Switch)", LeftMargin, _nextY,
            "Hold the Record button to dictate, release to stop.\nDead man's switch style dictation.");
        _nextY += RowHeight + 5;

        // Audio Feedback section
        AddSectionDivider("Audio Feedback");

        // Start Beep row
        _startBeepCheck = AddCheckBox("Start Beep", LeftMargin, _nextY,
            "Plays an audio beep when dictation starts.");
        _startVolumeSlider = AddTrackBar(LeftMargin + 110, _nextY, 130, 0, 100, 50);
        _startVolLabel = AddLabel("50%", LeftMargin + 250, _nextY + 3);
        _startVolumeSlider.ValueChanged += (s, e) => _startVolLabel.Text = $"{_startVolumeSlider.Value}%";
        _nextY += RowHeight;

        // Stop Beep row
        _stopBeepCheck = AddCheckBox("Stop Beep", LeftMargin, _nextY,
            "Plays an audio beep when dictation stops.");
        _stopVolumeSlider = AddTrackBar(LeftMargin + 110, _nextY, 130, 0, 100, 50);
        _stopVolLabel = AddLabel("50%", LeftMargin + 250, _nextY + 3);
        _stopVolumeSlider.ValueChanged += (s, e) => _stopVolLabel.Text = $"{_stopVolumeSlider.Value}%";
        _nextY += RowHeight;

        // Dictation pause
        AddLabel("Start Beep Pause:", LeftMargin, _nextY + 3);
        _dictationPauseNum = AddNumericUpDown(LeftMargin + 130, _nextY, 70, 100, 5000, 1000,
            "Delay before playing start beep, to avoid false triggers.\nRecommended: 800-1200ms.");
        AddLabel("ms", LeftMargin + 205, _nextY + 3);
        _nextY += RowHeight + 5;

        // Updates section
        AddSectionDivider("Updates");

        _autoUpdateCheck = AddCheckBox("Auto-update on startup", LeftMargin, _nextY,
            "Automatically check for and install updates on startup.");

        var checkUpdatesBtn = AddButton("Check Now", LeftMargin + 200, _nextY - 2, 100, 24, async (s, e) =>
        {
            var btn = (Button)s!;
            btn.Enabled = false;
            btn.Text = "Checking...";
            try
            {
                await _mainForm.CheckForUpdatesManualAsync();
            }
            finally
            {
                if (!btn.IsDisposed)
                {
                    btn.Text = "Check Now";
                    btn.Enabled = true;
                }
            }
        }, "Manually check for available updates now.");
        _nextY += RowHeight;

        UpdateHeight();
    }

    private void UpdateIndicatorSubState()
    {
        _hideIndicatorWhenNoStudyCheck.Enabled = _indicatorCheck.Checked;
        _hideIndicatorWhenNoStudyCheck.ForeColor = _indicatorCheck.Checked
            ? Color.FromArgb(180, 180, 180)
            : Color.FromArgb(100, 100, 100);
    }

    private void SetupHotkeyCapture(TextBox box)
    {
        // Ensure Alt key combinations reach the TextBox KeyDown handler
        box.PreviewKeyDown += (s, e) => e.IsInputKey = true;

        box.KeyDown += (s, e) =>
        {
            e.SuppressKeyPress = true;
            e.Handled = true;

            var parts = new System.Collections.Generic.List<string>();
            if (e.Control) parts.Add("Ctrl");
            if (e.Alt) parts.Add("Alt");
            if (e.Shift) parts.Add("Shift");

            // Only accept if a non-modifier key is pressed
            if (e.KeyCode != Keys.ControlKey && e.KeyCode != Keys.Menu &&
                e.KeyCode != Keys.ShiftKey && e.KeyCode != Keys.None)
            {
                parts.Add(KeyCodeToDisplayName(e.KeyCode));
                box.Text = string.Join("+", parts);
            }
        };

        box.Click += (s, e) =>
        {
            box.SelectAll();
        };
    }

    /// <summary>
    /// Convert WinForms Keys enum to display name matching KeyboardService.VKToName format.
    /// Keys.D0-D9 become "0"-"9" instead of "D0"-"D9".
    /// </summary>
    private static string KeyCodeToDisplayName(Keys keyCode)
    {
        if (keyCode >= Keys.D0 && keyCode <= Keys.D9)
            return ((char)('0' + (keyCode - Keys.D0))).ToString();
        if (keyCode >= Keys.NumPad0 && keyCode <= Keys.NumPad9)
            return ((char)('0' + (keyCode - Keys.NumPad0))).ToString();
        return keyCode.ToString();
    }

    public override void LoadSettings(Configuration config)
    {
        _ivHotkeyBox.Text = config.IvReportHotkey ?? "";
        _indicatorCheck.Checked = config.IndicatorEnabled;
        _hideIndicatorWhenNoStudyCheck.Checked = config.HideIndicatorWhenNoStudy;
        _autoStopCheck.Checked = config.AutoStopDictation;
        _deadManCheck.Checked = config.DeadManSwitch;
        _startBeepCheck.Checked = config.StartBeepEnabled;
        _stopBeepCheck.Checked = config.StopBeepEnabled;
        _startVolumeSlider.Value = VolumeToSlider(config.StartBeepVolume);
        _stopVolumeSlider.Value = VolumeToSlider(config.StopBeepVolume);
        _dictationPauseNum.Value = config.DictationPauseMs;
        _autoUpdateCheck.Checked = config.AutoUpdateEnabled;

        UpdateIndicatorSubState();
    }

    public override void SaveSettings(Configuration config)
    {
        config.IvReportHotkey = _ivHotkeyBox.Text;
        config.IndicatorEnabled = _indicatorCheck.Checked;
        config.HideIndicatorWhenNoStudy = _hideIndicatorWhenNoStudyCheck.Checked;
        config.AutoStopDictation = _autoStopCheck.Checked;
        config.DeadManSwitch = _deadManCheck.Checked;
        config.StartBeepEnabled = _startBeepCheck.Checked;
        config.StopBeepEnabled = _stopBeepCheck.Checked;
        config.StartBeepVolume = SliderToVolume(_startVolumeSlider.Value);
        config.StopBeepVolume = SliderToVolume(_stopVolumeSlider.Value);
        config.DictationPauseMs = (int)_dictationPauseNum.Value;
        config.AutoUpdateEnabled = _autoUpdateCheck.Checked;
    }

    // Volume slider uses logarithmic scale for natural feel
    private static double SliderToVolume(int sliderValue)
    {
        if (sliderValue <= 0) return 0;
        return (Math.Pow(100, sliderValue / 100.0) - 1) / 99.0;
    }

    private static int VolumeToSlider(double volume)
    {
        if (volume <= 0) return 0;
        return (int)Math.Round(50.0 * Math.Log10(volume * 99.0 + 1.0));
    }
}
