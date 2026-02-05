using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Dialog for configuring action-to-hotkey and mic button mappings.
/// </summary>
public class KeyMappingsDialog : Form
{
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    private readonly Configuration _config;
    private readonly ActionController _controller;
    private readonly bool _isHeadless;
    private readonly Dictionary<string, TextBox> _hotkeyBoxes = new();
    private readonly Dictionary<string, ComboBox> _micCombos = new();
    private ComboBox _deviceCombo = null!;
    private Label _deviceStatusLabel = null!;

    public KeyMappingsDialog(Configuration config, ActionController controller, bool isHeadless)
    {
        _config = config;
        _controller = controller;
        _isHeadless = isHeadless;

        InitializeUI();
        LoadSettings();
    }

    private Dictionary<string, ActionMapping> GetCurrentMappings()
    {
        return _deviceCombo.SelectedIndex == 1 ? _config.SpeechMikeActionMappings : _config.ActionMappings;
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        try
        {
            int value = 1;
            DwmSetWindowAttribute(Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
        }
        catch { }
    }

    private void InitializeUI()
    {
        Text = "Key Mappings";
        Size = new Size(500, 580);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.FromArgb(30, 30, 30);
        ForeColor = Color.White;

        var contentPanel = new Panel
        {
            Location = new Point(0, 0),
            Size = new Size(Width, Height - 80),
            AutoScroll = true,
            BackColor = Color.FromArgb(30, 30, 30)
        };
        Controls.Add(contentPanel);

        int y = 15;

        // Device selector
        contentPanel.Controls.Add(new Label
        {
            Text = "Configure mappings for:",
            Location = new Point(20, y + 3),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });

        _deviceCombo = new ComboBox
        {
            Location = new Point(170, y),
            Width = 130,
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _deviceCombo.Items.AddRange(new[] { "PowerMic", "SpeechMike" });
        _deviceCombo.SelectedIndex = _config.PreferredMicrophone == HidService.DeviceSpeechMike ? 1 : 0;
        _deviceCombo.SelectedIndexChanged += (s, e) => LoadSettings();
        contentPanel.Controls.Add(_deviceCombo);

        _deviceStatusLabel = new Label
        {
            Location = new Point(310, y + 3),
            AutoSize = true,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        UpdateDeviceStatus();
        contentPanel.Controls.Add(_deviceStatusLabel);

        y += 35;

        // Separator
        contentPanel.Controls.Add(new Label
        {
            Text = "─────────────────────────────────────────────────────",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(60, 60, 60)
        });
        y += 20;

        // Headers
        contentPanel.Controls.Add(new Label
        {
            Text = "Action",
            Location = new Point(20, y),
            Width = 150,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        });

        if (!_isHeadless)
        {
            contentPanel.Controls.Add(new Label
            {
                Text = "Hotkey",
                Location = new Point(180, y),
                Width = 120,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            });
        }

        contentPanel.Controls.Add(new Label
        {
            Text = "Mic Button",
            Location = new Point(_isHeadless ? 180 : 310, y),
            Width = 120,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        });

        y += 30;

        // Action descriptions for tooltips
        var actionDescriptions = new Dictionary<string, string>
        {
            [Actions.SystemBeep] = "Plays an audio beep to indicate dictation state change.",
            [Actions.GetPrior] = "Extract prior study from InteleViewer, format, and paste to Mosaic.",
            [Actions.CriticalFindings] = "Scrape Clario for exam note, extract contact info, and paste to Mosaic.\nHold Win key for debug mode.",
            [Actions.ShowReport] = "Copy current report from Mosaic and display in popup window.",
            [Actions.CaptureSeries] = "Use OCR to read series/image numbers from InteleViewer and paste to Mosaic.",
            [Actions.ToggleRecord] = "Toggle dictation on/off in Mosaic (sends Alt+R).",
            [Actions.ProcessReport] = "Process report with RadPair (sends Alt+P to Mosaic).",
            [Actions.SignReport] = "Sign/finalize the report (sends Alt+F to Mosaic).",
            [Actions.CreateImpression] = "Click the Create Impression button in Mosaic.",
            [Actions.DiscardStudy] = "Close current study without signing.",
            [Actions.ShowPickLists] = "Open pick list popup for quick text insertion.",
            [Actions.CycleWindowLevel] = "Cycle through window/level presets in InteleViewer.",
            [Actions.CreateCriticalNote] = "Create Critical Communication Note in Clario for current study."
        };

        var toolTip = new ToolTip
        {
            AutoPopDelay = 15000,
            InitialDelay = 300,
            ReshowDelay = 100,
            ShowAlways = true
        };

        foreach (var action in Actions.All)
        {
            if (action == Actions.None) continue;
            if (_isHeadless && action == Actions.CycleWindowLevel) continue;

            var lbl = new Label
            {
                Text = action,
                Location = new Point(20, y + 3),
                AutoSize = true,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9)
            };
            contentPanel.Controls.Add(lbl);

            if (actionDescriptions.TryGetValue(action, out var desc))
            {
                toolTip.SetToolTip(lbl, desc);
            }

            if (!_isHeadless)
            {
                var hotkeyBox = new TextBox
                {
                    Location = new Point(180, y),
                    Width = 120,
                    BackColor = Color.FromArgb(60, 60, 60),
                    ForeColor = Color.White,
                    ReadOnly = true,
                    Cursor = Cursors.Hand,
                    BorderStyle = BorderStyle.FixedSingle
                };
                SetupHotkeyCapture(hotkeyBox);
                contentPanel.Controls.Add(hotkeyBox);
                _hotkeyBoxes[action] = hotkeyBox;
            }

            var micCombo = new ComboBox
            {
                Location = new Point(_isHeadless ? 180 : 310, y),
                Width = 140,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            // Buttons will be populated by LoadSettings() based on selected device
            contentPanel.Controls.Add(micCombo);
            _micCombos[action] = micCombo;

            y += 30;
        }

        // Button panel
        var buttonPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            BackColor = Color.FromArgb(40, 40, 40)
        };
        Controls.Add(buttonPanel);

        var saveBtn = new Button
        {
            Text = "Save",
            Size = new Size(80, 30),
            Location = new Point(Width - 200, 10),
            BackColor = Color.FromArgb(51, 102, 0),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        saveBtn.Click += (s, e) => SaveAndClose();
        buttonPanel.Controls.Add(saveBtn);

        var cancelBtn = new Button
        {
            Text = "Cancel",
            Size = new Size(80, 30),
            Location = new Point(Width - 110, 10),
            BackColor = Color.FromArgb(102, 51, 0),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        cancelBtn.Click += (s, e) => Close();
        buttonPanel.Controls.Add(cancelBtn);
    }

    private void SetupHotkeyCapture(TextBox box)
    {
        box.KeyDown += (s, e) =>
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
                parts.Add(e.KeyCode.ToString());
                box.Text = string.Join("+", parts);
            }
        };

        box.Click += (s, e) => box.SelectAll();
    }

    private void UpdateDeviceStatus()
    {
        var connected = _controller.GetConnectedMicrophoneName();
        if (!string.IsNullOrEmpty(connected))
        {
            bool isSelected = (_deviceCombo.SelectedIndex == 0 && !connected.Contains("Speech")) ||
                              (_deviceCombo.SelectedIndex == 1 && connected.Contains("Speech"));
            _deviceStatusLabel.Text = isSelected ? $"(Connected: {connected})" : $"(Other connected: {connected})";
            _deviceStatusLabel.ForeColor = isSelected ? Color.LightGreen : Color.Yellow;
        }
        else
        {
            _deviceStatusLabel.Text = "(Not connected)";
            _deviceStatusLabel.ForeColor = Color.Gray;
        }
    }

    private string[] GetButtonsForSelectedDevice()
    {
        return _deviceCombo.SelectedIndex == 1
            ? HidService.SpeechMikeButtons
            : HidService.PowerMicButtons;
    }

    private void LoadSettings()
    {
        UpdateDeviceStatus();
        var micMappings = GetCurrentMappings();  // Device-specific for mic buttons
        var buttons = GetButtonsForSelectedDevice();

        foreach (var action in Actions.All)
        {
            if (action == Actions.None) continue;

            // Hotkeys are always from PowerMic mappings (shared across devices)
            if (_hotkeyBoxes.TryGetValue(action, out var hotkeyBox))
            {
                var hotkeyMapping = _config.ActionMappings.GetValueOrDefault(action);
                hotkeyBox.Text = hotkeyMapping?.Hotkey ?? "";
            }

            // Mic buttons are device-specific
            if (_micCombos.TryGetValue(action, out var micCombo))
            {
                var micMapping = micMappings.GetValueOrDefault(action);
                var currentSelection = micMapping?.MicButton ?? "";
                micCombo.Items.Clear();
                micCombo.Items.Add("");
                micCombo.Items.AddRange(buttons);

                // Select the current mapping if it exists in this device's button list
                micCombo.SelectedItem = micCombo.Items.Contains(currentSelection) ? currentSelection : "";
            }
        }
    }

    private void SaveAndClose()
    {
        // Save to both mappings - hotkeys are shared, mic buttons are per-device
        SaveMappingsForDevice(_config.ActionMappings, false);
        SaveMappingsForDevice(_config.SpeechMikeActionMappings, true);

        _config.Save();
        _controller.RefreshServices();
        DialogResult = DialogResult.OK;
        Close();
    }

    private void SaveMappingsForDevice(Dictionary<string, ActionMapping> targetMappings, bool isSpeechMike)
    {
        bool isCurrentDevice = (_deviceCombo.SelectedIndex == 1) == isSpeechMike;

        foreach (var action in Actions.All)
        {
            if (action == Actions.None) continue;

            string? hotkey = null;
            string? micButton = null;

            // Hotkeys are always taken from the UI (shared across devices)
            if (_hotkeyBoxes.TryGetValue(action, out var hotkeyBox))
            {
                hotkey = string.IsNullOrWhiteSpace(hotkeyBox.Text) ? null : hotkeyBox.Text;
            }

            // Mic buttons: only update from UI if this is the currently selected device
            if (isCurrentDevice && _micCombos.TryGetValue(action, out var micCombo))
            {
                micButton = micCombo.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(micButton)) micButton = null;
            }
            else
            {
                // Keep existing mic button for the other device
                micButton = targetMappings.GetValueOrDefault(action)?.MicButton;
            }

            if (hotkey != null || micButton != null)
            {
                targetMappings[action] = new ActionMapping { Hotkey = hotkey, MicButton = micButton };
            }
            else
            {
                targetMappings.Remove(action);
            }
        }
    }
}
