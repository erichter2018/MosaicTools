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

    public KeyMappingsDialog(Configuration config, ActionController controller, bool isHeadless)
    {
        _config = config;
        _controller = controller;
        _isHeadless = isHeadless;

        InitializeUI();
        LoadSettings();
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
        Size = new Size(500, 520);
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
            micCombo.Items.Add("");
            micCombo.Items.AddRange(HidService.AllButtons);
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

    private void LoadSettings()
    {
        foreach (var action in Actions.All)
        {
            if (action == Actions.None) continue;

            var mapping = _config.ActionMappings.GetValueOrDefault(action);

            if (_hotkeyBoxes.TryGetValue(action, out var hotkeyBox))
            {
                hotkeyBox.Text = mapping?.Hotkey ?? "";
            }

            if (_micCombos.TryGetValue(action, out var micCombo))
            {
                var mic = mapping?.MicButton ?? "";
                micCombo.SelectedItem = micCombo.Items.Contains(mic) ? mic : "";
            }
        }
    }

    private void SaveAndClose()
    {
        foreach (var action in Actions.All)
        {
            if (action == Actions.None) continue;

            string? hotkey = null;
            string? micButton = null;

            if (_hotkeyBoxes.TryGetValue(action, out var hotkeyBox))
            {
                hotkey = string.IsNullOrWhiteSpace(hotkeyBox.Text) ? null : hotkeyBox.Text;
            }

            if (_micCombos.TryGetValue(action, out var micCombo))
            {
                micButton = micCombo.SelectedItem?.ToString();
                if (string.IsNullOrWhiteSpace(micButton)) micButton = null;
            }

            if (hotkey != null || micButton != null)
            {
                _config.ActionMappings[action] = new ActionMapping { Hotkey = hotkey, MicButton = micButton };
            }
            else
            {
                _config.ActionMappings.Remove(action);
            }
        }

        _config.Save();
        _controller.RefreshServices();
        DialogResult = DialogResult.OK;
        Close();
    }
}
