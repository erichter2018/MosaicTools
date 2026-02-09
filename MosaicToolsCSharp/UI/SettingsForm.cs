using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Settings dialog matching Python's SettingsWindow.
/// </summary>
public class SettingsForm : Form
{
    // Dark mode title bar support
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    // Dark scrollbar support for native controls
    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
    private static extern int SetWindowTheme(IntPtr hWnd, string? pszSubAppName, string? pszSubIdList);

    private readonly Configuration _config;
    private readonly ActionController _controller;
    private readonly MainForm _mainForm;
    
    private TabControl _tabControl = null!;
    
    // General tab controls
    private TextBox _doctorNameBox = null!;
    private CheckBox _startBeepCheck = null!;
    private CheckBox _stopBeepCheck = null!;
    private TrackBar _startVolumeSlider = null!;
    private TrackBar _stopVolumeSlider = null!;
    private NumericUpDown _dictationPauseNum = null!;
    private CheckBox _floatingToolbarCheck = null!;
    private CheckBox _indicatorCheck = null!;
    private CheckBox _autoStopCheck = null!;
    private CheckBox _deadManCheck = null!;
    private Label _startVolLabel = null!;
    private Label _stopVolLabel = null!;
    private TextBox _ivHotkeyBox = null!;
    private TextBox _criticalTemplateBox = null!;
    private TextBox _seriesTemplateBox = null!;
    private TextBox _comparisonTemplateBox = null!;
    private CheckBox _separatePastedItemsCheck = null!;
    private ComboBox _reportFontFamilyCombo = null!;
    private NumericUpDown _reportFontSizeNumeric = null!;
    private CheckBox _autoUpdateCheck = null!;
    private CheckBox _hideIndicatorWhenNoStudyCheck = null!;
    private CheckBox _showTooltipsCheck = null!;

    // Tooltip system
    private ToolTip _settingsToolTip = null!;
    private List<Label> _tooltipLabels = new();

    // Advanced tab controls
    private CheckBox _scrollToBottomCheck = null!;
    private NumericUpDown _scrapeIntervalUpDown = null!;
    private CheckBox _showClinicalHistoryCheck = null!;
    private CheckBox _alwaysShowClinicalHistoryCheck = null!;
    private CheckBox _hideClinicalHistoryWhenNoStudyCheck = null!;
    private CheckBox _autoFixClinicalHistoryCheck = null!;
    private CheckBox _showDraftedIndicatorCheck = null!;
    private CheckBox _showTemplateMismatchCheck = null!;
    private CheckBox _genderCheckEnabledCheck = null!;
    private CheckBox _strokeDetectionEnabledCheck = null!;
    private CheckBox _strokeDetectionUseClinicalHistoryCheck = null!;
    private CheckBox _strokeClickToCreateNoteCheck = null!;
    private CheckBox _strokeAutoCreateNoteCheck = null!;
    private CheckBox _trackCriticalStudiesCheck = null!;
    private CheckBox _showImpressionCheck = null!;
    private CheckBox _showReportAfterProcessCheck = null!;
    private CheckBox _showLineCountToastCheck = null!;
    private NumericUpDown _scrollThreshold1 = null!;
    private NumericUpDown _scrollThreshold2 = null!;
    private NumericUpDown _scrollThreshold3 = null!;
    private CheckBox _ignoreInpatientDraftedCheck = null!;
    private RadioButton _ignoreInpatientAllXrRadio = null!;
    private RadioButton _ignoreInpatientChestOnlyRadio = null!;

    // InteleViewer Window/Level keys (null in headless mode)
    private TextBox? _windowLevelKeysBox;

    // Experimental tab controls
    private CheckBox _rvuCounterEnabledCheck = null!;
    private ComboBox _rvuDisplayModeCombo = null!;
    private CheckBox _rvuGoalEnabledCheck = null!;
    private NumericUpDown _rvuGoalValueBox = null!;
    private TextBox _rvuCounterPathBox = null!;
    private Label _rvuCounterStatusLabel = null!;
    private CheckBox _rvuMetricTotalCheck = null!;
    private CheckBox _rvuMetricPerHourCheck = null!;
    private CheckBox _rvuMetricCurrentHourCheck = null!;
    private CheckBox _rvuMetricPriorHourCheck = null!;
    private CheckBox _rvuMetricEstTotalCheck = null!;
    private ComboBox _rvuOverflowLayoutCombo = null!;
    private Label _rvuOverflowLayoutLabel = null!;
    private CheckBox _showReportChangesCheck = null!;
    private CheckBox _correlationEnabledCheck = null!;
    private CheckBox _pickListsEnabledCheck = null!;
    private CheckBox _pickListSkipSingleMatchCheck = null!;
    private CheckBox _pickListKeepOpenCheck = null!;
    private Label _pickListsCountLabel = null!;
    private Label _macrosCountLabel = null!;
    private Panel _reportChangesColorPanel = null!;
    private TrackBar _reportChangesAlphaSlider = null!;
    private Label _reportChangesAlphaLabel = null!;
    private RichTextBox _reportChangesPreview = null!;
    private CheckBox _reportTransparentCheck = null!;
    private TrackBar _reportTransparencySlider = null!;
    private Label _reportTransparencyLabel = null!;

    // Network monitor controls
    private CheckBox _connectivityMonitorEnabledCheck = null!;
    private NumericUpDown _connectivityIntervalUpDown = null!;
    private NumericUpDown _connectivityTimeoutUpDown = null!;

    public SettingsForm(Configuration config, ActionController controller, MainForm mainForm)
    {
        _config = config;
        _controller = controller;
        _mainForm = mainForm;

        InitializeUI();
        LoadSettings();

        // Apply dark scrollbar theme to all controls after UI is built
        ApplyDarkScrollbars(this);
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        // Enable dark mode title bar
        try
        {
            int value = 1;
            DwmSetWindowAttribute(Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
        }
        catch { /* Ignore on older Windows versions */ }
    }

    /// <summary>
    /// Recursively applies the Windows dark mode scrollbar theme to all controls.
    /// </summary>
    private void ApplyDarkScrollbars(Control parent)
    {
        foreach (Control c in parent.Controls)
        {
            try
            {
                if (c.IsHandleCreated)
                    SetWindowTheme(c.Handle, "DarkMode_Explorer", null);
                else
                    c.HandleCreated += (_, _) =>
                    {
                        try { SetWindowTheme(c.Handle, "DarkMode_Explorer", null); }
                        catch { }
                    };
            }
            catch { }

            if (c.Controls.Count > 0)
                ApplyDarkScrollbars(c);
        }
    }
    
    private void InitializeUI()
    {
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var versionStr = version != null ? $" v{version.Major}.{version.Minor}.{version.Build}" : "";
        Text = $"Mosaic Tools Settings{versionStr}";
        Size = new Size(500, 550);
        StartPosition = FormStartPosition.Manual;
        Location = new Point(_config.SettingsX, _config.SettingsY);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.FromArgb(30, 30, 30);
        ForeColor = Color.White;

        // Initialize tooltip system
        _settingsToolTip = new ToolTip
        {
            AutoPopDelay = 15000,
            InitialDelay = 300,
            ReshowDelay = 100,
            ShowAlways = true
        };

        // Tab control with owner-drawn dark theme
        _tabControl = new DarkTabControl
        {
            Dock = DockStyle.Fill
        };
        Controls.Add(_tabControl);
        
        // General tab (Profile, Desktop Mode, Updates)
        var generalTab = CreateGeneralTab();
        _tabControl.TabPages.Add(generalTab);

        // Keys tab (was Control Map)
        var keysTab = CreateControlMapTab();
        _tabControl.TabPages.Add(keysTab);

        // IV Buttons tab (was Button Studio)
        var ivButtonsTab = CreateButtonStudioTab();
        _tabControl.TabPages.Add(ivButtonsTab);

        // Text Automation tab (combines Templates, Macros, Pick Lists)
        var textAutomationTab = CreateTextAutomationTab();
        _tabControl.TabPages.Add(textAutomationTab);

        // Notifications tab (Clinical History, Alerts)
        var notificationsTab = CreateNotificationsTab();
        _tabControl.TabPages.Add(notificationsTab);

        // Behavior tab (Scraping, Scrolling, Focus)
        var behaviorTab = CreateBehaviorTab();
        _tabControl.TabPages.Add(behaviorTab);

        // Reference tab (AHK docs, Debug tips)
        var referenceTab = CreateReferenceTab();
        _tabControl.TabPages.Add(referenceTab);
        
        // Save/Cancel/Help buttons
        var buttonPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            BackColor = Color.FromArgb(40, 40, 40)
        };

        var helpBtn = new Button
        {
            Text = "Help",
            Size = new Size(80, 30),
            Location = new Point(10, 10),
            BackColor = Color.FromArgb(51, 51, 102),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        helpBtn.Click += (_, _) => ShowTabHelp();
        buttonPanel.Controls.Add(helpBtn);

        // Version label (centered)
        var appVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var versionLabel = new Label
        {
            Text = appVersion != null ? $"v{appVersion.Major}.{appVersion.Minor}.{appVersion.Build}" : "",
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 9),
            AutoSize = true
        };
        versionLabel.Location = new Point((Width - versionLabel.PreferredWidth) / 2, 15);
        buttonPanel.Controls.Add(versionLabel);

        var saveBtn = new Button
        {
            Text = "Save",
            Size = new Size(80, 30),
            Location = new Point(Width - 200, 10),
            BackColor = Color.FromArgb(51, 102, 0),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        saveBtn.Click += (_, _) => SaveAndClose();
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
        cancelBtn.Click += (_, _) => Close();
        buttonPanel.Controls.Add(cancelBtn);

        Controls.Add(buttonPanel);
    }
    
    private TabPage CreateGeneralTab()
    {
        var tab = new TabPage("General")
        {
            BackColor = Color.FromArgb(40, 40, 40),
            AutoScroll = true
        };

        int y = 10;
        int groupWidth = 440;

        // ========== PROFILE SECTION ==========
        var profileGroup = new GroupBox
        {
            Text = "Profile",
            Location = new Point(15, y),
            Size = new Size(groupWidth, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(profileGroup);

        var doctorNameLabel = new Label
        {
            Text = "Doctor Name:",
            Location = new Point(15, 25),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        profileGroup.Controls.Add(doctorNameLabel);
        _doctorNameBox = new TextBox
        {
            Location = new Point(120, 22),
            Width = 200,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        profileGroup.Controls.Add(_doctorNameBox);
        CreateTooltipLabel(profileGroup, _doctorNameBox, "Your name as it appears in Clario. Used to filter your own notes\nfrom Critical Findings results so the contact person is identified.");

        // Show Tooltips checkbox
        _showTooltipsCheck = new CheckBox
        {
            Text = "Show tooltips",
            Location = new Point(340, 24),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9),
            Checked = _config.ShowTooltips
        };
        _showTooltipsCheck.CheckedChanged += (s, e) => UpdateTooltipVisibility();
        profileGroup.Controls.Add(_showTooltipsCheck);

        y += 70;

        // ========== DESKTOP MODE SECTION (hidden in headless) ==========
        // Initialize controls regardless (for LoadSettings)
        _startBeepCheck = new CheckBox { AutoSize = true };
        _stopBeepCheck = new CheckBox { AutoSize = true };
        _startVolumeSlider = new TrackBar { Minimum = 0, Maximum = 100, TickStyle = TickStyle.None, AutoSize = false, Height = 20 };
        _stopVolumeSlider = new TrackBar { Minimum = 0, Maximum = 100, TickStyle = TickStyle.None, AutoSize = false, Height = 20 };
        _startVolLabel = new Label { AutoSize = true, ForeColor = Color.White };
        _stopVolLabel = new Label { AutoSize = true, ForeColor = Color.White };
        _dictationPauseNum = new NumericUpDown { Minimum = 100, Maximum = 5000, Increment = 100, BackColor = Color.FromArgb(60, 60, 60), ForeColor = Color.White };
        _ivHotkeyBox = new TextBox { BackColor = Color.FromArgb(60, 60, 60), ForeColor = Color.White, ReadOnly = true, Cursor = Cursors.Hand };
        _indicatorCheck = new CheckBox { AutoSize = true };
        _hideIndicatorWhenNoStudyCheck = new CheckBox { AutoSize = true };
        _autoStopCheck = new CheckBox { AutoSize = true };

        if (!App.IsHeadless)
        {
            var desktopGroup = new GroupBox
            {
                Text = "Desktop Mode",
                Location = new Point(15, y),
                Size = new Size(groupWidth, 290),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            tab.Controls.Add(desktopGroup);

            int dy = 22;

            // IV Report Hotkey
            desktopGroup.Controls.Add(new Label
            {
                Text = "IV Report Hotkey:",
                Location = new Point(15, dy),
                AutoSize = true,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9)
            });
            _ivHotkeyBox.Location = new Point(140, dy - 3);
            _ivHotkeyBox.Width = 120;
            SetupHotkeyCapture(_ivHotkeyBox);
            desktopGroup.Controls.Add(_ivHotkeyBox);
            CreateTooltipLabel(desktopGroup, _ivHotkeyBox, "Hotkey to copy report from InteleViewer.\nClick to capture a new key combination.");
            dy += 30;

            // Recording Indicator
            _indicatorCheck.Text = "Show Recording Indicator";
            _indicatorCheck.Location = new Point(15, dy);
            _indicatorCheck.ForeColor = Color.White;
            _indicatorCheck.Font = new Font("Segoe UI", 9);
            desktopGroup.Controls.Add(_indicatorCheck);
            CreateTooltipLabel(desktopGroup, _indicatorCheck, "Shows a small colored rectangle indicating dictation state\n(red = recording, gray = stopped).");
            dy += 22;

            _hideIndicatorWhenNoStudyCheck.Text = "Hide when no study open";
            _hideIndicatorWhenNoStudyCheck.Location = new Point(35, dy);
            _hideIndicatorWhenNoStudyCheck.ForeColor = Color.Gray;
            _hideIndicatorWhenNoStudyCheck.Font = new Font("Segoe UI", 9);
            _hideIndicatorWhenNoStudyCheck.CheckedChanged += (s, e) =>
            {
                _config.HideIndicatorWhenNoStudy = _hideIndicatorWhenNoStudyCheck.Checked;
                _mainForm.UpdateIndicatorVisibility();
            };
            desktopGroup.Controls.Add(_hideIndicatorWhenNoStudyCheck);
            CreateTooltipLabel(desktopGroup, _hideIndicatorWhenNoStudyCheck, "Hides the recording indicator when no study\nis open in Mosaic.");
            dy += 25;

            // Auto-Stop Dictation
            _autoStopCheck.Text = "Auto-Stop Dictation on Process";
            _autoStopCheck.Location = new Point(15, dy);
            _autoStopCheck.ForeColor = Color.White;
            _autoStopCheck.Font = new Font("Segoe UI", 9);
            desktopGroup.Controls.Add(_autoStopCheck);
            CreateTooltipLabel(desktopGroup, _autoStopCheck, "Automatically stops dictation when you press\nProcess Report.");
            dy += 30;

            // Audio Feedback section header
            desktopGroup.Controls.Add(new Label
            {
                Text = "Audio Feedback",
                Location = new Point(15, dy),
                AutoSize = true,
                ForeColor = Color.FromArgb(180, 180, 180),
                Font = new Font("Segoe UI", 8, FontStyle.Italic)
            });
            dy += 20;

            // Start Beep row
            _startBeepCheck.Text = "Start Beep";
            _startBeepCheck.Location = new Point(15, dy);
            _startBeepCheck.ForeColor = Color.White;
            _startBeepCheck.Font = new Font("Segoe UI", 9);
            desktopGroup.Controls.Add(_startBeepCheck);

            _startVolumeSlider.Location = new Point(120, dy - 3);
            _startVolumeSlider.Width = 150;
            _startVolumeSlider.ValueChanged += (s, e) => UpdateVolumeLabels();
            desktopGroup.Controls.Add(_startVolumeSlider);

            _startVolLabel.Location = new Point(275, dy + 3);
            desktopGroup.Controls.Add(_startVolLabel);
            CreateTooltipLabelAt(desktopGroup, 310, dy + 3, "Plays an audio beep when dictation starts.\nSlider adjusts volume (0-100%).");
            dy += 40;

            // Stop Beep row
            _stopBeepCheck.Text = "Stop Beep";
            _stopBeepCheck.Location = new Point(15, dy);
            _stopBeepCheck.ForeColor = Color.White;
            _stopBeepCheck.Font = new Font("Segoe UI", 9);
            desktopGroup.Controls.Add(_stopBeepCheck);

            _stopVolumeSlider.Location = new Point(120, dy - 3);
            _stopVolumeSlider.Width = 150;
            _stopVolumeSlider.ValueChanged += (s, e) => UpdateVolumeLabels();
            desktopGroup.Controls.Add(_stopVolumeSlider);

            _stopVolLabel.Location = new Point(275, dy + 3);
            desktopGroup.Controls.Add(_stopVolLabel);
            CreateTooltipLabelAt(desktopGroup, 310, dy + 3, "Plays an audio beep when dictation stops.\nSlider adjusts volume (0-100%).");
            dy += 40;

            // Dictation pause
            desktopGroup.Controls.Add(new Label
            {
                Text = "Start Beep Pause:",
                Location = new Point(15, dy + 2),
                AutoSize = true,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9)
            });
            _dictationPauseNum.Location = new Point(155, dy);
            _dictationPauseNum.Width = 70;
            desktopGroup.Controls.Add(_dictationPauseNum);
            var msLabel = new Label
            {
                Text = "ms",
                Location = new Point(230, dy + 2),
                AutoSize = true,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 9)
            };
            desktopGroup.Controls.Add(msLabel);
            CreateTooltipLabel(desktopGroup, msLabel, "Delay before playing start beep, to avoid\nfalse triggers. Recommended: 800-1200ms.");

            y += 300;
        }

        // ========== COMMON OPTIONS SECTION ==========
        var optionsGroup = new GroupBox
        {
            Text = "Options",
            Location = new Point(15, y),
            Size = new Size(groupWidth, 70),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(optionsGroup);

        _floatingToolbarCheck = new CheckBox
        {
            Text = "Show InteleViewer Buttons",
            Location = new Point(15, 22),
            ForeColor = Color.White,
            AutoSize = true,
            Font = new Font("Segoe UI", 9)
        };
        optionsGroup.Controls.Add(_floatingToolbarCheck);

        var configureLink = new LinkLabel
        {
            Text = "Configure",
            Location = new Point(200, 23),
            AutoSize = true,
            LinkColor = Color.FromArgb(100, 149, 237),
            Font = new Font("Segoe UI", 9)
        };
        configureLink.Click += (s, e) => _tabControl.SelectedIndex = 2;
        optionsGroup.Controls.Add(configureLink);
        CreateTooltipLabel(optionsGroup, configureLink, "Shows configurable buttons for InteleViewer shortcuts\n(window/level presets, zoom, etc.).");

        _deadManCheck = new CheckBox
        {
            Text = "Push-to-Talk (Dead Man's Switch)",
            Location = new Point(15, 44),
            ForeColor = Color.White,
            AutoSize = true,
            Font = new Font("Segoe UI", 9)
        };
        optionsGroup.Controls.Add(_deadManCheck);
        CreateTooltipLabel(optionsGroup, _deadManCheck, "Hold the Record button to dictate, release to stop.\nDead man's switch style dictation.");

        y += 80;

        // ========== UPDATES SECTION ==========
        var updatesGroup = new GroupBox
        {
            Text = "Updates",
            Location = new Point(15, y),
            Size = new Size(groupWidth, 55),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(updatesGroup);

        _autoUpdateCheck = new CheckBox
        {
            Text = "Auto-update",
            Location = new Point(15, 22),
            ForeColor = App.IsHeadless ? Color.Gray : Color.White,
            AutoSize = true,
            Enabled = !App.IsHeadless,
            Font = new Font("Segoe UI", 9)
        };
        updatesGroup.Controls.Add(_autoUpdateCheck);
        CreateTooltipLabel(updatesGroup, _autoUpdateCheck, "Automatically check for and install updates\non startup.");

        var checkUpdatesBtn = new Button
        {
            Text = "Check for Updates",
            Location = new Point(140, 19),
            Size = new Size(130, 25),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        checkUpdatesBtn.FlatAppearance.BorderColor = Color.Gray;
        checkUpdatesBtn.Click += async (s, e) =>
        {
            checkUpdatesBtn.Enabled = false;
            checkUpdatesBtn.Text = "Checking...";
            try
            {
                await _mainForm.CheckForUpdatesManualAsync();
            }
            finally
            {
                if (!checkUpdatesBtn.IsDisposed)
                {
                    checkUpdatesBtn.Text = "Check for Updates";
                    checkUpdatesBtn.Enabled = true;
                }
            }
        };
        updatesGroup.Controls.Add(checkUpdatesBtn);
        CreateTooltipLabel(updatesGroup, checkUpdatesBtn, "Manually check for available updates now.");

        return tab;
    }
    
    private TabPage CreateControlMapTab()
    {
        var tab = new TabPage("Keys")
        {
            BackColor = Color.FromArgb(40, 40, 40),
            AutoScroll = true
        };
        
        int y = 20;
        
        tab.Controls.Add(CreateLabel("Action", 20, y, 150));
        if (!App.IsHeadless)
        {
            tab.Controls.Add(CreateLabel("Hotkey", 180, y, 120));
        }
        tab.Controls.Add(CreateLabel("Mic Button", App.IsHeadless ? 180 : 310, y, 120));
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

        foreach (var action in Actions.All)
        {
            if (action == Actions.None) continue;

            // Hide Cycle Window/Level entirely in headless mode
            if (App.IsHeadless && action == Actions.CycleWindowLevel) continue;

            // Use AutoSize label so tooltip appears right after text
            var lbl = new Label
            {
                Text = action,
                Location = new Point(20, y),
                AutoSize = true,
                ForeColor = Color.White
            };
            tab.Controls.Add(lbl);

            // Add tooltip for action description right after the label text
            if (actionDescriptions.TryGetValue(action, out var desc))
            {
                CreateTooltipLabel(tab, lbl, desc, 2, 2);
            }

            var hotkeyBox = new TextBox
            {
                Name = $"hotkey_{action}",
                Location = new Point(180, y),
                Width = 120,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                Tag = action,
                ReadOnly = true,
                Cursor = Cursors.Hand
            };
            hotkeyBox.Text = _config.ActionMappings.GetValueOrDefault(action)?.Hotkey ?? "";
            SetupHotkeyCapture(hotkeyBox);

            // Hide hotkey in headless mode
            if (!App.IsHeadless)
            {
                tab.Controls.Add(hotkeyBox);
            }

            var micCombo = new ComboBox
            {
                Name = $"mic_{action}",
                Location = new Point(App.IsHeadless ? 180 : 310, y),
                Width = 140,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                Tag = action
            };
            micCombo.Items.Add("");
            micCombo.Items.AddRange(HidService.AllButtons);
            var currentMic = _config.ActionMappings.GetValueOrDefault(action)?.MicButton ?? "";
            micCombo.SelectedItem = micCombo.Items.Contains(currentMic) ? currentMic : "";
            tab.Controls.Add(micCombo);

            y += 30;
        }
        
        y += 20;
        var infoLabel = new Label
        {
            Text = "Hardcoded buttons for Nuance PowerMic 2:\n" +
                   "Skip Back - Process Report\n" +
                   "Skip Forward - Generate Impression\n" +
                   "Record Button - Start / Stop Dictation\n" +
                   "Checkmark - Sign Report",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.DarkGray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        tab.Controls.Add(infoLabel);
        
        return tab;
    }
    
    // Button Studio state (class-level to persist during session)
    private List<FloatingButtonDef> _studioButtons = new();
    private int _studioColumns = 2;
    private int _selectedButtonIdx = 0;
    private Panel _previewPanel = null!;
    private FlowLayoutPanel _buttonListPanel = null!;
    
    // Editor controls
    private RadioButton _typeSquareRadio = null!;
    private RadioButton _typeWideRadio = null!;
    private ComboBox _iconCombo = null!;
    private TextBox _labelBox = null!;
    private TextBox _keystrokeBox = null!;
    private Button _recButton = null!;
    private ComboBox _actionCombo = null!;
    private bool _updatingEditor = false;  // Guard flag to prevent feedback loop
    
    private static readonly string[] IconLibrary = {
        "", "↑", "↓", "←", "→", "↕", "↔", "↺", "↻", "⟲", "⟳",
        "+", "−", "×", "÷", "⊕", "⊖", "☰", "◎", "▶", "⏸", "⏹", 
        "⚙", "☀", "★", "✓", "✗", "✚", "⎚", "▦", "◐", "◑", "⇑", "⇓", "⇐", "⇒",
        "🔍", "🔎", "📋", "📌", "🔒", "🔓", "💾", "🗑", "📁", "📂",
        "◀", "▲", "▼", "◆", "○", "●", "□", "■", "△", "▷"
    };
    
    private Button _iconButton = null!;
    private Panel? _iconPickerPanel;
    
    private TabPage CreateButtonStudioTab()
    {
        var tab = new TabPage("IV Buttons")
        {
            BackColor = Color.FromArgb(40, 40, 40)
        };
        
        // Deep copy buttons config for editing
        _studioButtons = _config.FloatingButtons.Buttons
            .Select(b => new FloatingButtonDef { Type = b.Type, Icon = b.Icon, Label = b.Label, Keystroke = b.Keystroke, Action = b.Action })
            .ToList();
        _studioColumns = _config.FloatingButtons.Columns;
        _selectedButtonIdx = 0;
        
        // === Top Row: Columns selector ===
        var topPanel = new Panel
        {
            Location = new Point(10, 10),
            Size = new Size(460, 30),
            BackColor = Color.FromArgb(40, 40, 40)
        };
        tab.Controls.Add(topPanel);
        
        topPanel.Controls.Add(CreateLabel("Columns:", 0, 5, 60));
        var columnsNum = new NumericUpDown
        {
            Name = "btn_columns",
            Location = new Point(65, 3),
            Width = 50,
            Minimum = 1,
            Maximum = 3,
            Value = _studioColumns,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        columnsNum.ValueChanged += (_, _) =>
        {
            _studioColumns = (int)columnsNum.Value;
            RenderPreview();
        };
        topPanel.Controls.Add(columnsNum);
        CreateTooltipLabel(topPanel, columnsNum, "Number of button columns (1-3).\nMaximum 9 buttons total.");

        var maxLabel = CreateLabel("(max 9 buttons)", 125, 5, 120);
        maxLabel.ForeColor = Color.Gray;
        maxLabel.Font = new Font(maxLabel.Font.FontFamily, 8, FontStyle.Italic);
        topPanel.Controls.Add(maxLabel);
        
        // === Left Side: Live Preview + Button List ===
        var leftPanel = new Panel
        {
            Location = new Point(10, 45),
            Size = new Size(240, 350),
            BackColor = Color.FromArgb(40, 40, 40)
        };
        tab.Controls.Add(leftPanel);
        
        // Preview label
        var previewLabel = CreateLabel("Live Preview", 0, 0, 100);
        previewLabel.Font = new Font(previewLabel.Font, FontStyle.Bold);
        leftPanel.Controls.Add(previewLabel);
        
        // Preview grid - with scroll for many buttons
        _previewPanel = new Panel
        {
            Location = new Point(0, 22),
            Size = new Size(230, 180),
            BackColor = Color.FromArgb(51, 51, 51),
            BorderStyle = BorderStyle.FixedSingle,
            AutoScroll = true
        };
        leftPanel.Controls.Add(_previewPanel);
        
        // Button list label
        var listLabel = CreateLabel("Button List", 0, 210, 100);
        listLabel.Font = new Font(listLabel.Font, FontStyle.Bold);
        leftPanel.Controls.Add(listLabel);
        
        // Button list strip - taller with wrapping
        _buttonListPanel = new FlowLayoutPanel
        {
            Location = new Point(0, 232),
            Size = new Size(230, 60),
            BackColor = Color.FromArgb(51, 51, 51),
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            AutoScroll = false
        };
        leftPanel.Controls.Add(_buttonListPanel);
        
        // Control buttons - with spacing above
        var ctrlPanel = new FlowLayoutPanel
        {
            Location = new Point(0, 305),
            Size = new Size(230, 35),
            FlowDirection = FlowDirection.LeftToRight,
            BackColor = Color.FromArgb(40, 40, 40)
        };
        leftPanel.Controls.Add(ctrlPanel);
        
        var addBtn = new Button { Text = "+Add", Width = 50, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        addBtn.Click += (_, _) => AddButton();
        ctrlPanel.Controls.Add(addBtn);
        
        var delBtn = new Button { Text = "-Del", Width = 45, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        delBtn.Click += (_, _) => DeleteButton();
        ctrlPanel.Controls.Add(delBtn);
        
        var upBtn = new Button { Text = "▲", Width = 30, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        upBtn.Click += (_, _) => MoveButtonUp();
        ctrlPanel.Controls.Add(upBtn);
        
        var downBtn = new Button { Text = "▼", Width = 30, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        downBtn.Click += (_, _) => MoveButtonDown();
        ctrlPanel.Controls.Add(downBtn);
        
        var resetBtn = new Button { Text = "Reset", Width = 50, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.FromArgb(204, 0, 0) };
        resetBtn.Click += (_, _) => ResetToDefaults();
        ctrlPanel.Controls.Add(resetBtn);
        
        // === Right Side: Button Editor ===
        var editorPanel = new GroupBox
        {
            Text = "Button Editor",
            Location = new Point(260, 45),
            Size = new Size(210, 295),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(editorPanel);
        
        int ey = 25;
        
        // Type
        var typeLabel = new Label { Text = "Type:", Location = new Point(10, ey), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) };
        editorPanel.Controls.Add(typeLabel);
        
        _typeSquareRadio = new RadioButton { Text = "Square", Location = new Point(65, ey - 2), AutoSize = true, ForeColor = Color.White, Checked = true };
        _typeSquareRadio.CheckedChanged += (_, _) => { if (!_updatingEditor && _typeSquareRadio.Checked) ApplyEditorChanges(); };
        editorPanel.Controls.Add(_typeSquareRadio);
        
        _typeWideRadio = new RadioButton { Text = "Wide", Location = new Point(135, ey - 2), AutoSize = true, ForeColor = Color.White };
        _typeWideRadio.CheckedChanged += (_, _) => { if (!_updatingEditor && _typeWideRadio.Checked) ApplyEditorChanges(); };
        editorPanel.Controls.Add(_typeWideRadio);
        CreateTooltipLabel(editorPanel, _typeWideRadio, "Square: small button, good for icons.\nWide: full-width, good for text labels.");
        ey += 30;
        
        // Icon - button that opens grid picker
        var iconLabel = new Label { Text = "Icon:", Location = new Point(10, ey), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) };
        editorPanel.Controls.Add(iconLabel);
        
        _iconButton = new Button
        {
            Location = new Point(65, ey - 3),
            Size = new Size(50, 28),
            Text = "",
            Font = new Font("Segoe UI Symbol", 12),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            TextAlign = ContentAlignment.MiddleCenter,
            UseCompatibleTextRendering = true,
            Padding = new Padding(0)
        };
        _iconButton.FlatAppearance.BorderColor = Color.Gray;
        _iconButton.Click += (_, _) => ShowIconPicker();
        editorPanel.Controls.Add(_iconButton);
        
        // Clear icon button (X)
        var clearIconBtn = new Button
        {
            Location = new Point(120, ey),
            Size = new Size(22, 22),
            Text = "X",
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(100, 50, 50),
            ForeColor = Color.White,
            TextAlign = ContentAlignment.MiddleCenter,
            UseCompatibleTextRendering = true
        };
        clearIconBtn.FlatAppearance.BorderSize = 0;
        clearIconBtn.Click += (_, _) =>
        {
            _iconButton.Text = "";
            ApplyEditorChanges();
        };
        editorPanel.Controls.Add(clearIconBtn);
        
        // Keep old combo for compatibility but hide
        _iconCombo = new ComboBox { Visible = false };
        ey += 30;
        
        // Label
        var lblLabel = new Label { Text = "Label:", Location = new Point(10, ey), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) };
        editorPanel.Controls.Add(lblLabel);
        
        _labelBox = new TextBox
        {
            Location = new Point(65, ey - 2),
            Width = 120,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        _labelBox.TextChanged += (_, _) => { if (!_updatingEditor) ApplyEditorChanges(); };  // Auto-apply on change
        editorPanel.Controls.Add(_labelBox);
        ey += 30;
        
        // Keystroke
        var keyLabel = new Label { Text = "Key:", Location = new Point(10, ey), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) };
        editorPanel.Controls.Add(keyLabel);
        
        _keystrokeBox = new TextBox
        {
            Location = new Point(65, ey - 2),
            Width = 80,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            ReadOnly = true,
            Cursor = Cursors.Hand
        };
        _keystrokeBox.TextChanged += (_, _) => { if (!_updatingEditor) ApplyEditorChanges(); };
        SetupHotkeyCapture(_keystrokeBox);
        editorPanel.Controls.Add(_keystrokeBox);
        
        _recButton = new Button { Text = "Rec", Location = new Point(150, ey - 3), Width = 40, Height = 23, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        _recButton.Click += (_, _) => _keystrokeBox.Focus();
        editorPanel.Controls.Add(_recButton);
        ey += 30;

        // Action
        var actionLabel = new Label { Text = "Action:", Location = new Point(10, ey + 2), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) };
        editorPanel.Controls.Add(actionLabel);

        _actionCombo = new ComboBox
        {
            Location = new Point(65, ey - 1),
            Width = 125,
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _actionCombo.Items.Add("(keystroke)");
        foreach (var a in Actions.All)
        {
            if (a != Actions.None) _actionCombo.Items.Add(a);
        }
        _actionCombo.SelectedIndex = 0;
        _actionCombo.SelectedIndexChanged += (_, _) => { if (!_updatingEditor) ApplyEditorChanges(); };
        editorPanel.Controls.Add(_actionCombo);
        ey += 35;

        // Hint
        var hintLabel = new Label
        {
            Text = "Action overrides keystroke.\nKeystroke sends to InteleViewer.",
            Location = new Point(10, ey),
            Size = new Size(180, 35),
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        editorPanel.Controls.Add(hintLabel);
        ey += 50;
        
        // Apply button
        var applyBtn = new Button
        {
            Text = "Apply",
            Location = new Point(65, ey),
            Width = 80,
            Height = 28,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(51, 102, 0),
            ForeColor = Color.White
        };
        applyBtn.Click += (_, _) => ApplyEditorChanges();
        editorPanel.Controls.Add(applyBtn);
        
        // Initial render
        RenderPreview();
        RenderButtonList();
        UpdateEditorFromSelection();
        
        return tab;
    }
    
    private void RenderPreview()
    {
        var oldPreviewCtrls = _previewPanel.Controls.Cast<Control>().ToList();
        _previewPanel.Controls.Clear();
        foreach (var ctrl in oldPreviewCtrls) ctrl.Dispose();
        
        int btnSize = 50;
        int wideBtnHeight = 34;  // Match actual toolbar
        int padding = 3;
        int yPos = padding;  // Track Y position
        int col = 0;
        
        for (int i = 0; i < _studioButtons.Count; i++)
        {
            var btnCfg = _studioButtons[i];
            bool isSelected = (i == _selectedButtonIdx);
            bool isWide = btnCfg.Type == "wide";
            int btnHeight = isWide ? wideBtnHeight : btnSize;
            var borderColor = isSelected ? Color.Lime : Color.Gray;
            
            var borderPanel = new Panel
            {
                BackColor = borderColor,
                Size = isWide
                    ? new Size(_studioColumns * (btnSize + padding) - padding, btnHeight)
                    : new Size(btnSize, btnHeight)
            };
            
            if (isWide)
            {
                if (col > 0) { yPos += btnSize + padding; col = 0; }
                borderPanel.Location = new Point(padding, yPos);
                yPos += btnHeight + padding;
                col = 0;
            }
            else
            {
                borderPanel.Location = new Point(col * (btnSize + padding) + padding, yPos);
                col++;
                if (col >= _studioColumns) { col = 0; yPos += btnSize + padding; }
            }
            
            // Get display text - prefer icon, then label, then "?"
            string displayText;
            if (!string.IsNullOrWhiteSpace(btnCfg.Icon))
                displayText = btnCfg.Icon;
            else if (!string.IsNullOrWhiteSpace(btnCfg.Label))
                displayText = isWide ? btnCfg.Label : btnCfg.Label.Substring(0, Math.Min(4, btnCfg.Label.Length));
            else
                displayText = "?";
            
            var btn = new Button
            {
                Text = displayText,
                Font = isWide 
                    ? new Font("Segoe UI", 8, FontStyle.Bold) 
                    : new Font("Segoe UI Symbol", 12, FontStyle.Bold),  // Smaller fonts for preview
                ForeColor = Color.FromArgb(204, 204, 204),
                BackColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Padding = new Padding(0),
                Tag = i
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, _) => SelectButton((int)((Button)s!).Tag!);
            
            borderPanel.Padding = new Padding(1);
            borderPanel.Controls.Add(btn);
            _previewPanel.Controls.Add(borderPanel);
        }
    }
    
    private void RenderButtonList()
    {
        var oldBtnListCtrls = _buttonListPanel.Controls.Cast<Control>().ToList();
        _buttonListPanel.Controls.Clear();
        foreach (var ctrl in oldBtnListCtrls) ctrl.Dispose();
        
        for (int i = 0; i < _studioButtons.Count; i++)
        {
            var btnCfg = _studioButtons[i];
            bool isSelected = (i == _selectedButtonIdx);
            
            // Get display text - prefer icon, then label, then "?"
            string displayText;
            if (!string.IsNullOrWhiteSpace(btnCfg.Icon))
                displayText = btnCfg.Icon;
            else if (!string.IsNullOrWhiteSpace(btnCfg.Label))
                displayText = btnCfg.Label.Substring(0, Math.Min(3, btnCfg.Label.Length));
            else
                displayText = "?";
            
            var lbl = new Label
            {
                Text = $"{i + 1}: {displayText}",
                BackColor = isSelected ? Color.FromArgb(68, 68, 68) : Color.FromArgb(34, 34, 34),
                ForeColor = isSelected ? Color.White : Color.FromArgb(204, 204, 204),
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(45, 28),
                Margin = new Padding(2),
                Cursor = Cursors.Hand,
                Tag = i
            };
            lbl.Click += (s, _) => SelectButton((int)((Label)s!).Tag!);
            _buttonListPanel.Controls.Add(lbl);
        }
    }
    
    private void SelectButton(int idx)
    {
        _selectedButtonIdx = idx;
        
        // Preserve scroll position
        var scrollPos = _previewPanel.AutoScrollPosition;
        RenderPreview();
        _previewPanel.AutoScrollPosition = new Point(-scrollPos.X, -scrollPos.Y);
        
        RenderButtonList();
        UpdateEditorFromSelection();
    }
    
    private void UpdateEditorFromSelection()
    {
        if (_selectedButtonIdx < 0 || _selectedButtonIdx >= _studioButtons.Count) return;
        
        _updatingEditor = true;  // Prevent feedback loop
        try
        {
            var btn = _studioButtons[_selectedButtonIdx];
            _typeSquareRadio.Checked = btn.Type != "wide";
            _typeWideRadio.Checked = btn.Type == "wide";
            _iconButton.Text = btn.Icon ?? "";
            _labelBox.Text = btn.Label ?? "";
            _keystrokeBox.Text = btn.Keystroke ?? "";

            var action = btn.Action ?? "";
            if (!string.IsNullOrEmpty(action) && action != Actions.None)
            {
                var idx = _actionCombo.Items.IndexOf(action);
                _actionCombo.SelectedIndex = idx >= 0 ? idx : 0;
            }
            else
            {
                _actionCombo.SelectedIndex = 0; // (keystroke)
            }

            // Disable keystroke controls when an action is selected (action overrides keystroke)
            var actionSelected = _actionCombo.SelectedIndex > 0;
            _keystrokeBox.Enabled = !actionSelected;
            _recButton.Enabled = !actionSelected;
        }
        finally
        {
            _updatingEditor = false;
        }
    }
    
    private void ApplyEditorChanges()
    {
        if (_selectedButtonIdx < 0 || _selectedButtonIdx >= _studioButtons.Count) return;
        if (_updatingEditor) return;
        
        var btn = _studioButtons[_selectedButtonIdx];
        btn.Type = _typeWideRadio.Checked ? "wide" : "square";
        btn.Icon = _iconButton.Text;
        btn.Label = _labelBox.Text;
        btn.Keystroke = _keystrokeBox.Text;
        btn.Action = _actionCombo.SelectedIndex > 0 ? _actionCombo.SelectedItem?.ToString() ?? "" : "";

        // Disable keystroke controls when an action is selected (action overrides keystroke)
        var actionSelected = _actionCombo.SelectedIndex > 0;
        _keystrokeBox.Enabled = !actionSelected;
        _recButton.Enabled = !actionSelected;

        // Preserve scroll position
        var scrollPos = _previewPanel.AutoScrollPosition;
        RenderPreview();
        _previewPanel.AutoScrollPosition = new Point(-scrollPos.X, -scrollPos.Y);
        RenderButtonList();
    }
    
    private void ShowIconPicker()
    {
        // Close existing picker if open
        if (_iconPickerPanel != null)
        {
            var existingForm = _iconPickerPanel.FindForm();
            if (existingForm != null && existingForm != this)
            {
                existingForm.Close();
            }
            _iconPickerPanel = null;
            return;
        }
        
        // Create popup form (won't be clipped by parent form)
        int cols = 10;
        int iconSize = 28;
        int rows = (int)Math.Ceiling(IconLibrary.Length / (double)cols);
        
        var pickerForm = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            StartPosition = FormStartPosition.Manual,
            Size = new Size(cols * iconSize + 12, rows * iconSize + 12),
            BackColor = Color.FromArgb(45, 45, 45),
            ShowInTaskbar = false,
            TopMost = true
        };
        
        // Position below the icon button
        var buttonScreenPos = _iconButton.PointToScreen(Point.Empty);
        pickerForm.Location = new Point(buttonScreenPos.X, buttonScreenPos.Y + _iconButton.Height + 2);
        
        // Panel inside form
        _iconPickerPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(45, 45, 45)
        };
        pickerForm.Controls.Add(_iconPickerPanel);
        
        // Add icon buttons in grid
        for (int i = 0; i < IconLibrary.Length; i++)
        {
            var icon = IconLibrary[i];
            var btn = new Button
            {
                Text = string.IsNullOrEmpty(icon) ? "∅" : icon,
                Size = new Size(iconSize, iconSize),
                Location = new Point((i % cols) * iconSize + 6, (i / cols) * iconSize + 6),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Symbol", 11),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(60, 60, 60),
                Tag = icon
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(80, 80, 80);
            btn.Click += (s, _) =>
            {
                var selectedIcon = (string)((Button)s!).Tag!;
                _iconButton.Text = selectedIcon;
                pickerForm.Close();
                _iconPickerPanel = null;
                ApplyEditorChanges();
            };
            _iconPickerPanel.Controls.Add(btn);
        }
        
        // Close when clicking outside or losing focus
        pickerForm.Deactivate += (_, _) =>
        {
            pickerForm.Close();
            _iconPickerPanel = null;
        };
        
        pickerForm.Show(this);
    }
    
    private void AddButton()
    {
        if (_studioButtons.Count >= 9)
        {
            MessageBox.Show("Maximum 9 buttons allowed.", "Limit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        _studioButtons.Add(new FloatingButtonDef { Type = "square", Icon = "", Label = "", Keystroke = "" });
        _selectedButtonIdx = _studioButtons.Count - 1;
        RenderPreview();
        RenderButtonList();
        
        // Clear editor for new button
        _typeSquareRadio.Checked = true;
        _iconButton.Text = "";
        _labelBox.Text = "";
        _keystrokeBox.Text = "";
        UpdateEditorFromSelection();
    }

    private void DeleteButton()
    {
        if (_studioButtons.Count == 0) return;
        if (_selectedButtonIdx >= 0 && _selectedButtonIdx < _studioButtons.Count)
        {
            _studioButtons.RemoveAt(_selectedButtonIdx);
            _selectedButtonIdx = Math.Max(0, _selectedButtonIdx - 1);
            RenderPreview();
            RenderButtonList();
            UpdateEditorFromSelection();
        }
    }
    
    private void MoveButtonUp()
    {
        if (_selectedButtonIdx <= 0) return;
        (_studioButtons[_selectedButtonIdx], _studioButtons[_selectedButtonIdx - 1]) = 
            (_studioButtons[_selectedButtonIdx - 1], _studioButtons[_selectedButtonIdx]);
        _selectedButtonIdx--;
        RenderPreview();
        RenderButtonList();
    }
    
    private void MoveButtonDown()
    {
        if (_selectedButtonIdx >= _studioButtons.Count - 1) return;
        (_studioButtons[_selectedButtonIdx], _studioButtons[_selectedButtonIdx + 1]) = 
            (_studioButtons[_selectedButtonIdx + 1], _studioButtons[_selectedButtonIdx]);
        _selectedButtonIdx++;
        RenderPreview();
        RenderButtonList();
    }
    
    private void ResetToDefaults()
    {
        var defaults = FloatingButtonsConfig.Default;
        _studioButtons = defaults.Buttons
            .Select(b => new FloatingButtonDef { Type = b.Type, Icon = b.Icon, Label = b.Label, Keystroke = b.Keystroke, Action = b.Action })
            .ToList();
        _studioColumns = defaults.Columns;
        _selectedButtonIdx = 0;
        
        // Update columns spinner
        var buttonStudioTab = _tabControl.TabPages[2];
        foreach (Control ctrl in buttonStudioTab.Controls)
        {
            if (ctrl is Panel p)
            {
                foreach (Control c in p.Controls)
                {
                    if (c is NumericUpDown num && num.Name == "btn_columns")
                        num.Value = _studioColumns;
                }
            }
        }
        
        RenderPreview();
        RenderButtonList();
        UpdateEditorFromSelection();
    }

    // ========== TEXT AUTOMATION TAB (Templates + Macros + Pick Lists) ==========

    private TabPage CreateTextAutomationTab()
    {
        var tab = new TabPage("Text")
        {
            BackColor = Color.FromArgb(40, 40, 40),
            AutoScroll = true
        };

        int y = 10;
        int groupWidth = 445;

        // ========== REPORT POPUP FONT SECTION ==========
        var fontGroup = new GroupBox
        {
            Text = "Report Popup Font",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 55),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(fontGroup);

        fontGroup.Controls.Add(new Label
        {
            Text = "Font:",
            Location = new Point(10, 22),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });

        _reportFontFamilyCombo = new ComboBox
        {
            Location = new Point(50, 19),
            Width = 150,
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9),
            FlatStyle = FlatStyle.Flat
        };
        _reportFontFamilyCombo.Items.AddRange(new object[] { "Consolas", "Courier New", "Cascadia Mono", "Lucida Console", "Segoe UI", "Calibri", "Arial" });
        fontGroup.Controls.Add(_reportFontFamilyCombo);

        fontGroup.Controls.Add(new Label
        {
            Text = "Size:",
            Location = new Point(215, 22),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });

        _reportFontSizeNumeric = new NumericUpDown
        {
            Location = new Point(255, 19),
            Width = 60,
            Minimum = 7,
            Maximum = 24,
            DecimalPlaces = 1,
            Increment = 0.5m,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        fontGroup.Controls.Add(_reportFontSizeNumeric);

        var resetFontBtn = new Button
        {
            Text = "Reset",
            Location = new Point(325, 18),
            Size = new Size(55, 24),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 8)
        };
        resetFontBtn.Click += (s, e) =>
        {
            _reportFontFamilyCombo.SelectedItem = "Consolas";
            _reportFontSizeNumeric.Value = 11;
        };
        fontGroup.Controls.Add(resetFontBtn);

        y += 65;

        // Separate pasted items checkbox
        _separatePastedItemsCheck = new CheckBox
        {
            Text = "Separate pasted items with line break",
            Location = new Point(10, y),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        tab.Controls.Add(_separatePastedItemsCheck);
        y += 30;

        // ========== TEMPLATES SECTION ==========
        var templatesGroup = new GroupBox
        {
            Text = "Templates",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 295),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(templatesGroup);

        int ty = 20;
        templatesGroup.Controls.Add(new Label
        {
            Text = "Critical Findings:",
            Location = new Point(10, ty),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });
        ty += 18;

        _criticalTemplateBox = new TextBox
        {
            Location = new Point(10, ty),
            Width = 420,
            Height = 50,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        templatesGroup.Controls.Add(_criticalTemplateBox);
        ty += 55;

        var criticalPlaceholdersLabel = new Label
        {
            Text = "Placeholders: {name}, {time}, {date}",
            Location = new Point(10, ty),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        templatesGroup.Controls.Add(criticalPlaceholdersLabel);
        CreateTooltipLabel(templatesGroup, criticalPlaceholdersLabel, "Template for pasting critical findings.\n{name} = contact person, {time} = time, {date} = date.", 2, 0);
        ty += 20;

        templatesGroup.Controls.Add(new Label
        {
            Text = "Series/Image:",
            Location = new Point(10, ty),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });
        ty += 18;

        _seriesTemplateBox = new TextBox
        {
            Location = new Point(10, ty),
            Width = 420,
            Height = 25,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        templatesGroup.Controls.Add(_seriesTemplateBox);
        ty += 28;

        var seriesPlaceholdersLabel = new Label
        {
            Text = "Placeholders: {series}, {image}",
            Location = new Point(10, ty),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        templatesGroup.Controls.Add(seriesPlaceholdersLabel);
        CreateTooltipLabel(templatesGroup, seriesPlaceholdersLabel, "Template for series capture.\n{series} = series number, {image} = image number.", 2, 0);
        ty += 20;

        templatesGroup.Controls.Add(new Label
        {
            Text = "Get Prior Comparison:",
            Location = new Point(10, ty),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });
        ty += 18;

        _comparisonTemplateBox = new TextBox
        {
            Location = new Point(10, ty),
            Width = 420,
            Height = 50,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        templatesGroup.Controls.Add(_comparisonTemplateBox);
        ty += 55;

        var comparisonPlaceholdersLabel = new Label
        {
            Text = "Placeholders: {date}, {time}, {description}, {noimages}",
            Location = new Point(10, ty),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        templatesGroup.Controls.Add(comparisonPlaceholdersLabel);
        CreateTooltipLabel(templatesGroup, comparisonPlaceholdersLabel, "Template for Get Prior comparison line.\n{date} = study date, {time} = time (if recent),\n{description} = modality text, {noimages} = 'No Prior Images.' if applicable.", 2, 0);

        y += 305;

        // ========== MACROS SECTION ==========
        var macrosGroup = new GroupBox
        {
            Text = "Macros",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 90),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(macrosGroup);

        int my = 20;
        _macrosEnabledCheck = new CheckBox
        {
            Text = "Enable Macros",
            Location = new Point(10, my),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9),
            Checked = _config.MacrosEnabled
        };
        _macrosEnabledCheck.CheckedChanged += (s, e) =>
        {
            UpdateMacroStates();
        };
        macrosGroup.Controls.Add(_macrosEnabledCheck);

        var editMacrosBtn = new Button
        {
            Text = "Edit...",
            Location = new Point(130, my - 2),
            Size = new Size(60, 22),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        editMacrosBtn.FlatAppearance.BorderColor = Color.Gray;
        editMacrosBtn.Click += OnEditMacrosClick;
        macrosGroup.Controls.Add(editMacrosBtn);

        _macrosCountLabel = new Label
        {
            Text = GetMacrosCountText(),
            Location = new Point(200, my + 2),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        macrosGroup.Controls.Add(_macrosCountLabel);
        CreateTooltipLabel(macrosGroup, _macrosCountLabel, "Auto-insert text snippets based on study description.\nRequires Scrape Mosaic to be enabled.", 2, 0);
        my += 25;

        _macrosBlankLinesCheck = new CheckBox
        {
            Text = "Add blank lines before macro text",
            Location = new Point(30, my),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 9),
            Checked = _config.MacrosBlankLinesBefore
        };
        macrosGroup.Controls.Add(_macrosBlankLinesCheck);
        CreateTooltipLabel(macrosGroup, _macrosBlankLinesCheck, "Add 10 blank lines before macro text\nfor dictation space.");

        UpdateMacroStates();

        y += 100;

        // ========== PICK LISTS SECTION ==========
        var pickListsGroup = new GroupBox
        {
            Text = "Pick Lists",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 90),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(pickListsGroup);

        int py = 20;
        _pickListsEnabledCheck = new CheckBox
        {
            Text = "Enable Pick Lists",
            Location = new Point(10, py),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9),
            Checked = _config.PickListsEnabled
        };
        _pickListsEnabledCheck.CheckedChanged += (s, e) => UpdatePickListStates();
        pickListsGroup.Controls.Add(_pickListsEnabledCheck);

        var editPickListsBtn = new Button
        {
            Text = "Edit...",
            Location = new Point(145, py - 2),
            Size = new Size(60, 22),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        editPickListsBtn.FlatAppearance.BorderColor = Color.Gray;
        editPickListsBtn.Click += OnEditPickListsClick;
        pickListsGroup.Controls.Add(editPickListsBtn);

        _pickListsCountLabel = new Label
        {
            Text = GetPickListsCountText(),
            Location = new Point(215, py + 2),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        pickListsGroup.Controls.Add(_pickListsCountLabel);
        CreateTooltipLabel(pickListsGroup, _pickListsCountLabel, "Show pick list popup for studies matching\nconfigured triggers.", 2, 0);
        py += 25;

        _pickListSkipSingleMatchCheck = new CheckBox
        {
            Text = "Skip list when only one matches",
            Location = new Point(30, py),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 9),
            Checked = _config.PickListSkipSingleMatch
        };
        pickListsGroup.Controls.Add(_pickListSkipSingleMatchCheck);
        CreateTooltipLabel(pickListsGroup, _pickListSkipSingleMatchCheck, "Automatically select if only one pick list\nmatches the study.");
        py += 22;

        _pickListKeepOpenCheck = new CheckBox
        {
            Text = "Keep window open",
            Location = new Point(30, py),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 9),
            Checked = _config.PickListKeepOpen
        };
        pickListsGroup.Controls.Add(_pickListKeepOpenCheck);
        CreateTooltipLabel(pickListsGroup, _pickListKeepOpenCheck, "Keep pick list visible after inserting\n(use number keys for quick selection).");
        py += 22;

        pickListsGroup.Controls.Add(new Label
        {
            Text = "Assign \"Show Pick Lists\" action in Keys tab.",
            Location = new Point(10, py),
            AutoSize = true,
            ForeColor = Color.FromArgb(100, 100, 100),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        });

        UpdatePickListStates();

        return tab;
    }

    private TabPage CreateReferenceTab()
    {
        var tab = new TabPage("Reference")
        {
            BackColor = Color.FromArgb(40, 40, 40)
        };

        var infoBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(40, 40, 40),
            ForeColor = Color.LightGray,
            Font = new Font("Consolas", 9),
            Text = @"=== AHK / External Integration ===

Send Windows Messages to trigger actions from
AutoHotkey or other programs:

WM_TRIGGER_SCRAPE = 0x8001        # Critical Findings
WM_TRIGGER_BEEP = 0x8003          # System Beep
WM_TRIGGER_SHOW_REPORT = 0x8004   # Show Report
WM_TRIGGER_CAPTURE_SERIES = 0x8005 # Capture Series
WM_TRIGGER_GET_PRIOR = 0x8006     # Get Prior
WM_TRIGGER_TOGGLE_RECORD = 0x8007 # Toggle Record
WM_TRIGGER_PROCESS_REPORT = 0x8008 # Process Report
WM_TRIGGER_SIGN_REPORT = 0x8009   # Sign Report
WM_TRIGGER_OPEN_SETTINGS = 0x800A # Open Settings
WM_TRIGGER_CREATE_IMPRESSION = 0x800B # Create Impression
WM_TRIGGER_DISCARD_STUDY = 0x800C # Discard Study
WM_TRIGGER_CHECK_UPDATES = 0x800D # Check for Updates
WM_TRIGGER_SHOW_PICK_LISTS = 0x800E # Show Pick Lists
WM_TRIGGER_CREATE_CRITICAL_NOTE = 0x800F # Create Critical Note

Example AHK:
DetectHiddenWindows, On
PostMessage, 0x8001, 0, 0,, ahk_class WindowsForms

=== Critical Note Creation ===

Create a Critical Communication Note in Clario for any study:
• Map 'Create Critical Note' to hotkey/mic button
• Ctrl+Click on Clinical History window
• Right-click Clinical History → Create Critical Note
• Windows message 0x800F from AHK

=== Debug Tips ===

• Hold Win key while triggering Critical Findings
  to see raw data without pasting.

• Right-click Clinical History or Impression
  windows for context menu and debug options.

=== Config File ===
Settings: %LOCALAPPDATA%\MosaicTools\MosaicToolsSettings.json
"
        };
        tab.Controls.Add(infoBox);

        return tab;
    }

    // ========== NOTIFICATIONS TAB (Clinical History + Alerts) ==========

    private TabPage CreateNotificationsTab()
    {
        var tab = new TabPage("Alerts")
        {
            BackColor = Color.FromArgb(40, 40, 40),
            AutoScroll = true
        };

        int y = 10;
        int groupWidth = 445;

        // ========== CRITICAL STUDIES TRACKER SECTION ==========
        var criticalTrackerGroup = new GroupBox
        {
            Text = "Critical Studies Tracker",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 70),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(criticalTrackerGroup);

        int cy = 20;

        _trackCriticalStudiesCheck = new CheckBox
        {
            Text = "Track critical studies this session",
            Location = new Point(10, cy),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        criticalTrackerGroup.Controls.Add(_trackCriticalStudiesCheck);
        CreateTooltipLabel(criticalTrackerGroup, _trackCriticalStudiesCheck, "Show badge count of critical notes created\nthis session. Click badge to see list.");

        criticalTrackerGroup.Controls.Add(new Label
        {
            Text = "Shows count badge on main bar when critical notes are created",
            Location = new Point(30, cy + 22),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        });

        y += 80;

        // ========== NOTIFICATION BOX SECTION ==========
        var notificationGroup = new GroupBox
        {
            Text = "Notification Box",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 195),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(notificationGroup);

        int ny = 20;

        _showClinicalHistoryCheck = new CheckBox
        {
            Text = "Enable Notification Box",
            Location = new Point(10, ny),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        _showClinicalHistoryCheck.CheckedChanged += (s, e) => UpdateNotificationBoxStates();
        notificationGroup.Controls.Add(_showClinicalHistoryCheck);
        CreateTooltipLabel(notificationGroup, _showClinicalHistoryCheck, "Floating window showing clinical history and alerts.\nRequires Scrape Mosaic to be enabled.");
        ny += 22;

        notificationGroup.Controls.Add(new Label
        {
            Text = "Floating window for clinical history and alerts.",
            Location = new Point(30, ny),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        });
        ny += 22;

        // Clinical History Display options
        _alwaysShowClinicalHistoryCheck = new CheckBox
        {
            Text = "Show clinical history in notification box",
            Location = new Point(30, ny),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        _alwaysShowClinicalHistoryCheck.CheckedChanged += (s, e) => UpdateNotificationBoxStates();
        notificationGroup.Controls.Add(_alwaysShowClinicalHistoryCheck);
        CreateTooltipLabel(notificationGroup, _alwaysShowClinicalHistoryCheck, "Display clinical history text.\nUnchecked = alerts-only mode.");
        ny += 22;

        notificationGroup.Controls.Add(new Label
        {
            Text = "Unchecked = alerts-only mode (hidden until alert)",
            Location = new Point(50, ny),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        });
        ny += 22;

        _hideClinicalHistoryWhenNoStudyCheck = new CheckBox
        {
            Text = "Hide when no study open",
            Location = new Point(50, ny),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        _hideClinicalHistoryWhenNoStudyCheck.CheckedChanged += (s, e) =>
        {
            _config.HideClinicalHistoryWhenNoStudy = _hideClinicalHistoryWhenNoStudyCheck.Checked;
            _mainForm.UpdateClinicalHistoryVisibility();
        };
        notificationGroup.Controls.Add(_hideClinicalHistoryWhenNoStudyCheck);
        CreateTooltipLabel(notificationGroup, _hideClinicalHistoryWhenNoStudyCheck, "Hide notification box when no study\nis open in Mosaic.");
        ny += 22;

        _autoFixClinicalHistoryCheck = new CheckBox
        {
            Text = "Auto-paste corrected history",
            Location = new Point(50, ny),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        notificationGroup.Controls.Add(_autoFixClinicalHistoryCheck);
        CreateTooltipLabel(notificationGroup, _autoFixClinicalHistoryCheck, "Automatically paste corrected clinical history\nto Mosaic when malformed text is detected.");
        ny += 22;

        _showDraftedIndicatorCheck = new CheckBox
        {
            Text = "Show Drafted indicator",
            Location = new Point(50, ny),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        notificationGroup.Controls.Add(_showDraftedIndicatorCheck);

        var greenBorderLabel = new Label
        {
            Text = "(green border)",
            Location = new Point(220, ny + 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(100, 180, 100),
            Font = new Font("Segoe UI", 8)
        };
        notificationGroup.Controls.Add(greenBorderLabel);
        CreateTooltipLabel(notificationGroup, greenBorderLabel, "Green border when report has IMPRESSION section\n(indicates draft in progress).");

        y += 205;

        // ========== ALERTS SECTION ==========
        var alertsGroup = new GroupBox
        {
            Text = "Alert Triggers",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 195),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(alertsGroup);

        int ay = 20;

        alertsGroup.Controls.Add(new Label
        {
            Text = "These alerts appear even in alerts-only mode:",
            Location = new Point(10, ay),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 180, 180),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        });
        ay += 22;

        // Template Mismatch
        _showTemplateMismatchCheck = new CheckBox
        {
            Text = "Template mismatch",
            Location = new Point(10, ay),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        alertsGroup.Controls.Add(_showTemplateMismatchCheck);

        var redBorderLabel = new Label
        {
            Text = "(red border)",
            Location = new Point(170, ay + 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(220, 100, 100),
            Font = new Font("Segoe UI", 8)
        };
        alertsGroup.Controls.Add(redBorderLabel);
        CreateTooltipLabel(alertsGroup, redBorderLabel, "Red border when report template doesn't\nmatch study type.");
        ay += 22;

        // Gender Check
        _genderCheckEnabledCheck = new CheckBox
        {
            Text = "Gender check",
            Location = new Point(10, ay),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        alertsGroup.Controls.Add(_genderCheckEnabledCheck);

        var flashingRedLabel = new Label
        {
            Text = "(flashing red)",
            Location = new Point(170, ay + 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 100, 100),
            Font = new Font("Segoe UI", 8)
        };
        alertsGroup.Controls.Add(flashingRedLabel);
        CreateTooltipLabel(alertsGroup, flashingRedLabel, "Flashing red border for gender-specific terms\nin wrong patient.");
        ay += 22;

        // Stroke Detection
        _strokeDetectionEnabledCheck = new CheckBox
        {
            Text = "Stroke detection",
            Location = new Point(10, ay),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        _strokeDetectionEnabledCheck.CheckedChanged += (s, e) => UpdateNotificationBoxStates();
        alertsGroup.Controls.Add(_strokeDetectionEnabledCheck);

        var purpleBorderLabel = new Label
        {
            Text = "(purple border)",
            Location = new Point(170, ay + 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 130, 220),
            Font = new Font("Segoe UI", 8)
        };
        alertsGroup.Controls.Add(purpleBorderLabel);
        CreateTooltipLabel(alertsGroup, purpleBorderLabel, "Purple border for stroke-related studies.");
        ay += 22;

        _strokeDetectionUseClinicalHistoryCheck = new CheckBox
        {
            Text = "Also use clinical history keywords",
            Location = new Point(30, ay),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        alertsGroup.Controls.Add(_strokeDetectionUseClinicalHistoryCheck);

        var editKeywordsBtn = new Button
        {
            Text = "Edit...",
            Location = new Point(260, ay - 2),
            Size = new Size(50, 20),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 7),
            Cursor = Cursors.Hand
        };
        editKeywordsBtn.FlatAppearance.BorderColor = Color.Gray;
        editKeywordsBtn.Click += OnEditStrokeKeywordsClick;
        alertsGroup.Controls.Add(editKeywordsBtn);
        CreateTooltipLabel(alertsGroup, editKeywordsBtn, "Also check clinical history for stroke keywords.");
        ay += 22;

        _strokeClickToCreateNoteCheck = new CheckBox
        {
            Text = "Click alert to create Clario note",
            Location = new Point(30, ay),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        alertsGroup.Controls.Add(_strokeClickToCreateNoteCheck);
        CreateTooltipLabel(alertsGroup, _strokeClickToCreateNoteCheck, "Click purple alert to create Clario\ncommunication note.");
        ay += 22;

        _strokeAutoCreateNoteCheck = new CheckBox
        {
            Text = "Auto-create Clario note on Process Report",
            Location = new Point(30, ay),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        alertsGroup.Controls.Add(_strokeAutoCreateNoteCheck);
        CreateTooltipLabel(alertsGroup, _strokeAutoCreateNoteCheck, "Automatically create note when pressing\nProcess Report for stroke cases.");

        return tab;
    }

    // ========== BEHAVIOR TAB (Scraping, Scrolling, Focus) ==========

    private TabPage CreateBehaviorTab()
    {
        var tab = new TabPage("Behavior")
        {
            BackColor = Color.FromArgb(40, 40, 40),
            AutoScroll = true
        };

        int y = 10;
        int groupWidth = 445;

        // ========== BACKGROUND MONITORING SECTION ==========
        var monitoringGroup = new GroupBox
        {
            Text = "Background Monitoring",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 70),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(monitoringGroup);

        int my = 20;

        monitoringGroup.Controls.Add(new Label
        {
            Text = "Scrape Mosaic every",
            Location = new Point(10, my),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });

        _scrapeIntervalUpDown = new NumericUpDown
        {
            Location = new Point(160, my - 2),
            Width = 50,
            Minimum = 1,
            Maximum = 30,
            Value = 1,
            BackColor = Color.FromArgb(45, 45, 48),
            ForeColor = Color.White
        };
        monitoringGroup.Controls.Add(_scrapeIntervalUpDown);

        monitoringGroup.Controls.Add(new Label
        {
            Text = "seconds",
            Location = new Point(215, my),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });
        my += 25;

        monitoringGroup.Controls.Add(new Label
        {
            Text = "Keep this at 1s unless you are having massive performance degradation.",
            Location = new Point(30, my),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        });

        y += 80;

        // ========== REPORT PROCESSING SECTION ==========
        var processingGroup = new GroupBox
        {
            Text = "Report Processing",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 165),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(processingGroup);

        int py = 20;

        _scrollToBottomCheck = new CheckBox
        {
            Text = "Scroll to bottom on Process Report",
            Location = new Point(10, py),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        _scrollToBottomCheck.CheckedChanged += (s, e) => UpdateThresholdStates();
        processingGroup.Controls.Add(_scrollToBottomCheck);
        CreateTooltipLabel(processingGroup, _scrollToBottomCheck, "Scroll report to bottom after Process Report\nto show IMPRESSION section.");
        py += 22;

        _showLineCountToastCheck = new CheckBox
        {
            Text = "Show line count toast",
            Location = new Point(30, py),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        processingGroup.Controls.Add(_showLineCountToastCheck);
        CreateTooltipLabel(processingGroup, _showLineCountToastCheck, "Toast showing number of lines\nafter Process Report.");
        py += 25;

        // Smart Scroll Thresholds (compact horizontal layout)
        processingGroup.Controls.Add(new Label
        {
            Text = "Smart Scroll Thresholds (lines):",
            Location = new Point(30, py),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 180, 180),
            Font = new Font("Segoe UI", 8)
        });
        py += 18;

        processingGroup.Controls.Add(new Label { Text = "1 PgDn >", Location = new Point(30, py + 2), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) });
        _scrollThreshold1 = new NumericUpDown
        {
            Location = new Point(90, py),
            Width = 50,
            Minimum = 1,
            Maximum = 500,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        processingGroup.Controls.Add(_scrollThreshold1);

        processingGroup.Controls.Add(new Label { Text = "2 PgDn >", Location = new Point(155, py + 2), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) });
        _scrollThreshold2 = new NumericUpDown
        {
            Location = new Point(215, py),
            Width = 50,
            Minimum = 1,
            Maximum = 500,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        processingGroup.Controls.Add(_scrollThreshold2);

        processingGroup.Controls.Add(new Label { Text = "3 PgDn >", Location = new Point(280, py + 2), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) });
        _scrollThreshold3 = new NumericUpDown
        {
            Location = new Point(340, py),
            Width = 50,
            Minimum = 1,
            Maximum = 500,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        processingGroup.Controls.Add(_scrollThreshold3);
        CreateTooltipLabel(processingGroup, _scrollThreshold3, "Lines at which to add Page Down presses.\n1-2-3 PgDn at each threshold.");
        py += 30;

        // Threshold constraints
        _scrollThreshold1.ValueChanged += (s, e) =>
        {
            if (_scrollThreshold2.Value <= _scrollThreshold1.Value)
                _scrollThreshold2.Value = Math.Min(_scrollThreshold2.Maximum, _scrollThreshold1.Value + 1);
        };
        _scrollThreshold2.ValueChanged += (s, e) =>
        {
            if (_scrollThreshold1.Value >= _scrollThreshold2.Value)
                _scrollThreshold1.Value = Math.Max(_scrollThreshold1.Minimum, _scrollThreshold2.Value - 1);
            if (_scrollThreshold3.Value <= _scrollThreshold2.Value)
                _scrollThreshold3.Value = Math.Min(_scrollThreshold3.Maximum, _scrollThreshold2.Value + 1);
        };
        _scrollThreshold3.ValueChanged += (s, e) =>
        {
            if (_scrollThreshold2.Value >= _scrollThreshold3.Value)
                _scrollThreshold2.Value = Math.Max(_scrollThreshold2.Minimum, _scrollThreshold3.Value - 1);
        };

        _showImpressionCheck = new CheckBox
        {
            Text = "Show Impression popup after Process Report",
            Location = new Point(10, py),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        processingGroup.Controls.Add(_showImpressionCheck);
        CreateTooltipLabel(processingGroup, _showImpressionCheck, "Display impression text after Process Report.\nClicks to dismiss, auto-hides on sign.");
        py += 22;

        _showReportAfterProcessCheck = new CheckBox
        {
            Text = "Show Report overlay after Process Report",
            Location = new Point(10, py),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        processingGroup.Controls.Add(_showReportAfterProcessCheck);
        CreateTooltipLabel(processingGroup, _showReportAfterProcessCheck, "Automatically open report popup after Process Report.\nShows Changes or Rainbow highlighting if enabled.");

        y += 175;

        // ========== INTELEVIEWER SECTION ========== (skip in headless mode)
        if (!App.IsHeadless)
        {
            var inteleviewerGroup = new GroupBox
            {
                Text = "InteleViewer Integration",
                Location = new Point(10, y),
                Size = new Size(groupWidth, 95),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            tab.Controls.Add(inteleviewerGroup);

            int iy = 20;

            inteleviewerGroup.Controls.Add(new Label
            {
                Text = "Window/Level cycle keys:",
                Location = new Point(10, iy),
                AutoSize = true,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9)
            });

            _windowLevelKeysBox = new TextBox
            {
                Location = new Point(170, iy - 2),
                Size = new Size(200, 22),
                BackColor = Color.FromArgb(50, 50, 50),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9)
            };
            inteleviewerGroup.Controls.Add(_windowLevelKeysBox);
            CreateTooltipLabel(inteleviewerGroup, _windowLevelKeysBox, "Keys sent to InteleViewer for\nwindow/level cycling.");
            iy += 25;

            inteleviewerGroup.Controls.Add(new Label
            {
                Text = "Comma-separated keys sent to InteleViewer (e.g., F4, F5, F7, F6)",
                Location = new Point(10, iy),
                AutoSize = true,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8, FontStyle.Italic)
            });
            iy += 18;

            inteleviewerGroup.Controls.Add(new Label
            {
                Text = "Cycles through window/level presets when triggered",
                Location = new Point(10, iy),
                AutoSize = true,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8, FontStyle.Italic)
            });

            y += 105;
        }

        // ========== REPORT CHANGES SECTION ==========
        var reportChangesGroup = new GroupBox
        {
            Text = "Report Changes Highlighting",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 170),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(reportChangesGroup);

        int rcy = 20;

        _showReportChangesCheck = new CheckBox
        {
            Text = "Highlight report changes",
            Location = new Point(10, rcy),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        reportChangesGroup.Controls.Add(_showReportChangesCheck);

        _reportChangesColorPanel = new Panel
        {
            Location = new Point(190, rcy - 2),
            Size = new Size(40, 20),
            BorderStyle = BorderStyle.FixedSingle,
            Cursor = Cursors.Hand,
            BackColor = Color.FromArgb(144, 238, 144)
        };
        _reportChangesColorPanel.Click += OnReportChangesColorClick;
        reportChangesGroup.Controls.Add(_reportChangesColorPanel);

        _reportChangesAlphaSlider = new TrackBar
        {
            Location = new Point(240, rcy - 2),
            Size = new Size(100, 20),
            Minimum = 5,
            Maximum = 100,
            TickStyle = TickStyle.None,
            Value = 30,
            BackColor = Color.FromArgb(45, 45, 48),
            AutoSize = false
        };
        _reportChangesAlphaSlider.ValueChanged += (s, e) => { _reportChangesAlphaLabel.Text = $"{_reportChangesAlphaSlider.Value}%"; UpdateReportChangesPreview(); };
        reportChangesGroup.Controls.Add(_reportChangesAlphaSlider);

        _reportChangesAlphaLabel = new Label
        {
            Text = "30%",
            Location = new Point(345, rcy),
            AutoSize = true,
            ForeColor = Color.Gray
        };
        reportChangesGroup.Controls.Add(_reportChangesAlphaLabel);
        CreateTooltipLabel(reportChangesGroup, _reportChangesAlphaLabel, "Highlight new text in report popup with color.\nClick color box to pick, slider for transparency.");
        rcy += 25;

        _reportChangesPreview = new RichTextBox
        {
            Location = new Point(10, rcy),
            Size = new Size(420, 40),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            ReadOnly = true,
            BorderStyle = BorderStyle.FixedSingle
        };
        reportChangesGroup.Controls.Add(_reportChangesPreview);
        UpdateReportChangesPreview();
        rcy += 42;

        _correlationEnabledCheck = new CheckBox
        {
            Text = "Rainbow Mode (findings-impression correlation)",
            Location = new Point(10, rcy),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        reportChangesGroup.Controls.Add(_correlationEnabledCheck);
        CreateTooltipLabel(reportChangesGroup, _correlationEnabledCheck, "Color-codes matching concepts between Findings and Impression.\nEach impression item and its related findings are highlighted in the same color.\nClick report popup to cycle between Changes and Rainbow modes.");
        rcy += 25;

        _reportTransparentCheck = new CheckBox
        {
            Text = "Transparent overlay (see image through report)",
            Location = new Point(10, rcy),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        _reportTransparentCheck.CheckedChanged += (s, e) =>
        {
            _reportTransparencySlider.Enabled = _reportTransparentCheck.Checked;
        };
        reportChangesGroup.Controls.Add(_reportTransparentCheck);
        CreateTooltipLabel(reportChangesGroup, _reportTransparentCheck, "When enabled, the report popup background is semi-transparent\nso the radiology image shows through. Text stays fully readable.\nWhen disabled, uses the original opaque dark background.");
        rcy += 25;

        var transparencyCaption = new Label
        {
            Text = "Overlay transparency:",
            Location = new Point(30, rcy + 2),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        reportChangesGroup.Controls.Add(transparencyCaption);

        _reportTransparencySlider = new TrackBar
        {
            Location = new Point(170, rcy - 2),
            Size = new Size(175, 20),
            Minimum = 10,
            Maximum = 100,
            TickStyle = TickStyle.None,
            Value = 55,
            BackColor = Color.FromArgb(45, 45, 48),
            AutoSize = false
        };
        _reportTransparencySlider.ValueChanged += (s, e) => { _reportTransparencyLabel.Text = $"{_reportTransparencySlider.Value}%"; };
        reportChangesGroup.Controls.Add(_reportTransparencySlider);

        _reportTransparencyLabel = new Label
        {
            Text = "55%",
            Location = new Point(350, rcy + 2),
            AutoSize = true,
            ForeColor = Color.Gray
        };
        reportChangesGroup.Controls.Add(_reportTransparencyLabel);
        CreateTooltipLabel(reportChangesGroup, _reportTransparencyLabel, "Controls how see-through the report overlay is.\nLower values let more of the image show through.");

        y += 180;

        // ========== RVUCOUNTER SECTION ==========
        var rvuGroup = new GroupBox
        {
            Text = "RVUCounter Integration",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 185),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(rvuGroup);

        int ry = 20;

        _rvuCounterEnabledCheck = new CheckBox
        {
            Text = "Send study events to RVUCounter",
            Location = new Point(10, ry),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        _rvuCounterEnabledCheck.CheckedChanged += (s, e) =>
        {
            if (!_rvuCounterEnabledCheck.Checked)
            {
                var result = MessageBox.Show(
                    "Warning: Disabling RVUCounter integration will prevent MosaicTools from tracking your RVU counts.\n\n" +
                    "Only disable this if you know what you're doing and don't use RVUCounter.\n\n" +
                    "Are you sure you want to disable it?",
                    "Disable RVUCounter?",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    _rvuCounterEnabledCheck.Checked = true;
                }
            }
        };
        rvuGroup.Controls.Add(_rvuCounterEnabledCheck);
        CreateTooltipLabel(rvuGroup, _rvuCounterEnabledCheck, "When enabled, sends signed study events to RVUCounter\nso it can track your RVU productivity.");
        ry += 22;

        // Metrics checkboxes
        var metricsLabel = new Label
        {
            Text = "Display metrics:",
            Location = new Point(10, ry + 2),
            AutoSize = true,
            ForeColor = Color.LightGray,
            Font = new Font("Segoe UI", 9)
        };
        rvuGroup.Controls.Add(metricsLabel);
        ry += 20;

        _rvuMetricTotalCheck = new CheckBox { Text = "Total", Location = new Point(20, ry), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        _rvuMetricPerHourCheck = new CheckBox { Text = "RVU/h", Location = new Point(80, ry), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        _rvuMetricCurrentHourCheck = new CheckBox { Text = "This Hour", Location = new Point(150, ry), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        _rvuMetricPriorHourCheck = new CheckBox { Text = "Prev Hour", Location = new Point(235, ry), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        _rvuMetricEstTotalCheck = new CheckBox { Text = "Est Total", Location = new Point(325, ry), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        rvuGroup.Controls.Add(_rvuMetricTotalCheck);
        rvuGroup.Controls.Add(_rvuMetricPerHourCheck);
        rvuGroup.Controls.Add(_rvuMetricCurrentHourCheck);
        rvuGroup.Controls.Add(_rvuMetricPriorHourCheck);
        rvuGroup.Controls.Add(_rvuMetricEstTotalCheck);

        // Wire up checkbox changes to enable/disable layout combo
        EventHandler updateLayoutState = (s, e) => UpdateOverflowLayoutState();
        _rvuMetricTotalCheck.CheckedChanged += updateLayoutState;
        _rvuMetricPerHourCheck.CheckedChanged += updateLayoutState;
        _rvuMetricCurrentHourCheck.CheckedChanged += updateLayoutState;
        _rvuMetricPriorHourCheck.CheckedChanged += updateLayoutState;
        _rvuMetricEstTotalCheck.CheckedChanged += updateLayoutState;
        ry += 22;

        // Layout combo (for 3+ metrics)
        _rvuOverflowLayoutLabel = new Label
        {
            Text = "3+ layout:",
            Location = new Point(10, ry + 2),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        rvuGroup.Controls.Add(_rvuOverflowLayoutLabel);

        _rvuOverflowLayoutCombo = new ComboBox
        {
            Location = new Point(80, ry - 1),
            Size = new Size(110, 22),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Enabled = false
        };
        _rvuOverflowLayoutCombo.Items.AddRange(new object[] { "Horizontal", "Vertical Stack", "Hover Popup", "Carousel" });
        rvuGroup.Controls.Add(_rvuOverflowLayoutCombo);
        CreateTooltipLabel(rvuGroup, _rvuOverflowLayoutCombo, "Layout when 3+ metrics selected:\nHorizontal = wide bar, Vertical = stacked rows,\nHover = popup on mouse-over, Carousel = cycles every 4s.");

        // Keep legacy combo hidden but functional for save/load
        _rvuDisplayModeCombo = new ComboBox { Visible = false };
        _rvuDisplayModeCombo.Items.AddRange(new object[] { "Total", "RVU/h", "Both" });
        rvuGroup.Controls.Add(_rvuDisplayModeCombo);

        ry += 22;

        _rvuGoalEnabledCheck = new CheckBox
        {
            Text = "Goal:",
            Location = new Point(10, ry),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        rvuGroup.Controls.Add(_rvuGoalEnabledCheck);

        _rvuGoalValueBox = new NumericUpDown
        {
            Location = new Point(65, ry - 2),
            Size = new Size(55, 20),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            Minimum = 1,
            Maximum = 100,
            DecimalPlaces = 1,
            Increment = 0.5m,
            Value = 10
        };
        rvuGroup.Controls.Add(_rvuGoalValueBox);

        var rvuGoalSuffixLabel = new Label
        {
            Text = "/h (colors RVU/h red when below)",
            Location = new Point(122, ry + 2),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        rvuGroup.Controls.Add(rvuGoalSuffixLabel);
        ry += 24;

        _rvuCounterPathBox = new TextBox
        {
            Location = new Point(10, ry),
            Size = new Size(300, 20),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.LightGray,
            ReadOnly = true
        };
        rvuGroup.Controls.Add(_rvuCounterPathBox);

        var findRvuBtn = new Button
        {
            Text = "Find",
            Location = new Point(320, ry - 2),
            Size = new Size(45, 22),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        findRvuBtn.Click += OnFindRvuCounterClick;
        rvuGroup.Controls.Add(findRvuBtn);

        var browseRvuBtn = new Button
        {
            Text = "...",
            Location = new Point(370, ry - 2),
            Size = new Size(30, 22),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        browseRvuBtn.Click += OnBrowseDatabaseClick;
        rvuGroup.Controls.Add(browseRvuBtn);
        ry += 25;

        _rvuCounterStatusLabel = new Label
        {
            Text = "",
            Location = new Point(10, ry),
            Size = new Size(300, 18),
            ForeColor = Color.Gray
        };
        rvuGroup.Controls.Add(_rvuCounterStatusLabel);

        y += 195;

        // ========== INPATIENT XR HANDLING SECTION ==========
        var inpatientGroup = new GroupBox
        {
            Text = "Inpatient XR Handling",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 85),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(inpatientGroup);

        int ipy = 20;

        _ignoreInpatientDraftedCheck = new CheckBox
        {
            Text = "Ignore Inpatient Drafted (select all after auto-insertions)",
            Location = new Point(10, ipy),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        _ignoreInpatientDraftedCheck.CheckedChanged += (s, e) =>
        {
            _ignoreInpatientAllXrRadio.Enabled = _ignoreInpatientDraftedCheck.Checked;
            _ignoreInpatientChestOnlyRadio.Enabled = _ignoreInpatientDraftedCheck.Checked;
        };
        inpatientGroup.Controls.Add(_ignoreInpatientDraftedCheck);
        CreateTooltipLabel(inpatientGroup, _ignoreInpatientDraftedCheck, "Auto-select all text for inpatient XR studies\nafter macro/clinical history insertions.");
        ipy += 25;

        _ignoreInpatientAllXrRadio = new RadioButton
        {
            Text = "All Inpatient XR",
            Location = new Point(30, ipy),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9),
            Checked = true,
            Enabled = false
        };
        inpatientGroup.Controls.Add(_ignoreInpatientAllXrRadio);

        _ignoreInpatientChestOnlyRadio = new RadioButton
        {
            Text = "Inpatient Chest XR only",
            Location = new Point(170, ipy),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9),
            Enabled = false
        };
        inpatientGroup.Controls.Add(_ignoreInpatientChestOnlyRadio);
        CreateTooltipLabel(inpatientGroup, _ignoreInpatientChestOnlyRadio, "Apply to all inpatient X-ray studies\nor only inpatient chest X-rays.");

        y += 95;

        // ========== EXPERIMENTAL SECTION ==========
        var experimentalGroup = new GroupBox
        {
            Text = "Experimental",
            Location = new Point(10, y),
            Size = new Size(groupWidth, 140),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        tab.Controls.Add(experimentalGroup);

        int ey = 20;

        experimentalGroup.Controls.Add(new Label
        {
            Text = "⚠ May change or be removed",
            Location = new Point(10, ey),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 180, 50),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        });
        ey += 22;

        // Network Monitor (experimental)
        _connectivityMonitorEnabledCheck = new CheckBox
        {
            Text = "Network Monitor - show connectivity status dots",
            Location = new Point(10, ey),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        _connectivityMonitorEnabledCheck.CheckedChanged += (s, e) => UpdateNetworkSettingsStates();
        experimentalGroup.Controls.Add(_connectivityMonitorEnabledCheck);
        CreateTooltipLabel(experimentalGroup, _connectivityMonitorEnabledCheck, "Show connectivity dots in widget bar\n(experimental).");
        ey += 25;

        experimentalGroup.Controls.Add(new Label
        {
            Text = "Check every:",
            Location = new Point(30, ey + 2),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });
        _connectivityIntervalUpDown = new NumericUpDown
        {
            Location = new Point(120, ey),
            Width = 50,
            Minimum = 10,
            Maximum = 120,
            Value = 30,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White
        };
        experimentalGroup.Controls.Add(_connectivityIntervalUpDown);
        experimentalGroup.Controls.Add(new Label
        {
            Text = "sec",
            Location = new Point(175, ey + 2),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 9)
        });

        experimentalGroup.Controls.Add(new Label
        {
            Text = "Timeout:",
            Location = new Point(220, ey + 2),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        });
        _connectivityTimeoutUpDown = new NumericUpDown
        {
            Location = new Point(280, ey),
            Width = 50,
            Minimum = 1,
            Maximum = 10,
            Value = 5,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White
        };
        experimentalGroup.Controls.Add(_connectivityTimeoutUpDown);
        experimentalGroup.Controls.Add(new Label
        {
            Text = "sec",
            Location = new Point(335, ey + 2),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 9)
        });
        ey += 28;

        experimentalGroup.Controls.Add(new Label
        {
            Text = "Shows 4 status dots: Mirth, Mosaic, Clario, InteleViewer (placeholder IPs)",
            Location = new Point(30, ey),
            AutoSize = true,
            ForeColor = Color.FromArgb(100, 100, 100),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        });

        return tab;
    }

    // ========== Helper Methods ==========


    // Keep this stub for the actual Experimental tab in InitializeUI
    private TabPage CreateExperimentalTab_Stub()
    {
        var tab = new TabPage("Experimental")
        {
            BackColor = Color.FromArgb(40, 40, 40)
        };

        // Create scrollable container
        var scrollPanel = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = Color.FromArgb(40, 40, 40)
        };
        tab.Controls.Add(scrollPanel);

        int y = 20;

        // Warning label
        var warningLabel = new Label
        {
            Text = "⚠ Experimental features - may change or be removed",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 180, 50),
            Font = new Font("Segoe UI", 9, FontStyle.Italic)
        };
        scrollPanel.Controls.Add(warningLabel);
        y += 35;

        // ========== PICK LISTS SECTION (first for easier testing) ==========
        var pickListsHeader = new Label
        {
            Text = "Pick Lists",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        scrollPanel.Controls.Add(pickListsHeader);
        y += 25;

        // Enable Pick Lists checkbox
        _pickListsEnabledCheck = new CheckBox
        {
            Text = "Enable Pick Lists",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.White,
            Checked = _config.PickListsEnabled
        };
        _pickListsEnabledCheck.CheckedChanged += (s, e) => UpdatePickListStates();
        scrollPanel.Controls.Add(_pickListsEnabledCheck);
        y += 25;

        // Skip single match checkbox
        _pickListSkipSingleMatchCheck = new CheckBox
        {
            Text = "Skip list selector when only one matches",
            Location = new Point(40, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Checked = _config.PickListSkipSingleMatch
        };
        scrollPanel.Controls.Add(_pickListSkipSingleMatchCheck);
        y += 25;

        // Keep pick list open checkbox
        _pickListKeepOpenCheck = new CheckBox
        {
            Text = "Keep pick list window open",
            Location = new Point(40, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Checked = _config.PickListKeepOpen
        };
        var pickListToolTip = new ToolTip { AutoPopDelay = 10000, InitialDelay = 300 };
        pickListToolTip.SetToolTip(_pickListKeepOpenCheck, "When enabled, the pick list stays visible after inserting text.\nClick the popup to use number keys.");
        scrollPanel.Controls.Add(_pickListKeepOpenCheck);
        y += 25;

        // Edit Pick Lists button
        var editPickListsBtn = new Button
        {
            Text = "Edit Pick Lists...",
            Location = new Point(20, y),
            Size = new Size(130, 27),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        editPickListsBtn.FlatAppearance.BorderColor = Color.Gray;
        editPickListsBtn.Click += OnEditPickListsClick;
        scrollPanel.Controls.Add(editPickListsBtn);

        // Pick lists count label
        _pickListsCountLabel = new Label
        {
            Text = GetPickListsCountText(),
            Location = new Point(160, y + 5),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        scrollPanel.Controls.Add(_pickListsCountLabel);
        y += 35;

        // Hint about assigning hotkey
        var pickListsHint = new Label
        {
            Text = "Assign \"Show Pick Lists\" action to a hotkey or mic button in the Keys tab.",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(120, 120, 120),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        scrollPanel.Controls.Add(pickListsHint);
        y += 35;

        UpdatePickListStates();

        // ========== RVUCOUNTER SECTION ==========
        var rvuHeader = new Label
        {
            Text = "RVUCounter Integration",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        scrollPanel.Controls.Add(rvuHeader);
        y += 25;

        // RVUCounter enabled checkbox
        _rvuCounterEnabledCheck = new CheckBox
        {
            Text = "Read RVUCounter shift total",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.White
        };
        scrollPanel.Controls.Add(_rvuCounterEnabledCheck);
        y += 25;

        // Metrics checkboxes
        var metricsLabel2 = new Label
        {
            Text = "Display metrics:",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.LightGray
        };
        scrollPanel.Controls.Add(metricsLabel2);
        y += 20;

        _rvuMetricTotalCheck = new CheckBox { Text = "Total", Location = new Point(30, y), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        _rvuMetricPerHourCheck = new CheckBox { Text = "RVU/h", Location = new Point(95, y), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        _rvuMetricCurrentHourCheck = new CheckBox { Text = "This Hour", Location = new Point(170, y), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        _rvuMetricPriorHourCheck = new CheckBox { Text = "Prev Hour", Location = new Point(260, y), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        _rvuMetricEstTotalCheck = new CheckBox { Text = "Est Total", Location = new Point(350, y), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 8) };
        scrollPanel.Controls.Add(_rvuMetricTotalCheck);
        scrollPanel.Controls.Add(_rvuMetricPerHourCheck);
        scrollPanel.Controls.Add(_rvuMetricCurrentHourCheck);
        scrollPanel.Controls.Add(_rvuMetricPriorHourCheck);
        scrollPanel.Controls.Add(_rvuMetricEstTotalCheck);

        EventHandler updateLayoutState2 = (s, e) => UpdateOverflowLayoutState();
        _rvuMetricTotalCheck.CheckedChanged += updateLayoutState2;
        _rvuMetricPerHourCheck.CheckedChanged += updateLayoutState2;
        _rvuMetricCurrentHourCheck.CheckedChanged += updateLayoutState2;
        _rvuMetricPriorHourCheck.CheckedChanged += updateLayoutState2;
        _rvuMetricEstTotalCheck.CheckedChanged += updateLayoutState2;
        y += 24;

        // Layout combo for 3+ metrics
        _rvuOverflowLayoutLabel = new Label
        {
            Text = "3+ layout:",
            Location = new Point(20, y + 2),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        scrollPanel.Controls.Add(_rvuOverflowLayoutLabel);

        _rvuOverflowLayoutCombo = new ComboBox
        {
            Location = new Point(90, y - 1),
            Size = new Size(120, 22),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Enabled = false
        };
        _rvuOverflowLayoutCombo.Items.AddRange(new object[] { "Horizontal", "Vertical Stack", "Hover Popup", "Carousel" });
        scrollPanel.Controls.Add(_rvuOverflowLayoutCombo);

        // Hidden legacy combo
        _rvuDisplayModeCombo = new ComboBox { Visible = false };
        _rvuDisplayModeCombo.Items.AddRange(new object[] { "Total", "RVU/h", "Both" });
        scrollPanel.Controls.Add(_rvuDisplayModeCombo);
        y += 28;

        // Goal setting
        _rvuGoalEnabledCheck = new CheckBox
        {
            Text = "Color by goal:",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.White
        };
        scrollPanel.Controls.Add(_rvuGoalEnabledCheck);

        _rvuGoalValueBox = new NumericUpDown
        {
            Location = new Point(130, y - 2),
            Size = new Size(60, 22),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            Minimum = 1,
            Maximum = 100,
            DecimalPlaces = 1,
            Increment = 0.5m,
            Value = 10
        };
        scrollPanel.Controls.Add(_rvuGoalValueBox);

        var rvuGoalSuffixLabel = new Label
        {
            Text = "RVU/h (blue if met, red if below)",
            Location = new Point(195, y + 2),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        scrollPanel.Controls.Add(rvuGoalSuffixLabel);
        y += 28;

        // Path label
        var pathLabel = new Label
        {
            Text = "Database path:",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.LightGray
        };
        scrollPanel.Controls.Add(pathLabel);
        y += 22;

        // Path textbox (read-only display)
        _rvuCounterPathBox = new TextBox
        {
            Location = new Point(20, y),
            Size = new Size(410, 23),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.LightGray,
            ReadOnly = true
        };
        scrollPanel.Controls.Add(_rvuCounterPathBox);
        y += 30;

        // Find button (auto-search)
        var findBtn = new Button
        {
            Text = "Auto-Find",
            Location = new Point(20, y),
            Size = new Size(90, 27),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        findBtn.Click += OnFindRvuCounterClick;
        scrollPanel.Controls.Add(findBtn);

        // Browse button (manual select)
        var browseBtn = new Button
        {
            Text = "Browse...",
            Location = new Point(120, y),
            Size = new Size(90, 27),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        browseBtn.Click += OnBrowseDatabaseClick;
        scrollPanel.Controls.Add(browseBtn);
        y += 35;

        // Status label
        _rvuCounterStatusLabel = new Label
        {
            Text = "",
            Location = new Point(20, y),
            Size = new Size(410, 20),
            ForeColor = Color.Gray
        };
        scrollPanel.Controls.Add(_rvuCounterStatusLabel);
        y += 40;

        // Report Changes section header
        var changesHeader = new Label
        {
            Text = "Report Changes Highlighting",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };
        scrollPanel.Controls.Add(changesHeader);
        y += 25;

        // Show report changes checkbox
        _showReportChangesCheck = new CheckBox
        {
            Text = "Highlight changes when viewing report after Process",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.White,
            Checked = _config.ShowReportChanges
        };
        scrollPanel.Controls.Add(_showReportChangesCheck);
        y += 25;

        // Explanation label
        var changesHintLabel = new Label
        {
            Text = "Shows what you dictated by highlighting differences from the original template.",
            Location = new Point(40, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        scrollPanel.Controls.Add(changesHintLabel);
        y += 25;

        // Color picker row
        var colorLabel = new Label
        {
            Text = "Highlight color:",
            Location = new Point(20, y + 3),
            AutoSize = true,
            ForeColor = Color.LightGray
        };
        scrollPanel.Controls.Add(colorLabel);

        // Color preview panel (clickable)
        _reportChangesColorPanel = new Panel
        {
            Location = new Point(120, y),
            Size = new Size(60, 24),
            BorderStyle = BorderStyle.FixedSingle,
            Cursor = Cursors.Hand
        };
        try
        {
            _reportChangesColorPanel.BackColor = ColorTranslator.FromHtml(_config.ReportChangesColor);
        }
        catch
        {
            _reportChangesColorPanel.BackColor = Color.FromArgb(144, 238, 144); // Light green
        }
        _reportChangesColorPanel.Click += OnReportChangesColorClick;
        scrollPanel.Controls.Add(_reportChangesColorPanel);

        // Color hex label
        var colorHexLabel = new Label
        {
            Text = _config.ReportChangesColor,
            Location = new Point(190, y + 3),
            AutoSize = true,
            ForeColor = Color.Gray,
            Tag = "colorHex"
        };
        scrollPanel.Controls.Add(colorHexLabel);
        y += 35;

        // Opacity slider row
        var opacityLabel = new Label
        {
            Text = "Opacity:",
            Location = new Point(20, y + 3),
            AutoSize = true,
            ForeColor = Color.LightGray
        };
        scrollPanel.Controls.Add(opacityLabel);

        _reportChangesAlphaSlider = new TrackBar
        {
            Location = new Point(90, y),
            Size = new Size(150, 30),
            Minimum = 5,
            Maximum = 100,
            TickFrequency = 10,
            Value = Math.Clamp(_config.ReportChangesAlpha, 5, 100),
            BackColor = Color.FromArgb(45, 45, 48)
        };
        _reportChangesAlphaSlider.ValueChanged += (s, e) =>
        {
            _reportChangesAlphaLabel.Text = $"{_reportChangesAlphaSlider.Value}%";
            UpdateReportChangesPreview();
        };
        scrollPanel.Controls.Add(_reportChangesAlphaSlider);

        _reportChangesAlphaLabel = new Label
        {
            Text = $"{_reportChangesAlphaSlider.Value}%",
            Location = new Point(245, y + 3),
            AutoSize = true,
            ForeColor = Color.Gray
        };
        scrollPanel.Controls.Add(_reportChangesAlphaLabel);
        y += 40;

        // Live preview
        var previewLabel = new Label
        {
            Text = "Preview:",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.LightGray
        };
        scrollPanel.Controls.Add(previewLabel);
        y += 20;

        _reportChangesPreview = new RichTextBox
        {
            Location = new Point(20, y),
            Size = new Size(340, 60),
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.Gainsboro,
            Font = new Font("Consolas", 10),
            ReadOnly = true,
            BorderStyle = BorderStyle.FixedSingle,
            Text = "Normal text. This text was dictated. More normal text."
        };
        scrollPanel.Controls.Add(_reportChangesPreview);

        // Apply initial preview highlighting
        UpdateReportChangesPreview();

        return tab;
    }

    private void OnReportChangesColorClick(object? sender, EventArgs e)
    {
        using var colorDialog = new ColorDialog
        {
            Color = _reportChangesColorPanel.BackColor,
            FullOpen = true
        };

        if (colorDialog.ShowDialog(this) == DialogResult.OK)
        {
            _reportChangesColorPanel.BackColor = colorDialog.Color;
            _config.ReportChangesColor = ColorTranslator.ToHtml(colorDialog.Color);

            // Update hex label
            foreach (Control c in _reportChangesColorPanel.Parent!.Controls)
            {
                if (c is Label lbl && lbl.Tag?.ToString() == "colorHex")
                {
                    lbl.Text = _config.ReportChangesColor;
                    break;
                }
            }

            // Update preview
            UpdateReportChangesPreview();
        }
    }

    private void UpdateReportChangesPreview()
    {
        if (_reportChangesPreview == null) return;

        // Reset text and formatting
        _reportChangesPreview.Text = "Normal text. This text was dictated. More normal text.";
        _reportChangesPreview.SelectAll();
        _reportChangesPreview.SelectionBackColor = _reportChangesPreview.BackColor;
        _reportChangesPreview.Select(0, 0);

        // Calculate blended highlight color
        var baseColor = _reportChangesColorPanel.BackColor;
        var bg = _reportChangesPreview.BackColor;
        float alpha = _reportChangesAlphaSlider.Value / 100f;
        var highlightColor = Color.FromArgb(
            (int)(bg.R + (baseColor.R - bg.R) * alpha),
            (int)(bg.G + (baseColor.G - bg.G) * alpha),
            (int)(bg.B + (baseColor.B - bg.B) * alpha)
        );

        // Highlight the "dictated" portion
        const string highlightText = "This text was dictated.";
        int idx = _reportChangesPreview.Text.IndexOf(highlightText);
        if (idx >= 0)
        {
            _reportChangesPreview.Select(idx, highlightText.Length);
            _reportChangesPreview.SelectionBackColor = highlightColor;
            _reportChangesPreview.Select(0, 0);
        }
    }

    private void OnEditPickListsClick(object? sender, EventArgs e)
    {
        try
        {
            var editor = new PickListEditorForm(_config);
            var result = editor.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                // Config was saved in the editor form
                _pickListsCountLabel.Text = GetPickListsCountText();
            }
            editor.Dispose();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error opening Pick List Editor:\n{ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void UpdatePickListStates()
    {
        bool enabled = _pickListsEnabledCheck.Checked;
        _pickListSkipSingleMatchCheck.Enabled = enabled;
        _pickListKeepOpenCheck.Enabled = enabled;
    }

    private string GetPickListsCountText()
    {
        var count = _config.PickLists.Count;
        var enabledCount = _config.PickLists.Count(pl => pl.Enabled);
        if (count == 0)
            return "No pick lists configured";
        if (enabledCount == count)
            return $"{count} pick list{(count == 1 ? "" : "s")} configured";
        return $"{enabledCount} of {count} pick list{(count == 1 ? "" : "s")} enabled";
    }

    private void UpdateNetworkSettingsStates()
    {
        bool enabled = _connectivityMonitorEnabledCheck.Checked;
        _connectivityIntervalUpDown.Enabled = enabled;
        _connectivityTimeoutUpDown.Enabled = enabled;
    }

    private void OnEditMacrosClick(object? sender, EventArgs e)
    {
        using (var editor = new MacroEditorForm(_config))
        {
            if (editor.ShowDialog() == DialogResult.OK)
            {
                _macrosCountLabel.Text = GetMacrosCountText();
            }
        }
    }

    private string GetMacrosCountText()
    {
        var count = _config.Macros.Count;
        var enabledCount = _config.Macros.Count(m => m.Enabled);
        if (count == 0)
            return "No macros configured";
        if (enabledCount == count)
            return $"{count} macro{(count == 1 ? "" : "s")} configured";
        return $"{enabledCount} of {count} macro{(count == 1 ? "" : "s")} enabled";
    }

    private void OnEditStrokeKeywordsClick(object? sender, EventArgs e)
    {
        var currentKeywords = string.Join(Environment.NewLine, _config.StrokeClinicalHistoryKeywords);
        var input = InputBox.Show(
            "Enter stroke detection keywords (one per line):\n\nThese keywords are checked against the clinical history when 'Also use clinical history keywords' is enabled.",
            "Edit Stroke Keywords",
            currentKeywords,
            true);

        if (input != null)
        {
            var keywords = input.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(k => k.Trim())
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .ToList();
            _config.StrokeClinicalHistoryKeywords = keywords;
            _config.Save();
        }
    }

    private void UpdateOverflowLayoutState()
    {
        int checkedCount = 0;
        if (_rvuMetricTotalCheck.Checked) checkedCount++;
        if (_rvuMetricPerHourCheck.Checked) checkedCount++;
        if (_rvuMetricCurrentHourCheck.Checked) checkedCount++;
        if (_rvuMetricPriorHourCheck.Checked) checkedCount++;
        if (_rvuMetricEstTotalCheck.Checked) checkedCount++;

        bool enable = checkedCount >= 3;
        _rvuOverflowLayoutCombo.Enabled = enable;
        _rvuOverflowLayoutLabel.ForeColor = enable ? Color.LightGray : Color.Gray;
    }

    private void OnFindRvuCounterClick(object? sender, EventArgs e)
    {
        string? foundDbPath = null;

        // 1. Look in same directory as MosaicTools
        var appDir = AppContext.BaseDirectory;
        var parentDir = Path.GetDirectoryName(appDir.TrimEnd(Path.DirectorySeparatorChar));

        if (!string.IsNullOrEmpty(parentDir))
        {
            // Check sibling folders
            foreach (var dir in Directory.GetDirectories(parentDir))
            {
                var dbPath = Path.Combine(dir, "data", "rvu_records.db");
                if (File.Exists(dbPath))
                {
                    foundDbPath = dbPath;
                    break;
                }
            }
        }

        // 2. Try Desktop
        if (foundDbPath == null)
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            foreach (var dir in Directory.GetDirectories(desktop))
            {
                var dbPath = Path.Combine(dir, "data", "rvu_records.db");
                if (File.Exists(dbPath))
                {
                    foundDbPath = dbPath;
                    break;
                }
            }
        }

        if (foundDbPath != null)
        {
            _rvuCounterPathBox.Text = foundDbPath;
            _config.RvuCounterPath = foundDbPath;
            _rvuCounterStatusLabel.Text = "✓ Database found";
            _rvuCounterStatusLabel.ForeColor = Color.LightGreen;
        }
        else
        {
            MessageBox.Show(
                "Could not find RVUCounter automatically.\nUse 'Browse...' to select the database file manually.",
                "RVUCounter Not Found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }

    private void OnBrowseDatabaseClick(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Select RVUCounter Database",
            Filter = "SQLite Database (*.db)|*.db|All Files (*.*)|*.*",
            FileName = "rvu_records.db"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            _rvuCounterPathBox.Text = dialog.FileName;
            _config.RvuCounterPath = dialog.FileName;
            _rvuCounterStatusLabel.Text = "✓ Database selected";
            _rvuCounterStatusLabel.ForeColor = Color.LightGreen;
        }
    }

    // ========== Helper Methods for Tab State Management ==========

    private void UpdateThresholdStates()
    {
        bool thresholdsEnabled = _scrollToBottomCheck.Checked;
        _scrollThreshold1.Enabled = thresholdsEnabled;
        _scrollThreshold2.Enabled = thresholdsEnabled;
        _scrollThreshold3.Enabled = thresholdsEnabled;
        _showLineCountToastCheck.Enabled = thresholdsEnabled;

        // Update notification box states
        UpdateNotificationBoxStates();
    }

    private void UpdateNotificationBoxStates()
    {
        bool notificationBoxEnabled = _showClinicalHistoryCheck.Checked;
        bool alwaysShowEnabled = notificationBoxEnabled && _alwaysShowClinicalHistoryCheck.Checked;

        // "Always show" and its children depend on notification box being enabled
        _alwaysShowClinicalHistoryCheck.Enabled = notificationBoxEnabled;

        // These settings only make sense in "always show" mode
        _hideClinicalHistoryWhenNoStudyCheck.Enabled = alwaysShowEnabled;
        _autoFixClinicalHistoryCheck.Enabled = alwaysShowEnabled;
        _showDraftedIndicatorCheck.Enabled = alwaysShowEnabled;

        // Alerts are enabled when notification box is enabled (regardless of always-show)
        _showTemplateMismatchCheck.Enabled = notificationBoxEnabled;
        _genderCheckEnabledCheck.Enabled = notificationBoxEnabled;
        _strokeDetectionEnabledCheck.Enabled = notificationBoxEnabled;
        _strokeDetectionUseClinicalHistoryCheck.Enabled = notificationBoxEnabled && _strokeDetectionEnabledCheck.Checked;
    }

    // Old CreateTemplatesTab removed - content moved to CreateTextAutomationTab

    // Macros tab state
    private CheckBox _macrosEnabledCheck = null!;
    private CheckBox _macrosBlankLinesCheck = null!;

    private void UpdateMacroStates()
    {
        bool enabled = _macrosEnabledCheck.Checked;
        _macrosBlankLinesCheck.Enabled = enabled;
        _macrosBlankLinesCheck.ForeColor = enabled ? Color.White : Color.Gray;
    }

    private static Label CreateLabel(string text, int x, int y, int width)
    {
        return new Label
        {
            Text = text,
            Location = new Point(x, y),
            Width = width,
            ForeColor = Color.White
        };
    }

    /// <summary>
    /// Creates a "?" tooltip label positioned to the right of the anchor control.
    /// </summary>
    private Label CreateTooltipLabel(Control parent, Control anchor, string tooltip, int offsetX = 2, int offsetY = 0)
    {
        var label = new Label
        {
            Text = "?",
            Size = new Size(14, 14),
            Location = new Point(anchor.Right + offsetX, anchor.Top + offsetY),
            ForeColor = Color.FromArgb(100, 150, 200),
            Font = new Font("Segoe UI", 7, FontStyle.Bold),
            Cursor = Cursors.Help,
            Visible = _config.ShowTooltips
        };
        parent.Controls.Add(label);
        _settingsToolTip.SetToolTip(label, tooltip);
        _tooltipLabels.Add(label);
        return label;
    }

    /// <summary>
    /// Creates a "?" tooltip label at an absolute position.
    /// </summary>
    private Label CreateTooltipLabelAt(Control parent, int x, int y, string tooltip)
    {
        var label = new Label
        {
            Text = "?",
            Size = new Size(14, 14),
            Location = new Point(x, y),
            ForeColor = Color.FromArgb(100, 150, 200),
            Font = new Font("Segoe UI", 7, FontStyle.Bold),
            Cursor = Cursors.Help,
            Visible = _config.ShowTooltips
        };
        parent.Controls.Add(label);
        _settingsToolTip.SetToolTip(label, tooltip);
        _tooltipLabels.Add(label);
        return label;
    }

    /// <summary>
    /// Updates visibility of all tooltip labels based on the ShowTooltips setting.
    /// </summary>
    private void UpdateTooltipVisibility()
    {
        bool visible = _showTooltipsCheck?.Checked ?? _config.ShowTooltips;
        foreach (var label in _tooltipLabels)
        {
            label.Visible = visible;
        }
    }

    private void SetupHotkeyCapture(TextBox tb)
    {
        tb.KeyDown += (s, e) =>
        {
            e.SuppressKeyPress = true;

            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
            {
                tb.Text = "";
                return;
            }

            // Ignore modifier-only presses
            if (e.KeyCode == Keys.ControlKey || e.KeyCode == Keys.ShiftKey || 
                e.KeyCode == Keys.Menu || e.KeyCode is Keys.LWin or Keys.RWin)
            {
                return;
            }

            var mods = new List<string>();
            if (e.Control) mods.Add("ctrl");
            if (e.Shift) mods.Add("shift");
            if (e.Alt) mods.Add("alt");

            string keyName = e.KeyCode.ToString().ToLower();
            
            // Map common keys to standard names
            keyName = keyName switch
            {
                "return" => "enter",
                "next" => "pagedown",
                "prior" => "pageup",
                "oemperiod" => ".",
                "oemcomma" => ",",
                "oemminus" => "-",
                "oemplus" => "=",
                "oemquestion" => "/",
                "oem5" => "\\",
                "oemopenbrackets" => "[",
                "oem6" => "]",
                "oem1" => ";",
                "oemquotes" => "'",
                "oemtilde" => "`",
                _ => keyName
            };
            
            // Handle D1-D0
            if (keyName.Length == 2 && keyName[0] == 'd' && char.IsDigit(keyName[1]))
                keyName = keyName.Substring(1);

            string result = mods.Count > 0 ? string.Join("+", mods) + "+" + keyName : keyName;

            if (Configuration.IsHotkeyRestricted(result))
            {
                MessageBox.Show(
                    $"The hotkey '{result}' is reserved by Mosaic and cannot be used as a trigger for this tool to avoid feedback loops and conflicts.\n\n" +
                    "Please choose a different combination (e.g. including Shift or using different keys).",
                    "Restricted Hotkey",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                tb.Text = "";
            }
            else
            {
                tb.Text = result;
            }
        };
    }
    
    private void LoadSettings()
    {
        _doctorNameBox.Text = _config.DoctorName;
        _startBeepCheck.Checked = _config.StartBeepEnabled;
        _stopBeepCheck.Checked = _config.StopBeepEnabled;
        _startVolumeSlider.Value = VolumeToSlider(_config.StartBeepVolume);
        _stopVolumeSlider.Value = VolumeToSlider(_config.StopBeepVolume);
        _dictationPauseNum.Value = _config.DictationPauseMs;
        _ivHotkeyBox.Text = _config.IvReportHotkey;
        _floatingToolbarCheck.Checked = _config.FloatingToolbarEnabled;
        _indicatorCheck.Checked = _config.IndicatorEnabled;
        _autoStopCheck.Checked = _config.AutoStopDictation;
        _deadManCheck.Checked = _config.DeadManSwitch;
        _autoUpdateCheck.Checked = App.IsHeadless ? false : _config.AutoUpdateEnabled;
        _hideIndicatorWhenNoStudyCheck.Checked = _config.HideIndicatorWhenNoStudy;
        _criticalTemplateBox.Text = _config.CriticalFindingsTemplate;
        _seriesTemplateBox.Text = _config.SeriesImageTemplate;
        _comparisonTemplateBox.Text = _config.ComparisonTemplate;
        _separatePastedItemsCheck.Checked = _config.SeparatePastedItems;
        _reportFontFamilyCombo.SelectedItem = _config.ReportPopupFontFamily;
        if (_reportFontFamilyCombo.SelectedIndex < 0) _reportFontFamilyCombo.SelectedIndex = 0;
        _reportFontSizeNumeric.Value = (decimal)Math.Clamp(_config.ReportPopupFontSize, 7f, 24f);

        // Advanced tab
        _scrollToBottomCheck.Checked = _config.ScrollToBottomOnProcess;
        _showLineCountToastCheck.Checked = _config.ShowLineCountToast;
        _scrapeIntervalUpDown.Value = Math.Clamp(_config.ScrapeIntervalSeconds, 1, 30);
        _showClinicalHistoryCheck.Checked = _config.ShowClinicalHistory;
        _alwaysShowClinicalHistoryCheck.Checked = _config.AlwaysShowClinicalHistory;
        _hideClinicalHistoryWhenNoStudyCheck.Checked = _config.HideClinicalHistoryWhenNoStudy;
        _autoFixClinicalHistoryCheck.Checked = _config.AutoFixClinicalHistory;
        _showDraftedIndicatorCheck.Checked = _config.ShowDraftedIndicator;
        _showTemplateMismatchCheck.Checked = _config.ShowTemplateMismatch;
        _genderCheckEnabledCheck.Checked = _config.GenderCheckEnabled;
        _strokeDetectionEnabledCheck.Checked = _config.StrokeDetectionEnabled;
        _strokeDetectionUseClinicalHistoryCheck.Checked = _config.StrokeDetectionUseClinicalHistory;
        _strokeClickToCreateNoteCheck.Checked = _config.StrokeClickToCreateNote;
        _strokeAutoCreateNoteCheck.Checked = _config.StrokeAutoCreateNote;
        _trackCriticalStudiesCheck.Checked = _config.TrackCriticalStudies;
        _showImpressionCheck.Checked = _config.ShowImpression;
        _showReportAfterProcessCheck.Checked = _config.ShowReportAfterProcess;
        _scrollThreshold1.Value = _config.ScrollThreshold1;
        _scrollThreshold2.Value = _config.ScrollThreshold2;
        _scrollThreshold3.Value = _config.ScrollThreshold3;

        // InteleViewer (not in headless mode)
        if (_windowLevelKeysBox != null)
            _windowLevelKeysBox.Text = string.Join(", ", _config.WindowLevelKeys);

        // Network monitor
        _connectivityMonitorEnabledCheck.Checked = _config.ConnectivityMonitorEnabled;
        _connectivityIntervalUpDown.Value = Math.Clamp(_config.ConnectivityCheckIntervalSeconds, 10, 120);
        _connectivityTimeoutUpDown.Value = Math.Clamp(_config.ConnectivityTimeoutMs / 1000, 1, 10);
        UpdateNetworkSettingsStates();

        // Ignore Inpatient Drafted
        _ignoreInpatientDraftedCheck.Checked = _config.IgnoreInpatientDrafted;
        _ignoreInpatientAllXrRadio.Checked = _config.IgnoreInpatientDraftedMode == 0;
        _ignoreInpatientChestOnlyRadio.Checked = _config.IgnoreInpatientDraftedMode == 1;
        _ignoreInpatientAllXrRadio.Enabled = _config.IgnoreInpatientDrafted;
        _ignoreInpatientChestOnlyRadio.Enabled = _config.IgnoreInpatientDrafted;

        // RVUCounter and Report Changes (no longer in Experimental section)
        _rvuCounterEnabledCheck.Checked = _config.RvuCounterEnabled;
        _rvuDisplayModeCombo.SelectedIndex = (int)_config.RvuDisplayMode;
        // Load metric flags into checkboxes
        _rvuMetricTotalCheck.Checked = _config.RvuMetrics.HasFlag(RvuMetric.Total);
        _rvuMetricPerHourCheck.Checked = _config.RvuMetrics.HasFlag(RvuMetric.PerHour);
        _rvuMetricCurrentHourCheck.Checked = _config.RvuMetrics.HasFlag(RvuMetric.CurrentHour);
        _rvuMetricPriorHourCheck.Checked = _config.RvuMetrics.HasFlag(RvuMetric.PriorHour);
        _rvuMetricEstTotalCheck.Checked = _config.RvuMetrics.HasFlag(RvuMetric.EstimatedTotal);
        _rvuOverflowLayoutCombo.SelectedIndex = (int)_config.RvuOverflowLayout;
        UpdateOverflowLayoutState();
        _rvuGoalEnabledCheck.Checked = _config.RvuGoalEnabled;
        _rvuGoalValueBox.Value = (decimal)Math.Clamp(_config.RvuGoalPerHour, 1, 100);
        _rvuCounterPathBox.Text = _config.RvuCounterPath;
        _showReportChangesCheck.Checked = _config.ShowReportChanges;
        _correlationEnabledCheck.Checked = _config.CorrelationEnabled;
        try
        {
            _reportChangesColorPanel.BackColor = ColorTranslator.FromHtml(_config.ReportChangesColor);
        }
        catch
        {
            _reportChangesColorPanel.BackColor = Color.FromArgb(144, 238, 144);
        }
        _reportChangesAlphaSlider.Value = Math.Clamp(_config.ReportChangesAlpha, 5, 100);
        _reportChangesAlphaLabel.Text = $"{_reportChangesAlphaSlider.Value}%";
        _reportTransparentCheck.Checked = _config.ReportPopupTransparent;
        _reportTransparencySlider.Value = Math.Clamp(_config.ReportPopupTransparency, 10, 100);
        _reportTransparencyLabel.Text = $"{_reportTransparencySlider.Value}%";
        _reportTransparencySlider.Enabled = _reportTransparentCheck.Checked;
        _pickListsEnabledCheck.Checked = _config.PickListsEnabled;
        _pickListSkipSingleMatchCheck.Checked = _config.PickListSkipSingleMatch;
        _pickListKeepOpenCheck.Checked = _config.PickListKeepOpen;
        _pickListsCountLabel.Text = GetPickListsCountText();
        UpdatePickListStates();
        if (!string.IsNullOrEmpty(_config.RvuCounterPath))
        {
            if (File.Exists(_config.RvuCounterPath))
            {
                _rvuCounterStatusLabel.Text = "✓ Database found";
                _rvuCounterStatusLabel.ForeColor = Color.LightGreen;
            }
            else
            {
                _rvuCounterStatusLabel.Text = "✗ Database not found at configured path";
                _rvuCounterStatusLabel.ForeColor = Color.Salmon;
            }
        }

        UpdateThresholdStates();

        UpdateVolumeLabels();
    }
    
    private void UpdateVolumeLabels()
    {
        _startVolLabel.Text = $"{_startVolumeSlider.Value}%";
        _stopVolLabel.Text = $"{_stopVolumeSlider.Value}%";
    }

    private static double SliderToVolume(int sliderValue)
    {
        if (sliderValue <= 0) return 0;
        // Logarithmic feel: v = (100^(s/100) - 1) / 99.0
        return (Math.Pow(100, sliderValue / 100.0) - 1) / 99.0;
    }

    private static int VolumeToSlider(double volume)
    {
        if (volume <= 0) return 0;
        // inverse: s = 50 * log10(v * 99 + 1)
        return (int)Math.Round(50.0 * Math.Log10(volume * 99.0 + 1.0));
    }
    
    private void ShowTabHelp()
    {
        string title = "Help";
        string content = "";

        switch (_tabControl.SelectedIndex)
        {
            case 0: // General
                title = "General Settings Help";
                content = @"GENERAL SETTINGS
═══════════════════════════════════════

DOCTOR NAME
Your last name (e.g., ""Smith""). This is used when parsing Clario exam notes for critical findings. When the note contains text like ""Discussed with Smith and Jones"", the tool uses your name to identify that Jones is the contact person, not you.

BEEP SETTINGS
These control the audio feedback when dictation starts and stops.

• Start Beep Enabled: Play a tone when dictation begins.
• Start Volume: How loud the start beep is (0-100%).
• Stop Beep Enabled: Play a tone when dictation ends.
• Stop Volume: How loud the stop beep is (0-100%).
• Start Beep Pause: Delay (in milliseconds) before playing the start beep. This helps avoid the beep being recorded at the beginning of your dictation. Recommended: 800-1200ms.

IV REPORT HOTKEY
The keyboard shortcut used to open a report in InteleViewer. This is used by the ""Get Prior"" action. Default is ""v"" (the standard InteleViewer shortcut).

FLOATING TOOLBAR
When enabled, shows a small toolbar with customizable buttons that can send keystrokes to InteleViewer (for window/level presets, zoom, etc.). Configure the buttons in the ""IV Buttons"" tab.

RECORDING INDICATOR
When enabled, shows a small colored dot on screen that indicates whether dictation is currently active (recording) or not.

AUTO-STOP DICTATION ON PROCESS
When enabled, automatically stops dictation when you trigger ""Process Report"". This prevents accidentally continuing to dictate while reviewing the processed report.

PUSH-TO-TALK (DEAD MAN'S SWITCH)
When enabled, you must hold down the Record Button on your PowerMic to dictate. Releasing the button stops recording. This is useful if you prefer push-to-talk style dictation rather than toggle on/off.

AUTO-UPDATE
When enabled, the app automatically checks for updates on startup. If a newer version is available, it downloads in the background and prompts you to restart. Click ""Check for Updates"" to manually check at any time.";
                break;

            case 1: // Keys
                title = "Keys Help";
                content = @"KEYS
═══════════════════════════════════════

This tab lets you assign keyboard shortcuts (hotkeys) and PowerMic buttons to various actions.

AVAILABLE ACTIONS

• System Beep: Plays an audio tone to indicate dictation state change. Useful if you want a separate button just for audio feedback.

• Get Prior: Extracts the comparison study information from InteleViewer and pastes it into Mosaic. Position your cursor in InteleViewer on the prior study, then trigger this action.

• Critical Findings: Scrapes the Clario worklist for exam notes and extracts critical findings information (contact name, time, etc.) and pastes it into Mosaic using your template. Hold the Windows key while triggering to enter debug mode (shows raw data without pasting).

• Show Report: Copies the current report from Mosaic and displays it in a popup window. Useful for reviewing while looking at images.

• Capture Series/Image: Uses OCR to read the series and image numbers from the yellow selection box in InteleViewer and pastes them into Mosaic using your template.

• Start/Stop Recording: Toggles dictation on/off in Mosaic.

• Process Report: Sends Alt+P to Mosaic to process the report with RadPair.

• Sign Report: Sends Alt+F to Mosaic to sign/finalize the report.

• Create Impression: Clicks the Create Impression button in Mosaic. Useful when Skip Forward is mapped to another action.

• Discard Study: Closes the current study without signing. Useful for studies you've already read or don't need to report.

• Show Pick Lists: Opens the pick list popup for quick text insertion. Configure pick lists in the Pick Lists tab.

• Cycle Window/Level: Cycles through configured window/level presets in InteleViewer. Configure presets in the IV Buttons tab.

• Create Critical Note: Creates a Critical Communication Note in Clario for the current study. Works for both stroke and non-stroke cases. Can also be triggered by Ctrl+Click on the Clinical History window or via the right-click context menu.

HOTKEY COLUMN
Click on a hotkey field and press your desired key combination. Use Backspace or Delete to clear. Some hotkeys (like Alt+P, Alt+R) are restricted because they conflict with Mosaic's built-in shortcuts.

MIC BUTTON COLUMN
Select which PowerMic II button triggers this action. Note: Some buttons have hardcoded Mosaic functions (Skip Back = Process, Checkmark = Sign) - the tool works alongside these.";
                break;

            case 2: // IV Buttons
                title = "IV Buttons Help";
                content = @"IV BUTTONS
═══════════════════════════════════════

Create custom buttons for the floating toolbar. These buttons send keystrokes to InteleViewer for quick access to window/level presets, zoom controls, and other shortcuts.

COLUMNS
Set how many columns of buttons to display (1-3). More columns = wider toolbar.

LIVE PREVIEW
Shows how your toolbar will look. Click any button to select it for editing.

BUTTON LIST
Quick reference showing all buttons. Click to select.

BUTTON CONTROLS
• +Add: Add a new button (maximum 9 buttons).
• -Del: Delete the selected button.
• ▲/▼: Move the selected button up or down in the order.
• Reset: Restore the default button configuration.

BUTTON EDITOR

• Type:
  - Square: A small square button, good for icons.
  - Wide: A full-width button that spans all columns, good for text labels.

• Icon: Click the icon button to open a picker with common symbols. Click X to clear the icon.

• Label: Text label for the button. For square buttons, only the first few characters show. For wide buttons, the full label is displayed.

• Key: The keystroke to send when clicked. Click the field and press your desired key combination (e.g., Ctrl+V for vertical flip in InteleViewer).

TIPS FOR INTELEVIEWER
Configure custom shortcuts in InteleViewer under Utilities > User Preferences > Keyboard Shortcuts. Then assign those same shortcuts to your toolbar buttons here.

Common InteleViewer shortcuts:
• Ctrl+V: Vertical flip
• Ctrl+H: Horizontal flip
• , and . : Rotate
• - : Zoom out";
                break;

            case 3: // Text Automation
                title = "Text Automation Help";
                content = @"TEXT AUTOMATION
═══════════════════════════════════════

Customize the text that gets inserted into your reports.

CRITICAL FINDINGS TEMPLATE
═══════════════════════════
This template is used when you trigger the ""Critical Findings"" action. The tool scrapes Clario for exam note information and fills in the placeholders.

Available placeholders:
• {name} - The contact person's name (Dr./Nurse who was notified)
• {time} - The time of communication (e.g., ""2:30 PM EST"")
• {date} - The date (e.g., ""01/15/2026"")

Example template:
""Critical findings were discussed with and acknowledged by {name} at {time} on {date}.""

SERIES/IMAGE TEMPLATE
═════════════════════
This template is used when you trigger ""Capture Series/Image"". The tool uses OCR to read the series and image numbers from your screen.

Available placeholders:
• {series} - The series number
• {image} - The image number

Example: ""(series {series}, image {image})"" → ""(series 3, image 142)""

MACROS
══════
Macros are text snippets automatically inserted when you open a new study. This saves time for common report sections.

• Enable Macros: Turn macros on/off globally. Requires ""Scrape Mosaic"" to be enabled.
• Add blank lines: Insert 10 blank lines before macro text for dictation space.
• Click Edit to manage your macros. Each macro can match based on required terms and optional ""any of"" terms in the study description.

PICK LISTS
══════════
Pick lists provide quick text insertion via popup menus. Assign ""Show Pick Lists"" to a hotkey or mic button to trigger.

• Enable Pick Lists: Turn pick lists on/off.
• Skip single match: Auto-select when only one pick list matches the study.
• Keep window open: Leave the popup visible after inserting (use number keys for quick selection).

TIPS
• Hold Win key while triggering Critical Findings for debug mode.
• Macros can reference pick lists using syntax like {picklist:Name}.";
                break;

            case 4: // Alerts
                title = "Alerts Help";
                content = @"ALERTS
═══════════════════════════════════════

This tab configures the notification box and visual alerts that appear while reading studies.

CRITICAL STUDIES TRACKER
═══════════════════════════
Track critical studies this session - shows a count badge on the main widget bar when you create critical communication notes. Click the badge to see a list of all critical studies and double-click to open them in Clario.

NOTIFICATION BOX
════════════════
A floating window that shows clinical history and alerts. Requires ""Scrape Mosaic"" to be enabled in the Behavior tab.

• Enable Notification Box: Turn the floating window on/off.
• Show clinical history: Display the clinical history text. When unchecked, the window only appears when alerts are triggered (alerts-only mode).
• Hide when no study open: Hide the window when no study is active in Mosaic.
• Auto-paste corrected history: Automatically paste cleaned/corrected clinical history when malformed text is detected.
• Show Drafted indicator: Green border when report has IMPRESSION section (draft in progress).

ALERT TRIGGERS
══════════════
These alerts appear as colored borders around the notification box, even in alerts-only mode:

• Template mismatch (red border): The report template doesn't match the study description. For example, ordered ""CT Chest Abdomen Pelvis"" but template is ""CT Abdomen Pelvis"".

• Gender check (flashing red): Report contains gender-specific terms (uterus, prostate, etc.) that don't match the patient's documented gender.

• Stroke detection (purple border): Study appears to be stroke-related based on description or clinical history keywords.

STROKE ALERT OPTIONS
────────────────────
• Also use clinical history keywords: Check clinical history text for stroke-related terms in addition to the study description.
• Click alert to create Clario note: Click the purple border to automatically create a Critical Communication Note in Clario.
• Auto-create on Process Report: Automatically create the Clario note when you press Process Report for stroke cases.";
                break;

            case 5: // Behavior
                title = "Behavior Settings Help";
                content = @"BEHAVIOR SETTINGS
═══════════════════════════════════════

These settings control background features and behaviors.

RESTORE FOCUS AFTER ACTION
When enabled, after an action completes (like pasting text into Mosaic), focus returns to the window you were previously using. This is helpful if you trigger actions from InteleViewer and want to continue working there.

SCRAPE MOSAIC
═════════════
Enables background monitoring of the Mosaic report editor. This powers several sub-features below. The tool reads the report content every few seconds.

ACCESSION TRACKING
When you switch to a new study, a toast notification shows the new accession number. This also resets all tracking state (auto-fix, template mismatch, etc.) so each study starts fresh.

  SHOW CLINICAL HISTORY
  ─────────────────────
  Displays a floating window showing the clinical history from the current report. The history is automatically cleaned (removes duplicates, junk text, fixes formatting).

  • Drag the ⋯ handle to reposition the window.
  • Left-click to paste the cleaned history into Mosaic's transcript box.
  • Right-click to copy debug info to clipboard.
  • The window auto-updates when you switch studies.

  TEXT COLOR INDICATOR
  • Yellow text: The displayed history was cleaned/fixed and differs from what's in the final report. This means the fix hasn't been processed into the report yet.
  • White text: The displayed history matches the final report (fix was processed, or no fix was needed).

    AUTO-FIX
    When enabled, automatically pastes the cleaned history into Mosaic when malformed text is detected. This saves you from manually clicking to paste.

    INDICATE DRAFTED STATUS
    Shows a GREEN border around the clinical history window when the current study has a drafted report. This helps you quickly see if you've already started a report.

    WARN ON TEMPLATE MISMATCH
    Shows a RED border around the clinical history window when the study description doesn't match the report template. For example, if the order says ""CT Chest Abdomen Pelvis"" but the template is ""CT Abdomen Pelvis"", you'll see a red warning.

    This helps catch cases where Mosaic selected the wrong default template.

  SHOW IMPRESSION
  ───────────────
  Displays a floating window showing the Impression section after you process a report. This lets you review your impression while looking at images.

  • Click the impression window to dismiss it.
  • It auto-hides when you sign the report.

DEBUGGING THE FLOATING WINDOWS
══════════════════════════════
Right-click on either the Clinical History or Impression floating window to copy debug information to your clipboard. A small ""Debug copied!"" message appears.

For Clinical History, the debug info includes:
• The study description (from the order)
• The template name (from the report)
• Whether a mismatch was detected

This is useful for reporting issues or understanding why a template mismatch warning appeared.

SCROLL TO BOTTOM ON PROCESS
═══════════════════════════
When enabled, after processing a report, the tool sends Page Down keys to scroll down in the report editor. This helps you see the Impression section without manual scrolling.

  SHOW LINE COUNT TOAST
  Shows a notification with the report line count and how many Page Down keys were sent.

  SMART SCROLL THRESHOLDS
  Configure how many Page Down keys are sent based on report length:
  • Short reports (< threshold 1): No scrolling
  • Medium reports: 1 Page Down
  • Long reports: 2 Page Downs
  • Very long reports (> threshold 3): 3 Page Downs

  Adjust these thresholds based on your typical report lengths and screen size.";
                break;

            case 6: // Reference
                title = "Reference Help";
                content = @"REFERENCE
═══════════════════════════════════════

AHK / EXTERNAL INTEGRATION
This shows how to trigger Mosaic Tools actions from external programs.

WINDOWS MESSAGE CODES
• 0x8001 - Critical Findings (hold Win key for debug mode)
• 0x8003 - System Beep
• 0x8004 - Show Report
• 0x8005 - Capture Series/Image
• 0x8006 - Get Prior
• 0x8007 - Toggle Recording
• 0x8008 - Process Report
• 0x8009 - Sign Report
• 0x800A - Open Settings
• 0x800B - Create Impression
• 0x800C - Discard Study
• 0x800D - Check for Updates
• 0x800E - Show Pick Lists
• 0x800F - Create Critical Note

AUTOHOTKEY EXAMPLE
DetectHiddenWindows, On
PostMessage, 0x8001, 0, 0,, ahk_class WindowsForms

CREATE CRITICAL NOTE
Creates a Critical Communication Note in Clario for the current study.
Works for both stroke and non-stroke cases.
Triggers:
• Map to hotkey or mic button in Actions tab
• Ctrl+Click on Clinical History window
• Right-click Clinical History window → Create Critical Note
• Windows message 0x800F from AHK

DEBUG TIPS
• Hold Win key while triggering Critical Findings to see raw data.
• Right-click Clinical History or Impression windows for context menu.

SETTINGS FILE
%LOCALAPPDATA%\MosaicTools\MosaicToolsSettings.json";
                break;

            default:
                content = "No help available for this tab.";
                break;
        }

        // Show help in a scrollable window with formatted text
        var helpForm = new Form
        {
            Text = title,
            Size = new Size(550, 500),
            StartPosition = FormStartPosition.CenterParent,
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.White,
            FormBorderStyle = FormBorderStyle.Sizable,
            MinimumSize = new Size(400, 300)
        };

        var richTextBox = new RichTextBox
        {
            ReadOnly = true,
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(40, 40, 40),
            ForeColor = Color.FromArgb(200, 200, 200),
            Font = new Font("Segoe UI", 10),
            BorderStyle = BorderStyle.None,
            DetectUrls = false
        };

        // Format the content with highlighted headers
        FormatHelpText(richTextBox, content);

        // Deselect text
        richTextBox.SelectionStart = 0;
        richTextBox.SelectionLength = 0;

        var closeBtn = new Button
        {
            Text = "Close",
            Size = new Size(80, 30),
            Dock = DockStyle.Bottom,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        closeBtn.Click += (_, _) => helpForm.Close();

        helpForm.Controls.Add(richTextBox);
        helpForm.Controls.Add(closeBtn);
        helpForm.ShowDialog(this);
    }

    /// <summary>
    /// Format help text with highlighted section headers.
    /// </summary>
    private static void FormatHelpText(RichTextBox rtb, string content)
    {
        var headerColor = Color.FromArgb(100, 180, 255); // Bright blue for main headers
        var subHeaderColor = Color.FromArgb(255, 200, 100); // Orange/gold for sub-headers
        var normalColor = Color.FromArgb(200, 200, 200); // Normal text

        var lines = content.Split('\n');

        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();

            // Main headers: ALL CAPS lines (at least 3 caps words, or contains ═)
            if (trimmed.Length > 0 && (trimmed.Contains('═') || IsAllCapsHeader(trimmed)))
            {
                rtb.SelectionColor = headerColor;
                rtb.SelectionFont = new Font(rtb.Font, FontStyle.Bold);
                rtb.AppendText(line + "\n");
            }
            // Sub-headers: Lines with ─ or indented ALL CAPS
            else if (trimmed.Contains('─') || (line.StartsWith("  ") && IsAllCapsHeader(trimmed)))
            {
                rtb.SelectionColor = subHeaderColor;
                rtb.SelectionFont = new Font(rtb.Font, FontStyle.Bold);
                rtb.AppendText(line + "\n");
            }
            else
            {
                rtb.SelectionColor = normalColor;
                rtb.SelectionFont = new Font(rtb.Font, FontStyle.Regular);
                rtb.AppendText(line + "\n");
            }
        }
    }

    /// <summary>
    /// Check if a line is an ALL CAPS header (most characters are uppercase letters).
    /// </summary>
    private static bool IsAllCapsHeader(string line)
    {
        if (string.IsNullOrWhiteSpace(line)) return false;

        int upperCount = 0;
        int letterCount = 0;

        foreach (char c in line)
        {
            if (char.IsLetter(c))
            {
                letterCount++;
                if (char.IsUpper(c)) upperCount++;
            }
        }

        // At least 3 letters and 90%+ uppercase
        return letterCount >= 3 && upperCount >= letterCount * 0.9;
    }

    private void SaveAndClose()
    {
        // General settings
        _config.DoctorName = _doctorNameBox.Text.Trim();
        _config.ShowTooltips = _showTooltipsCheck.Checked;
        _config.StartBeepEnabled = _startBeepCheck.Checked;
        _config.StopBeepEnabled = _stopBeepCheck.Checked;
        _config.StartBeepVolume = SliderToVolume(_startVolumeSlider.Value);
        _config.StopBeepVolume = SliderToVolume(_stopVolumeSlider.Value);
        _config.DictationPauseMs = (int)_dictationPauseNum.Value;
        _config.IvReportHotkey = _ivHotkeyBox.Text.Trim();
        _config.FloatingToolbarEnabled = _floatingToolbarCheck.Checked;
        _config.IndicatorEnabled = _indicatorCheck.Checked;
        _config.AutoStopDictation = _autoStopCheck.Checked;
        _config.DeadManSwitch = _deadManCheck.Checked;
        _config.AutoUpdateEnabled = App.IsHeadless ? false : _autoUpdateCheck.Checked;
        _config.HideIndicatorWhenNoStudy = _hideIndicatorWhenNoStudyCheck.Checked;
        _config.CriticalFindingsTemplate = _criticalTemplateBox.Text.Trim();
        _config.SeriesImageTemplate = _seriesTemplateBox.Text.Trim();
        _config.ComparisonTemplate = _comparisonTemplateBox.Text.Trim();
        _config.SeparatePastedItems = _separatePastedItemsCheck.Checked;
        _config.ReportPopupFontFamily = _reportFontFamilyCombo.SelectedItem?.ToString() ?? "Consolas";
        _config.ReportPopupFontSize = (float)_reportFontSizeNumeric.Value;

        // Advanced tab
        _config.ScrollToBottomOnProcess = _scrollToBottomCheck.Checked;
        _config.ShowLineCountToast = _showLineCountToastCheck.Checked;
        _config.ScrapeIntervalSeconds = (int)_scrapeIntervalUpDown.Value;
        _config.ShowClinicalHistory = _showClinicalHistoryCheck.Checked;
        _config.AlwaysShowClinicalHistory = _alwaysShowClinicalHistoryCheck.Checked;
        _config.HideClinicalHistoryWhenNoStudy = _hideClinicalHistoryWhenNoStudyCheck.Checked;
        _config.AutoFixClinicalHistory = _autoFixClinicalHistoryCheck.Checked;
        _config.ShowDraftedIndicator = _showDraftedIndicatorCheck.Checked;
        _config.ShowTemplateMismatch = _showTemplateMismatchCheck.Checked;
        _config.GenderCheckEnabled = _genderCheckEnabledCheck.Checked;
        _config.StrokeDetectionEnabled = _strokeDetectionEnabledCheck.Checked;
        _config.StrokeDetectionUseClinicalHistory = _strokeDetectionUseClinicalHistoryCheck.Checked;
        _config.StrokeClickToCreateNote = _strokeClickToCreateNoteCheck.Checked;
        _config.StrokeAutoCreateNote = _strokeAutoCreateNoteCheck.Checked;
        _config.TrackCriticalStudies = _trackCriticalStudiesCheck.Checked;
        _config.ShowImpression = _showImpressionCheck.Checked;
        _config.ShowReportAfterProcess = _showReportAfterProcessCheck.Checked;
        _config.ScrollThreshold1 = (int)_scrollThreshold1.Value;
        _config.ScrollThreshold2 = (int)_scrollThreshold2.Value;
        _config.ScrollThreshold3 = (int)_scrollThreshold3.Value;

        // Macros (macros list is saved in MacroEditorForm)
        _config.MacrosEnabled = _macrosEnabledCheck.Checked;
        _config.MacrosBlankLinesBefore = _macrosBlankLinesCheck.Checked;

        // InteleViewer (not in headless mode)
        if (_windowLevelKeysBox != null)
            _config.WindowLevelKeys = _windowLevelKeysBox.Text
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();

        // Network monitor
        _config.ConnectivityMonitorEnabled = _connectivityMonitorEnabledCheck.Checked;
        _config.ConnectivityCheckIntervalSeconds = (int)_connectivityIntervalUpDown.Value;
        _config.ConnectivityTimeoutMs = (int)_connectivityTimeoutUpDown.Value * 1000;

        // Ignore Inpatient Drafted
        _config.IgnoreInpatientDrafted = _ignoreInpatientDraftedCheck.Checked;
        _config.IgnoreInpatientDraftedMode = _ignoreInpatientChestOnlyRadio.Checked ? 1 : 0;

        // RVUCounter and Report Changes
        _config.RvuCounterEnabled = _rvuCounterEnabledCheck.Checked;
        _config.RvuDisplayMode = (RvuDisplayMode)_rvuDisplayModeCombo.SelectedIndex;
        // Save metric flags from checkboxes
        var metrics = RvuMetric.None;
        if (_rvuMetricTotalCheck.Checked) metrics |= RvuMetric.Total;
        if (_rvuMetricPerHourCheck.Checked) metrics |= RvuMetric.PerHour;
        if (_rvuMetricCurrentHourCheck.Checked) metrics |= RvuMetric.CurrentHour;
        if (_rvuMetricPriorHourCheck.Checked) metrics |= RvuMetric.PriorHour;
        if (_rvuMetricEstTotalCheck.Checked) metrics |= RvuMetric.EstimatedTotal;
        if (metrics == RvuMetric.None) metrics = RvuMetric.Total; // Default to Total if nothing checked
        _config.RvuMetrics = metrics;
        _config.RvuOverflowLayout = (RvuOverflowLayout)Math.Max(0, _rvuOverflowLayoutCombo.SelectedIndex);
        _config.RvuGoalEnabled = _rvuGoalEnabledCheck.Checked;
        _config.RvuGoalPerHour = (double)_rvuGoalValueBox.Value;
        _config.RvuCounterPath = _rvuCounterPathBox.Text;
        _config.ShowReportChanges = _showReportChangesCheck.Checked;
        _config.CorrelationEnabled = _correlationEnabledCheck.Checked;
        _config.ReportChangesColor = ColorTranslator.ToHtml(_reportChangesColorPanel.BackColor);
        _config.ReportChangesAlpha = _reportChangesAlphaSlider.Value;
        _config.ReportPopupTransparent = _reportTransparentCheck.Checked;
        _config.ReportPopupTransparency = _reportTransparencySlider.Value;
        _config.PickListsEnabled = _pickListsEnabledCheck.Checked;
        _config.PickListSkipSingleMatch = _pickListSkipSingleMatchCheck.Checked;
        _config.PickListKeepOpen = _pickListKeepOpenCheck.Checked;

        // Position
        _config.SettingsX = this.Location.X;
        _config.SettingsY = this.Location.Y;
        
        // Control Map
        var controlMapTab = _tabControl.TabPages[1];
        foreach (Control ctrl in controlMapTab.Controls)
        {
            if (ctrl is TextBox tb && tb.Name.StartsWith("hotkey_"))
            {
                var action = (string)tb.Tag!;
                if (!_config.ActionMappings.ContainsKey(action))
                    _config.ActionMappings[action] = new ActionMapping();
                _config.ActionMappings[action].Hotkey = tb.Text.Trim();
            }
            else if (ctrl is ComboBox cb && cb.Name.StartsWith("mic_"))
            {
                var action = (string)cb.Tag!;
                if (!_config.ActionMappings.ContainsKey(action))
                    _config.ActionMappings[action] = new ActionMapping();
                _config.ActionMappings[action].MicButton = cb.SelectedItem?.ToString() ?? "";
            }
        }
        
        // Button Studio - save full config from working copy
        _config.FloatingButtons.Columns = _studioColumns;
        _config.FloatingButtons.Buttons = _studioButtons
            .Select(b => new FloatingButtonDef { Type = b.Type, Icon = b.Icon, Label = b.Label, Keystroke = b.Keystroke, Action = b.Action })
            .ToList();
        
        // Save
        _config.Save();
        
        // Refresh services immediately
        _controller.RefreshServices();
        
        // Apply changes
        _mainForm.ToggleFloatingToolbar(_config.FloatingToolbarEnabled);
        _mainForm.UpdateIndicatorVisibility();  // Respects "hide when no study" setting
        _mainForm.UpdateClinicalHistoryVisibility();  // Respects "hide when no study" setting
        _mainForm.RefreshFloatingToolbar(_config.FloatingButtons);
        _mainForm.RefreshRvuLayout();  // Apply RVU display mode / enable changes
        _mainForm.RefreshConnectivityService();  // Apply network monitor settings

        Close();
    }
}

/// <summary>
/// Custom TabControl with dark theme support.
/// </summary>
public class DarkTabControl : TabControl
{
    public DarkTabControl()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.DoubleBuffer, true);
        DrawMode = TabDrawMode.OwnerDrawFixed;
        SizeMode = TabSizeMode.Fixed;
        ItemSize = new Size(62, 26);
        Padding = new Point(4, 3);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        // Fill entire background
        using var bgBrush = new SolidBrush(Color.FromArgb(30, 30, 30));
        e.Graphics.FillRectangle(bgBrush, ClientRectangle);

        // Draw each tab
        for (int i = 0; i < TabCount; i++)
        {
            DrawTab(e.Graphics, i);
        }

        // Draw selected tab page border
        if (SelectedTab != null)
        {
            var pageRect = SelectedTab.Bounds;
            pageRect.Inflate(1, 1);
        }
    }

    private void DrawTab(Graphics g, int index)
    {
        var tabRect = GetTabRect(index);
        var isSelected = SelectedIndex == index;

        // Tab background
        using var tabBrush = new SolidBrush(isSelected
            ? Color.FromArgb(50, 50, 50)
            : Color.FromArgb(35, 35, 35));
        g.FillRectangle(tabBrush, tabRect);

        // Tab text
        using var textBrush = new SolidBrush(isSelected ? Color.White : Color.FromArgb(160, 160, 160));
        using var textFormat = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center
        };
        g.DrawString(TabPages[index].Text, Font, textBrush, tabRect, textFormat);

        // Selected indicator line
        if (isSelected)
        {
            using var linePen = new Pen(Color.FromArgb(80, 140, 200), 2);
            g.DrawLine(linePen, tabRect.Left + 3, tabRect.Bottom - 1, tabRect.Right - 3, tabRect.Bottom - 1);
        }
    }

    protected override void OnDrawItem(DrawItemEventArgs e)
    {
        // Handled in OnPaint
    }
}
