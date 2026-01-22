using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Settings dialog matching Python's SettingsWindow.
/// </summary>
public class SettingsForm : Form
{
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
    private CheckBox _autoUpdateCheck = null!;
    private CheckBox _hideIndicatorWhenNoStudyCheck = null!;

    // Advanced tab controls
    private CheckBox _restoreFocusCheck = null!;
    private CheckBox _scrollToBottomCheck = null!;
    private CheckBox _scrapeMosaicCheck = null!;
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
    private CheckBox _showImpressionCheck = null!;
    private CheckBox _showLineCountToastCheck = null!;
    private NumericUpDown _scrollThreshold1 = null!;
    private NumericUpDown _scrollThreshold2 = null!;
    private NumericUpDown _scrollThreshold3 = null!;

    // Experimental tab controls
    private CheckBox _rvuCounterEnabledCheck = null!;
    private TextBox _rvuCounterPathBox = null!;
    private Label _rvuCounterStatusLabel = null!;
    private CheckBox _showReportChangesCheck = null!;
    private CheckBox _pickListsEnabledCheck = null!;
    private CheckBox _pickListSkipSingleMatchCheck = null!;
    private Label _pickListsCountLabel = null!;
    private Label? _pickListsActionLabel;
    private TextBox? _pickListsActionHotkey;
    private ComboBox? _pickListsActionMic;
    private Panel _reportChangesColorPanel = null!;
    private TrackBar _reportChangesAlphaSlider = null!;
    private Label _reportChangesAlphaLabel = null!;
    private RichTextBox _reportChangesPreview = null!;

    public SettingsForm(Configuration config, ActionController controller, MainForm mainForm)
    {
        _config = config;
        _controller = controller;
        _mainForm = mainForm;
        
        InitializeUI();
        LoadSettings();
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
        
        // Tab control
        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill,
            Appearance = TabAppearance.FlatButtons
        };
        Controls.Add(_tabControl);
        
        // General tab
        var generalTab = CreateGeneralTab();
        _tabControl.TabPages.Add(generalTab);

        // Keys tab (was Control Map)
        var keysTab = CreateControlMapTab();
        _tabControl.TabPages.Add(keysTab);

        // IV Buttons tab (was Button Studio)
        var ivButtonsTab = CreateButtonStudioTab();
        _tabControl.TabPages.Add(ivButtonsTab);

        // Templates tab
        var templatesTab = CreateTemplatesTab();
        _tabControl.TabPages.Add(templatesTab);

        // Macros tab
        var macrosTab = CreateMacrosTab();
        _tabControl.TabPages.Add(macrosTab);

        // Advanced tab
        var advancedTab = CreateAdvancedTab();
        _tabControl.TabPages.Add(advancedTab);

        // Experimental tab
        var experimentalTab = CreateExperimentalTab();
        _tabControl.TabPages.Add(experimentalTab);

        // AHK tab (last)
        var ahkTab = CreateAhkTab();
        _tabControl.TabPages.Add(ahkTab);
        
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
            BackColor = Color.FromArgb(40, 40, 40)
        };
        
        int y = 20;
        int labelWidth = 180;
        
        // Doctor Name
        tab.Controls.Add(CreateLabel("Doctor Name:", 20, y, labelWidth));
        _doctorNameBox = new TextBox
        {
            Location = new Point(200, y),
            Width = 200,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        tab.Controls.Add(_doctorNameBox);
        y += 35;
        
        // Start/Stop Beep options (hide in headless mode - no audio feedback needed)
        _startBeepCheck = new CheckBox
        {
            Text = "Start Beep Enabled",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };

        var startVolLabel = CreateLabel("Start Volume:", 40, y + 25, labelWidth);
        _startVolumeSlider = new TrackBar
        {
            Location = new Point(200, y + 20),
            Width = 200,
            Minimum = 0,
            Maximum = 100,
            TickFrequency = 10
        };
        _startVolumeSlider.ValueChanged += (s, e) => UpdateVolumeLabels();

        _startVolLabel = new Label
        {
            Location = new Point(410, y + 27),
            AutoSize = true,
            ForeColor = Color.White
        };

        _stopBeepCheck = new CheckBox
        {
            Text = "Stop Beep Enabled",
            Location = new Point(20, y + 65),
            ForeColor = Color.White,
            AutoSize = true
        };

        var stopVolLabel = CreateLabel("Stop Volume:", 40, y + 90, labelWidth);
        _stopVolumeSlider = new TrackBar
        {
            Location = new Point(200, y + 85),
            Width = 200,
            Minimum = 0,
            Maximum = 100,
            TickFrequency = 10
        };
        _stopVolumeSlider.ValueChanged += (s, e) => UpdateVolumeLabels();

        _stopVolLabel = new Label
        {
            Location = new Point(410, y + 92),
            AutoSize = true,
            ForeColor = Color.White
        };

        var dictPauseLabel = CreateLabel("Start Beep Pause (ms):", 20, y + 135, labelWidth);
        _dictationPauseNum = new NumericUpDown
        {
            Location = new Point(200, y + 135),
            Width = 100,
            Minimum = 100,
            Maximum = 5000,
            Increment = 100,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };

        // Only show beep options in non-headless mode
        if (!App.IsHeadless)
        {
            tab.Controls.Add(_startBeepCheck);
            tab.Controls.Add(startVolLabel);
            tab.Controls.Add(_startVolumeSlider);
            tab.Controls.Add(_startVolLabel);
            tab.Controls.Add(_stopBeepCheck);
            tab.Controls.Add(stopVolLabel);
            tab.Controls.Add(_stopVolumeSlider);
            tab.Controls.Add(_stopVolLabel);
            tab.Controls.Add(dictPauseLabel);
            tab.Controls.Add(_dictationPauseNum);
            y += 170;
        }
        
        // IV Report Hotkey (hide in headless mode)
        var ivHotkeyLabel = CreateLabel("IV Report Hotkey:", 20, y, labelWidth);
        _ivHotkeyBox = new TextBox
        {
            Location = new Point(200, y),
            Width = 150,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            ReadOnly = true,
            Cursor = Cursors.Hand
        };
        SetupHotkeyCapture(_ivHotkeyBox);
        
        if (!App.IsHeadless)
        {
            tab.Controls.Add(ivHotkeyLabel);
            tab.Controls.Add(_ivHotkeyBox);
            y += 40;
        }
        
        // Checkboxes
        _floatingToolbarCheck = new CheckBox
        {
            Text = "Show Floating Toolbar",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_floatingToolbarCheck);
        y += 25;
        
        // Show Recording Indicator (hide in headless mode)
        _indicatorCheck = new CheckBox
        {
            Text = "Show Recording Indicator",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        if (!App.IsHeadless)
        {
            tab.Controls.Add(_indicatorCheck);
            y += 25;
        }

        // Hide indicator when no study
        _hideIndicatorWhenNoStudyCheck = new CheckBox
        {
            Text = "Hide indicator when no study open",
            Location = new Point(40, y),
            ForeColor = Color.Gray,
            AutoSize = true
        };
        _hideIndicatorWhenNoStudyCheck.CheckedChanged += (s, e) =>
        {
            _config.HideIndicatorWhenNoStudy = _hideIndicatorWhenNoStudyCheck.Checked;
            _mainForm.UpdateIndicatorVisibility();
        };
        if (!App.IsHeadless)
        {
            tab.Controls.Add(_hideIndicatorWhenNoStudyCheck);
            y += 25;
        }

        // Auto-Stop Dictation (hide in headless mode)
        _autoStopCheck = new CheckBox
        {
            Text = "Auto-Stop Dictation on Process",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        if (!App.IsHeadless)
        {
            tab.Controls.Add(_autoStopCheck);
            y += 25;
        }
        
        _deadManCheck = new CheckBox
        {
            Text = "Push-to-Talk (Dead Man's Switch)",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_deadManCheck);
        y += 35;

        // Auto-update section
        _autoUpdateCheck = new CheckBox
        {
            Text = "Auto-update",
            Location = new Point(20, y),
            ForeColor = App.IsHeadless ? Color.Gray : Color.White,
            AutoSize = true,
            Enabled = !App.IsHeadless // Disable auto-update in headless mode
        };
        tab.Controls.Add(_autoUpdateCheck);

        var checkUpdatesBtn = new Button
        {
            Text = "Check for Updates",
            Location = new Point(200, y - 3),
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
        tab.Controls.Add(checkUpdatesBtn);

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
        
        foreach (var action in Actions.All)
        {
            if (action == Actions.None) continue;

            // Check if this is the Pick Lists action (we'll create but maybe hide it)
            var isPickListAction = action == Actions.ShowPickLists;
            var hidePickList = isPickListAction && !_config.PickListsEnabled;

            var lbl = CreateLabel(action, 20, y, 150);
            lbl.Visible = !hidePickList;
            tab.Controls.Add(lbl);

            var hotkeyBox = new TextBox
            {
                Location = new Point(180, y),
                Width = 120,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                Tag = action,
                ReadOnly = true,
                Cursor = Cursors.Hand,
                Visible = !hidePickList
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
                Location = new Point(App.IsHeadless ? 180 : 310, y),
                Width = 140,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                Tag = action,
                Visible = !hidePickList
            };
            micCombo.Items.Add("");
            micCombo.Items.AddRange(HidService.AllButtons);
            var currentMic = _config.ActionMappings.GetValueOrDefault(action)?.MicButton ?? "";
            micCombo.SelectedItem = micCombo.Items.Contains(currentMic) ? currentMic : "";
            tab.Controls.Add(micCombo);

            // Store references to Pick Lists action controls for dynamic show/hide
            if (isPickListAction)
            {
                _pickListsActionLabel = lbl;
                _pickListsActionHotkey = hotkeyBox;
                _pickListsActionMic = micCombo;
            }

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
            .Select(b => new FloatingButtonDef { Type = b.Type, Icon = b.Icon, Label = b.Label, Keystroke = b.Keystroke })
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
            Size = new Size(210, 260),
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
        
        var recBtn = new Button { Text = "Rec", Location = new Point(150, ey - 3), Width = 40, Height = 23, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        recBtn.Click += (_, _) => _keystrokeBox.Focus();
        editorPanel.Controls.Add(recBtn);
        ey += 35;
        
        // Hint
        var hintLabel = new Label
        {
            Text = "Use saved shortcuts for\nInteleViewer - Utilities |\nUser Preferences",
            Location = new Point(10, ey),
            Size = new Size(180, 45),
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
        _previewPanel.Controls.Clear();
        
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
        _buttonListPanel.Controls.Clear();
        
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
            .Select(b => new FloatingButtonDef { Type = b.Type, Icon = b.Icon, Label = b.Label, Keystroke = b.Keystroke })
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
    
    private TabPage CreateAhkTab()
    {
        var tab = new TabPage("AHK")
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
            Text = @"=== AHK Integration ===

Send Windows Messages to trigger actions:

WM_TRIGGER_SCRAPE = 0x0401        # Critical Findings (Win key = debug)
WM_TRIGGER_BEEP = 0x0403          # System Beep
WM_TRIGGER_SHOW_REPORT = 0x0404   # Show Report
WM_TRIGGER_CAPTURE_SERIES = 0x0405 # Capture Series
WM_TRIGGER_GET_PRIOR = 0x0406     # Get Prior
WM_TRIGGER_TOGGLE_RECORD = 0x0407 # Toggle Record
WM_TRIGGER_PROCESS_REPORT = 0x0408 # Process Report
WM_TRIGGER_SIGN_REPORT = 0x0409   # Sign Report
WM_TRIGGER_OPEN_SETTINGS = 0x040A # Open Settings

Example AHK:
DetectHiddenWindows, On
PostMessage, 0x0401, 0, 0,, ahk_class WindowsForms

=== Config File ===
Settings stored in: MosaicToolsSettings.json
"
        };
        tab.Controls.Add(infoBox);

        return tab;
    }

    private TabPage CreateExperimentalTab()
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
        y += 22;

        // Warning about current version
        var rvuWarningLabel = new Label
        {
            Text = "Don't use this with current version of RVUCounter, may lead to unexpected results.",
            Location = new Point(40, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(200, 100, 100),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        scrollPanel.Controls.Add(rvuWarningLabel);
        y += 25;

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

        // Toggle visibility of Pick Lists action controls in the Keys tab
        if (_pickListsActionLabel != null) _pickListsActionLabel.Visible = enabled;
        if (_pickListsActionHotkey != null) _pickListsActionHotkey.Visible = enabled;
        if (_pickListsActionMic != null) _pickListsActionMic.Visible = enabled;
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

    private TabPage CreateAdvancedTab()
    {
        var tab = new TabPage("Advanced")
        {
            BackColor = Color.FromArgb(40, 40, 40),
            AutoScroll = true
        };

        int y = 20;

        // Warning label
        var warningLabel = new Label
        {
            Text = "Change these settings if you know what you're doing.",
            Location = new Point(20, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 180, 50),
            Font = new Font("Segoe UI", 9, FontStyle.Italic)
        };
        tab.Controls.Add(warningLabel);
        y += 25;

        // Debug info blurb
        var debugBlurb = new Label
        {
            Text = "Tip: Right-click the Clinical History or Impression windows to copy debug info.\nIf extraction isn't working correctly, send this to the author via Teams for finetuning.",
            Location = new Point(20, y),
            Size = new Size(420, 35),
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(debugBlurb);
        y += 45;

        // Restore Focus checkbox
        _restoreFocusCheck = new CheckBox
        {
            Text = "Restore Focus After Action",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_restoreFocusCheck);
        y += 25;

        var restoreFocusHint = new Label
        {
            Text = "Return focus to the previous window after Mosaic actions complete.",
            Location = new Point(40, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(restoreFocusHint);
        y += 30;

        // Scrape Mosaic checkbox with interval selector on same line
        _scrapeMosaicCheck = new CheckBox
        {
            Text = "Scrape Mosaic every",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        _scrapeMosaicCheck.CheckedChanged += (s, e) =>
        {
            // If scraping is disabled, also disable macros (they depend on scraping)
            if (!_scrapeMosaicCheck.Checked && _macrosEnabledCheck != null && _macrosEnabledCheck.Checked)
            {
                _macrosEnabledCheck.Checked = false;
            }
        };
        tab.Controls.Add(_scrapeMosaicCheck);

        _scrapeIntervalUpDown = new NumericUpDown
        {
            Location = new Point(_scrapeMosaicCheck.Right + 5, y - 2),
            Width = 50,
            Minimum = 1,
            Maximum = 30,
            Value = 3,
            BackColor = Color.FromArgb(45, 45, 48),
            ForeColor = Color.White
        };
        tab.Controls.Add(_scrapeIntervalUpDown);

        var secondsLabel = new Label
        {
            Text = "seconds",
            Location = new Point(_scrapeIntervalUpDown.Right + 5, y + 2),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(secondsLabel);
        y += 25;

        var scrapeHint = new Label
        {
            Text = "Enable background scraping of Mosaic report editor.",
            Location = new Point(40, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(scrapeHint);
        y += 30;

        // Notification Box checkbox (master toggle, formerly "Show Clinical History")
        _showClinicalHistoryCheck = new CheckBox
        {
            Text = "Notification Box",
            Location = new Point(40, y), // Indented under Scrape Mosaic
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_showClinicalHistoryCheck);
        y += 25;

        var notificationHint = new Label
        {
            Text = "Floating window for clinical history and alerts (template/gender/stroke).",
            Location = new Point(60, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(notificationHint);
        y += 25;

        // Always show clinical history checkbox (child of Notification Box)
        _alwaysShowClinicalHistoryCheck = new CheckBox
        {
            Text = "Always show clinical history",
            Location = new Point(60, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        _alwaysShowClinicalHistoryCheck.CheckedChanged += (s, e) => UpdateNotificationBoxStates();
        tab.Controls.Add(_alwaysShowClinicalHistoryCheck);
        y += 25;

        var alwaysShowHint = new Label
        {
            Text = "Unchecked = alerts-only mode (box hidden until alert).",
            Location = new Point(80, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(alwaysShowHint);
        y += 25;

        // Hide when no study checkbox (child of Always show)
        _hideClinicalHistoryWhenNoStudyCheck = new CheckBox
        {
            Text = "Hide when no study open",
            Location = new Point(80, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        _hideClinicalHistoryWhenNoStudyCheck.CheckedChanged += (s, e) =>
        {
            _config.HideClinicalHistoryWhenNoStudy = _hideClinicalHistoryWhenNoStudyCheck.Checked;
            _mainForm.UpdateClinicalHistoryVisibility();
        };
        tab.Controls.Add(_hideClinicalHistoryWhenNoStudyCheck);
        y += 25;

        // Auto-paste checkbox (child of Always show)
        _autoFixClinicalHistoryCheck = new CheckBox
        {
            Text = "Auto-paste fixed history",
            Location = new Point(80, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_autoFixClinicalHistoryCheck);
        y += 25;

        var autoFixHint = new Label
        {
            Text = "Auto-pastes cleaned history when malformed.\nClick box to paste manually.",
            Location = new Point(100, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(autoFixHint);
        y += 35;

        // Show Drafted Indicator checkbox (child of Always show)
        _showDraftedIndicatorCheck = new CheckBox
        {
            Text = "Indicate Drafted Status",
            Location = new Point(80, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_showDraftedIndicatorCheck);
        y += 25;

        var draftedHint = new Label
        {
            Text = "Green border when study is drafted.",
            Location = new Point(100, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(draftedHint);
        y += 30;

        // Alerts section header
        var alertsLabel = new Label
        {
            Text = "Alerts (appear even in alerts-only mode):",
            Location = new Point(60, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(200, 200, 200),
            Font = new Font("Segoe UI", 9, FontStyle.Italic)
        };
        tab.Controls.Add(alertsLabel);
        y += 25;

        // Show Template Mismatch checkbox (alert)
        _showTemplateMismatchCheck = new CheckBox
        {
            Text = "Template mismatch",
            Location = new Point(80, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_showTemplateMismatchCheck);

        var templateColorLabel = new Label
        {
            Text = "(red border)",
            Location = new Point(_showTemplateMismatchCheck.Right + 5, y + 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(220, 100, 100),
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(templateColorLabel);
        y += 25;

        // Gender Check checkbox (alert)
        _genderCheckEnabledCheck = new CheckBox
        {
            Text = "Gender check",
            Location = new Point(80, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_genderCheckEnabledCheck);

        var genderColorLabel = new Label
        {
            Text = "(flashing red)",
            Location = new Point(_genderCheckEnabledCheck.Right + 5, y + 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 100, 100),
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(genderColorLabel);
        y += 25;

        // Stroke Detection checkbox (alert)
        _strokeDetectionEnabledCheck = new CheckBox
        {
            Text = "Stroke detection",
            Location = new Point(80, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_strokeDetectionEnabledCheck);

        var strokeColorLabel = new Label
        {
            Text = "(purple border)",
            Location = new Point(_strokeDetectionEnabledCheck.Right + 5, y + 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 130, 220),
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(strokeColorLabel);
        y += 25;

        // Also use clinical history checkbox (subset of Stroke Detection)
        _strokeDetectionUseClinicalHistoryCheck = new CheckBox
        {
            Text = "Also use clinical history keywords",
            Location = new Point(100, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_strokeDetectionUseClinicalHistoryCheck);
        y += 25;

        var strokeHistoryHint = new Label
        {
            Text = "Keywords: stroke, CVA, TIA, etc.",
            Location = new Point(120, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(strokeHistoryHint);
        y += 25;

        // Click to create critical note checkbox (subset of Stroke Detection)
        _strokeClickToCreateNoteCheck = new CheckBox
        {
            Text = "Click to create critical note",
            Location = new Point(100, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_strokeClickToCreateNoteCheck);
        y += 25;

        // Auto-create on Process Report checkbox (subset of Stroke Detection)
        _strokeAutoCreateNoteCheck = new CheckBox
        {
            Text = "Auto-create on Process Report",
            Location = new Point(100, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_strokeAutoCreateNoteCheck);
        y += 30;

        // Show Impression checkbox (subset of Scrape Mosaic)
        _showImpressionCheck = new CheckBox
        {
            Text = "Show Impression",
            Location = new Point(40, y), // Indented under Scrape Mosaic
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_showImpressionCheck);
        y += 25;

        var impressionHint = new Label
        {
            Text = "Display impression in a floating box after Process Report.",
            Location = new Point(60, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(impressionHint);
        y += 35;

        // Scroll to Bottom checkbox (Second)
        _scrollToBottomCheck = new CheckBox
        {
            Text = "Scroll to Bottom on Process",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_scrollToBottomCheck);
        y += 25;

        var scrollHint = new Label
        {
            Text = "After 'Process Report', send 3 Page Down keys to scroll down the report.",
            Location = new Point(40, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(scrollHint);
        y += 35;


        // Show Line Count Toast (Subset)
        _showLineCountToastCheck = new CheckBox
        {
            Text = "Show Line Count Toast",
            Location = new Point(40, y), // Indented
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_showLineCountToastCheck);
        y += 30;
        
        // Smart Scroll Thresholds
        var threshGroup = new GroupBox
        {
            Text = "Smart Scroll Thresholds (Lines)",
            Location = new Point(20, y),
            Size = new Size(420, 100), // Narrower to avoid horizontal scroll
            ForeColor = Color.White
        };
        tab.Controls.Add(threshGroup);

        int rowY = 50;
        int inputWidth = 55;

        // Block 1
        threshGroup.Controls.Add(CreateLabel("1 PgDn >", 10, rowY+2, 55));
        _scrollThreshold1 = new NumericUpDown
        {
            Location = new Point(68, rowY),
            Width = inputWidth,
            Minimum = 1,
            Maximum = 500,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        threshGroup.Controls.Add(_scrollThreshold1);

        // Block 2
        threshGroup.Controls.Add(CreateLabel("2 PgDn >", 140, rowY+2, 55));
        _scrollThreshold2 = new NumericUpDown
        {
            Location = new Point(198, rowY),
            Width = inputWidth,
            Minimum = 1,
            Maximum = 500,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        threshGroup.Controls.Add(_scrollThreshold2);

        // Block 3
        threshGroup.Controls.Add(CreateLabel("3 PgDn >", 270, rowY+2, 55));
        _scrollThreshold3 = new NumericUpDown
        {
            Location = new Point(328, rowY),
            Width = inputWidth,
            Minimum = 1,
            Maximum = 500,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        threshGroup.Controls.Add(_scrollThreshold3);
        
        // Event handlers for enable/disable logic
        _scrollToBottomCheck.CheckedChanged += (_, _) => UpdateThresholdStates();
        _scrapeMosaicCheck.CheckedChanged += (_, _) => UpdateThresholdStates();
        _showClinicalHistoryCheck.CheckedChanged += (_, _) => UpdateThresholdStates();
        _strokeDetectionEnabledCheck.CheckedChanged += (_, _) => UpdateNotificationBoxStates();

        // Enforce logical constraints (T1 < T2 < T3)
        // Using explicit updates to ensure propagation

        _scrollThreshold1.ValueChanged += (_, _) =>
        {
            // Pushing Up: T1 -> T2 -> T3
            if (_scrollThreshold2.Value <= _scrollThreshold1.Value)
            {
                _scrollThreshold2.Value = Math.Min(_scrollThreshold2.Maximum, _scrollThreshold1.Value + 1);
                // The change to T2 will trigger T2's handler, but we can be explicit if needed
            }
        };

        _scrollThreshold2.ValueChanged += (_, _) =>
        {
            // Pushing Down: T2 -> T1
            if (_scrollThreshold1.Value >= _scrollThreshold2.Value)
            {
                 _scrollThreshold1.Value = Math.Max(_scrollThreshold1.Minimum, _scrollThreshold2.Value - 1);
            }

            // Pushing Up: T2 -> T3
            if (_scrollThreshold3.Value <= _scrollThreshold2.Value)
            {
                _scrollThreshold3.Value = Math.Min(_scrollThreshold3.Maximum, _scrollThreshold2.Value + 1);
            }
        };

        _scrollThreshold3.ValueChanged += (_, _) =>
        {
            // Pushing Down: T3 -> T2 -> T1
            if (_scrollThreshold2.Value >= _scrollThreshold3.Value)
            {
                _scrollThreshold2.Value = Math.Max(_scrollThreshold2.Minimum, _scrollThreshold3.Value - 1);
            }
        };

        y += 115; // After thresholds group

        // Debug tip section
        var debugTipLabel = new Label
        {
            Text = "Debug Tip: Right-click on the Clinical History or Impression floating\n" +
                   "windows to copy debug information to the clipboard.",
            Location = new Point(20, y),
            Size = new Size(420, 35),
            ForeColor = Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        tab.Controls.Add(debugTipLabel);

        return tab;
    }

    private void UpdateThresholdStates()
    {
        bool thresholdsEnabled = _scrollToBottomCheck.Checked && _scrapeMosaicCheck.Checked;
        _scrollThreshold1.Enabled = thresholdsEnabled;
        _scrollThreshold2.Enabled = thresholdsEnabled;
        _scrollThreshold3.Enabled = thresholdsEnabled;
        _showLineCountToastCheck.Enabled = thresholdsEnabled;

        // Clinical history and impression depend on scrape mosaic being enabled
        _showClinicalHistoryCheck.Enabled = _scrapeMosaicCheck.Checked;
        _showImpressionCheck.Enabled = _scrapeMosaicCheck.Checked;

        // Update notification box states
        UpdateNotificationBoxStates();
    }

    private void UpdateNotificationBoxStates()
    {
        bool notificationBoxEnabled = _scrapeMosaicCheck.Checked && _showClinicalHistoryCheck.Checked;
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

    private TabPage CreateTemplatesTab()
    {
        var tab = new TabPage("Templates")
        {
            BackColor = Color.FromArgb(40, 40, 40),
            AutoScroll = true
        };
        
        int y = 20;
        tab.Controls.Add(CreateLabel("Critical Findings Template:", 20, y, 400));
        y += 25;
        
        _criticalTemplateBox = new TextBox
        {
            Location = new Point(20, y),
            Width = 410,
            Height = 75,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        tab.Controls.Add(_criticalTemplateBox);
        y += 85;
        
        var helpLabel = new Label
        {
            Text = "Placeholders:\n" +
                   "{name} - Referring Clinician (Dr. / Nurse)\n" +
                   "{time} - Timestamp and Timezone\n" +
                   "{date} - Date of finding\n\n" +
                   "Example:\n" +
                   "Critical findings discussed with {name} at {time} on {date}.",
            Location = new Point(20, y),
            Size = new Size(410, 140),
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        tab.Controls.Add(helpLabel);
        y += 140;
        
        // Separator - more visible
        var sep = new Panel
        {
            Location = new Point(20, y - 10),
            Size = new Size(410, 1),
            BackColor = Color.FromArgb(100, 100, 100)
        };
        tab.Controls.Add(sep);
        
        tab.Controls.Add(CreateLabel("Series/Image Template:", 20, y, 400));
        y += 25;
        
        _seriesTemplateBox = new TextBox
        {
            Location = new Point(20, y),
            Width = 410,
            Height = 40,
            Multiline = true,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        tab.Controls.Add(_seriesTemplateBox);
        y += 45;
        
        var seriesHelpLabel = new Label
        {
            Text = "Placeholders: {series}, {image}\n" +
                   "Example: (series {series}, image {image})",
            Location = new Point(20, y),
            Size = new Size(410, 60),
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        tab.Controls.Add(seriesHelpLabel);
        
        // Add a bottom spacer for comfortable scrolling
        var bottomSpacer = new Label { Location = new Point(20, y + 80), Size = new Size(1, 10) };
        tab.Controls.Add(bottomSpacer);

        return tab;
    }

    // Macros tab state
    private CheckBox _macrosEnabledCheck = null!;
    private CheckBox _macrosBlankLinesCheck = null!;
    private Panel _macroListPanel = null!;
    private TextBox _macroNameBox = null!;
    private TextBox _macroCriteriaRequiredBox = null!;
    private TextBox _macroCriteriaAnyOfBox = null!;
    private TextBox _macroCriteriaExcludeBox = null!;
    private TextBox _macroTextBox = null!;
    private List<MacroConfig> _macros = new();
    private int _selectedMacroIdx = -1;
    private bool _updatingMacroEditor = false;

    private TabPage CreateMacrosTab()
    {
        var tab = new TabPage("Macros")
        {
            BackColor = Color.FromArgb(40, 40, 40)
        };

        int y = 12;

        // Top row: Enable + blank lines
        _macrosEnabledCheck = new CheckBox
        {
            Text = "Enable Macros",
            Location = new Point(15, y),
            AutoSize = true,
            ForeColor = Color.White,
            Checked = _config.MacrosEnabled
        };
        _macrosEnabledCheck.CheckedChanged += (s, e) =>
        {
            UpdateMacroStates();
            // Auto-enable Scrape Mosaic when macros are enabled (required for study change detection)
            if (_macrosEnabledCheck.Checked && !_scrapeMosaicCheck.Checked)
            {
                _scrapeMosaicCheck.Checked = true;
            }
        };
        tab.Controls.Add(_macrosEnabledCheck);

        _macrosBlankLinesCheck = new CheckBox
        {
            Text = "Add blank lines before (dictation space)",
            Location = new Point(140, y),
            AutoSize = true,
            ForeColor = Color.White,
            Checked = _config.MacrosBlankLinesBefore
        };
        tab.Controls.Add(_macrosBlankLinesCheck);
        y += 28;

        // Left column: Macro list
        var listLabel = new Label
        {
            Text = "Macros",
            Location = new Point(15, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 180, 180),
            Font = new Font("Segoe UI", 8)
        };
        tab.Controls.Add(listLabel);

        var addBtn = new Button
        {
            Text = "+",
            Location = new Point(100, y - 2),
            Size = new Size(24, 20),
            BackColor = Color.FromArgb(60, 100, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 8, FontStyle.Bold)
        };
        addBtn.FlatAppearance.BorderSize = 0;
        addBtn.Click += (s, e) => AddMacro();
        tab.Controls.Add(addBtn);

        var removeBtn = new Button
        {
            Text = "-",
            Location = new Point(126, y - 2),
            Size = new Size(24, 20),
            BackColor = Color.FromArgb(100, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 8, FontStyle.Bold)
        };
        removeBtn.FlatAppearance.BorderSize = 0;
        removeBtn.Click += (s, e) => RemoveMacro();
        tab.Controls.Add(removeBtn);
        y += 20;

        _macroListPanel = new Panel
        {
            Location = new Point(15, y),
            Size = new Size(135, 230),
            BackColor = Color.FromArgb(30, 30, 30),
            AutoScroll = true,
            BorderStyle = BorderStyle.FixedSingle
        };
        tab.Controls.Add(_macroListPanel);

        // Right column: Editor (full height)
        var editorX = 165;
        var editorY = 40;

        var nameLabel = new Label { Text = "Name", Location = new Point(editorX, editorY), AutoSize = true, ForeColor = Color.FromArgb(180, 180, 180), Font = new Font("Segoe UI", 8) };
        tab.Controls.Add(nameLabel);
        editorY += 16;

        _macroNameBox = new TextBox
        {
            Location = new Point(editorX, editorY),
            Width = 275,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        _macroNameBox.TextChanged += (s, e) => ApplyMacroEditorChanges();
        tab.Controls.Add(_macroNameBox);
        editorY += 26;

        var requiredLabel = new Label { Text = "Required (all must match)", Location = new Point(editorX, editorY), AutoSize = true, ForeColor = Color.FromArgb(180, 180, 180), Font = new Font("Segoe UI", 8) };
        tab.Controls.Add(requiredLabel);
        editorY += 16;

        _macroCriteriaRequiredBox = new TextBox
        {
            Location = new Point(editorX, editorY),
            Width = 275,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        _macroCriteriaRequiredBox.TextChanged += (s, e) => ApplyMacroEditorChanges();
        tab.Controls.Add(_macroCriteriaRequiredBox);
        editorY += 26;

        var anyOfLabel = new Label { Text = "Any of (optional, at least one)", Location = new Point(editorX, editorY), AutoSize = true, ForeColor = Color.FromArgb(180, 180, 180), Font = new Font("Segoe UI", 8) };
        tab.Controls.Add(anyOfLabel);
        editorY += 16;

        _macroCriteriaAnyOfBox = new TextBox
        {
            Location = new Point(editorX, editorY),
            Width = 275,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        _macroCriteriaAnyOfBox.TextChanged += (s, e) => ApplyMacroEditorChanges();
        tab.Controls.Add(_macroCriteriaAnyOfBox);
        editorY += 26;

        var excludeLabel = new Label { Text = "Exclude (none must match)", Location = new Point(editorX, editorY), AutoSize = true, ForeColor = Color.FromArgb(180, 180, 180), Font = new Font("Segoe UI", 8) };
        tab.Controls.Add(excludeLabel);
        editorY += 16;

        _macroCriteriaExcludeBox = new TextBox
        {
            Location = new Point(editorX, editorY),
            Width = 275,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        _macroCriteriaExcludeBox.TextChanged += (s, e) => ApplyMacroEditorChanges();
        tab.Controls.Add(_macroCriteriaExcludeBox);
        editorY += 26;

        var textLabel = new Label { Text = "Macro Text", Location = new Point(editorX, editorY), AutoSize = true, ForeColor = Color.FromArgb(180, 180, 180), Font = new Font("Segoe UI", 8) };
        tab.Controls.Add(textLabel);
        editorY += 16;

        _macroTextBox = new TextBox
        {
            Location = new Point(editorX, editorY),
            Size = new Size(275, 100),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Consolas", 9)
        };
        _macroTextBox.TextChanged += (s, e) => ApplyMacroEditorChanges();
        tab.Controls.Add(_macroTextBox);
        editorY += 104;

        // Note about case insensitivity
        var noteLabel = new Label
        {
            Text = "Separate criteria with commas. Not case sensitive.",
            Location = new Point(editorX, editorY),
            AutoSize = true,
            ForeColor = Color.FromArgb(120, 120, 120),
            Font = new Font("Segoe UI", 7.5f)
        };
        tab.Controls.Add(noteLabel);

        // Load macros
        _macros = _config.Macros.Select(m => new MacroConfig
        {
            Enabled = m.Enabled,
            Name = m.Name,
            CriteriaRequired = m.CriteriaRequired,
            CriteriaAnyOf = m.CriteriaAnyOf,
            CriteriaExclude = m.CriteriaExclude,
            Text = m.Text
        }).ToList();

        RenderMacroList();
        UpdateMacroStates();

        return tab;
    }

    private void RenderMacroList()
    {
        _macroListPanel.Controls.Clear();
        int y = 1;

        for (int i = 0; i < _macros.Count; i++)
        {
            var macro = _macros[i];
            var idx = i;
            bool isSelected = (i == _selectedMacroIdx);
            bool isGlobal = string.IsNullOrWhiteSpace(macro.CriteriaRequired) && string.IsNullOrWhiteSpace(macro.CriteriaAnyOf);

            var itemPanel = new Panel
            {
                Location = new Point(1, y),
                Size = new Size(115, 32),
                BackColor = isSelected ? Color.FromArgb(50, 60, 80) : Color.FromArgb(40, 40, 40),
                Cursor = Cursors.Hand
            };

            var cb = new CheckBox
            {
                Location = new Point(2, 8),
                Size = new Size(16, 16),
                Checked = macro.Enabled,
                FlatStyle = FlatStyle.Flat
            };
            cb.CheckedChanged += (s, e) => { _macros[idx].Enabled = cb.Checked; };
            itemPanel.Controls.Add(cb);

            var nameStr = string.IsNullOrWhiteSpace(macro.Name) ? "(unnamed)" : macro.Name;
            if (nameStr.Length > 12) nameStr = nameStr.Substring(0, 10) + "..";

            var nameLabel = new Label
            {
                Text = (isGlobal ? "● " : "◆ ") + nameStr,
                Location = new Point(18, 2),
                Size = new Size(95, 15),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 7.5f, FontStyle.Bold)
            };
            nameLabel.ForeColor = isGlobal ? Color.FromArgb(140, 200, 140) : Color.White;
            itemPanel.Controls.Add(nameLabel);

            var criteriaStr = macro.GetCriteriaDisplayString();
            if (criteriaStr.Length > 14) criteriaStr = criteriaStr.Substring(0, 12) + "..";

            var criteriaLabel = new Label
            {
                Text = criteriaStr,
                Location = new Point(18, 16),
                Size = new Size(95, 14),
                ForeColor = Color.FromArgb(130, 130, 130),
                Font = new Font("Segoe UI", 7)
            };
            itemPanel.Controls.Add(criteriaLabel);

            void SelectThis(object? s, EventArgs e) => SelectMacro(idx);
            itemPanel.Click += SelectThis;
            nameLabel.Click += SelectThis;
            criteriaLabel.Click += SelectThis;

            _macroListPanel.Controls.Add(itemPanel);
            y += 33;
        }
    }

    private void SelectMacro(int idx)
    {
        _selectedMacroIdx = idx;
        RenderMacroList();
        UpdateMacroEditorFromSelection();
    }

    private void UpdateMacroEditorFromSelection()
    {
        _updatingMacroEditor = true;
        try
        {
            if (_selectedMacroIdx < 0 || _selectedMacroIdx >= _macros.Count)
            {
                _macroNameBox.Text = "";
                _macroCriteriaRequiredBox.Text = "";
                _macroCriteriaAnyOfBox.Text = "";
                _macroCriteriaExcludeBox.Text = "";
                _macroTextBox.Text = "";
                _macroNameBox.Enabled = false;
                _macroCriteriaRequiredBox.Enabled = false;
                _macroCriteriaAnyOfBox.Enabled = false;
                _macroCriteriaExcludeBox.Enabled = false;
                _macroTextBox.Enabled = false;
                return;
            }

            var macro = _macros[_selectedMacroIdx];
            _macroNameBox.Text = macro.Name;
            _macroCriteriaRequiredBox.Text = macro.CriteriaRequired;
            _macroCriteriaAnyOfBox.Text = macro.CriteriaAnyOf;
            _macroCriteriaExcludeBox.Text = macro.CriteriaExclude;
            _macroTextBox.Text = macro.Text;
            _macroNameBox.Enabled = true;
            _macroCriteriaRequiredBox.Enabled = true;
            _macroCriteriaAnyOfBox.Enabled = true;
            _macroCriteriaExcludeBox.Enabled = true;
            _macroTextBox.Enabled = true;
        }
        finally
        {
            _updatingMacroEditor = false;
        }
    }

    private void ApplyMacroEditorChanges()
    {
        if (_updatingMacroEditor) return;
        if (_selectedMacroIdx < 0 || _selectedMacroIdx >= _macros.Count) return;

        var macro = _macros[_selectedMacroIdx];
        macro.Name = _macroNameBox.Text;
        macro.CriteriaRequired = _macroCriteriaRequiredBox.Text;
        macro.CriteriaAnyOf = _macroCriteriaAnyOfBox.Text;
        macro.CriteriaExclude = _macroCriteriaExcludeBox.Text;
        macro.Text = _macroTextBox.Text;

        RenderMacroList();
    }

    private void AddMacro()
    {
        _macros.Add(new MacroConfig
        {
            Enabled = true,
            Name = "New Macro",
            CriteriaRequired = "",
            CriteriaExclude = "",
            CriteriaAnyOf = "",
            Text = ""
        });
        _selectedMacroIdx = _macros.Count - 1;
        RenderMacroList();
        UpdateMacroEditorFromSelection();
    }

    private void RemoveMacro()
    {
        if (_selectedMacroIdx < 0 || _selectedMacroIdx >= _macros.Count) return;

        _macros.RemoveAt(_selectedMacroIdx);
        if (_selectedMacroIdx >= _macros.Count)
            _selectedMacroIdx = _macros.Count - 1;

        RenderMacroList();
        UpdateMacroEditorFromSelection();
    }

    private void UpdateMacroStates()
    {
        bool enabled = _macrosEnabledCheck.Checked;
        _macrosBlankLinesCheck.Enabled = enabled;
        _macroListPanel.Enabled = enabled;
        _macroNameBox.Enabled = enabled && _selectedMacroIdx >= 0;
        _macroCriteriaRequiredBox.Enabled = enabled && _selectedMacroIdx >= 0;
        _macroCriteriaAnyOfBox.Enabled = enabled && _selectedMacroIdx >= 0;
        _macroTextBox.Enabled = enabled && _selectedMacroIdx >= 0;
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
        
        // Advanced tab
        _restoreFocusCheck.Checked = _config.RestoreFocusAfterAction;
        _scrollToBottomCheck.Checked = _config.ScrollToBottomOnProcess;
        _showLineCountToastCheck.Checked = _config.ShowLineCountToast;
        _scrapeMosaicCheck.Checked = _config.ScrapeMosaicEnabled;
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
        _showImpressionCheck.Checked = _config.ShowImpression;
        _scrollThreshold1.Value = _config.ScrollThreshold1;
        _scrollThreshold2.Value = _config.ScrollThreshold2;
        _scrollThreshold3.Value = _config.ScrollThreshold3;

        // Experimental tab
        _rvuCounterEnabledCheck.Checked = _config.RvuCounterEnabled;
        _rvuCounterPathBox.Text = _config.RvuCounterPath;
        _showReportChangesCheck.Checked = _config.ShowReportChanges;
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
        _pickListsEnabledCheck.Checked = _config.PickListsEnabled;
        _pickListSkipSingleMatchCheck.Checked = _config.PickListSkipSingleMatch;
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

            case 3: // Templates
                title = "Templates Help";
                content = @"TEMPLATES
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

Result:
""Critical findings were discussed with and acknowledged by Dr. Jones at 2:30 PM EST on 01/15/2026.""

DEBUGGING CRITICAL FINDINGS
If the critical findings aren't being extracted correctly:
1. Hold the Windows key on your keyboard
2. While holding Win, trigger Critical Findings (mic button or hotkey)
3. Debug mode activates - a window shows the raw scraped text and formatted result
4. This helps identify if the note format is different than expected

SERIES/IMAGE TEMPLATE
═════════════════════
This template is used when you trigger ""Capture Series/Image"". The tool uses OCR to read the series and image numbers from your screen.

Available placeholders:
• {series} - The series number
• {image} - The image number

Example template:
""(series {series}, image {image})""

Result:
""(series 3, image 142)""

TIPS
• Position the yellow selection box in InteleViewer so the series/image info is visible in the header area.
• The OCR looks for patterns like ""S: 3"" or ""Series: 3"" and ""I: 142"" or ""Image: 142"".";
                break;

            case 4: // Macros
                title = "Macros Help";
                content = @"MACROS
═══════════════════════════════════════

Macros are text snippets that are automatically inserted when you open a new study. This saves time for common report sections you frequently type.

ENABLE MACROS
Turn macros on or off globally.

ADD BLANK LINES BEFORE MACROS
When enabled, inserts 10 blank lines before your macro text. This gives you space to dictate above the macro content.

CREATING MACROS
Click + to add a new macro. Each macro has:

• Name: A label to identify the macro (e.g., ""CT Abdomen Template"")

• Required: Comma-separated terms that must ALL appear in the study description.
  Example: ""CT, abdomen"" matches ""CT ABDOMEN PELVIS WITH CONTRAST""

• Any of: Comma-separated terms where at least ONE must appear.
  Example: ""head, brain"" matches studies with either term.

• Macro Text: The actual text to insert.

MATCHING LOGIC
• Both criteria blank = Global macro (applies to all studies)
• Required only = All terms must match
• Any of only = At least one term must match
• Both = All Required terms AND at least one Any of term

MULTIPLE MACROS
Multiple macros can match the same study. All matching macros are inserted together. Blank lines (if enabled) are added only once.";
                break;

            case 5: // Advanced
                title = "Advanced Settings Help";
                content = @"ADVANCED SETTINGS
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

            case 6: // Experimental
                title = "Experimental Features Help";
                content = @"EXPERIMENTAL FEATURES
═══════════════════════════════════════

Features in this tab are experimental and may change or be removed in future versions.

RVUCOUNTER INTEGRATION
Reads the current shift RVU total from a local RVUCounter installation.

To set up:
1. Click 'Find RVUCounter' to locate your RVUCounter installation
2. The tool will search common locations (Desktop, same folder as MosaicTools)
3. If not found automatically, browse to select the RVUCounter folder
4. Enable 'Read RVUCounter shift total' to activate the feature

The RVUCounter database is read in read-only mode, so it won't interfere with RVUCounter's normal operation.";
                break;

            case 7: // AHK
                title = "AHK Integration Help";
                content = @"AHK (AUTOHOTKEY) INTEGRATION
═══════════════════════════════════════

This tab shows how to trigger Mosaic Tools actions from external programs like AutoHotkey scripts.

HOW IT WORKS
Mosaic Tools listens for Windows Messages. You can send these messages from any program to trigger actions without needing to set up hotkeys.

WINDOWS MESSAGE CODES
• 0x0401 - Critical Findings (hold Win key for debug mode)
• 0x0403 - System Beep
• 0x0404 - Show Report
• 0x0405 - Capture Series/Image
• 0x0406 - Get Prior
• 0x0407 - Toggle Recording
• 0x0408 - Process Report
• 0x0409 - Sign Report
• 0x040A - Open Settings

AUTOHOTKEY EXAMPLE
To trigger Critical Findings from an AHK script:

DetectHiddenWindows, On
PostMessage, 0x0401, 0, 0,, ahk_class WindowsForms

This can be useful if you want to:
• Trigger actions from other applications
• Create complex macros that combine multiple tools
• Use foot pedals or other input devices that only support AHK

SETTINGS FILE LOCATION
Your settings are saved to:
%LOCALAPPDATA%\MosaicTools\MosaicToolsSettings.json

You can back up this file or copy it to other workstations.";
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
        
        // Advanced tab
        _config.RestoreFocusAfterAction = _restoreFocusCheck.Checked;
        _config.ScrollToBottomOnProcess = _scrollToBottomCheck.Checked;
        _config.ShowLineCountToast = _showLineCountToastCheck.Checked;
        _config.ScrapeMosaicEnabled = _scrapeMosaicCheck.Checked;
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
        _config.ShowImpression = _showImpressionCheck.Checked;
        _config.ScrollThreshold1 = (int)_scrollThreshold1.Value;
        _config.ScrollThreshold2 = (int)_scrollThreshold2.Value;
        _config.ScrollThreshold3 = (int)_scrollThreshold3.Value;

        // Macros
        _config.MacrosEnabled = _macrosEnabledCheck.Checked;
        _config.MacrosBlankLinesBefore = _macrosBlankLinesCheck.Checked;
        _config.Macros = _macros.Select(m => new MacroConfig
        {
            Enabled = m.Enabled,
            Name = m.Name,
            CriteriaRequired = m.CriteriaRequired,
            CriteriaAnyOf = m.CriteriaAnyOf,
            CriteriaExclude = m.CriteriaExclude,
            Text = m.Text
        }).ToList();

        // Experimental
        _config.RvuCounterEnabled = _rvuCounterEnabledCheck.Checked;
        _config.RvuCounterPath = _rvuCounterPathBox.Text;
        _config.ShowReportChanges = _showReportChangesCheck.Checked;
        _config.ReportChangesColor = ColorTranslator.ToHtml(_reportChangesColorPanel.BackColor);
        _config.ReportChangesAlpha = _reportChangesAlphaSlider.Value;
        _config.PickListsEnabled = _pickListsEnabledCheck.Checked;
        _config.PickListSkipSingleMatch = _pickListSkipSingleMatchCheck.Checked;

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
            .Select(b => new FloatingButtonDef { Type = b.Type, Icon = b.Icon, Label = b.Label, Keystroke = b.Keystroke })
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

        Close();
    }
}
