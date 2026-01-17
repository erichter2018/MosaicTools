using System;
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
        Text = "Mosaic Tools Settings";
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
        
        // Control Map tab
        var controlMapTab = CreateControlMapTab();
        _tabControl.TabPages.Add(controlMapTab);
        
        // Button Studio tab
        var buttonStudioTab = CreateButtonStudioTab();
        _tabControl.TabPages.Add(buttonStudioTab);

        // Templates tab
        var templatesTab = CreateTemplatesTab();
        _tabControl.TabPages.Add(templatesTab);
        
        // Advanced tab
        var advancedTab = CreateAdvancedTab();
        _tabControl.TabPages.Add(advancedTab);
        
        // Save/Cancel buttons
        var buttonPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            BackColor = Color.FromArgb(40, 40, 40)
        };
        
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
        
        // Start Beep
        _startBeepCheck = new CheckBox
        {
            Text = "Start Beep Enabled",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_startBeepCheck);
        y += 25;
        
        tab.Controls.Add(CreateLabel("Start Volume:", 40, y, labelWidth));
        _startVolumeSlider = new TrackBar
        {
            Location = new Point(200, y - 5),
            Width = 200,
            Minimum = 0,
            Maximum = 100,
            TickFrequency = 10
        };
        _startVolumeSlider.ValueChanged += (s, e) => UpdateVolumeLabels();
        tab.Controls.Add(_startVolumeSlider);
        
        _startVolLabel = new Label
        {
            Location = new Point(410, y + 2),
            AutoSize = true,
            ForeColor = Color.White
        };
        tab.Controls.Add(_startVolLabel);
        y += 40;
        
        // Stop Beep
        _stopBeepCheck = new CheckBox
        {
            Text = "Stop Beep Enabled",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_stopBeepCheck);
        y += 25;
        
        tab.Controls.Add(CreateLabel("Stop Volume:", 40, y, labelWidth));
        _stopVolumeSlider = new TrackBar
        {
            Location = new Point(200, y - 5),
            Width = 200,
            Minimum = 0,
            Maximum = 100,
            TickFrequency = 10
        };
        _stopVolumeSlider.ValueChanged += (s, e) => UpdateVolumeLabels();
        tab.Controls.Add(_stopVolumeSlider);
        
        _stopVolLabel = new Label
        {
            Location = new Point(410, y + 2),
            AutoSize = true,
            ForeColor = Color.White
        };
        tab.Controls.Add(_stopVolLabel);
        y += 45;
        
        // Dictation Pause
        tab.Controls.Add(CreateLabel("Start Beep Pause (ms):", 20, y, labelWidth));
        _dictationPauseNum = new NumericUpDown
        {
            Location = new Point(200, y),
            Width = 100,
            Minimum = 100,
            Maximum = 5000,
            Increment = 100,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        tab.Controls.Add(_dictationPauseNum);
        y += 35;
        
        // IV Report Hotkey
        tab.Controls.Add(CreateLabel("IV Report Hotkey:", 20, y, labelWidth));
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
        tab.Controls.Add(_ivHotkeyBox);
        y += 40;
        
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
        
        _indicatorCheck = new CheckBox
        {
            Text = "Show Recording Indicator",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_indicatorCheck);
        y += 25;
        
        _autoStopCheck = new CheckBox
        {
            Text = "Auto-Stop Dictation on Process",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_autoStopCheck);
        y += 25;
        
        _deadManCheck = new CheckBox
        {
            Text = "Push-to-Talk (Dead Man's Switch)",
            Location = new Point(20, y),
            ForeColor = Color.White,
            AutoSize = true
        };
        tab.Controls.Add(_deadManCheck);
        
        return tab;
    }
    
    private TabPage CreateControlMapTab()
    {
        var tab = new TabPage("Control Map")
        {
            BackColor = Color.FromArgb(40, 40, 40),
            AutoScroll = true
        };
        
        int y = 20;
        
        tab.Controls.Add(CreateLabel("Action", 20, y, 150));
        tab.Controls.Add(CreateLabel("Hotkey", 180, y, 120));
        tab.Controls.Add(CreateLabel("Mic Button", 310, y, 120));
        y += 30;
        
        foreach (var action in Actions.All)
        {
            if (action == Actions.None) continue;
            
            var lbl = CreateLabel(action, 20, y, 150);
            tab.Controls.Add(lbl);
            
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
            tab.Controls.Add(hotkeyBox);
            
            var micCombo = new ComboBox
            {
                Name = $"mic_{action}",
                Location = new Point(310, y),
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
        var tab = new TabPage("Button Studio")
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
    
    private TabPage CreateAdvancedTab()
    {
        var tab = new TabPage("Advanced")
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

WM_TRIGGER_SCRAPE = 0x0401        # Critical Findings
WM_TRIGGER_DEBUG = 0x0402         # Debug Scrape
WM_TRIGGER_BEEP = 0x0403          # System Beep
WM_TRIGGER_SHOW_REPORT = 0x0404   # Show Report
WM_TRIGGER_CAPTURE_SERIES = 0x0405 # Capture Series
WM_TRIGGER_GET_PRIOR = 0x0406     # Get Prior
WM_TRIGGER_TOGGLE_RECORD = 0x0407 # Toggle Record
WM_TRIGGER_PROCESS_REPORT = 0x0408 # Process Report
WM_TRIGGER_SIGN_REPORT = 0x0409   # Sign Report

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
            Height = 150,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9)
        };
        tab.Controls.Add(_criticalTemplateBox);
        y += 160;
        
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

            tb.Text = mods.Count > 0 ? string.Join("+", mods) + "+" + keyName : keyName;
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
        _criticalTemplateBox.Text = _config.CriticalFindingsTemplate;
        _seriesTemplateBox.Text = _config.SeriesImageTemplate;
        
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
        _config.CriticalFindingsTemplate = _criticalTemplateBox.Text.Trim();
        _config.SeriesImageTemplate = _seriesTemplateBox.Text.Trim();

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
        _mainForm.ToggleIndicator(_config.IndicatorEnabled);
        _mainForm.RefreshFloatingToolbar(_config.FloatingButtons);
        
        Close();
    }
}
