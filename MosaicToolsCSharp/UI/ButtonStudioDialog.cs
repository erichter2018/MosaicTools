using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Dialog for configuring InteleViewer floating toolbar buttons.
/// </summary>
public class ButtonStudioDialog : Form
{
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    private readonly Configuration _config;
    private List<FloatingButtonDef> _buttons = new();
    private int _columns = 2;
    private int _selectedIdx = 0;
    private bool _updatingEditor = false;

    private Panel _previewPanel = null!;
    private FlowLayoutPanel _buttonListPanel = null!;
    private NumericUpDown _columnsNum = null!;

    // Editor controls
    private RadioButton _typeSquareRadio = null!;
    private RadioButton _typeWideRadio = null!;
    private Button _iconButton = null!;
    private TextBox _labelBox = null!;
    private TextBox _keystrokeBox = null!;
    private ComboBox _actionCombo = null!;
    private Panel? _iconPickerPanel;

    private static readonly string[] IconLibrary = {
        "", "↑", "↓", "←", "→", "↕", "↔", "↺", "↻", "⟲", "⟳",
        "+", "−", "×", "÷", "⊕", "⊖", "☰", "◎", "▶", "⏸", "⏹",
        "⚙", "☀", "★", "✓", "✗", "✚", "⎚", "▦", "◐", "◑", "⇑", "⇓", "⇐", "⇒",
        "🔍", "🔎", "📋", "📌", "🔒", "🔓", "💾", "🗑", "📁", "📂",
        "◀", "▲", "▼", "◆", "○", "●", "□", "■", "△", "▷"
    };

    public ButtonStudioDialog(Configuration config)
    {
        _config = config;

        // Deep copy buttons
        _buttons = config.FloatingButtons.Buttons
            .Select(b => new FloatingButtonDef { Type = b.Type, Icon = b.Icon, Label = b.Label, Keystroke = b.Keystroke, Action = b.Action })
            .ToList();
        _columns = config.FloatingButtons.Columns;

        InitializeUI();
        RenderPreview();
        RenderButtonList();
        LoadEditorFromSelection();
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
        Text = "InteleViewer Button Studio";
        Size = new Size(520, 480);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.FromArgb(30, 30, 30);
        ForeColor = Color.White;

        // Top row - columns selector
        var topPanel = new Panel
        {
            Location = new Point(10, 10),
            Size = new Size(490, 30),
            BackColor = Color.FromArgb(30, 30, 30)
        };
        Controls.Add(topPanel);

        topPanel.Controls.Add(new Label
        {
            Text = "Columns:",
            Location = new Point(0, 5),
            AutoSize = true,
            ForeColor = Color.White
        });

        _columnsNum = new NumericUpDown
        {
            Location = new Point(65, 3),
            Width = 50,
            Minimum = 1,
            Maximum = 3,
            Value = _columns,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        _columnsNum.ValueChanged += (s, e) =>
        {
            _columns = (int)_columnsNum.Value;
            RenderPreview();
        };
        topPanel.Controls.Add(_columnsNum);

        topPanel.Controls.Add(new Label
        {
            Text = "(max 9 buttons)",
            Location = new Point(125, 5),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        });

        // Left panel - preview and button list
        var leftPanel = new Panel
        {
            Location = new Point(10, 45),
            Size = new Size(240, 330),
            BackColor = Color.FromArgb(30, 30, 30)
        };
        Controls.Add(leftPanel);

        leftPanel.Controls.Add(new Label
        {
            Text = "Live Preview",
            Location = new Point(0, 0),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        });

        _previewPanel = new Panel
        {
            Location = new Point(0, 22),
            Size = new Size(230, 160),
            BackColor = Color.FromArgb(51, 51, 51),
            BorderStyle = BorderStyle.FixedSingle,
            AutoScroll = true
        };
        leftPanel.Controls.Add(_previewPanel);

        leftPanel.Controls.Add(new Label
        {
            Text = "Button List",
            Location = new Point(0, 190),
            AutoSize = true,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        });

        _buttonListPanel = new FlowLayoutPanel
        {
            Location = new Point(0, 212),
            Size = new Size(230, 60),
            BackColor = Color.FromArgb(51, 51, 51),
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            AutoScroll = false
        };
        leftPanel.Controls.Add(_buttonListPanel);

        // Control buttons
        var ctrlPanel = new FlowLayoutPanel
        {
            Location = new Point(0, 280),
            Size = new Size(230, 35),
            FlowDirection = FlowDirection.LeftToRight,
            BackColor = Color.FromArgb(30, 30, 30)
        };
        leftPanel.Controls.Add(ctrlPanel);

        var addBtn = new Button { Text = "+Add", Width = 50, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        addBtn.Click += (s, e) => AddButton();
        ctrlPanel.Controls.Add(addBtn);

        var delBtn = new Button { Text = "-Del", Width = 45, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        delBtn.Click += (s, e) => DeleteButton();
        ctrlPanel.Controls.Add(delBtn);

        var upBtn = new Button { Text = "▲", Width = 30, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        upBtn.Click += (s, e) => MoveButtonUp();
        ctrlPanel.Controls.Add(upBtn);

        var downBtn = new Button { Text = "▼", Width = 30, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        downBtn.Click += (s, e) => MoveButtonDown();
        ctrlPanel.Controls.Add(downBtn);

        var resetBtn = new Button { Text = "Reset", Width = 50, Height = 25, FlatStyle = FlatStyle.Flat, ForeColor = Color.FromArgb(204, 0, 0) };
        resetBtn.Click += (s, e) => ResetToDefaults();
        ctrlPanel.Controls.Add(resetBtn);

        // Right panel - editor
        var editorPanel = new GroupBox
        {
            Text = "Button Editor",
            Location = new Point(260, 45),
            Size = new Size(240, 270),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        Controls.Add(editorPanel);

        int ey = 25;

        // Type
        editorPanel.Controls.Add(new Label { Text = "Type:", Location = new Point(10, ey), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) });
        _typeSquareRadio = new RadioButton { Text = "Square", Location = new Point(65, ey - 2), AutoSize = true, ForeColor = Color.White, Checked = true };
        _typeSquareRadio.CheckedChanged += (s, e) => { if (!_updatingEditor && _typeSquareRadio.Checked) ApplyEditorChanges(); };
        editorPanel.Controls.Add(_typeSquareRadio);

        _typeWideRadio = new RadioButton { Text = "Wide", Location = new Point(140, ey - 2), AutoSize = true, ForeColor = Color.White };
        _typeWideRadio.CheckedChanged += (s, e) => { if (!_updatingEditor && _typeWideRadio.Checked) ApplyEditorChanges(); };
        editorPanel.Controls.Add(_typeWideRadio);
        ey += 30;

        // Icon
        editorPanel.Controls.Add(new Label { Text = "Icon:", Location = new Point(10, ey), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) });
        _iconButton = new Button
        {
            Location = new Point(65, ey - 3),
            Size = new Size(50, 28),
            Text = "",
            Font = new Font("Segoe UI Symbol", 12),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            TextAlign = ContentAlignment.MiddleCenter
        };
        _iconButton.FlatAppearance.BorderColor = Color.Gray;
        _iconButton.Click += (s, e) => ShowIconPicker();
        editorPanel.Controls.Add(_iconButton);

        var clearIconBtn = new Button
        {
            Location = new Point(120, ey),
            Size = new Size(22, 22),
            Text = "X",
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(100, 50, 50),
            ForeColor = Color.White
        };
        clearIconBtn.FlatAppearance.BorderSize = 0;
        clearIconBtn.Click += (s, e) =>
        {
            _iconButton.Text = "";
            ApplyEditorChanges();
        };
        editorPanel.Controls.Add(clearIconBtn);
        ey += 35;

        // Label
        editorPanel.Controls.Add(new Label { Text = "Label:", Location = new Point(10, ey), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) });
        _labelBox = new TextBox
        {
            Location = new Point(65, ey - 2),
            Width = 150,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White
        };
        _labelBox.TextChanged += (s, e) => { if (!_updatingEditor) ApplyEditorChanges(); };
        editorPanel.Controls.Add(_labelBox);
        ey += 30;

        // Keystroke
        editorPanel.Controls.Add(new Label { Text = "Key:", Location = new Point(10, ey), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) });
        _keystrokeBox = new TextBox
        {
            Location = new Point(65, ey - 2),
            Width = 100,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            ReadOnly = true,
            Cursor = Cursors.Hand
        };
        SetupKeystrokeCapture(_keystrokeBox);
        editorPanel.Controls.Add(_keystrokeBox);

        var recBtn = new Button { Text = "Rec", Location = new Point(170, ey - 3), Width = 40, Height = 23, FlatStyle = FlatStyle.Flat, ForeColor = Color.White };
        recBtn.Click += (s, e) => _keystrokeBox.Focus();
        editorPanel.Controls.Add(recBtn);
        ey += 30;

        // Action
        editorPanel.Controls.Add(new Label { Text = "Action:", Location = new Point(10, ey), Width = 50, ForeColor = Color.White, Font = new Font("Segoe UI", 9) });
        _actionCombo = new ComboBox
        {
            Location = new Point(65, ey - 1),
            Width = 150,
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
        _actionCombo.SelectedIndexChanged += (s, e) => { if (!_updatingEditor) ApplyEditorChanges(); };
        editorPanel.Controls.Add(_actionCombo);
        ey += 35;

        // Hint
        editorPanel.Controls.Add(new Label
        {
            Text = "Action overrides keystroke.\nKeystroke sends to InteleViewer.",
            Location = new Point(10, ey),
            Size = new Size(210, 35),
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8)
        });

        // Bottom buttons
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

    private void SetupKeystrokeCapture(TextBox box)
    {
        box.KeyDown += (s, e) =>
        {
            e.SuppressKeyPress = true;
            e.Handled = true;

            if (e.KeyCode == Keys.Escape)
            {
                box.Text = "";
                ApplyEditorChanges();
                return;
            }

            var key = e.KeyCode.ToString();
            if (e.KeyCode != Keys.ControlKey && e.KeyCode != Keys.Menu &&
                e.KeyCode != Keys.ShiftKey && e.KeyCode != Keys.None)
            {
                box.Text = key;
                ApplyEditorChanges();
            }
        };
        box.Click += (s, e) => box.SelectAll();
    }

    private void ShowIconPicker()
    {
        if (_iconPickerPanel != null)
        {
            Controls.Remove(_iconPickerPanel);
            _iconPickerPanel.Dispose();
            _iconPickerPanel = null;
            return;
        }

        _iconPickerPanel = new Panel
        {
            Location = new Point(260, 320),
            Size = new Size(240, 100),
            BackColor = Color.FromArgb(50, 50, 50),
            BorderStyle = BorderStyle.FixedSingle,
            AutoScroll = true
        };

        var flow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            WrapContents = true
        };
        _iconPickerPanel.Controls.Add(flow);

        foreach (var icon in IconLibrary)
        {
            var btn = new Button
            {
                Text = icon,
                Size = new Size(28, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                Font = new Font("Segoe UI Symbol", 11),
                Margin = new Padding(2)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += (s, e) =>
            {
                _iconButton.Text = icon;
                ApplyEditorChanges();
                Controls.Remove(_iconPickerPanel);
                _iconPickerPanel.Dispose();
                _iconPickerPanel = null;
            };
            flow.Controls.Add(btn);
        }

        Controls.Add(_iconPickerPanel);
        _iconPickerPanel.BringToFront();
    }

    private void RenderPreview()
    {
        var oldControls = _previewPanel.Controls.Cast<Control>().ToList();
        _previewPanel.Controls.Clear();
        foreach (var ctrl in oldControls) ctrl.Dispose();

        int btnSize = 50;
        int wideWidth = btnSize * 2 + 5;
        int spacing = 5;
        int x = spacing, y = spacing;
        int colIdx = 0;

        foreach (var def in _buttons)
        {
            bool isWide = def.Type == "wide";
            int btnWidth = isWide ? wideWidth : btnSize;

            if (isWide && colIdx > 0)
            {
                x = spacing;
                y += btnSize + spacing;
                colIdx = 0;
            }

            var btn = new Button
            {
                Location = new Point(x, y),
                Size = new Size(btnWidth, btnSize),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 70),
                ForeColor = Color.White,
                Font = new Font("Segoe UI Symbol", 10)
            };
            btn.FlatAppearance.BorderColor = Color.Gray;

            string display = !string.IsNullOrEmpty(def.Icon) ? def.Icon : def.Label ?? "";
            btn.Text = display.Length > 6 ? display.Substring(0, 6) : display;

            _previewPanel.Controls.Add(btn);

            if (isWide)
            {
                x = spacing;
                y += btnSize + spacing;
                colIdx = 0;
            }
            else
            {
                x += btnWidth + spacing;
                colIdx++;
                if (colIdx >= _columns)
                {
                    x = spacing;
                    y += btnSize + spacing;
                    colIdx = 0;
                }
            }
        }
    }

    private void RenderButtonList()
    {
        var oldControls = _buttonListPanel.Controls.Cast<Control>().ToList();
        _buttonListPanel.Controls.Clear();
        foreach (var ctrl in oldControls) ctrl.Dispose();

        for (int i = 0; i < _buttons.Count; i++)
        {
            var def = _buttons[i];
            var idx = i;

            var btn = new Button
            {
                Size = new Size(24, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = i == _selectedIdx ? Color.FromArgb(80, 80, 120) : Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                Font = new Font("Segoe UI Symbol", 9),
                Text = !string.IsNullOrEmpty(def.Icon) ? def.Icon : (def.Label?.Length > 0 ? def.Label[0].ToString() : "?"),
                Margin = new Padding(2)
            };
            btn.FlatAppearance.BorderColor = i == _selectedIdx ? Color.CornflowerBlue : Color.Gray;
            btn.Click += (s, e) =>
            {
                _selectedIdx = idx;
                RenderButtonList();
                LoadEditorFromSelection();
            };
            _buttonListPanel.Controls.Add(btn);
        }
    }

    private void LoadEditorFromSelection()
    {
        if (_selectedIdx < 0 || _selectedIdx >= _buttons.Count) return;

        _updatingEditor = true;
        var def = _buttons[_selectedIdx];

        _typeSquareRadio.Checked = def.Type != "wide";
        _typeWideRadio.Checked = def.Type == "wide";
        _iconButton.Text = def.Icon ?? "";
        _labelBox.Text = def.Label ?? "";
        _keystrokeBox.Text = def.Keystroke ?? "";

        if (!string.IsNullOrEmpty(def.Action) && _actionCombo.Items.Contains(def.Action))
            _actionCombo.SelectedItem = def.Action;
        else
            _actionCombo.SelectedIndex = 0;

        _updatingEditor = false;
    }

    private void ApplyEditorChanges()
    {
        if (_selectedIdx < 0 || _selectedIdx >= _buttons.Count) return;

        var def = _buttons[_selectedIdx];
        def.Type = _typeWideRadio.Checked ? "wide" : "square";
        def.Icon = _iconButton.Text;
        def.Label = _labelBox.Text;
        def.Keystroke = _keystrokeBox.Text;
        def.Action = _actionCombo.SelectedIndex > 0 ? _actionCombo.SelectedItem?.ToString() : null;

        RenderPreview();
        RenderButtonList();
    }

    private void AddButton()
    {
        if (_buttons.Count >= 9) return;

        _buttons.Add(new FloatingButtonDef { Type = "square", Icon = "★", Label = "New" });
        _selectedIdx = _buttons.Count - 1;
        RenderPreview();
        RenderButtonList();
        LoadEditorFromSelection();
    }

    private void DeleteButton()
    {
        if (_buttons.Count <= 1 || _selectedIdx < 0) return;

        _buttons.RemoveAt(_selectedIdx);
        _selectedIdx = Math.Min(_selectedIdx, _buttons.Count - 1);
        RenderPreview();
        RenderButtonList();
        LoadEditorFromSelection();
    }

    private void MoveButtonUp()
    {
        if (_selectedIdx <= 0) return;

        (_buttons[_selectedIdx], _buttons[_selectedIdx - 1]) = (_buttons[_selectedIdx - 1], _buttons[_selectedIdx]);
        _selectedIdx--;
        RenderPreview();
        RenderButtonList();
    }

    private void MoveButtonDown()
    {
        if (_selectedIdx >= _buttons.Count - 1) return;

        (_buttons[_selectedIdx], _buttons[_selectedIdx + 1]) = (_buttons[_selectedIdx + 1], _buttons[_selectedIdx]);
        _selectedIdx++;
        RenderPreview();
        RenderButtonList();
    }

    private void ResetToDefaults()
    {
        var result = MessageBox.Show("Reset to default buttons?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (result != DialogResult.Yes) return;

        _buttons = FloatingButtonsConfig.Default.Buttons.ToList();
        _columns = 2;
        _columnsNum.Value = 2;
        _selectedIdx = 0;
        RenderPreview();
        RenderButtonList();
        LoadEditorFromSelection();
    }

    private void SaveAndClose()
    {
        _config.FloatingButtons.Buttons = _buttons;
        _config.FloatingButtons.Columns = _columns;
        _config.Save();
        DialogResult = DialogResult.OK;
        Close();
    }
}
