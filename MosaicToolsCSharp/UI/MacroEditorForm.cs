using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Editor form for managing macros (study-triggered text insertion).
/// </summary>
public class MacroEditorForm : Form
{
    private readonly Configuration _config;
    private List<MacroConfig> _macros;
    private MacroConfig? _selectedMacro;
    private bool _suppressEvents;

    // Left panel
    private ListBox _listBox = null!;
    private TextBox _nameBox = null!;
    private CheckBox _enabledCheck = null!;

    // Right panel - criteria
    private TextBox _criteriaRequiredBox = null!;
    private TextBox _criteriaAnyOfBox = null!;
    private TextBox _criteriaExcludeBox = null!;
    private TextBox _textBox = null!;

    public MacroEditorForm(Configuration config)
    {
        _config = config;
        _macros = config.Macros.Select(CloneMacro).ToList();
        InitializeUI();
        RefreshListBox();
    }

    private MacroConfig CloneMacro(MacroConfig m)
    {
        return new MacroConfig
        {
            Enabled = m.Enabled,
            Name = m.Name,
            CriteriaRequired = m.CriteriaRequired,
            CriteriaAnyOf = m.CriteriaAnyOf,
            CriteriaExclude = m.CriteriaExclude,
            Text = m.Text
        };
    }

    private void InitializeUI()
    {
        Text = "Macro Editor";
        Size = new Size(750, 550);
        MinimumSize = new Size(650, 450);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Color.FromArgb(30, 30, 30);
        ForeColor = Color.White;
        Padding = new Padding(15);

        // Bottom buttons
        var saveBtn = new Button
        {
            Text = "Save && Close",
            Size = new Size(100, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            BackColor = Color.FromArgb(51, 102, 0),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        saveBtn.FlatAppearance.BorderSize = 0;
        saveBtn.Click += (s, e) => SaveAndClose();
        Controls.Add(saveBtn);

        var cancelBtn = new Button
        {
            Text = "Cancel",
            Size = new Size(80, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            BackColor = Color.FromArgb(102, 51, 0),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        cancelBtn.FlatAppearance.BorderSize = 0;
        cancelBtn.Click += (s, e) => Close();
        Controls.Add(cancelBtn);

        // Position buttons
        Resize += (s, e) =>
        {
            cancelBtn.Location = new Point(ClientSize.Width - cancelBtn.Width - 15, ClientSize.Height - cancelBtn.Height - 15);
            saveBtn.Location = new Point(cancelBtn.Left - saveBtn.Width - 10, cancelBtn.Top);
        };
        cancelBtn.Location = new Point(ClientSize.Width - cancelBtn.Width - 15, ClientSize.Height - cancelBtn.Height - 15);
        saveBtn.Location = new Point(cancelBtn.Left - saveBtn.Width - 10, cancelBtn.Top);

        CreateLeftPanel();
        CreateRightPanel();
    }

    private void CreateLeftPanel()
    {
        int x = 15, y = 15;

        // Header
        var header = CreateLabel("MACROS", x, y, true);
        Controls.Add(header);

        // Buttons
        var addBtn = CreateButton("+", x + 80, y - 3, 26);
        addBtn.Click += (s, e) => AddMacro();
        Controls.Add(addBtn);

        var removeBtn = CreateButton("-", x + 110, y - 3, 26);
        removeBtn.Click += (s, e) => RemoveMacro();
        Controls.Add(removeBtn);

        var cloneBtn = CreateButton("Clone", x + 140, y - 3, 50);
        cloneBtn.Click += (s, e) => CloneMacro();
        Controls.Add(cloneBtn);

        var moveUpBtn = CreateButton("^", x + 195, y - 3, 26);
        moveUpBtn.Click += (s, e) => MoveMacro(-1);
        Controls.Add(moveUpBtn);

        var moveDownBtn = CreateButton("v", x + 225, y - 3, 26);
        moveDownBtn.Click += (s, e) => MoveMacro(1);
        Controls.Add(moveDownBtn);
        y += 25;

        // List box
        _listBox = new ListBox
        {
            Location = new Point(x, y),
            Size = new Size(200, 350),
            BackColor = Color.FromArgb(45, 45, 45),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            DrawMode = DrawMode.OwnerDrawFixed,
            ItemHeight = 36,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
        };
        _listBox.DrawItem += ListBox_DrawItem;
        _listBox.SelectedIndexChanged += ListBox_SelectedIndexChanged;
        Controls.Add(_listBox);
    }

    private void CreateRightPanel()
    {
        int x = 235, y = 15;
        int rightWidth = 480;

        // Header
        var header = CreateLabel("MACRO PROPERTIES", x, y, true);
        Controls.Add(header);
        y += 25;

        // Enabled
        _enabledCheck = new CheckBox
        {
            Text = "Enabled",
            Location = new Point(x, y),
            AutoSize = true,
            ForeColor = Color.White
        };
        _enabledCheck.CheckedChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedMacro == null) return;
            _selectedMacro.Enabled = _enabledCheck.Checked;
            RefreshListBox();
        };
        Controls.Add(_enabledCheck);
        y += 28;

        // Name
        Controls.Add(CreateLabel("Name:", x, y + 2));
        _nameBox = CreateTextBox(x + 50, y, rightWidth - 50);
        _nameBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _nameBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedMacro == null) return;
            _selectedMacro.Name = _nameBox.Text;
            RefreshListBox();
        };
        Controls.Add(_nameBox);
        y += 30;

        // Study Criteria section
        Controls.Add(CreateLabel("STUDY CRITERIA (match study description)", x, y, true, 8));
        y += 22;

        // Required
        Controls.Add(CreateLabel("Required (all must match):", x, y + 2));
        y += 18;
        _criteriaRequiredBox = CreateTextBox(x, y, rightWidth);
        _criteriaRequiredBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _criteriaRequiredBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedMacro == null) return;
            _selectedMacro.CriteriaRequired = _criteriaRequiredBox.Text;
            RefreshListBox();
        };
        Controls.Add(_criteriaRequiredBox);
        y += 28;

        // Any of
        Controls.Add(CreateLabel("Any of (at least one must match):", x, y + 2));
        y += 18;
        _criteriaAnyOfBox = CreateTextBox(x, y, rightWidth);
        _criteriaAnyOfBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _criteriaAnyOfBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedMacro == null) return;
            _selectedMacro.CriteriaAnyOf = _criteriaAnyOfBox.Text;
            RefreshListBox();
        };
        Controls.Add(_criteriaAnyOfBox);
        y += 28;

        // Exclude
        Controls.Add(CreateLabel("Exclude (none must match):", x, y + 2));
        y += 18;
        _criteriaExcludeBox = CreateTextBox(x, y, rightWidth);
        _criteriaExcludeBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _criteriaExcludeBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedMacro == null) return;
            _selectedMacro.CriteriaExclude = _criteriaExcludeBox.Text;
        };
        Controls.Add(_criteriaExcludeBox);
        y += 30;

        // Hint
        var hintLabel = CreateLabel("Separate multiple keywords with commas. Leave all empty for global macro.", x, y);
        hintLabel.ForeColor = Color.FromArgb(100, 100, 100);
        hintLabel.Font = new Font("Segoe UI", 8, FontStyle.Italic);
        Controls.Add(hintLabel);
        y += 22;

        // Text section
        Controls.Add(CreateLabel("TEXT TO INSERT", x, y, true, 8));
        y += 20;

        _textBox = new TextBox
        {
            Location = new Point(x, y),
            Size = new Size(rightWidth, 150),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Consolas", 9),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        _textBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedMacro == null) return;
            _selectedMacro.Text = _textBox.Text;
        };
        Controls.Add(_textBox);
    }

    private Label CreateLabel(string text, int x, int y, bool bold = false, int size = 9)
    {
        return new Label
        {
            Text = text,
            Location = new Point(x, y),
            AutoSize = true,
            ForeColor = bold ? Color.FromArgb(180, 180, 180) : Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular)
        };
    }

    private TextBox CreateTextBox(int x, int y, int width)
    {
        return new TextBox
        {
            Location = new Point(x, y),
            Width = width,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
    }

    private Button CreateButton(string text, int x, int y, int width)
    {
        var btn = new Button
        {
            Text = text,
            Location = new Point(x, y),
            Size = new Size(width, 22),
            BackColor = Color.FromArgb(55, 55, 55),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 8)
        };
        btn.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
        return btn;
    }

    private void ListBox_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= _macros.Count) return;

        var macro = _macros[e.Index];
        e.DrawBackground();

        var isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        using var bgBrush = new SolidBrush(isSelected ? Color.FromArgb(50, 80, 110) : Color.FromArgb(45, 45, 45));
        e.Graphics.FillRectangle(bgBrush, e.Bounds);

        // Checkbox indicator
        var checkRect = new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 10, 14, 14);
        using var checkBrush = new SolidBrush(macro.Enabled ? Color.FromArgb(50, 120, 50) : Color.FromArgb(70, 70, 70));
        e.Graphics.FillRectangle(checkBrush, checkRect);
        if (macro.Enabled)
        {
            using var pen = new Pen(Color.White, 2);
            e.Graphics.DrawLine(pen, checkRect.X + 3, checkRect.Y + 7, checkRect.X + 6, checkRect.Y + 10);
            e.Graphics.DrawLine(pen, checkRect.X + 6, checkRect.Y + 10, checkRect.X + 11, checkRect.Y + 4);
        }

        // Global or criteria-based indicator
        bool isGlobal = string.IsNullOrWhiteSpace(macro.CriteriaRequired) && string.IsNullOrWhiteSpace(macro.CriteriaAnyOf);
        var indicatorColor = isGlobal ? Color.FromArgb(140, 200, 140) : Color.FromArgb(100, 180, 255);

        // Name
        var name = string.IsNullOrEmpty(macro.Name) ? "(unnamed)" : macro.Name;
        using var nameFont = new Font("Segoe UI", 9, FontStyle.Bold);
        using var nameBrush = new SolidBrush(macro.Enabled ? indicatorColor : Color.Gray);
        e.Graphics.DrawString(name, nameFont, nameBrush, e.Bounds.X + 26, e.Bounds.Y + 3);

        // Criteria summary
        var criteria = macro.GetCriteriaDisplayString();
        if (criteria.Length > 28) criteria = criteria.Substring(0, 26) + "...";
        using var criteriaFont = new Font("Segoe UI", 8);
        using var criteriaBrush = new SolidBrush(Color.FromArgb(110, 110, 110));
        e.Graphics.DrawString(criteria, criteriaFont, criteriaBrush, e.Bounds.X + 26, e.Bounds.Y + 19);
    }

    private void ListBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_listBox.SelectedIndex >= 0 && _listBox.SelectedIndex < _macros.Count)
        {
            _selectedMacro = _macros[_listBox.SelectedIndex];
            LoadMacroProperties();
        }
        else
        {
            _selectedMacro = null;
            ClearMacroProperties();
        }
        UpdateButtonStates();
    }

    private void RefreshListBox()
    {
        var idx = _listBox.SelectedIndex;
        _listBox.Items.Clear();
        foreach (var macro in _macros)
            _listBox.Items.Add(macro);
        if (idx >= 0 && idx < _listBox.Items.Count)
            _listBox.SelectedIndex = idx;
        else if (_listBox.Items.Count > 0)
            _listBox.SelectedIndex = 0;
    }

    private void LoadMacroProperties()
    {
        if (_selectedMacro == null) return;
        _suppressEvents = true;
        _enabledCheck.Checked = _selectedMacro.Enabled;
        _nameBox.Text = _selectedMacro.Name;
        _criteriaRequiredBox.Text = _selectedMacro.CriteriaRequired;
        _criteriaAnyOfBox.Text = _selectedMacro.CriteriaAnyOf;
        _criteriaExcludeBox.Text = _selectedMacro.CriteriaExclude;
        _textBox.Text = _selectedMacro.Text;
        _suppressEvents = false;
    }

    private void ClearMacroProperties()
    {
        _suppressEvents = true;
        _enabledCheck.Checked = false;
        _nameBox.Text = "";
        _criteriaRequiredBox.Text = "";
        _criteriaAnyOfBox.Text = "";
        _criteriaExcludeBox.Text = "";
        _textBox.Text = "";
        _suppressEvents = false;
    }

    private void UpdateButtonStates()
    {
        var hasMacro = _selectedMacro != null;
        _enabledCheck.Enabled = hasMacro;
        _nameBox.Enabled = hasMacro;
        _criteriaRequiredBox.Enabled = hasMacro;
        _criteriaAnyOfBox.Enabled = hasMacro;
        _criteriaExcludeBox.Enabled = hasMacro;
        _textBox.Enabled = hasMacro;
    }

    private void AddMacro()
    {
        var newMacro = new MacroConfig
        {
            Enabled = true,
            Name = "New Macro",
            CriteriaRequired = "",
            CriteriaAnyOf = "",
            CriteriaExclude = "",
            Text = ""
        };
        _macros.Add(newMacro);
        RefreshListBox();
        _listBox.SelectedIndex = _macros.Count - 1;
        _nameBox.Focus();
        _nameBox.SelectAll();
    }

    private void RemoveMacro()
    {
        if (_selectedMacro == null) return;
        var idx = _macros.IndexOf(_selectedMacro);
        _macros.Remove(_selectedMacro);
        RefreshListBox();
        if (_macros.Count > 0)
            _listBox.SelectedIndex = Math.Min(idx, _macros.Count - 1);
    }

    private void MoveMacro(int dir)
    {
        if (_selectedMacro == null) return;
        var idx = _macros.IndexOf(_selectedMacro);
        var newIdx = idx + dir;
        if (newIdx < 0 || newIdx >= _macros.Count) return;

        _macros.RemoveAt(idx);
        _macros.Insert(newIdx, _selectedMacro);
        RefreshListBox();
        _listBox.SelectedIndex = newIdx;
    }

    private void CloneMacro()
    {
        if (_selectedMacro == null) return;
        var clone = CloneMacro(_selectedMacro);
        clone.Name += " (Copy)";
        _macros.Add(clone);
        RefreshListBox();
        _listBox.SelectedIndex = _macros.Count - 1;
    }

    private void SaveAndClose()
    {
        _config.Macros = _macros;
        _config.Save();
        DialogResult = DialogResult.OK;
        Close();
    }
}
