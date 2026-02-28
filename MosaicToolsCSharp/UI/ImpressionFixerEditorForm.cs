using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Editor form for managing impression fixer entries (quick-insert/replace impression text).
/// </summary>
public class ImpressionFixerEditorForm : Form
{
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

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

    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private readonly Configuration _config;
    private List<ImpressionFixerEntry> _entries;
    private ImpressionFixerEntry? _selectedEntry;
    private bool _suppressEvents;

    // Left panel
    private ListBox _listBox = null!;
    private CheckBox _enabledCheck = null!;

    // Right panel
    private TextBox _blurbBox = null!;
    private RadioButton _insertRadio = null!;
    private RadioButton _replaceRadio = null!;
    private CheckBox _requireComparisonCheck = null!;
    private Label _maxCompWeeksLabel = null!;
    private NumericUpDown _maxCompWeeksBox = null!;
    private TextBox _criteriaRequiredBox = null!;
    private TextBox _criteriaAnyOfBox = null!;
    private TextBox _criteriaExcludeBox = null!;
    private TextBox _textBox = null!;

    public ImpressionFixerEditorForm(Configuration config)
    {
        _config = config;
        _entries = config.ImpressionFixers.Select(CloneEntry).ToList();
        InitializeUI();
        RefreshListBox();
    }

    private ImpressionFixerEntry CloneEntry(ImpressionFixerEntry e)
    {
        return new ImpressionFixerEntry
        {
            Id = Guid.NewGuid().ToString("N")[..8],
            Enabled = e.Enabled,
            Blurb = e.Blurb,
            Text = e.Text,
            ReplaceMode = e.ReplaceMode,
            RequireComparison = e.RequireComparison,
            MaxComparisonWeeks = e.MaxComparisonWeeks,
            CriteriaRequired = e.CriteriaRequired,
            CriteriaAnyOf = e.CriteriaAnyOf,
            CriteriaExclude = e.CriteriaExclude
        };
    }

    private void InitializeUI()
    {
        Text = "Impression Fixer Editor";
        Size = new Size(750, 680);
        MinimumSize = new Size(650, 580);
        StartPosition = FormStartPosition.CenterScreen;
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

        // Backup/Restore buttons (bottom left)
        var backupBtn = new Button
        {
            Text = "Backup...",
            Size = new Size(80, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        backupBtn.FlatAppearance.BorderSize = 0;
        backupBtn.Click += (s, e) => BackupEntries();
        Controls.Add(backupBtn);

        var restoreBtn = new Button
        {
            Text = "Restore...",
            Size = new Size(80, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        restoreBtn.FlatAppearance.BorderSize = 0;
        restoreBtn.Click += (s, e) => RestoreEntries();
        Controls.Add(restoreBtn);

        // Position buttons
        Resize += (s, e) =>
        {
            cancelBtn.Location = new Point(ClientSize.Width - cancelBtn.Width - 15, ClientSize.Height - cancelBtn.Height - 15);
            saveBtn.Location = new Point(cancelBtn.Left - saveBtn.Width - 10, cancelBtn.Top);
            backupBtn.Location = new Point(15, ClientSize.Height - backupBtn.Height - 15);
            restoreBtn.Location = new Point(backupBtn.Right + 5, backupBtn.Top);
        };
        cancelBtn.Location = new Point(ClientSize.Width - cancelBtn.Width - 15, ClientSize.Height - cancelBtn.Height - 15);
        saveBtn.Location = new Point(cancelBtn.Left - saveBtn.Width - 10, cancelBtn.Top);
        backupBtn.Location = new Point(15, ClientSize.Height - backupBtn.Height - 15);
        restoreBtn.Location = new Point(backupBtn.Right + 5, backupBtn.Top);

        CreateLeftPanel();
        CreateRightPanel();
    }

    private void CreateLeftPanel()
    {
        int x = 15, y = 15;

        var header = CreateLabel("ENTRIES", x, y, true);
        Controls.Add(header);

        var addBtn = CreateButton("+", x + 80, y - 3, 26);
        addBtn.Click += (s, e) => AddEntry();
        Controls.Add(addBtn);

        var removeBtn = CreateButton("-", x + 110, y - 3, 26);
        removeBtn.Click += (s, e) => RemoveEntry();
        Controls.Add(removeBtn);

        var cloneBtn = CreateButton("Clone", x + 140, y - 3, 50);
        cloneBtn.Click += (s, e) => CloneEntry();
        Controls.Add(cloneBtn);

        var moveUpBtn = CreateButton("^", x + 195, y - 3, 26);
        moveUpBtn.Click += (s, e) => MoveEntry(-1);
        Controls.Add(moveUpBtn);

        var moveDownBtn = CreateButton("v", x + 225, y - 3, 26);
        moveDownBtn.Click += (s, e) => MoveEntry(1);
        Controls.Add(moveDownBtn);
        y += 25;

        _listBox = new ListBox
        {
            Location = new Point(x, y),
            Size = new Size(230, 480),
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
        int x = 300, y = 15;
        int rightWidth = 415;

        var header = CreateLabel("ENTRY PROPERTIES", x, y, true);
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
            if (_suppressEvents || _selectedEntry == null) return;
            _selectedEntry.Enabled = _enabledCheck.Checked;
            RefreshListBox();
        };
        Controls.Add(_enabledCheck);
        y += 28;

        // Blurb
        Controls.Add(CreateLabel("Blurb (short label):", x, y + 2));
        _blurbBox = CreateTextBox(x + 130, y, 150);
        _blurbBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedEntry == null) return;
            _selectedEntry.Blurb = _blurbBox.Text;
            RefreshListBox();
        };
        Controls.Add(_blurbBox);
        y += 30;

        // Mode radio buttons
        Controls.Add(CreateLabel("Mode:", x, y + 2));
        _insertRadio = new RadioButton
        {
            Text = "Insert (append)",
            Location = new Point(x + 50, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(0, 200, 200) // cyan
        };
        _insertRadio.CheckedChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedEntry == null || !_insertRadio.Checked) return;
            _selectedEntry.ReplaceMode = false;
            RefreshListBox();
        };
        Controls.Add(_insertRadio);

        _replaceRadio = new RadioButton
        {
            Text = "Replace (overwrite)",
            Location = new Point(x + 190, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 180, 0) // amber
        };
        _replaceRadio.CheckedChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedEntry == null || !_replaceRadio.Checked) return;
            _selectedEntry.ReplaceMode = true;
            RefreshListBox();
        };
        Controls.Add(_replaceRadio);
        y += 28;

        // Require Comparison
        _requireComparisonCheck = new CheckBox
        {
            Text = "Only show when COMPARISON has a prior date",
            Location = new Point(x, y),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 180, 180)
        };
        _requireComparisonCheck.CheckedChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedEntry == null) return;
            _selectedEntry.RequireComparison = _requireComparisonCheck.Checked;
            _maxCompWeeksLabel.Enabled = _requireComparisonCheck.Checked && _selectedEntry != null;
            _maxCompWeeksBox.Enabled = _requireComparisonCheck.Checked && _selectedEntry != null;
        };
        Controls.Add(_requireComparisonCheck);
        y += 26;

        // Max comparison age
        _maxCompWeeksLabel = CreateLabel("Max age (weeks):", x + 25, y + 2);
        _maxCompWeeksLabel.ForeColor = Color.FromArgb(140, 140, 140);
        Controls.Add(_maxCompWeeksLabel);
        _maxCompWeeksBox = new NumericUpDown
        {
            Location = new Point(x + 145, y),
            Width = 60,
            Minimum = 0,
            Maximum = 520,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        _maxCompWeeksBox.ValueChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedEntry == null) return;
            _selectedEntry.MaxComparisonWeeks = (int)_maxCompWeeksBox.Value;
        };
        Controls.Add(_maxCompWeeksBox);

        var maxCompHint = CreateLabel("0 = no limit", x + 210, y + 2);
        maxCompHint.ForeColor = Color.FromArgb(100, 100, 100);
        maxCompHint.Font = new Font("Segoe UI", 8, FontStyle.Italic);
        Controls.Add(maxCompHint);
        y += 30;

        // Study Criteria section
        Controls.Add(CreateLabel("STUDY CRITERIA (match study description)", x, y, true, 8));
        y += 22;

        Controls.Add(CreateLabel("Required (all must match):", x, y + 2));
        y += 18;
        _criteriaRequiredBox = CreateTextBox(x, y, rightWidth);
        _criteriaRequiredBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _criteriaRequiredBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedEntry == null) return;
            _selectedEntry.CriteriaRequired = _criteriaRequiredBox.Text;
            RefreshListBox();
        };
        Controls.Add(_criteriaRequiredBox);
        y += 28;

        Controls.Add(CreateLabel("Any of (at least one must match):", x, y + 2));
        y += 18;
        _criteriaAnyOfBox = CreateTextBox(x, y, rightWidth);
        _criteriaAnyOfBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _criteriaAnyOfBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedEntry == null) return;
            _selectedEntry.CriteriaAnyOf = _criteriaAnyOfBox.Text;
            RefreshListBox();
        };
        Controls.Add(_criteriaAnyOfBox);
        y += 28;

        Controls.Add(CreateLabel("Exclude (none must match):", x, y + 2));
        y += 18;
        _criteriaExcludeBox = CreateTextBox(x, y, rightWidth);
        _criteriaExcludeBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _criteriaExcludeBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedEntry == null) return;
            _selectedEntry.CriteriaExclude = _criteriaExcludeBox.Text;
        };
        Controls.Add(_criteriaExcludeBox);
        y += 30;

        var hintLabel = CreateLabel("Separate multiple keywords with commas. Leave all empty for all studies.", x, y);
        hintLabel.ForeColor = Color.FromArgb(100, 100, 100);
        hintLabel.Font = new Font("Segoe UI", 8, FontStyle.Italic);
        Controls.Add(hintLabel);
        y += 22;

        // Text section
        Controls.Add(CreateLabel("IMPRESSION TEXT", x, y, true, 8));
        y += 20;

        _textBox = new TextBox
        {
            Location = new Point(x, y),
            Size = new Size(rightWidth, 130),
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
            if (_suppressEvents || _selectedEntry == null) return;
            _selectedEntry.Text = _textBox.Text;
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
        if (e.Index < 0 || e.Index >= _entries.Count) return;

        var entry = _entries[e.Index];
        e.DrawBackground();

        var isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        using var bgBrush = new SolidBrush(isSelected ? Color.FromArgb(50, 80, 110) : Color.FromArgb(45, 45, 45));
        e.Graphics.FillRectangle(bgBrush, e.Bounds);

        // Checkbox indicator
        var checkRect = new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 10, 14, 14);
        using var checkBrush = new SolidBrush(entry.Enabled ? Color.FromArgb(50, 120, 50) : Color.FromArgb(70, 70, 70));
        e.Graphics.FillRectangle(checkBrush, checkRect);
        if (entry.Enabled)
        {
            using var pen = new Pen(Color.White, 2);
            e.Graphics.DrawLine(pen, checkRect.X + 3, checkRect.Y + 7, checkRect.X + 6, checkRect.Y + 10);
            e.Graphics.DrawLine(pen, checkRect.X + 6, checkRect.Y + 10, checkRect.X + 11, checkRect.Y + 4);
        }

        // Mode-based color: cyan for insert, amber for replace
        Color modeColor = entry.ReplaceMode
            ? Color.FromArgb(255, 180, 0)
            : Color.FromArgb(0, 200, 200);

        // Mode prefix + blurb
        var prefix = entry.ReplaceMode ? "=" : "+";
        var blurb = string.IsNullOrEmpty(entry.Blurb) ? "(unnamed)" : entry.Blurb;
        var displayName = $"{prefix}{blurb}";
        using var nameFont = new Font("Segoe UI", 9, FontStyle.Bold);
        using var nameBrush = new SolidBrush(entry.Enabled ? modeColor : Color.Gray);
        e.Graphics.DrawString(displayName, nameFont, nameBrush, e.Bounds.X + 26, e.Bounds.Y + 3);

        // Criteria summary
        var criteria = entry.GetCriteriaDisplayString();
        if (criteria.Length > 28) criteria = criteria.Substring(0, 26) + "...";
        using var criteriaFont = new Font("Segoe UI", 8);
        using var criteriaBrush = new SolidBrush(Color.FromArgb(110, 110, 110));
        e.Graphics.DrawString(criteria, criteriaFont, criteriaBrush, e.Bounds.X + 26, e.Bounds.Y + 19);
    }

    private void ListBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_listBox.SelectedIndex >= 0 && _listBox.SelectedIndex < _entries.Count)
        {
            _selectedEntry = _entries[_listBox.SelectedIndex];
            LoadEntryProperties();
        }
        else
        {
            _selectedEntry = null;
            ClearEntryProperties();
        }
        UpdateButtonStates();
    }

    private void RefreshListBox()
    {
        var idx = _listBox.SelectedIndex;
        _listBox.Items.Clear();
        foreach (var entry in _entries)
            _listBox.Items.Add(entry);
        if (idx >= 0 && idx < _listBox.Items.Count)
            _listBox.SelectedIndex = idx;
        else if (_listBox.Items.Count > 0)
            _listBox.SelectedIndex = 0;
    }

    private void LoadEntryProperties()
    {
        if (_selectedEntry == null) return;
        _suppressEvents = true;
        _enabledCheck.Checked = _selectedEntry.Enabled;
        _blurbBox.Text = _selectedEntry.Blurb;
        if (_selectedEntry.ReplaceMode)
            _replaceRadio.Checked = true;
        else
            _insertRadio.Checked = true;
        _requireComparisonCheck.Checked = _selectedEntry.RequireComparison;
        _maxCompWeeksBox.Value = Math.Clamp(_selectedEntry.MaxComparisonWeeks, 0, 520);
        _maxCompWeeksLabel.Enabled = _selectedEntry.RequireComparison;
        _maxCompWeeksBox.Enabled = _selectedEntry.RequireComparison;
        _criteriaRequiredBox.Text = _selectedEntry.CriteriaRequired;
        _criteriaAnyOfBox.Text = _selectedEntry.CriteriaAnyOf;
        _criteriaExcludeBox.Text = _selectedEntry.CriteriaExclude;
        _textBox.Text = _selectedEntry.Text;
        _suppressEvents = false;
    }

    private void ClearEntryProperties()
    {
        _suppressEvents = true;
        _enabledCheck.Checked = false;
        _blurbBox.Text = "";
        _insertRadio.Checked = true;
        _requireComparisonCheck.Checked = false;
        _maxCompWeeksBox.Value = 0;
        _maxCompWeeksLabel.Enabled = false;
        _maxCompWeeksBox.Enabled = false;
        _criteriaRequiredBox.Text = "";
        _criteriaAnyOfBox.Text = "";
        _criteriaExcludeBox.Text = "";
        _textBox.Text = "";
        _suppressEvents = false;
    }

    private void UpdateButtonStates()
    {
        var hasEntry = _selectedEntry != null;
        _enabledCheck.Enabled = hasEntry;
        _blurbBox.Enabled = hasEntry;
        _insertRadio.Enabled = hasEntry;
        _replaceRadio.Enabled = hasEntry;
        _requireComparisonCheck.Enabled = hasEntry;
        _maxCompWeeksLabel.Enabled = hasEntry && (_selectedEntry?.RequireComparison ?? false);
        _maxCompWeeksBox.Enabled = hasEntry && (_selectedEntry?.RequireComparison ?? false);
        _criteriaRequiredBox.Enabled = hasEntry;
        _criteriaAnyOfBox.Enabled = hasEntry;
        _criteriaExcludeBox.Enabled = hasEntry;
        _textBox.Enabled = hasEntry;
    }

    private void AddEntry()
    {
        var newEntry = new ImpressionFixerEntry
        {
            Enabled = true,
            Blurb = "New",
            Text = "",
            ReplaceMode = false,
            RequireComparison = false
        };
        _entries.Add(newEntry);
        RefreshListBox();
        _listBox.SelectedIndex = _entries.Count - 1;
        _blurbBox.Focus();
        _blurbBox.SelectAll();
    }

    private void RemoveEntry()
    {
        if (_selectedEntry == null) return;
        var idx = _entries.IndexOf(_selectedEntry);
        _entries.Remove(_selectedEntry);
        RefreshListBox();
        if (_entries.Count > 0)
            _listBox.SelectedIndex = Math.Min(idx, _entries.Count - 1);
    }

    private void MoveEntry(int dir)
    {
        if (_selectedEntry == null) return;
        var idx = _entries.IndexOf(_selectedEntry);
        var newIdx = idx + dir;
        if (newIdx < 0 || newIdx >= _entries.Count) return;

        _entries.RemoveAt(idx);
        _entries.Insert(newIdx, _selectedEntry);
        RefreshListBox();
        _listBox.SelectedIndex = newIdx;
    }

    private void CloneEntry()
    {
        if (_selectedEntry == null) return;
        var clone = CloneEntry(_selectedEntry);
        clone.Blurb += " (Copy)";
        _entries.Add(clone);
        RefreshListBox();
        _listBox.SelectedIndex = _entries.Count - 1;
    }

    #region Backup/Restore

    private void BackupEntries()
    {
        if (_entries.Count == 0)
        {
            MessageBox.Show("No entries to backup.", "Backup", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Title = "Backup Impression Fixers",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            DefaultExt = "json",
            FileName = $"ImpressionFixers_Backup_{DateTime.Now:yyyy-MM-dd}.json"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        try
        {
            var json = JsonSerializer.Serialize(_entries, _jsonOptions);
            File.WriteAllText(dialog.FileName, json);
            MessageBox.Show($"Backed up {_entries.Count} entry/entries to:\n{dialog.FileName}",
                "Backup Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save backup:\n{ex.Message}", "Backup Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void RestoreEntries()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Restore Impression Fixers",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            DefaultExt = "json"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        List<ImpressionFixerEntry> imported;
        try
        {
            var json = File.ReadAllText(dialog.FileName);
            imported = JsonSerializer.Deserialize<List<ImpressionFixerEntry>>(json, _jsonOptions) ?? new();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to read backup file:\n{ex.Message}", "Restore Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (imported.Count == 0)
        {
            MessageBox.Show("No entries found in the backup file.", "Restore", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var message = $"Found {imported.Count} entry/entries in backup.\n\nImport all?";
        if (MessageBox.Show(message, "Confirm Restore", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        foreach (var item in imported)
        {
            var clone = CloneEntry(item);
            _entries.Add(clone);
        }

        RefreshListBox();
        if (_entries.Count > 0)
            _listBox.SelectedIndex = _entries.Count - 1;

        MessageBox.Show($"Imported {imported.Count} entry/entries.",
            "Restore Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    #endregion

    private void SaveAndClose()
    {
        _config.ImpressionFixers = _entries;
        _config.Save();
        DialogResult = DialogResult.OK;
        Close();
    }
}
