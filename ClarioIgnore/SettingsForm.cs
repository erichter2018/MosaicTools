using System;
using System.Drawing;
using System.Windows.Forms;

namespace ClarioIgnore;

public class SettingsForm : Form
{
    private ListBox _rulesList = null!;
    private TextBox _nameBox = null!;
    private TextBox _requiredBox = null!;
    private TextBox _anyOfBox = null!;
    private TextBox _excludeBox = null!;
    private CheckBox _enabledCheck = null!;
    private CheckBox _includePriorityCheck = null!;
    private Button _addButton = null!;
    private Button _deleteButton = null!;
    private Button _saveButton = null!;
    private Button _testButton = null!;
    private NumericUpDown _intervalUpDown = null!;
    private NumericUpDown _maxFilesUpDown = null!;
    private CheckBox _notificationsCheck = null!;

    private int _selectedIndex = -1;

    public SettingsForm()
    {
        InitializeComponent();
        LoadRules();
    }

    private void InitializeComponent()
    {
        Text = "ClarioIgnore Settings";
        Size = new Size(600, 500);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        // Rules list
        var rulesLabel = new Label { Text = "Skip Rules:", Location = new Point(10, 10), AutoSize = true };
        Controls.Add(rulesLabel);

        _rulesList = new ListBox
        {
            Location = new Point(10, 30),
            Size = new Size(200, 200)
        };
        _rulesList.SelectedIndexChanged += RulesList_SelectedIndexChanged;
        Controls.Add(_rulesList);

        // Add/Delete buttons
        _addButton = new Button { Text = "Add", Location = new Point(10, 235), Size = new Size(60, 25) };
        _addButton.Click += AddButton_Click;
        Controls.Add(_addButton);

        _deleteButton = new Button { Text = "Delete", Location = new Point(75, 235), Size = new Size(60, 25) };
        _deleteButton.Click += DeleteButton_Click;
        Controls.Add(_deleteButton);

        // Rule editor panel
        var editorPanel = new Panel { Location = new Point(220, 30), Size = new Size(360, 230) };
        Controls.Add(editorPanel);

        int y = 0;

        var enabledLabel = new Label { Text = "Enabled:", Location = new Point(0, y), AutoSize = true };
        editorPanel.Controls.Add(enabledLabel);
        _enabledCheck = new CheckBox { Location = new Point(100, y), Checked = true };
        editorPanel.Controls.Add(_enabledCheck);
        y += 25;

        var nameLabel = new Label { Text = "Rule Name:", Location = new Point(0, y), AutoSize = true };
        editorPanel.Controls.Add(nameLabel);
        _nameBox = new TextBox { Location = new Point(100, y), Size = new Size(250, 20) };
        editorPanel.Controls.Add(_nameBox);
        y += 30;

        var requiredLabel = new Label { Text = "Required:", Location = new Point(0, y), AutoSize = true };
        editorPanel.Controls.Add(requiredLabel);
        _requiredBox = new TextBox { Location = new Point(100, y), Size = new Size(250, 20) };
        editorPanel.Controls.Add(_requiredBox);
        y += 25;

        var reqHint = new Label
        {
            Text = "ALL terms must match (comma-separated)",
            Location = new Point(100, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font(Font.FontFamily, 7)
        };
        editorPanel.Controls.Add(reqHint);
        y += 20;

        var anyOfLabel = new Label { Text = "Any Of:", Location = new Point(0, y), AutoSize = true };
        editorPanel.Controls.Add(anyOfLabel);
        _anyOfBox = new TextBox { Location = new Point(100, y), Size = new Size(250, 20) };
        editorPanel.Controls.Add(_anyOfBox);
        y += 25;

        var anyHint = new Label
        {
            Text = "At least ONE must match (comma-separated)",
            Location = new Point(100, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font(Font.FontFamily, 7)
        };
        editorPanel.Controls.Add(anyHint);
        y += 20;

        var excludeLabel = new Label { Text = "Exclude:", Location = new Point(0, y), AutoSize = true };
        editorPanel.Controls.Add(excludeLabel);
        _excludeBox = new TextBox { Location = new Point(100, y), Size = new Size(250, 20) };
        editorPanel.Controls.Add(_excludeBox);
        y += 25;

        var exclHint = new Label
        {
            Text = "NONE of these can match (comma-separated)",
            Location = new Point(100, y),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font(Font.FontFamily, 7)
        };
        editorPanel.Controls.Add(exclHint);
        y += 25;

        _includePriorityCheck = new CheckBox
        {
            Text = "Include priority in matching (e.g., STAT IP, Routine)",
            Location = new Point(0, y),
            AutoSize = true
        };
        editorPanel.Controls.Add(_includePriorityCheck);

        // General settings
        var settingsGroup = new GroupBox
        {
            Text = "General Settings",
            Location = new Point(10, 270),
            Size = new Size(570, 80)
        };
        Controls.Add(settingsGroup);

        var intervalLabel = new Label { Text = "Poll interval (seconds):", Location = new Point(10, 25), AutoSize = true };
        settingsGroup.Controls.Add(intervalLabel);

        _intervalUpDown = new NumericUpDown
        {
            Location = new Point(150, 23),
            Size = new Size(60, 20),
            Minimum = 5,
            Maximum = 120,
            Value = Configuration.Instance.PollIntervalSeconds
        };
        settingsGroup.Controls.Add(_intervalUpDown);

        var maxFilesLabel = new Label { Text = "Max files to monitor:", Location = new Point(230, 25), AutoSize = true };
        settingsGroup.Controls.Add(maxFilesLabel);

        _maxFilesUpDown = new NumericUpDown
        {
            Location = new Point(360, 23),
            Size = new Size(60, 20),
            Minimum = 1,
            Maximum = 50,
            Value = Configuration.Instance.MaxFilesToMonitor
        };
        settingsGroup.Controls.Add(_maxFilesUpDown);

        _notificationsCheck = new CheckBox
        {
            Text = "Show notifications when studies are skipped",
            Location = new Point(10, 50),
            AutoSize = true,
            Checked = Configuration.Instance.ShowNotifications
        };
        settingsGroup.Controls.Add(_notificationsCheck);

        // Test and Save buttons
        _testButton = new Button
        {
            Text = "Test Rules",
            Location = new Point(10, 360),
            Size = new Size(100, 30)
        };
        _testButton.Click += TestButton_Click;
        Controls.Add(_testButton);

        _saveButton = new Button
        {
            Text = "Save && Close",
            Location = new Point(480, 360),
            Size = new Size(100, 30)
        };
        _saveButton.Click += SaveButton_Click;
        Controls.Add(_saveButton);

        // Status/help text
        var helpLabel = new Label
        {
            Text = "Rules are checked every poll interval. Studies matching any enabled rule will be skipped.\n" +
                   "Example: Required='US', AnyOf='VENOUS,DOPPLER', Exclude='ARTERIAL' matches any US study\n" +
                   "that contains either VENOUS or DOPPLER, but not ARTERIAL.",
            Location = new Point(10, 400),
            Size = new Size(570, 50),
            ForeColor = Color.DimGray
        };
        Controls.Add(helpLabel);
    }

    private void LoadRules()
    {
        _rulesList.Items.Clear();
        foreach (var rule in Configuration.Instance.SkipRules)
        {
            var prefix = rule.Enabled ? "✓ " : "✗ ";
            _rulesList.Items.Add(prefix + rule.Name);
        }
    }

    private void RulesList_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_selectedIndex >= 0 && _selectedIndex < Configuration.Instance.SkipRules.Count)
        {
            // Save current edits before switching
            SaveCurrentRule();
        }

        _selectedIndex = _rulesList.SelectedIndex;

        if (_selectedIndex >= 0 && _selectedIndex < Configuration.Instance.SkipRules.Count)
        {
            var rule = Configuration.Instance.SkipRules[_selectedIndex];
            _enabledCheck.Checked = rule.Enabled;
            _nameBox.Text = rule.Name;
            _requiredBox.Text = rule.CriteriaRequired;
            _anyOfBox.Text = rule.CriteriaAnyOf;
            _excludeBox.Text = rule.CriteriaExclude;
            _includePriorityCheck.Checked = rule.IncludePriority;
        }
    }

    private void SaveCurrentRule()
    {
        if (_selectedIndex >= 0 && _selectedIndex < Configuration.Instance.SkipRules.Count)
        {
            var rule = Configuration.Instance.SkipRules[_selectedIndex];
            rule.Enabled = _enabledCheck.Checked;
            rule.Name = _nameBox.Text.Trim();
            rule.CriteriaRequired = _requiredBox.Text.Trim();
            rule.CriteriaAnyOf = _anyOfBox.Text.Trim();
            rule.CriteriaExclude = _excludeBox.Text.Trim();
            rule.IncludePriority = _includePriorityCheck.Checked;
        }
    }

    private void AddButton_Click(object? sender, EventArgs e)
    {
        var newRule = new SkipRule
        {
            Name = "New Rule",
            Enabled = true
        };
        Configuration.Instance.SkipRules.Add(newRule);
        LoadRules();
        _rulesList.SelectedIndex = _rulesList.Items.Count - 1;
    }

    private void DeleteButton_Click(object? sender, EventArgs e)
    {
        if (_selectedIndex >= 0 && _selectedIndex < Configuration.Instance.SkipRules.Count)
        {
            Configuration.Instance.SkipRules.RemoveAt(_selectedIndex);
            _selectedIndex = -1;
            LoadRules();

            // Clear editor
            _enabledCheck.Checked = false;
            _nameBox.Text = "";
            _requiredBox.Text = "";
            _anyOfBox.Text = "";
            _excludeBox.Text = "";
        }
    }

    private void TestButton_Click(object? sender, EventArgs e)
    {
        SaveCurrentRule();
        LoadRules(); // Refresh display

        using var service = new ClarioService();
        var matches = service.GetMatchingItems();

        if (matches.Count == 0)
        {
            MessageBox.Show("No worklist items match any enabled rules.\n\n" +
                           "(Make sure Clario is open with studies in the worklist)",
                           "Test Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            var message = $"Found {matches.Count} matching item(s):\n\n";
            foreach (var (item, rule) in matches)
            {
                message += $"• {item.Procedure}\n  → Rule: {rule.Name}\n\n";
            }
            MessageBox.Show(message, "Test Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void SaveButton_Click(object? sender, EventArgs e)
    {
        SaveCurrentRule();
        LoadRules(); // Refresh display

        Configuration.Instance.PollIntervalSeconds = (int)_intervalUpDown.Value;
        Configuration.Instance.MaxFilesToMonitor = (int)_maxFilesUpDown.Value;
        Configuration.Instance.ShowNotifications = _notificationsCheck.Checked;
        Configuration.Instance.Save();

        DialogResult = DialogResult.OK;
        Close();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        SaveCurrentRule();
        base.OnFormClosing(e);
    }
}
