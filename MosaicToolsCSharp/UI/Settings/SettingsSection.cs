using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Base class for settings sections. Provides consistent styling,
/// search matching, and tooltip registration.
/// </summary>
public abstract class SettingsSection : Panel
{
    public string Title { get; }
    public abstract string SectionId { get; }

    protected readonly List<string> _searchTerms = new();
    protected readonly ToolTip _toolTip;
    protected int _nextY = 15;
    protected const int LeftMargin = 20;
    protected const int RightMargin = 20;
    protected const int RowHeight = 28;
    protected const int SubRowHeight = 24;

    private Label? _titleLabel;

    protected SettingsSection(string title, ToolTip toolTip)
    {
        Title = title;
        _toolTip = toolTip;
        _searchTerms.Add(title.ToLowerInvariant());

        BackColor = Color.FromArgb(45, 45, 48);
        Padding = new Padding(0, 0, 0, 15);

        // Title header
        _titleLabel = new Label
        {
            Text = title,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.White,
            Location = new Point(LeftMargin, 12),
            AutoSize = true
        };
        Controls.Add(_titleLabel);
        _nextY = 45;
    }

    /// <summary>
    /// Check if this section matches a search query.
    /// Matches against title and all registered search terms.
    /// </summary>
    public bool MatchesSearch(string query)
    {
        if (string.IsNullOrWhiteSpace(query)) return true;
        var q = query.ToLowerInvariant();
        return _searchTerms.Any(term => term.Contains(q));
    }

    /// <summary>
    /// Calculate the required height based on content.
    /// </summary>
    public void UpdateHeight()
    {
        Height = _nextY + 10;
    }

    #region Layout Helpers

    protected Label AddLabel(string text, int x, int y, bool isSubItem = false)
    {
        var label = new Label
        {
            Text = text,
            Location = new Point(x, y),
            AutoSize = true,
            ForeColor = isSubItem ? Color.FromArgb(200, 200, 200) : Color.White,
            Font = new Font("Segoe UI", 9)
        };
        Controls.Add(label);
        _searchTerms.Add(text.ToLowerInvariant());
        return label;
    }

    protected CheckBox AddCheckBox(string text, int x, int y, string? tooltip = null, bool isSubItem = false)
    {
        var cb = new CheckBox
        {
            Text = text,
            Location = new Point(x, y),
            AutoSize = true,
            ForeColor = isSubItem ? Color.FromArgb(180, 180, 180) : Color.White,
            Font = new Font("Segoe UI", 9)
        };
        Controls.Add(cb);
        _searchTerms.Add(text.ToLowerInvariant());

        if (!string.IsNullOrEmpty(tooltip))
        {
            _toolTip.SetToolTip(cb, tooltip);
        }
        return cb;
    }

    protected TextBox AddTextBox(int x, int y, int width, string? tooltip = null)
    {
        var tb = new TextBox
        {
            Location = new Point(x, y),
            Width = width,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        Controls.Add(tb);

        if (!string.IsNullOrEmpty(tooltip))
        {
            _toolTip.SetToolTip(tb, tooltip);
        }
        return tb;
    }

    protected TextBox AddMultilineTextBox(int x, int y, int width, int height, string? tooltip = null)
    {
        var tb = new TextBox
        {
            Location = new Point(x, y),
            Size = new Size(width, height),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };
        Controls.Add(tb);

        if (!string.IsNullOrEmpty(tooltip))
        {
            _toolTip.SetToolTip(tb, tooltip);
        }
        return tb;
    }

    protected NumericUpDown AddNumericUpDown(int x, int y, int width, decimal min, decimal max, decimal value, string? tooltip = null)
    {
        var nud = new NumericUpDown
        {
            Location = new Point(x, y),
            Width = width,
            Minimum = min,
            Maximum = max,
            Value = Math.Max(min, Math.Min(max, value)),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        Controls.Add(nud);

        if (!string.IsNullOrEmpty(tooltip))
        {
            _toolTip.SetToolTip(nud, tooltip);
        }
        return nud;
    }

    protected TrackBar AddTrackBar(int x, int y, int width, int min, int max, int value, string? tooltip = null)
    {
        var tb = new TrackBar
        {
            Location = new Point(x, y - 5),
            Size = new Size(width, 25),
            Minimum = min,
            Maximum = max,
            Value = Math.Max(min, Math.Min(max, value)),
            TickStyle = TickStyle.None,
            BackColor = Color.FromArgb(45, 45, 48),
            AutoSize = false
        };
        Controls.Add(tb);

        if (!string.IsNullOrEmpty(tooltip))
        {
            _toolTip.SetToolTip(tb, tooltip);
        }
        return tb;
    }

    protected ComboBox AddComboBox(int x, int y, int width, string[] items, string? tooltip = null)
    {
        var cb = new ComboBox
        {
            Location = new Point(x, y),
            Width = width,
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        cb.Items.AddRange(items);
        if (cb.Items.Count > 0) cb.SelectedIndex = 0;
        Controls.Add(cb);

        if (!string.IsNullOrEmpty(tooltip))
        {
            _toolTip.SetToolTip(cb, tooltip);
        }
        return cb;
    }

    protected Button AddButton(string text, int x, int y, int width, int height, EventHandler onClick, string? tooltip = null)
    {
        var btn = new Button
        {
            Text = text,
            Location = new Point(x, y),
            Size = new Size(width, height),
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
        btn.Click += onClick;
        Controls.Add(btn);

        if (!string.IsNullOrEmpty(tooltip))
        {
            _toolTip.SetToolTip(btn, tooltip);
        }
        return btn;
    }

    protected Panel AddColorPanel(int x, int y, int width, int height, Color color, EventHandler onClick, string? tooltip = null)
    {
        var panel = new Panel
        {
            Location = new Point(x, y),
            Size = new Size(width, height),
            BackColor = color,
            BorderStyle = BorderStyle.FixedSingle,
            Cursor = Cursors.Hand
        };
        panel.Click += onClick;
        Controls.Add(panel);

        if (!string.IsNullOrEmpty(tooltip))
        {
            _toolTip.SetToolTip(panel, tooltip);
        }
        return panel;
    }

    protected void AddSectionDivider(string text)
    {
        _nextY += 8;
        var label = new Label
        {
            Text = text,
            Location = new Point(LeftMargin, _nextY),
            AutoSize = true,
            ForeColor = Color.FromArgb(140, 140, 140),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        Controls.Add(label);
        _searchTerms.Add(text.ToLowerInvariant());
        _nextY += 22;
    }

    protected void AddHintLabel(string text, int x)
    {
        var label = new Label
        {
            Text = text,
            Location = new Point(x, _nextY),
            AutoSize = true,
            ForeColor = Color.FromArgb(120, 120, 120),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        Controls.Add(label);
        _nextY += 18;
    }

    protected void AddRow(string labelText, Control control, string? tooltip = null, int labelWidth = 150)
    {
        var label = AddLabel(labelText, LeftMargin, _nextY + 3);
        control.Location = new Point(LeftMargin + labelWidth, _nextY);
        if (!Controls.Contains(control)) Controls.Add(control);

        if (!string.IsNullOrEmpty(tooltip))
        {
            _toolTip.SetToolTip(label, tooltip);
            _toolTip.SetToolTip(control, tooltip);
        }

        _nextY += RowHeight;
    }

    #endregion

    /// <summary>
    /// Load settings from configuration into controls.
    /// </summary>
    public abstract void LoadSettings(Services.Configuration config);

    /// <summary>
    /// Save control values back to configuration.
    /// </summary>
    public abstract void SaveSettings(Services.Configuration config);
}
