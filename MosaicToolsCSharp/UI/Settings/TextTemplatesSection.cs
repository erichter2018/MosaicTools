using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Text & Templates settings: Report font, templates, macros, pick lists.
/// </summary>
public class TextTemplatesSection : SettingsSection
{
    public override string SectionId => "text";

    private readonly ComboBox _reportFontFamilyCombo;
    private readonly NumericUpDown _reportFontSizeNumeric;
    private readonly CheckBox _separatePastedItemsCheck;
    private readonly TextBox _criticalTemplateBox;
    private readonly TextBox _seriesTemplateBox;
    private readonly TextBox _comparisonTemplateBox;
    private readonly CheckBox _macrosEnabledCheck;
    private readonly CheckBox _macrosBlankLinesCheck;
    private readonly Label _macrosCountLabel;
    private readonly CheckBox _pickListsEnabledCheck;
    private readonly CheckBox _pickListSkipSingleMatchCheck;
    private readonly CheckBox _pickListKeepOpenCheck;
    private readonly Label _pickListsCountLabel;

    private readonly Configuration _config;

    public TextTemplatesSection(ToolTip toolTip, Configuration config) : base("Text & Templates", toolTip)
    {
        _config = config;

        // Report Popup Font
        AddSectionDivider("Report Popup Font");

        AddLabel("Font:", LeftMargin, _nextY + 3);
        _reportFontFamilyCombo = AddComboBox(LeftMargin + 50, _nextY, 140,
            new[] { "Consolas", "Courier New", "Cascadia Mono", "Lucida Console", "Segoe UI", "Calibri", "Arial" },
            "Font family for report popup.");

        AddLabel("Size:", LeftMargin + 205, _nextY + 3);
        _reportFontSizeNumeric = AddNumericUpDown(LeftMargin + 245, _nextY, 55, 7, 24, 11,
            "Font size for report popup.");
        _reportFontSizeNumeric.DecimalPlaces = 1;
        _reportFontSizeNumeric.Increment = 0.5m;

        var resetFontBtn = AddButton("Reset", LeftMargin + 310, _nextY - 2, 55, 24, (s, e) =>
        {
            _reportFontFamilyCombo.SelectedItem = "Consolas";
            _reportFontSizeNumeric.Value = 11;
        });
        _nextY += RowHeight + 5;

        _separatePastedItemsCheck = AddCheckBox("Separate pasted items with line break", LeftMargin, _nextY,
            "Add a blank line between items when pasting multiple things.");
        _nextY += RowHeight + 10;

        // Templates
        AddSectionDivider("Templates");

        AddLabel("Critical Findings:", LeftMargin, _nextY);
        _nextY += 18;
        _criticalTemplateBox = AddMultilineTextBox(LeftMargin, _nextY, 400, 45,
            "Template for pasting critical findings.\nPlaceholders: {name}, {time}, {date}");
        _nextY += 50;
        AddHintLabel("Placeholders: {name}, {time}, {date}", LeftMargin);

        AddLabel("Series/Image:", LeftMargin, _nextY);
        _nextY += 18;
        _seriesTemplateBox = AddTextBox(LeftMargin, _nextY, 400,
            "Template for series capture.\nPlaceholders: {series}, {image}");
        _nextY += 28;
        AddHintLabel("Placeholders: {series}, {image}", LeftMargin);

        AddLabel("Get Prior Comparison:", LeftMargin, _nextY);
        _nextY += 18;
        _comparisonTemplateBox = AddMultilineTextBox(LeftMargin, _nextY, 400, 45,
            "Template for Get Prior comparison line.\nPlaceholders: {date}, {time}, {description}, {noimages}");
        _nextY += 50;
        AddHintLabel("Placeholders: {date}, {time}, {description}, {noimages}", LeftMargin);
        _nextY += 5;

        // Macros
        AddSectionDivider("Macros");

        _macrosEnabledCheck = AddCheckBox("Enable Macros", LeftMargin, _nextY,
            "Auto-insert text snippets based on study description.\nRequires Scrape Mosaic to be enabled.");

        var editMacrosBtn = AddButton("Edit...", LeftMargin + 130, _nextY - 2, 60, 22, OnEditMacrosClick,
            "Edit macro definitions.");

        _macrosCountLabel = AddLabel(GetMacrosCountText(), LeftMargin + 200, _nextY + 2);
        _macrosCountLabel.ForeColor = Color.Gray;
        _macrosCountLabel.Font = new Font("Segoe UI", 8);
        _nextY += SubRowHeight;

        _macrosBlankLinesCheck = AddCheckBox("Add blank lines before macro text", LeftMargin + 25, _nextY,
            "Add 10 blank lines before macro text for dictation space.", true);
        _macrosEnabledCheck.CheckedChanged += (s, e) => UpdateMacroStates();
        _nextY += RowHeight + 5;

        // Pick Lists
        AddSectionDivider("Pick Lists");

        _pickListsEnabledCheck = AddCheckBox("Enable Pick Lists", LeftMargin, _nextY,
            "Show pick list popup for studies matching configured triggers.");

        var editPickListsBtn = AddButton("Edit...", LeftMargin + 140, _nextY - 2, 60, 22, OnEditPickListsClick,
            "Edit pick list definitions.");

        _pickListsCountLabel = AddLabel(GetPickListsCountText(), LeftMargin + 210, _nextY + 2);
        _pickListsCountLabel.ForeColor = Color.Gray;
        _pickListsCountLabel.Font = new Font("Segoe UI", 8);
        _nextY += SubRowHeight;

        _pickListSkipSingleMatchCheck = AddCheckBox("Skip list when only one matches", LeftMargin + 25, _nextY,
            "Automatically select if only one pick list matches the study.", true);
        _nextY += SubRowHeight;

        _pickListKeepOpenCheck = AddCheckBox("Keep window open", LeftMargin + 25, _nextY,
            "Keep pick list visible after inserting (use number keys for quick selection).", true);
        _pickListsEnabledCheck.CheckedChanged += (s, e) => UpdatePickListStates();
        _nextY += RowHeight;

        AddHintLabel("Assign \"Show Pick Lists\" action in Keys dialog.", LeftMargin);

        UpdateHeight();
    }

    private void UpdateMacroStates()
    {
        bool enabled = _macrosEnabledCheck.Checked;
        _macrosBlankLinesCheck.Enabled = enabled;
        _macrosBlankLinesCheck.ForeColor = enabled ? Color.FromArgb(180, 180, 180) : Color.FromArgb(100, 100, 100);
    }

    private void UpdatePickListStates()
    {
        bool enabled = _pickListsEnabledCheck.Checked;
        _pickListSkipSingleMatchCheck.Enabled = enabled;
        _pickListKeepOpenCheck.Enabled = enabled;
        _pickListSkipSingleMatchCheck.ForeColor = enabled ? Color.FromArgb(180, 180, 180) : Color.FromArgb(100, 100, 100);
        _pickListKeepOpenCheck.ForeColor = enabled ? Color.FromArgb(180, 180, 180) : Color.FromArgb(100, 100, 100);
    }

    private string GetMacrosCountText()
    {
        var count = _config.Macros.Count;
        return $"({count} macro{(count == 1 ? "" : "s")})";
    }

    private string GetPickListsCountText()
    {
        var count = _config.PickLists.Count;
        return $"({count} list{(count == 1 ? "" : "s")})";
    }

    private void OnEditMacrosClick(object? sender, EventArgs e)
    {
        using var editor = new MacroEditorForm(_config);
        editor.ShowDialog(FindForm());
        _macrosCountLabel.Text = GetMacrosCountText();
    }

    private void OnEditPickListsClick(object? sender, EventArgs e)
    {
        using var editor = new PickListEditorForm(_config);
        editor.ShowDialog(FindForm());
        _pickListsCountLabel.Text = GetPickListsCountText();
    }

    public override void LoadSettings(Configuration config)
    {
        _reportFontFamilyCombo.SelectedItem = config.ReportPopupFontFamily;
        if (_reportFontFamilyCombo.SelectedIndex < 0) _reportFontFamilyCombo.SelectedIndex = 0;

        _reportFontSizeNumeric.Value = (decimal)config.ReportPopupFontSize;
        _separatePastedItemsCheck.Checked = config.SeparatePastedItems;

        _criticalTemplateBox.Text = config.CriticalFindingsTemplate ?? "";
        _seriesTemplateBox.Text = config.SeriesImageTemplate ?? "";
        _comparisonTemplateBox.Text = config.ComparisonTemplate ?? "";

        _macrosEnabledCheck.Checked = config.MacrosEnabled;
        _macrosBlankLinesCheck.Checked = config.MacrosBlankLinesBefore;

        _pickListsEnabledCheck.Checked = config.PickListsEnabled;
        _pickListSkipSingleMatchCheck.Checked = config.PickListSkipSingleMatch;
        _pickListKeepOpenCheck.Checked = config.PickListKeepOpen;

        UpdateMacroStates();
        UpdatePickListStates();
    }

    public override void SaveSettings(Configuration config)
    {
        config.ReportPopupFontFamily = _reportFontFamilyCombo.SelectedItem?.ToString() ?? "Consolas";
        config.ReportPopupFontSize = (float)_reportFontSizeNumeric.Value;
        config.SeparatePastedItems = _separatePastedItemsCheck.Checked;

        config.CriticalFindingsTemplate = _criticalTemplateBox.Text;
        config.SeriesImageTemplate = _seriesTemplateBox.Text;
        config.ComparisonTemplate = _comparisonTemplateBox.Text;

        config.MacrosEnabled = _macrosEnabledCheck.Checked;
        config.MacrosBlankLinesBefore = _macrosBlankLinesCheck.Checked;

        config.PickListsEnabled = _pickListsEnabledCheck.Checked;
        config.PickListSkipSingleMatch = _pickListSkipSingleMatchCheck.Checked;
        config.PickListKeepOpen = _pickListKeepOpenCheck.Checked;
    }
}
