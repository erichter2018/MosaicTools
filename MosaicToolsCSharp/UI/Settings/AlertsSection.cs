using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Alerts settings: Critical studies tracker, notification box, alert triggers.
/// </summary>
public class AlertsSection : SettingsSection
{
    public override string SectionId => "alerts";

    private readonly CheckBox _trackCriticalStudiesCheck;
    private readonly CheckBox _showClinicalHistoryCheck;
    private readonly CheckBox _alwaysShowClinicalHistoryCheck;
    private readonly CheckBox _hideClinicalHistoryWhenNoStudyCheck;
    private readonly CheckBox _autoFixClinicalHistoryCheck;
    private readonly CheckBox _showDraftedIndicatorCheck;
    private readonly CheckBox _showTemplateMismatchCheck;
    private readonly CheckBox _genderCheckEnabledCheck;
    private readonly CheckBox _strokeDetectionEnabledCheck;
    private readonly CheckBox _strokeDetectionUseClinicalHistoryCheck;
    private readonly CheckBox _strokeClickToCreateNoteCheck;
    private readonly CheckBox _strokeAutoCreateNoteCheck;

    private readonly MainForm _mainForm;
    private readonly Configuration _config;
    private List<string>? _pendingStrokeKeywords;

    public AlertsSection(ToolTip toolTip, MainForm mainForm, Configuration config) : base("Alerts", toolTip)
    {
        _mainForm = mainForm;
        _config = config;

        // Critical Studies Tracker
        AddSectionDivider("Critical Studies Tracker");

        _trackCriticalStudiesCheck = AddCheckBox("Track critical studies this session", LeftMargin, _nextY,
            "Show badge count of critical notes created this session.\nClick badge to see list.");
        _nextY += SubRowHeight;
        AddHintLabel("Shows count badge on main bar when critical notes are created", LeftMargin + 25);
        _nextY += 5;

        // Notification Box
        AddSectionDivider("Notification Box");

        _showClinicalHistoryCheck = AddCheckBox("Enable Notification Box", LeftMargin, _nextY,
            "Floating window showing clinical history and alerts.\nRequires Scrape Mosaic to be enabled.");
        _showClinicalHistoryCheck.CheckedChanged += (s, e) => UpdateNotificationBoxStates();
        _nextY += SubRowHeight;
        AddHintLabel("Floating window for clinical history and alerts", LeftMargin + 25);
        _nextY += 3;

        _alwaysShowClinicalHistoryCheck = AddCheckBox("Show clinical history in notification box", LeftMargin + 25, _nextY,
            "Display clinical history text. Unchecked = alerts-only mode.", true);
        _alwaysShowClinicalHistoryCheck.CheckedChanged += (s, e) => UpdateNotificationBoxStates();
        _nextY += SubRowHeight;
        AddHintLabel("Unchecked = alerts-only mode (hidden until alert)", LeftMargin + 50);
        _nextY += 3;

        _hideClinicalHistoryWhenNoStudyCheck = AddCheckBox("Hide when no study open", LeftMargin + 50, _nextY,
            "Hides the notification box when no study is open in Mosaic.", true);
        _hideClinicalHistoryWhenNoStudyCheck.CheckedChanged += (s, e) => _mainForm.UpdateClinicalHistoryVisibility();
        _nextY += SubRowHeight;

        _autoFixClinicalHistoryCheck = AddCheckBox("Auto-paste corrected history", LeftMargin + 50, _nextY,
            "Automatically paste corrected clinical history to Mosaic\nwhen malformed text is detected.", true);
        _nextY += SubRowHeight;

        _showDraftedIndicatorCheck = AddCheckBox("Show Drafted indicator", LeftMargin + 50, _nextY,
            "Green border when report has IMPRESSION section (indicates draft in progress).", true);
        var greenLabel = AddLabel("(green border)", LeftMargin + 230, _nextY + 2);
        greenLabel.ForeColor = Color.FromArgb(100, 180, 100);
        greenLabel.Font = new Font("Segoe UI", 8);
        _nextY += RowHeight + 5;

        // Alert Triggers
        AddSectionDivider("Alert Triggers");
        AddHintLabel("These alerts appear even in alerts-only mode:", LeftMargin);
        _nextY += 3;

        // Template Mismatch
        _showTemplateMismatchCheck = AddCheckBox("Template mismatch", LeftMargin, _nextY,
            "Red border when report template doesn't match study type.");
        var redLabel = AddLabel("(red border)", LeftMargin + 160, _nextY + 2);
        redLabel.ForeColor = Color.FromArgb(220, 100, 100);
        redLabel.Font = new Font("Segoe UI", 8);
        _nextY += SubRowHeight;

        // Gender Check
        _genderCheckEnabledCheck = AddCheckBox("Gender check", LeftMargin, _nextY,
            "Flashing red border for gender-specific terms in wrong patient.");
        var flashingRedLabel = AddLabel("(flashing red)", LeftMargin + 160, _nextY + 2);
        flashingRedLabel.ForeColor = Color.FromArgb(255, 100, 100);
        flashingRedLabel.Font = new Font("Segoe UI", 8);
        _nextY += SubRowHeight;

        // Stroke Detection
        _strokeDetectionEnabledCheck = AddCheckBox("Stroke detection", LeftMargin, _nextY,
            "Purple border for stroke-related studies.");
        _strokeDetectionEnabledCheck.CheckedChanged += (s, e) => UpdateStrokeSubStates();
        var purpleLabel = AddLabel("(purple border)", LeftMargin + 160, _nextY + 2);
        purpleLabel.ForeColor = Color.FromArgb(180, 130, 220);
        purpleLabel.Font = new Font("Segoe UI", 8);
        _nextY += SubRowHeight;

        _strokeDetectionUseClinicalHistoryCheck = AddCheckBox("Also use clinical history keywords", LeftMargin + 25, _nextY,
            "Also check clinical history for stroke keywords.", true);
        var editKeywordsBtn = AddButton("Edit...", LeftMargin + 270, _nextY - 2, 50, 20, OnEditStrokeKeywordsClick,
            "Edit the list of stroke-related keywords.");
        editKeywordsBtn.Font = new Font("Segoe UI", 7);
        _nextY += SubRowHeight;

        _strokeClickToCreateNoteCheck = AddCheckBox("Click alert to create Clario note", LeftMargin + 25, _nextY,
            "Click purple alert to create Clario communication note.", true);
        _nextY += SubRowHeight;

        _strokeAutoCreateNoteCheck = AddCheckBox("Auto-create Clario note on Process Report", LeftMargin + 25, _nextY,
            "Automatically create note when pressing Process Report for stroke cases.", true);
        _nextY += RowHeight;

        UpdateHeight();
    }

    private void UpdateNotificationBoxStates()
    {
        bool enabled = _showClinicalHistoryCheck.Checked;
        _alwaysShowClinicalHistoryCheck.Enabled = enabled;
        _hideClinicalHistoryWhenNoStudyCheck.Enabled = enabled && _alwaysShowClinicalHistoryCheck.Checked;
        _autoFixClinicalHistoryCheck.Enabled = enabled;
        _showDraftedIndicatorCheck.Enabled = enabled;

        var subColor = enabled ? Color.FromArgb(180, 180, 180) : Color.FromArgb(100, 100, 100);
        _alwaysShowClinicalHistoryCheck.ForeColor = subColor;
        _hideClinicalHistoryWhenNoStudyCheck.ForeColor = enabled && _alwaysShowClinicalHistoryCheck.Checked
            ? Color.FromArgb(180, 180, 180) : Color.FromArgb(100, 100, 100);
        _autoFixClinicalHistoryCheck.ForeColor = subColor;
        _showDraftedIndicatorCheck.ForeColor = subColor;
    }

    private void UpdateStrokeSubStates()
    {
        bool enabled = _strokeDetectionEnabledCheck.Checked;
        _strokeDetectionUseClinicalHistoryCheck.Enabled = enabled;
        _strokeClickToCreateNoteCheck.Enabled = enabled;
        _strokeAutoCreateNoteCheck.Enabled = enabled;

        var subColor = enabled ? Color.FromArgb(180, 180, 180) : Color.FromArgb(100, 100, 100);
        _strokeDetectionUseClinicalHistoryCheck.ForeColor = subColor;
        _strokeClickToCreateNoteCheck.ForeColor = subColor;
        _strokeAutoCreateNoteCheck.ForeColor = subColor;
    }

    private void OnEditStrokeKeywordsClick(object? sender, EventArgs e)
    {
        var currentKeywords = string.Join(Environment.NewLine, _config.StrokeClinicalHistoryKeywords ?? new List<string>());
        var input = InputBox.Show("Enter stroke-related keywords (one per line):",
            "Stroke Keywords",
            currentKeywords,
            multiline: true);

        if (input != null && input != currentKeywords)
        {
            _pendingStrokeKeywords = input
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim().ToLowerInvariant())
                .Where(s => !string.IsNullOrEmpty(s))
                .Distinct()
                .ToList();
        }
    }

    public override void LoadSettings(Configuration config)
    {
        _trackCriticalStudiesCheck.Checked = config.TrackCriticalStudies;
        _showClinicalHistoryCheck.Checked = config.ShowClinicalHistory;
        _alwaysShowClinicalHistoryCheck.Checked = config.AlwaysShowClinicalHistory;
        _hideClinicalHistoryWhenNoStudyCheck.Checked = config.HideClinicalHistoryWhenNoStudy;
        _autoFixClinicalHistoryCheck.Checked = config.AutoFixClinicalHistory;
        _showDraftedIndicatorCheck.Checked = config.ShowDraftedIndicator;
        _showTemplateMismatchCheck.Checked = config.ShowTemplateMismatch;
        _genderCheckEnabledCheck.Checked = config.GenderCheckEnabled;
        _strokeDetectionEnabledCheck.Checked = config.StrokeDetectionEnabled;
        _strokeDetectionUseClinicalHistoryCheck.Checked = config.StrokeDetectionUseClinicalHistory;
        _strokeClickToCreateNoteCheck.Checked = config.StrokeClickToCreateNote;
        _strokeAutoCreateNoteCheck.Checked = config.StrokeAutoCreateNote;

        UpdateNotificationBoxStates();
        UpdateStrokeSubStates();
    }

    public override void SaveSettings(Configuration config)
    {
        config.TrackCriticalStudies = _trackCriticalStudiesCheck.Checked;
        config.ShowClinicalHistory = _showClinicalHistoryCheck.Checked;
        config.AlwaysShowClinicalHistory = _alwaysShowClinicalHistoryCheck.Checked;
        config.HideClinicalHistoryWhenNoStudy = _hideClinicalHistoryWhenNoStudyCheck.Checked;
        config.AutoFixClinicalHistory = _autoFixClinicalHistoryCheck.Checked;
        config.ShowDraftedIndicator = _showDraftedIndicatorCheck.Checked;
        config.ShowTemplateMismatch = _showTemplateMismatchCheck.Checked;
        config.GenderCheckEnabled = _genderCheckEnabledCheck.Checked;
        config.StrokeDetectionEnabled = _strokeDetectionEnabledCheck.Checked;
        config.StrokeDetectionUseClinicalHistory = _strokeDetectionUseClinicalHistoryCheck.Checked;
        config.StrokeClickToCreateNote = _strokeClickToCreateNoteCheck.Checked;
        config.StrokeAutoCreateNote = _strokeAutoCreateNoteCheck.Checked;

        if (_pendingStrokeKeywords != null)
            config.StrokeClinicalHistoryKeywords = _pendingStrokeKeywords;
    }
}
