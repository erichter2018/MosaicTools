using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Profile settings: Doctor name, timezone, tooltip preferences.
/// </summary>
public class ProfileSection : SettingsSection
{
    public override string SectionId => "profile";

    private readonly TextBox _doctorNameBox;
    private readonly ComboBox _timezoneCombo;
    private readonly CheckBox _showTooltipsCheck;

    // Timezone options: null value means auto-detect/keep original
    private static readonly (string Display, string? Value)[] TimezoneOptions =
    {
        ("(Keep original)", null),
        ("Eastern Time", "Eastern Time"),
        ("Central Time", "Central Time"),
        ("Mountain Time", "Mountain Time"),
        ("Pacific Time", "Pacific Time")
    };

    public ProfileSection(ToolTip toolTip) : base("Profile", toolTip)
    {
        AddLabel("Doctor Name:", LeftMargin, _nextY + 3);
        _doctorNameBox = AddTextBox(LeftMargin + 110, _nextY, 200,
            "Your name as it appears in Clario. Used to filter your own notes\nfrom Critical Findings results so the contact person is identified.");
        _nextY += RowHeight + 5;

        AddLabel("Timezone:", LeftMargin, _nextY + 3);
        _timezoneCombo = AddComboBox(LeftMargin + 110, _nextY, 150,
            TimezoneOptions.Select(tz => tz.Display).ToArray(),
            "Timezone for Critical Findings output.\n(Keep original) preserves the timezone from the note.\nSelect a specific timezone to convert all times.");
        _nextY += RowHeight + 5;

        _showTooltipsCheck = AddCheckBox("Show tooltips throughout settings", LeftMargin, _nextY,
            "Display helpful tooltips when hovering over settings.");
        _nextY += RowHeight + 5;

        // Mic Gain Calibrator button
        AddButton("Mic Gain Calibrator", LeftMargin, _nextY, 160, 28, (s, e) =>
        {
            using var form = new AudioSetupForm();
            form.ShowDialog(FindForm());
        }, "Measure microphone levels and calculate the optimal Windows input gain for voice recognition.");
        _searchTerms.Add("mic gain calibrator");
        _searchTerms.Add("audio setup");
        _nextY += RowHeight + 5;

        // Acknowledgments
        AddLabel("Thank you to Brian Mowell for providing this feature.", LeftMargin, _nextY, isSubItem: true);
        _nextY += RowHeight + 5;

        AddLabel("Thank you to Dr. Tony Maung for providing code for the GetPrior function.", LeftMargin, _nextY, isSubItem: true);
        _nextY += RowHeight;

        UpdateHeight();
    }

    public CheckBox ShowTooltipsCheck => _showTooltipsCheck;

    public override void LoadSettings(Configuration config)
    {
        _doctorNameBox.Text = config.DoctorName ?? "";
        _showTooltipsCheck.Checked = config.ShowTooltips;

        // Find matching timezone option
        int tzIndex = 0;
        for (int i = 0; i < TimezoneOptions.Length; i++)
        {
            if (TimezoneOptions[i].Value == config.TargetTimezone)
            {
                tzIndex = i;
                break;
            }
        }
        _timezoneCombo.SelectedIndex = tzIndex;
    }

    public override void SaveSettings(Configuration config)
    {
        config.DoctorName = _doctorNameBox.Text.Trim();
        config.ShowTooltips = _showTooltipsCheck.Checked;

        // Map selected index to timezone value (null for "Keep original")
        int idx = _timezoneCombo.SelectedIndex;
        config.TargetTimezone = idx >= 0 && idx < TimezoneOptions.Length ? TimezoneOptions[idx].Value : null;
    }
}
