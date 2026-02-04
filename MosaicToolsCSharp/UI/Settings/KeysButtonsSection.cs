using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Keys & Buttons section: Links to open Keys and IV Buttons dialogs.
/// </summary>
public class KeysButtonsSection : SettingsSection
{
    public override string SectionId => "keys";

    private readonly CheckBox _floatingToolbarCheck;
    private readonly Configuration _config;
    private readonly ActionController _controller;

    public KeysButtonsSection(ToolTip toolTip, Configuration config, ActionController controller, bool isHeadless) : base("Keys & Buttons", toolTip)
    {
        _config = config;
        _controller = controller;

        // Keys Configuration
        AddSectionDivider("Action Mappings");

        AddLabel("Configure hotkeys and PowerMic button mappings for actions.", LeftMargin, _nextY);
        _nextY += SubRowHeight;

        var openKeysBtn = AddButton("Open Keys Configuration...", LeftMargin, _nextY, 200, 28, OnOpenKeysClick,
            "Configure hotkeys and mic button mappings for all actions.");
        _nextY += RowHeight + 10;

        // IV Buttons (non-headless only)
        if (!isHeadless)
        {
            AddSectionDivider("InteleViewer Buttons");

            _floatingToolbarCheck = AddCheckBox("Show InteleViewer Buttons toolbar", LeftMargin, _nextY,
                "Shows configurable buttons for InteleViewer shortcuts\n(window/level presets, zoom, etc.).");
            _nextY += SubRowHeight;

            var openButtonsBtn = AddButton("Open Button Studio...", LeftMargin, _nextY, 200, 28, OnOpenButtonsClick,
                "Configure the InteleViewer floating button toolbar.");
            _nextY += RowHeight + 10;
        }
        else
        {
            _floatingToolbarCheck = new CheckBox { Visible = false };
        }

        // Hardcoded buttons info
        AddSectionDivider("Hardcoded PowerMic Buttons");

        var infoLabel = AddLabel(
            "Skip Back → Process Report\n" +
            "Skip Forward → Generate Impression\n" +
            "Record Button → Start / Stop Dictation\n" +
            "Checkmark → Sign Report",
            LeftMargin, _nextY);
        infoLabel.ForeColor = Color.DarkGray;
        infoLabel.Font = new Font("Segoe UI", 8, FontStyle.Italic);
        infoLabel.AutoSize = true;
        _nextY += 70;

        UpdateHeight();
    }

    private void OnOpenKeysClick(object? sender, EventArgs e)
    {
        using var dialog = new KeyMappingsDialog(_config, _controller, App.IsHeadless);
        dialog.ShowDialog();
    }

    private void OnOpenButtonsClick(object? sender, EventArgs e)
    {
        using var dialog = new ButtonStudioDialog(_config);
        dialog.ShowDialog();
    }

    public override void LoadSettings(Configuration config)
    {
        _floatingToolbarCheck.Checked = config.FloatingToolbarEnabled;
    }

    public override void SaveSettings(Configuration config)
    {
        config.FloatingToolbarEnabled = _floatingToolbarCheck.Checked;
    }
}
