using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace MosaicTools.UI;

/// <summary>
/// Popup shown after an update (or fresh install) with a summary of new features.
/// </summary>
public class WhatsNewForm : Form
{
    private readonly string? _lastSeenVersion;
    private readonly string _currentVersion;

    public WhatsNewForm(string? lastSeenVersion, string currentVersion)
    {
        _lastSeenVersion = lastSeenVersion;
        _currentVersion = currentVersion;

        // Form properties - dark theme
        Text = "What's New in Mosaic Tools";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(30, 30, 30);
        ForeColor = Color.White;
        Size = new Size(420, 320);

        // Header label
        var headerLabel = new Label
        {
            Text = "What's New",
            Font = new Font("Segoe UI", 14, FontStyle.Bold),
            ForeColor = Color.FromArgb(100, 180, 255),
            BackColor = Color.Transparent,
            AutoSize = true,
            Location = new Point(20, 15)
        };
        Controls.Add(headerLabel);

        // Version label
        var versionLabel = new Label
        {
            Text = $"v{_currentVersion}",
            Font = new Font("Segoe UI", 10),
            ForeColor = Color.FromArgb(150, 150, 150),
            BackColor = Color.Transparent,
            AutoSize = true,
            Location = new Point(headerLabel.Right + 10, 20)
        };
        Controls.Add(versionLabel);

        // Content text box
        var contentBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Segoe UI", 10),
            ForeColor = Color.FromArgb(220, 220, 220),
            BackColor = Color.FromArgb(45, 45, 45),
            BorderStyle = BorderStyle.None,
            Location = new Point(20, 50),
            Size = new Size(365, 165)
        };
        contentBox.Text = GetRelevantChanges();
        contentBox.SelectionStart = 0;
        contentBox.SelectionLength = 0;
        Controls.Add(contentBox);

        // OK button - positioned below textbox with padding
        var okButton = new Button
        {
            Text = "OK",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.White,
            BackColor = Color.FromArgb(0, 120, 215),
            FlatStyle = FlatStyle.Flat,
            Size = new Size(100, 35),
            Location = new Point((ClientSize.Width - 100) / 2, 230),
            Cursor = Cursors.Hand
        };
        okButton.FlatAppearance.BorderSize = 0;
        okButton.Click += (_, _) => Close();
        Controls.Add(okButton);

        // Allow Enter or Escape to close
        AcceptButton = okButton;
        CancelButton = okButton;
    }

    /// <summary>
    /// Get changelog entries for versions newer than lastSeenVersion.
    /// </summary>
    private string GetRelevantChanges()
    {
        var whatsNewContent = LoadWhatsNewResource();
        if (string.IsNullOrWhiteSpace(whatsNewContent))
            return "Welcome to Mosaic Tools!";

        // Parse the content into version sections
        var lines = whatsNewContent.Split('\n');
        var result = new StringBuilder();
        string? currentVersionSection = null;
        bool includeSection = false;

        foreach (var rawLine in lines)
        {
            var line = rawLine.TrimEnd('\r');

            // Check if this is a version header (just a version number on its own line)
            if (IsVersionLine(line))
            {
                currentVersionSection = line.Trim();
                // Include if this version is newer than lastSeenVersion
                includeSection = ShouldIncludeVersion(currentVersionSection);

                if (includeSection)
                {
                    if (result.Length > 0)
                        result.AppendLine();
                    result.AppendLine($"v{currentVersionSection}");
                }
            }
            else if (includeSection && !string.IsNullOrWhiteSpace(line))
            {
                result.AppendLine(line);
            }
        }

        if (result.Length == 0)
            return "Thanks for using Mosaic Tools!";

        return result.ToString().TrimEnd();
    }

    /// <summary>
    /// Check if a line is a version number header.
    /// </summary>
    private static bool IsVersionLine(string line)
    {
        var trimmed = line.Trim();
        if (string.IsNullOrEmpty(trimmed))
            return false;

        // Version lines are just numbers and dots (e.g., "2.5.5")
        foreach (var c in trimmed)
        {
            if (!char.IsDigit(c) && c != '.')
                return false;
        }
        return trimmed.Contains('.');
    }

    /// <summary>
    /// Check if a version should be included based on lastSeenVersion.
    /// </summary>
    private bool ShouldIncludeVersion(string versionStr)
    {
        // If no last seen version, show everything (fresh install)
        if (string.IsNullOrWhiteSpace(_lastSeenVersion))
            return true;

        // Parse and compare versions
        if (!Version.TryParse(versionStr, out var version))
            return true; // Include if we can't parse

        if (!Version.TryParse(_lastSeenVersion, out var lastSeen))
            return true; // Include if we can't parse last seen

        return version > lastSeen;
    }

    /// <summary>
    /// Load the embedded WhatsNew.txt resource.
    /// </summary>
    private static string? LoadWhatsNewResource()
    {
        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "MosaicTools.WhatsNew.txt";

            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
            {
                Services.Logger.Trace($"WhatsNew resource not found: {resourceName}");
                return null;
            }

            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        catch (Exception ex)
        {
            Services.Logger.Trace($"Error loading WhatsNew resource: {ex.Message}");
            return null;
        }
    }
}
