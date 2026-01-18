using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Floating window displaying clinical history from Clario scrape.
/// </summary>
public class ClinicalHistoryForm : Form
{
    private readonly Configuration _config;
    private readonly Label _contentLabel;

    // Drag state
    private Point _dragStart;
    private bool _dragging;

    // For center-based positioning
    private bool _initialPositionSet = false;

    // Border colors
    private static readonly Color NormalBorderColor = Color.FromArgb(60, 60, 60);
    private static readonly Color DraftedBorderColor = Color.FromArgb(0, 180, 0); // Green
    private static readonly Color TemplateMismatchBorderColor = Color.FromArgb(220, 50, 50); // Red

    // Template matching state for debug
    private string? _lastDescription;
    private string? _lastTemplateName;
    private bool _templateMismatch = false;

    public ClinicalHistoryForm(Configuration config)
    {
        _config = config;

        // Form properties - frameless, topmost, black background
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(60, 60, 60); // Border color
        StartPosition = FormStartPosition.Manual;
        Padding = new Padding(1); // 1px border

        // Use TableLayoutPanel for proper auto-sizing
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.Black,
            ColumnCount = 1,
            RowCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(0),
            Margin = new Padding(0)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 16)); // Drag bar
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Content
        Controls.Add(layout);

        // Drag bar
        var dragBar = new Label
        {
            Text = "⋯",
            Font = new Font("Segoe UI", 8),
            ForeColor = Color.FromArgb(102, 102, 102),
            BackColor = Color.Black,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Cursor = Cursors.SizeAll,
            Margin = new Padding(0)
        };
        dragBar.MouseDown += OnDragStart;
        dragBar.MouseMove += OnDragMove;
        dragBar.MouseUp += OnDragEnd;
        layout.Controls.Add(dragBar, 0, 0);

        // Content label - bold
        _contentLabel = new Label
        {
            Text = "(No clinical history)",
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(200, 200, 200),
            BackColor = Color.Black,
            AutoSize = true,
            Margin = new Padding(10, 5, 10, 15),
            MaximumSize = new Size(0, 0) // No limit
        };
        layout.Controls.Add(_contentLabel, 0, 1);

        // Make form auto-size to content
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;

        // Context menu for drag bar only (Close option)
        var menu = new ContextMenuStrip();
        menu.Items.Add("Close", null, (_, _) => Close());
        dragBar.ContextMenuStrip = menu;

        // Right-click on content label copies debug info
        _contentLabel.MouseDown += OnContentLabelMouseDown;

        // Position from center after form sizes itself
        Shown += (_, _) => PositionFromCenter();
        SizeChanged += (_, _) => { if (_initialPositionSet) RepositionToCenter(); };
    }

    /// <summary>
    /// Position the form so its center is at the saved coordinates.
    /// </summary>
    private void PositionFromCenter()
    {
        Location = new Point(
            _config.ClinicalHistoryX - Width / 2,
            _config.ClinicalHistoryY - Height / 2
        );
        _initialPositionSet = true;
    }

    /// <summary>
    /// When size changes, keep the center in the same place.
    /// </summary>
    private void RepositionToCenter()
    {
        if (_dragging) return; // Don't reposition while dragging
        Location = new Point(
            _config.ClinicalHistoryX - Width / 2,
            _config.ClinicalHistoryY - Height / 2
        );
    }

    /// <summary>
    /// Update the displayed clinical history text.
    /// </summary>
    public void SetClinicalHistory(string? text)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetClinicalHistory(text));
            return;
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            _contentLabel.Text = "(No clinical history)";
            _contentLabel.ForeColor = Color.FromArgb(128, 128, 128);
        }
        else
        {
            _contentLabel.Text = text;
            _contentLabel.ForeColor = Color.FromArgb(200, 200, 200);
        }
    }

    /// <summary>
    /// Set the drafted state - shows green border when true (unless template mismatch).
    /// </summary>
    public void SetDraftedState(bool isDrafted)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetDraftedState(isDrafted));
            return;
        }

        // Template mismatch (red) overrides drafted (green)
        if (_templateMismatch)
        {
            BackColor = TemplateMismatchBorderColor;
        }
        else
        {
            BackColor = isDrafted ? DraftedBorderColor : NormalBorderColor;
        }
    }

    /// <summary>
    /// Set the template mismatch state - shows red border when true.
    /// This overrides the drafted (green) state.
    /// </summary>
    public void SetTemplateMismatchState(bool isMismatch, string? description, string? templateName)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetTemplateMismatchState(isMismatch, description, templateName));
            return;
        }

        _templateMismatch = isMismatch;
        _lastDescription = description;
        _lastTemplateName = templateName;

        // Update border color immediately
        if (isMismatch)
        {
            BackColor = TemplateMismatchBorderColor;
        }
        // If not mismatch, the border will be set by SetDraftedState
    }

    /// <summary>
    /// Handle right-click on content label - copy debug info to clipboard.
    /// </summary>
    private void OnContentLabelMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Right)
        {
            CopyDebugInfoToClipboard();
        }
    }

    /// <summary>
    /// Copy template matching debug info to clipboard.
    /// </summary>
    private void CopyDebugInfoToClipboard()
    {
        var debugInfo = $"=== Template Matching Debug ===\r\n" +
                        $"Description: {_lastDescription ?? "(none)"}\r\n" +
                        $"Template: {_lastTemplateName ?? "(none)"}\r\n" +
                        $"Mismatch: {_templateMismatch}";

        try
        {
            Clipboard.SetText(debugInfo);
            Services.Logger.Trace("Template debug info copied to clipboard");
            ShowCopiedToast();
        }
        catch (Exception ex)
        {
            Services.Logger.Trace($"Failed to copy debug info: {ex.Message}");
        }
    }

    /// <summary>
    /// Show a brief toast indicating debug info was copied.
    /// </summary>
    private void ShowCopiedToast()
    {
        var toast = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            ShowInTaskbar = false,
            TopMost = true,
            BackColor = Color.FromArgb(51, 51, 51),
            Size = new Size(120, 30),
            StartPosition = FormStartPosition.Manual,
            Location = new Point(Location.X + Width / 2 - 60, Location.Y - 35)
        };

        var label = new Label
        {
            Text = "Debug copied!",
            ForeColor = Color.White,
            BackColor = Color.FromArgb(51, 51, 51),
            Font = new Font("Segoe UI", 9),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter
        };
        toast.Controls.Add(label);
        toast.Show();

        var timer = new System.Windows.Forms.Timer { Interval = 1500 };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            if (!toast.IsDisposed) toast.Close();
        };
        timer.Start();
    }

    /// <summary>
    /// Remove non-printable and special characters from text.
    /// </summary>
    private static string CleanText(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var sb = new System.Text.StringBuilder();
        foreach (char c in text)
        {
            // Keep letters, digits, basic punctuation, and whitespace
            if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) || c == '-' || c == '/' || c == '\'')
            {
                sb.Append(c);
            }
            else if (char.IsControl(c))
            {
                // Replace control chars with space
                sb.Append(' ');
            }
            // Skip other weird unicode characters
        }

        // Collapse multiple spaces
        var cleaned = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();

        // Remove repeating phrases
        return RemoveRepeatingPhrases(cleaned);
    }

    /// <summary>
    /// Detect and remove repeating phrases in text.
    /// Handles cases like:
    /// - "Chest Pain >18 Chest Pain >18 Chest Pain >18" -> "Chest Pain >18"
    /// - "A; B A; B B A; B" -> "A; B"
    /// </summary>
    private static string RemoveRepeatingPhrases(string text)
    {
        if (string.IsNullOrEmpty(text) || text.Length < 10)
            return text;

        // Strategy: Find phrases that appear multiple times and keep only one instance
        // Try different phrase lengths, starting with longer (more specific) phrases

        string result = text;
        int maxPhraseLen = text.Length / 2;

        // First pass: find and remove exact duplicate phrases (longer phrases first)
        for (int phraseLen = maxPhraseLen; phraseLen >= 5; phraseLen--)
        {
            // Slide a window to find repeating phrases
            for (int startPos = 0; startPos <= result.Length - phraseLen; startPos++)
            {
                string phrase = result.Substring(startPos, phraseLen).Trim();

                // Skip if phrase is too short after trimming or starts/ends mid-word
                if (phrase.Length < 5) continue;

                // Count occurrences of this phrase
                int count = 0;
                int idx = 0;
                while ((idx = result.IndexOf(phrase, idx, StringComparison.OrdinalIgnoreCase)) >= 0)
                {
                    count++;
                    idx += phrase.Length;
                }

                // If phrase appears 2+ times, keep only first occurrence
                if (count >= 2)
                {
                    // Remove all but first occurrence
                    int firstOccurrence = result.IndexOf(phrase, StringComparison.OrdinalIgnoreCase);
                    string beforeFirst = result.Substring(0, firstOccurrence + phrase.Length);
                    string afterFirst = result.Substring(firstOccurrence + phrase.Length);

                    // Remove subsequent occurrences from afterFirst
                    while (afterFirst.IndexOf(phrase, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        int pos = afterFirst.IndexOf(phrase, StringComparison.OrdinalIgnoreCase);
                        afterFirst = afterFirst.Remove(pos, phrase.Length);
                    }

                    result = beforeFirst + afterFirst;

                    // Clean up multiple spaces
                    result = Regex.Replace(result, @"\s+", " ").Trim();

                    // Restart with the cleaned result
                    break;
                }
            }
        }

        // Second pass: check if whole string is just repetitions (simpler patterns)
        for (int patternLen = 3; patternLen <= result.Length / 2; patternLen++)
        {
            string pattern = result.Substring(0, patternLen);

            // Build expected string if pattern repeats
            var expected = new System.Text.StringBuilder();
            while (expected.Length < result.Length)
            {
                expected.Append(pattern);
            }

            // Compare (allowing for partial at end)
            string expectedStr = expected.ToString().Substring(0, result.Length);
            if (string.Equals(result, expectedStr, StringComparison.OrdinalIgnoreCase))
            {
                return pattern.Trim();
            }
        }

        return result;
    }

    /// <summary>
    /// Extract clinical history from raw Clario scrape text.
    /// Returns text after "CLINICAL HISTORY" up to the next ALL-CAPS section.
    /// Deduplicates if the same lines appear twice.
    /// </summary>
    public static string? ExtractClinicalHistory(string? rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText))
        {
            Services.Logger.Trace("ExtractClinicalHistory: rawText is null or empty");
            return null;
        }

        Services.Logger.Trace($"ExtractClinicalHistory: Input length={rawText.Length}");

        // Check if CLINICAL HISTORY even exists in the text
        if (!rawText.Contains("CLINICAL HISTORY", StringComparison.OrdinalIgnoreCase))
        {
            Services.Logger.Trace("ExtractClinicalHistory: 'CLINICAL HISTORY' not found in text");
            return null;
        }

        // Common section headers that follow CLINICAL HISTORY
        var sectionHeaders = @"TECHNIQUE|FINDINGS|IMPRESSION|COMPARISON|EXAM|PROCEDURE|INDICATION|CONCLUSION|RECOMMENDATION";

        // Find "CLINICAL HISTORY" section - look for content until next section header or end
        var match = Regex.Match(rawText,
            $@"CLINICAL HISTORY[:\s]*\n?(.+?)(?=\n\s*({sectionHeaders})\s*[:\n]|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!match.Success || match.Groups.Count < 2)
        {
            Services.Logger.Trace("ExtractClinicalHistory: Regex did not match");
            return null;
        }

        Services.Logger.Trace($"ExtractClinicalHistory: Regex matched, group1 length={match.Groups[1].Value.Length}");

        var content = match.Groups[1].Value.Trim();

        if (string.IsNullOrWhiteSpace(content))
            return null;

        // Deduplicate: split into lines and remove consecutive duplicates
        var lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var uniqueLines = new System.Collections.Generic.List<string>();

        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
                continue;

            // Check if this line (or very similar) already exists
            bool isDuplicate = false;
            foreach (var existing in uniqueLines)
            {
                if (string.Equals(existing, trimmed, StringComparison.OrdinalIgnoreCase))
                {
                    isDuplicate = true;
                    break;
                }
            }

            if (!isDuplicate)
                uniqueLines.Add(trimmed);
        }

        if (uniqueLines.Count == 0)
            return null;

        // Join and clean the text
        return CleanText(string.Join(" ", uniqueLines));
    }

    public void EnsureOnTop()
    {
        if (InvokeRequired)
        {
            Invoke(EnsureOnTop);
            return;
        }
        NativeWindows.ForceTopMost(this.Handle);
    }

    #region Drag Logic

    private void OnDragStart(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            _dragging = true;
            _dragStart = e.Location;
        }
    }

    private void OnDragMove(object? sender, MouseEventArgs e)
    {
        if (_dragging)
        {
            Location = new Point(
                Location.X + e.X - _dragStart.X,
                Location.Y + e.Y - _dragStart.Y
            );
        }
    }

    private void OnDragEnd(object? sender, MouseEventArgs e)
    {
        _dragging = false;
        // Save center coordinates, not top-left
        _config.ClinicalHistoryX = Location.X + Width / 2;
        _config.ClinicalHistoryY = Location.Y + Height / 2;
        _config.Save();
    }

    #endregion
}
