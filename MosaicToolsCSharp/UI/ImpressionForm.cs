using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Floating window displaying the Impression from the report.
/// Shown after ProcessReport, hidden on SignReport.
/// </summary>
public class ImpressionForm : Form
{
    private readonly Configuration _config;
    private readonly Label _contentLabel;

    // Drag state
    private Point _dragStart;
    private bool _dragging;

    // For center-based positioning
    private bool _initialPositionSet = false;

    public ImpressionForm(Configuration config)
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

        // Content label - bold, allows multiple lines
        _contentLabel = new Label
        {
            Text = "(Waiting for impression...)",
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(200, 200, 200),
            BackColor = Color.Black,
            AutoSize = true,
            Margin = new Padding(10, 5, 10, 15),
            MaximumSize = new Size(500, 0) // Limit width, allow height to grow
        };
        layout.Controls.Add(_contentLabel, 0, 1);

        // Make form auto-size to content
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;

        // Context menu for drag bar only (Close option)
        var menu = new ContextMenuStrip();
        menu.Items.Add("Close", null, (_, _) => Close());
        dragBar.ContextMenuStrip = menu;

        // Left click to dismiss, right click to copy debug
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
            _config.ImpressionX - Width / 2,
            _config.ImpressionY - Height / 2
        );
        _initialPositionSet = true;
    }

    /// <summary>
    /// When size changes, keep the center in the same place.
    /// </summary>
    private void RepositionToCenter()
    {
        if (_dragging) return;
        Location = new Point(
            _config.ImpressionX - Width / 2,
            _config.ImpressionY - Height / 2
        );
    }

    /// <summary>
    /// Update the displayed impression text.
    /// </summary>
    public void SetImpression(string? text)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetImpression(text));
            return;
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            _contentLabel.Text = "(Waiting for impression...)";
            _contentLabel.ForeColor = Color.FromArgb(128, 128, 128);
        }
        else
        {
            _contentLabel.Text = text;
            _contentLabel.ForeColor = Color.FromArgb(200, 200, 200);
        }
    }

    /// <summary>
    /// Extract impression from raw report text.
    /// Returns text after "IMPRESSION" up to the next ALL-CAPS section or end.
    /// </summary>
    public static string? ExtractImpression(string? rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText))
            return null;

        // Check if IMPRESSION exists in the text
        if (!rawText.Contains("IMPRESSION", StringComparison.OrdinalIgnoreCase))
            return null;

        // Common section headers that might follow IMPRESSION
        var sectionHeaders = @"TECHNIQUE|FINDINGS|CLINICAL HISTORY|COMPARISON|EXAM|PROCEDURE|INDICATION|CONCLUSION|RECOMMENDATION|SIGNATURE|ELECTRONICALLY SIGNED";

        // Find "IMPRESSION" section - look for content until next section header or end
        var match = Regex.Match(rawText,
            $@"IMPRESSION[:\s]*\n?(.+?)(?=\n\s*({sectionHeaders})\s*[:\n]|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!match.Success || match.Groups.Count < 2)
            return null;

        var content = match.Groups[1].Value.Trim();

        if (string.IsNullOrWhiteSpace(content))
            return null;

        // Collapse whitespace to single spaces, clean special characters
        content = Regex.Replace(content, @"[\r\n\t ]+", " ").Trim();
        content = CleanText(content);

        // Add line breaks before consecutive numbered items (1., 2., 3., etc.)
        content = FormatNumberedItems(content);

        return content;
    }

    /// <summary>
    /// Insert line breaks before consecutively numbered items.
    /// Only breaks on expected next number to avoid breaking on things like "2.5 cm".
    /// </summary>
    private static string FormatNumberedItems(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // Check if text starts with "1."
        if (!Regex.IsMatch(text, @"^\s*1\."))
            return text;

        var result = new System.Text.StringBuilder();
        int expectedNumber = 1;
        int i = 0;

        while (i < text.Length)
        {
            // Look for the expected number pattern (e.g., "2." when expecting 2)
            string nextPattern = $"{expectedNumber}.";

            if (i + nextPattern.Length <= text.Length)
            {
                string substring = text.Substring(i, nextPattern.Length);

                if (substring == nextPattern)
                {
                    // Check that it's followed by a space or letter (not another digit like "2.5")
                    bool isValidNumberedItem = (i + nextPattern.Length >= text.Length) ||
                        char.IsWhiteSpace(text[i + nextPattern.Length]) ||
                        char.IsLetter(text[i + nextPattern.Length]);

                    if (isValidNumberedItem)
                    {
                        // Add newline before this number (except for first item)
                        if (expectedNumber > 1)
                        {
                            result.Append('\n');
                        }
                        result.Append(nextPattern);
                        i += nextPattern.Length;
                        expectedNumber++;
                        continue;
                    }
                }
            }

            result.Append(text[i]);
            i++;
        }

        return result.ToString();
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
            // Keep letters, digits, basic punctuation, and spaces
            if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || c == ' ' || c == '-' || c == '/' || c == '\'')
            {
                sb.Append(c);
            }
            else if (char.IsWhiteSpace(c) || char.IsControl(c))
            {
                sb.Append(' ');
            }
            // Skip other weird unicode characters
        }

        // Collapse multiple spaces
        return Regex.Replace(sb.ToString(), @" +", " ").Trim();
    }

    /// <summary>
    /// Handle mouse clicks on content label - left to dismiss, right to copy debug.
    /// </summary>
    private void OnContentLabelMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            Close();
        }
        else if (e.Button == MouseButtons.Right)
        {
            CopyDebugInfoToClipboard();
        }
    }

    /// <summary>
    /// Copy impression debug info to clipboard.
    /// </summary>
    private void CopyDebugInfoToClipboard()
    {
        var debugInfo = $"=== Impression Debug ===\r\n" +
                        $"Content: {_contentLabel.Text}";

        try
        {
            Clipboard.SetText(debugInfo);
            Services.Logger.Trace("Impression debug info copied to clipboard");
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
        _config.ImpressionX = Location.X + Width / 2;
        _config.ImpressionY = Location.Y + Height / 2;
        _config.Save();
    }

    #endregion
}
