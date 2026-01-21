using System;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MosaicTools.Services;
using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;

namespace MosaicTools.UI;

/// <summary>
/// Report popup window.
/// Matches Python's ReportPopupWindow.
/// Supports diff highlighting when baseline is provided.
/// </summary>
public class ReportPopupForm : Form
{
    private readonly Configuration _config;
    private readonly RichTextBox _richTextBox;
    private string? _baselineReport;
    private string _currentReportText;

    // Drag state
    private Point _dragStart;
    private bool _dragging;

    // Resize prevention
    private bool _isResizing;

    public ReportPopupForm(Configuration config, string reportText, string? baselineReport = null)
    {
        _config = config;
        _baselineReport = baselineReport;
        _currentReportText = reportText;

        // Form properties
        FormBorderStyle = FormBorderStyle.None;
        Text = "";
        ShowInTaskbar = false;
        TopMost = true;
        DoubleBuffered = true;
        BackColor = Color.FromArgb(30, 30, 30);
        StartPosition = FormStartPosition.Manual;

        AutoSize = false;
        AutoScroll = false;

        int width = _config.ReportPopupWidth < 600 ? 600 : _config.ReportPopupWidth;
        int padding = 20;

        _richTextBox = new RichTextBox
        {
            Width = width - (padding * 2), // Initial width to force wrapping calc
            Location = new Point(padding, padding),
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.White,
            Font = new Font("Consolas", 11),
            Cursor = Cursors.Hand,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.None,
            DetectUrls = false,
            ShortcutsEnabled = false
        };

        // Hook event BEFORE setting text so it fires on assignment
        _richTextBox.ContentsResized += RichTextBox_ContentsResized;

        // 1. Set Base Color (Slightly less bright)
        _richTextBox.ForeColor = Color.Gainsboro; // Off-white/Light Gray

        // 2. Set Text (formatted for display)
        _richTextBox.Text = FormatReportText(reportText);

        Controls.Add(_richTextBox);

        // Initial Layout (fallback)
        this.ClientSize = new Size(width, 200); // Temporary size

        // Initial Location
        Location = new Point(_config.ReportPopupX, _config.ReportPopupY);

        // Config Save
        LocationChanged += (_, _) =>
        {
            _config.ReportPopupX = Location.X;
            _config.ReportPopupY = Location.Y;
        };

        FormClosed += (_, _) => _config.Save();

        // Actions
        SetupInteractions(this);
        SetupInteractions(_richTextBox);

        this.Load += (s,e) =>
        {
            this.Activate();
            ActiveControl = null; // Hide caret

            // Format keywords and diff highlighting AFTER load to ensure layout context
            ApplyFormattingAndDiff();

            // Force Resize Calculation manually
            PerformResize();
        };
    }

    /// <summary>
    /// Update the report text and re-apply diff highlighting.
    /// Called when Process Report is pressed while popup is open.
    /// </summary>
    public void UpdateReport(string newReportText, string? baseline = null)
    {
        _currentReportText = newReportText;
        if (baseline != null)
        {
            _baselineReport = baseline;
        }

        // Reset all formatting first
        _richTextBox.Text = FormatReportText(newReportText);
        _richTextBox.SelectAll();
        _richTextBox.SelectionFont = _richTextBox.Font;
        _richTextBox.SelectionColor = Color.Gainsboro;
        _richTextBox.SelectionBackColor = _richTextBox.BackColor;
        _richTextBox.Select(0, 0);

        // Re-apply formatting and diff
        ApplyFormattingAndDiff();

        // Force resize
        PerformResize();

        Logger.Trace($"ReportPopup updated: {newReportText.Length} chars, baseline={_baselineReport?.Length ?? 0} chars");
    }

    private void ApplyFormattingAndDiff()
    {
        // Format keywords
        FormatKeywords(_richTextBox, new[] { "IMPRESSION:", "FINDINGS:" });

        // Apply diff highlighting if baseline provided
        if (!string.IsNullOrEmpty(_baselineReport))
        {
            ApplyDiffHighlighting(_richTextBox, _baselineReport, _currentReportText);
        }
    }

    private void FormatKeywords(RichTextBox rtb, string[] keywords)
    {
        // Define Highlight Style
        // +2 points largest
        float baseSize = rtb.Font.Size;
        Font highlightFont = new Font(rtb.Font.FontFamily, baseSize + 2, FontStyle.Bold);
        Color highlightColor = Color.White;

        foreach (var word in keywords)
        {
            int startIndex = 0;
            while (startIndex < rtb.TextLength)
            {
                // Find next instance (Case Sensitive + Whole Word)
                int foundIndex = rtb.Find(word, startIndex, RichTextBoxFinds.WholeWord | RichTextBoxFinds.MatchCase);
                if (foundIndex == -1) break;

                // Select and Format
                rtb.Select(foundIndex, word.Length);
                rtb.SelectionColor = highlightColor;
                rtb.SelectionFont = highlightFont;

                // Move past this instance
                startIndex = foundIndex + word.Length;
            }
        }

        // Reset selection
        rtb.Select(0, 0);
    }

    private void ApplyDiffHighlighting(RichTextBox rtb, string baseline, string current)
    {
        try
        {
            // Parse the highlight color from config and blend with background for subtlety
            Color highlightColor;
            try
            {
                var baseColor = ColorTranslator.FromHtml(_config.ReportChangesColor);
                // Blend with dark background using configurable opacity
                var bg = BackColor;
                float alpha = _config.ReportChangesAlpha / 100f;
                highlightColor = Color.FromArgb(
                    (int)(bg.R + (baseColor.R - bg.R) * alpha),
                    (int)(bg.G + (baseColor.G - bg.G) * alpha),
                    (int)(bg.B + (baseColor.B - bg.B) * alpha)
                );
            }
            catch
            {
                // Subtle fallback - very dark green tint
                highlightColor = Color.FromArgb(58, 82, 58);
            }

            // Sentence-level diff: find sentences in current that don't exist in baseline
            var baselineSentences = SplitIntoSentences(baseline);
            var currentSentences = SplitIntoSentences(current);

            // Create set of normalized baseline sentences for lookup
            var baselineSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in baselineSentences)
            {
                var normalized = NormalizeSentence(s);
                if (!string.IsNullOrWhiteSpace(normalized))
                    baselineSet.Add(normalized);
            }

            // Find new sentences (in current but not in baseline)
            // Strip section heading prefixes for comparison
            var newSentences = new List<(string original, string content)>();
            foreach (var sentence in currentSentences)
            {
                var content = StripSectionHeading(sentence);
                var normalized = NormalizeSentence(content);
                if (!string.IsNullOrWhiteSpace(normalized) && !baselineSet.Contains(normalized))
                {
                    newSentences.Add((sentence.Trim(), content.Trim()));
                }
            }

            // Also add stripped versions of baseline sentences for comparison
            foreach (var s in baselineSentences)
            {
                var content = StripSectionHeading(s);
                var normalized = NormalizeSentence(content);
                if (!string.IsNullOrWhiteSpace(normalized))
                    baselineSet.Add(normalized);
            }

            // Re-filter with stripped baseline
            newSentences = newSentences.Where(x =>
                !baselineSet.Contains(NormalizeSentence(x.content))).ToList();

            // Highlight new sentences in the RTB
            string rtbText = rtb.Text;
            int highlightCount = 0;

            foreach (var (original, content) in newSentences)
            {
                if (string.IsNullOrWhiteSpace(content)) continue;
                if (content.Length < 3) continue; // Skip very short
                if (content.EndsWith(":")) continue; // Skip section headers
                if (IsSectionHeading(content)) continue; // Skip capitalized section headings

                // Find the content part in RTB text (not the heading prefix)
                int idx = rtbText.IndexOf(content, StringComparison.Ordinal);
                if (idx >= 0)
                {
                    rtb.Select(idx, content.Length);
                    rtb.SelectionBackColor = highlightColor;
                    highlightCount++;
                }
            }

            // Reset selection
            rtb.Select(0, 0);

            Logger.Trace($"Applied diff highlighting: {highlightCount} new sentences highlighted (from {newSentences.Count} detected)");
        }
        catch (Exception ex)
        {
            Logger.Trace($"Diff highlighting error: {ex.Message}");
        }
    }

    private static List<string> SplitIntoSentences(string text)
    {
        var sentences = new List<string>();
        if (string.IsNullOrEmpty(text)) return sentences;

        // Normalize line endings
        text = text.Replace("\r\n", "\n").Replace("\r", "\n");

        var current = new System.Text.StringBuilder();

        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            current.Append(c);

            // Sentence ends at . ! ? followed by whitespace or end of string
            bool isSentenceEnd = (c == '.' || c == '!' || c == '?') &&
                                 (i == text.Length - 1 || char.IsWhiteSpace(text[i + 1]));

            // Also split on newlines to handle list items
            bool isNewline = c == '\n';

            if (isSentenceEnd || isNewline)
            {
                var sentence = current.ToString().Trim();
                if (!string.IsNullOrWhiteSpace(sentence))
                {
                    sentences.Add(sentence);
                }
                current.Clear();
            }
        }

        // Add remaining text
        var remaining = current.ToString().Trim();
        if (!string.IsNullOrWhiteSpace(remaining))
        {
            sentences.Add(remaining);
        }

        return sentences;
    }

    private static string NormalizeSentence(string sentence)
    {
        if (string.IsNullOrEmpty(sentence)) return "";
        // Lowercase and collapse whitespace for comparison
        return System.Text.RegularExpressions.Regex.Replace(
            sentence.ToLowerInvariant().Trim(), @"\s+", " ");
    }

    /// <summary>
    /// Check if text is a section heading (mostly uppercase, may end with colon).
    /// Examples: "FINDINGS:", "LUNGS AND PLEURA:", "IMPRESSION:"
    /// </summary>
    private static bool IsSectionHeading(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;

        var trimmed = text.Trim().TrimEnd(':');
        if (trimmed.Length < 2) return false;

        // Count uppercase letters vs total letters
        int upperCount = 0;
        int letterCount = 0;
        foreach (char c in trimmed)
        {
            if (char.IsLetter(c))
            {
                letterCount++;
                if (char.IsUpper(c)) upperCount++;
            }
        }

        // If mostly uppercase (>80%), it's likely a section heading
        return letterCount > 0 && (double)upperCount / letterCount > 0.8;
    }

    /// <summary>
    /// Strip section heading prefix from a sentence.
    /// "LUNGS AND PLEURA: Some text here." -> "Some text here."
    /// </summary>
    private static string StripSectionHeading(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return text;

        var trimmed = text.Trim();
        int colonIdx = trimmed.IndexOf(':');
        if (colonIdx < 0 || colonIdx >= trimmed.Length - 1) return trimmed;

        // Check if text before colon is a section heading (mostly uppercase)
        var prefix = trimmed.Substring(0, colonIdx);
        if (IsSectionHeading(prefix + ":"))
        {
            // Return everything after the colon, trimmed
            return trimmed.Substring(colonIdx + 1).Trim();
        }

        return trimmed;
    }

    private void RichTextBox_ContentsResized(object? sender, ContentsResizedEventArgs e)
    {
        if (_isResizing) return;

        // CRASH FIX: Check if handle is created.
        if (!this.IsHandleCreated) return; // Ignore event if not ready

        // Use BeginInvoke to decouple from the layout cycle
        this.BeginInvoke(new Action(() =>
        {
            if (_isResizing) return;
            _isResizing = true;
            try
            {
                PerformResize();
            }
            finally
            {
                _isResizing = false;
            }
        }));
    }

    private void PerformResize()
    {
        // AUTHORITY: Ask the control for the position of the last character
        int contentHeight = 0;
        int textLength = _richTextBox.TextLength;

        if (textLength > 0)
        {
            // Get pixel position of last char
            Point pt = _richTextBox.GetPositionFromCharIndex(textLength - 1);
            // Height = Y pos + Line Height + Padding
            // Use Font.Height for the last line height approx (even if bolded, it's safer)
            // Getting specific line height is harder, but this is usually close enough.
            // We use the Base Font height + a buffer.
            contentHeight = pt.Y + _richTextBox.Font.Height + 10;
        }
        else
        {
            contentHeight = _richTextBox.Font.Height;
        }

        int padding = 20;
        int requiredInnerHeight = contentHeight;
        int requiredTotalHeight = requiredInnerHeight + (padding * 2);

        // Safety check for empty/small results
        if (requiredTotalHeight < 100) requiredTotalHeight = 100;

        // Get Max Height
        Rectangle workArea = Screen.FromControl(this).WorkingArea;
        int maxHeight = workArea.Height - 50;

        int finalFormHeight;
        bool needsScroll;

        if (requiredTotalHeight > maxHeight)
        {
            finalFormHeight = maxHeight;
            needsScroll = true;
        }
        else
        {
            finalFormHeight = requiredTotalHeight;
            needsScroll = false;
        }

        // Apply Scrollbars FIRST
        if (needsScroll && _richTextBox.ScrollBars != RichTextBoxScrollBars.Vertical)
        {
            _richTextBox.ScrollBars = RichTextBoxScrollBars.Vertical;
        }
        else if (!needsScroll && _richTextBox.ScrollBars != RichTextBoxScrollBars.None)
        {
            _richTextBox.ScrollBars = RichTextBoxScrollBars.None;
        }

        // Apply Form Size
        if (this.ClientSize.Height != finalFormHeight)
        {
            this.ClientSize = new Size(this.ClientSize.Width, finalFormHeight);
        }

        // Apply RTB Size
        int rtbHeight = finalFormHeight - (padding * 2);
        if (_richTextBox.Height != rtbHeight)
        {
            _richTextBox.Height = rtbHeight;
        }
    }

    private void SetupInteractions(Control control)
    {
        Point formPosOnMouseDown = Point.Empty;

        control.MouseDown += (s, e) =>
        {
            if (e.Button == MouseButtons.Left)
            {
                formPosOnMouseDown = this.Location;

                // Capture offset relative to Form
                _dragStart = new Point(e.X, e.Y);

                // If dragging via RTB, defer to OS Drag Logic
                if (control != this) {
                    NativeMethods.ReleaseCapture();
                    NativeMethods.SendMessage(this.Handle, NativeMethods.WM_NCLBUTTONDOWN, NativeMethods.HT_CAPTION, 0);

                    // After WM_NCLBUTTONDOWN returns (drag ended), check if form moved
                    // If it didn't move, it was a click - close the form
                    if (this.Location == formPosOnMouseDown)
                    {
                        Close();
                    }
                    _dragging = false;
                }
                else
                {
                    _dragging = true;
                }
            }
            if (e.Button == MouseButtons.Right) Close();
        };

        control.MouseMove += (s, e) =>
        {
            if (_dragging)
            {
                Point currentScreenPos = Cursor.Position;
                Location = new Point(
                    currentScreenPos.X - _dragStart.X,
                    currentScreenPos.Y - _dragStart.Y
                );
            }
        };

        control.MouseUp += (s, e) =>
        {
            if (_dragging && e.Button == MouseButtons.Left)
            {
                // Check if form moved - if not, it was a click
                if (this.Location == formPosOnMouseDown)
                {
                    Close();
                }
            }
            _dragging = false;
        };
    }

    /// <summary>
    /// Format scraped report text for display.
    /// - Single blank line between major sections
    /// - Content in most sections joined into paragraphs
    /// - FINDINGS: subsections on own lines with blank lines between, sentences joined
    /// - IMPRESSION: numbered items on own lines
    /// </summary>
    private static string FormatReportText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return text;

        // Normalize line endings
        text = text.Replace("\r\n", "\n").Replace("\r", "\n");

        var lines = text.Split('\n');
        var outputLines = new List<string>();

        // Major section headers
        var majorSections = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "EXAM:", "COMPARISON:", "CLINICAL HISTORY:", "FINDINGS:", "IMPRESSION:",
            "TECHNIQUE:", "INDICATION:", "PROCEDURE:", "CONCLUSION:", "RECOMMENDATION:"
        };

        string currentSection = "";
        var pendingContent = new List<string>();

        void FlushContent(bool asJoined = true)
        {
            if (pendingContent.Count > 0)
            {
                if (asJoined)
                    outputLines.Add(string.Join(" ", pendingContent));
                else
                    outputLines.AddRange(pendingContent);
                pendingContent.Clear();
            }
        }

        foreach (var rawLine in lines)
        {
            // Normalize whitespace within the line
            var line = Regex.Replace(rawLine.Trim(), @"\s+", " ");

            // Skip empty lines from input
            if (string.IsNullOrWhiteSpace(line))
                continue;

            bool isMajorSection = majorSections.Contains(line);
            bool isSubsectionHeader = !isMajorSection && IsSubsectionHeader(line);
            bool isNumberedItem = Regex.IsMatch(line, @"^\d+\.");

            if (isMajorSection)
            {
                FlushContent();

                // Blank line before major sections (except first)
                if (outputLines.Count > 0)
                    outputLines.Add("");

                outputLines.Add(line);
                currentSection = line.TrimEnd(':').ToUpperInvariant();
            }
            else if (currentSection == "FINDINGS")
            {
                if (isSubsectionHeader)
                {
                    // Flush previous subsection
                    FlushContent();

                    // Blank line before new subsection (but not right after FINDINGS:)
                    if (outputLines.Count > 0 && outputLines[outputLines.Count - 1] != "FINDINGS:")
                        outputLines.Add("");

                    pendingContent.Add(line);
                }
                else
                {
                    // Continue accumulating content (either in subsection or standalone)
                    pendingContent.Add(line);
                }
            }
            else if (currentSection == "IMPRESSION")
            {
                if (isNumberedItem)
                {
                    FlushContent();
                    outputLines.Add(line);
                }
                else
                {
                    // Non-numbered content in impression - accumulate
                    pendingContent.Add(line);
                }
            }
            else
            {
                // All other sections: accumulate content to join as paragraph
                pendingContent.Add(line);
            }
        }

        FlushContent();

        return string.Join(Environment.NewLine, outputLines);
    }

    /// <summary>
    /// Check if a line is a subsection header (ALL CAPS text followed by colon and content).
    /// Examples: "LUNGS AND PLEURA: Some text", "HEART AND MEDIASTINUM: More text"
    /// </summary>
    private static bool IsSubsectionHeader(string line)
    {
        // Must contain a colon
        int colonIdx = line.IndexOf(':');
        if (colonIdx <= 0) return false;

        // Get the part before the colon
        string prefix = line.Substring(0, colonIdx);

        // Must be at least 2 chars and mostly uppercase
        if (prefix.Length < 2) return false;

        int upperCount = 0;
        int letterCount = 0;
        foreach (char c in prefix)
        {
            if (char.IsLetter(c))
            {
                letterCount++;
                if (char.IsUpper(c)) upperCount++;
            }
        }

        // Consider it a subsection if >80% uppercase
        return letterCount > 0 && (double)upperCount / letterCount > 0.8;
    }

    private static class NativeMethods
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
    }
}
