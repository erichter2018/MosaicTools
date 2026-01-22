using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Alert types for the notification box (alerts-only mode).
/// </summary>
public enum AlertType
{
    TemplateMismatch,
    GenderMismatch,
    StrokeDetected
}

/// <summary>
/// Floating window displaying clinical history from Clario scrape.
/// Can also operate in "alerts only" mode where it only appears when alerts trigger.
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

    // Border colors - Priority order: Green (drafted) < Purple (stroke) < Flashing Red (gender) < Solid Red (template mismatch)
    private static readonly Color NormalBorderColor = Color.FromArgb(60, 60, 60);
    private static readonly Color DraftedBorderColor = Color.FromArgb(0, 180, 0); // Green - drafted status
    private static readonly Color StrokeBorderColor = Color.FromArgb(140, 80, 200); // Purple - stroke case
    private static readonly Color TemplateMismatchBorderColor = Color.FromArgb(220, 50, 50); // Red - template mismatch

    // Text colors
    private static readonly Color NormalTextColor = Color.FromArgb(200, 200, 200); // White-ish
    private static readonly Color FixedTextColor = Color.FromArgb(255, 255, 120); // Light yellow
    private static readonly Color EmptyTextColor = Color.FromArgb(128, 128, 128); // Gray

    // Template matching state for debug
    private string? _lastDescription;
    private string? _lastTemplateName;
    private bool _templateMismatch = false;

    // Auto-fix tracking - stores the accession we've already auto-fixed to prevent loops
    private string? _lastAutoFixedAccession;
    private DateTime _lastAutoFixTime = DateTime.MinValue;

    // Track displayed text and whether it was "fixed" from original
    private string? _currentDisplayedText;
    private bool _currentTextWasFixed = false;

    // Gender check - terms that are impossible for the opposite gender
    // Female-only terms (flag if patient is Male)
    private static readonly string[] FemaleOnlyTerms = {
        "uterus", "uterine", "ovary", "ovaries", "ovarian",
        "fallopian", "endometrium", "endometrial",
        "vagina", "vaginal", "vulva", "vulvar",
        "adnexa", "adnexal", "pregnancy", "pregnant", "gravid", "gestational",
        "placenta", "placental", "fetus", "fetal"
    };

    // Male-only terms (flag if patient is Female)
    private static readonly string[] MaleOnlyTerms = {
        "prostate", "prostatic", "seminal vesicle", "seminal vesicles",
        "testicle", "testis", "testes", "testicular",
        "scrotum", "scrotal", "epididymis", "epididymal",
        "spermatic cord", "penis", "penile", "vas deferens"
    };

    // Gender warning state
    private bool _genderWarningActive = false;
    private string? _genderWarningText;
    private string? _savedClinicalHistoryText;

    // Stroke detection state
    private bool _strokeDetected = false;
    private System.Windows.Forms.Timer? _blinkTimer;
    private bool _blinkState = false;

    // Alert-only mode state (when AlwaysShowClinicalHistory is false)
    private bool _showingAlert = false;
    private AlertType? _currentAlertType = null;

    // Tooltip for border color explanation
    private readonly ToolTip _borderTooltip;

    // Callback for stroke note creation (set by MainForm)
    private Func<bool>? _onStrokeNoteClick;

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

        // Tooltip for border color explanation
        _borderTooltip = new ToolTip
        {
            InitialDelay = 500,
            AutoPopDelay = 10000,
            ReshowDelay = 200
        };
        UpdateBorderTooltip();
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
    public void SetClinicalHistory(string? text, bool wasFixed = false)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetClinicalHistory(text, wasFixed));
            return;
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            // If text is null, we couldn't find CLINICAL HISTORY section - persist old text
            // If text is empty string, section exists but is empty - show "(No clinical history)"
            if (text == null && !string.IsNullOrWhiteSpace(_currentDisplayedText))
                return;

            _contentLabel.Text = "(No clinical history)";
            _contentLabel.ForeColor = EmptyTextColor;
            _currentDisplayedText = null;
            _currentTextWasFixed = false;
        }
        else
        {
            _contentLabel.Text = text;
            _currentDisplayedText = text;
            _currentTextWasFixed = wasFixed;
            // Set color based on fixed state - will be updated by UpdateTextColorFromFinalReport
            _contentLabel.ForeColor = wasFixed ? FixedTextColor : NormalTextColor;
        }
    }

    /// <summary>
    /// Set clinical history with auto-fix support.
    /// If auto-fix is enabled and the text was modified (malformed), automatically paste to Mosaic.
    /// </summary>
    public void SetClinicalHistoryWithAutoFix(string? preCleaned, string? cleaned, string? accession = null)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetClinicalHistoryWithAutoFix(preCleaned, cleaned, accession));
            return;
        }

        // Check if text was actually fixed (preCleaned differs from cleaned)
        bool wasFixed = !string.IsNullOrWhiteSpace(preCleaned) &&
                        !string.IsNullOrWhiteSpace(cleaned) &&
                        !string.Equals(preCleaned, cleaned, StringComparison.Ordinal);

        // Always update the display with fixed state
        SetClinicalHistory(cleaned, wasFixed);

        // Check if auto-fix should trigger
        if (!_config.AutoFixClinicalHistory)
        {
            Logger.Trace("Auto-fix: disabled in config");
            return;
        }

        if (string.IsNullOrWhiteSpace(cleaned))
        {
            Logger.Trace("Auto-fix: cleaned is empty");
            return;
        }

        Logger.Trace($"Auto-fix check: wasFixed={wasFixed}, accession='{accession}', preCleaned='{preCleaned?.Substring(0, Math.Min(50, preCleaned?.Length ?? 0))}...'");

        if (!wasFixed)
            return;

        // Prevent loops: don't re-fix the same accession
        if (!string.IsNullOrEmpty(accession) && string.Equals(_lastAutoFixedAccession, accession, StringComparison.Ordinal))
        {
            Logger.Trace($"Auto-fix: already fixed accession {accession}");
            return;
        }

        // All conditions met - trigger auto-fix
        Logger.Trace($"Auto-fix triggered for accession '{accession}': preCleaned='{preCleaned}' -> cleaned='{cleaned}'");
        _lastAutoFixedAccession = accession;
        _lastAutoFixTime = DateTime.Now;

        // Trigger the paste (this runs on background thread)
        PasteClinicalHistoryToMosaic();
    }

    /// <summary>
    /// Reset auto-fix tracking. Call when study changes to allow re-fixing.
    /// </summary>
    public void ResetAutoFixTracking()
    {
        _lastAutoFixedAccession = null;
    }

    /// <summary>
    /// Called when study changes (accession changes). Resets all tracking state and clears display.
    /// </summary>
    /// <param name="isNewStudy">True if switching to a new study (show "Loading..."), false if no study (show "No clinical history")</param>
    public void OnStudyChanged(bool isNewStudy = true)
    {
        if (InvokeRequired)
        {
            Invoke(() => OnStudyChanged(isNewStudy));
            return;
        }

        Logger.Trace($"ClinicalHistoryForm: Study changed (isNewStudy={isNewStudy}) - resetting state");

        // Reset auto-fix tracking
        _lastAutoFixedAccession = null;
        _lastAutoFixTime = DateTime.MinValue;

        // Reset template mismatch and stroke state
        _templateMismatch = false;
        _strokeDetected = false;
        _lastDescription = null;
        _lastTemplateName = null;
        BackColor = NormalBorderColor;
        UpdateBorderTooltip();

        // Clear displayed text immediately so old study's history doesn't persist
        _contentLabel.Text = isNewStudy ? "(Loading...)" : "(No clinical history)";
        _contentLabel.ForeColor = EmptyTextColor;
        _currentDisplayedText = null;
        _currentTextWasFixed = false;
    }

    /// <summary>
    /// Update text color based on whether displayed history matches the final report.
    /// If they match, text is white. If different (we fixed it), text is yellow.
    /// </summary>
    public void UpdateTextColorFromFinalReport(string? finalReportText)
    {
        if (InvokeRequired)
        {
            Invoke(() => UpdateTextColorFromFinalReport(finalReportText));
            return;
        }

        // If no displayed text or it wasn't fixed, nothing to do
        if (string.IsNullOrWhiteSpace(_currentDisplayedText))
            return;

        // If it wasn't marked as fixed, keep normal color
        if (!_currentTextWasFixed)
        {
            _contentLabel.ForeColor = NormalTextColor;
            return;
        }

        // Extract clinical history from final report
        var reportHistory = ExtractClinicalHistoryFromReport(finalReportText);

        // Compare with displayed text (normalize for comparison)
        bool matches = CompareHistoryText(_currentDisplayedText, reportHistory);

        if (matches)
        {
            // Final report now has our fixed text - show white
            _contentLabel.ForeColor = NormalTextColor;
            Logger.Trace("Clinical history matches final report - text color: white");
        }
        else
        {
            // Still different - show yellow
            _contentLabel.ForeColor = FixedTextColor;
        }
    }

    /// <summary>
    /// Extract clinical history text from a final report.
    /// </summary>
    private static string? ExtractClinicalHistoryFromReport(string? reportText)
    {
        if (string.IsNullOrWhiteSpace(reportText))
            return null;

        // Look for CLINICAL HISTORY section
        var match = Regex.Match(reportText,
            @"CLINICAL HISTORY[:\s]*(.*?)(?=\b(?:TECHNIQUE|FINDINGS|IMPRESSION|COMPARISON|PROCEDURE|CONCLUSION|RECOMMENDATION)\b|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!match.Success || match.Groups.Count < 2)
            return null;

        return match.Groups[1].Value.Trim();
    }

    /// <summary>
    /// Compare two clinical history texts, normalizing whitespace and punctuation.
    /// </summary>
    private static bool CompareHistoryText(string? text1, string? text2)
    {
        if (string.IsNullOrWhiteSpace(text1) || string.IsNullOrWhiteSpace(text2))
            return false;

        // Normalize: lowercase, collapse whitespace, remove trailing punctuation
        var norm1 = NormalizeForComparison(text1);
        var norm2 = NormalizeForComparison(text2);

        return string.Equals(norm1, norm2, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Normalize text for comparison - collapse whitespace, remove trailing punctuation.
    /// </summary>
    private static string NormalizeForComparison(string text)
    {
        // Collapse whitespace
        var result = Regex.Replace(text, @"\s+", " ").Trim();
        // Remove trailing period if present
        result = result.TrimEnd('.', ',', ';', ' ');
        return result;
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

        // Priority order: Green (drafted) < Purple (stroke) < Flashing Red (gender) < Solid Red (template mismatch)
        // Note: Flashing red (gender warning) is handled by UpdateBlinkDisplay and overrides this
        if (_templateMismatch)
        {
            // Solid red - highest priority (except gender which handles itself)
            BackColor = TemplateMismatchBorderColor;
        }
        else if (_strokeDetected)
        {
            // Purple - stroke case
            BackColor = StrokeBorderColor;
        }
        else
        {
            // Green (drafted) or normal
            BackColor = isDrafted ? DraftedBorderColor : NormalBorderColor;
        }
        UpdateBorderTooltip();
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
            UpdateBorderTooltip();
        }
        // If not mismatch, the border will be set by SetDraftedState
    }

    /// <summary>
    /// Set the stroke detection state - shows purple border when true.
    /// Overrides green (drafted) but is overridden by red (template mismatch) and flashing red (gender).
    /// </summary>
    public void SetStrokeState(bool isStroke)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetStrokeState(isStroke));
            return;
        }

        _strokeDetected = isStroke;

        // Update border color immediately if stroke detected and no higher priority state
        if (isStroke && !_templateMismatch && !_genderWarningActive)
        {
            BackColor = StrokeBorderColor;
        }
        UpdateBorderTooltip();
    }

    /// <summary>
    /// Returns whether stroke is currently detected.
    /// </summary>
    public bool IsStrokeDetected => _strokeDetected;

    /// <summary>
    /// Set the callback for creating a stroke critical note.
    /// Called when user clicks the notification box during stroke case.
    /// </summary>
    public void SetStrokeNoteClickCallback(Func<bool> callback)
    {
        _onStrokeNoteClick = callback;
    }

    /// <summary>
    /// Returns whether currently showing an alert (as opposed to clinical history).
    /// </summary>
    public bool IsShowingAlert => _showingAlert;

    /// <summary>
    /// Show an alert in the notification box (for alerts-only mode).
    /// This replaces the clinical history text with alert content.
    /// </summary>
    public void ShowAlertOnly(AlertType type, string details)
    {
        if (InvokeRequired)
        {
            Invoke(() => ShowAlertOnly(type, details));
            return;
        }

        _showingAlert = true;
        _currentAlertType = type;

        // Set appropriate border color and text based on alert type
        switch (type)
        {
            case AlertType.TemplateMismatch:
                BackColor = TemplateMismatchBorderColor;
                _contentLabel.BackColor = Color.Black;
                _contentLabel.ForeColor = Color.FromArgb(255, 150, 150); // Light red
                _contentLabel.Text = $"TEMPLATE MISMATCH\n{details}";
                break;

            case AlertType.GenderMismatch:
                // Gender mismatch uses the blinking red (handled by SetGenderWarning)
                // This is called for completeness but SetGenderWarning handles the display
                break;

            case AlertType.StrokeDetected:
                BackColor = StrokeBorderColor;
                _contentLabel.BackColor = Color.Black;
                _contentLabel.ForeColor = Color.FromArgb(200, 150, 255); // Light purple
                _contentLabel.Text = $"STROKE PROTOCOL\n{details}";
                break;
        }

        UpdateBorderTooltip();
    }

    /// <summary>
    /// Clear alert display and return to clinical history mode.
    /// </summary>
    public void ClearAlert()
    {
        if (InvokeRequired)
        {
            Invoke(ClearAlert);
            return;
        }

        _showingAlert = false;
        _currentAlertType = null;

        // Restore clinical history display
        if (!string.IsNullOrEmpty(_currentDisplayedText))
        {
            _contentLabel.Text = _currentDisplayedText;
            _contentLabel.ForeColor = _currentTextWasFixed ? FixedTextColor : NormalTextColor;
            _contentLabel.BackColor = Color.Black;
        }
        else
        {
            _contentLabel.Text = "(No clinical history)";
            _contentLabel.ForeColor = EmptyTextColor;
            _contentLabel.BackColor = Color.Black;
        }

        BackColor = NormalBorderColor;
        UpdateBorderTooltip();
    }

    /// <summary>
    /// Get alert text for the given alert type and current state.
    /// Used by ActionController to get alert details for display.
    /// </summary>
    public string GetAlertText(AlertType type)
    {
        switch (type)
        {
            case AlertType.TemplateMismatch:
                if (_lastDescription != null && _lastTemplateName != null)
                    return $"Study: {_lastDescription}\nTemplate: {_lastTemplateName}";
                return "Study/template mismatch detected";

            case AlertType.StrokeDetected:
                return "Study flagged as stroke protocol";

            case AlertType.GenderMismatch:
                return _genderWarningText ?? "Gender mismatch detected";

            default:
                return "";
        }
    }

    /// <summary>
    /// Update the tooltip to explain the current border color.
    /// </summary>
    private void UpdateBorderTooltip()
    {
        string tooltip;

        if (_genderWarningActive)
        {
            tooltip = "Red (flashing): Gender mismatch - report contains terms that don't match patient gender";
        }
        else if (_templateMismatch)
        {
            tooltip = "Red: Template mismatch - study description doesn't match the report template";
        }
        else if (_strokeDetected)
        {
            if (_config.StrokeClickToCreateNote)
            {
                tooltip = "Purple: Stroke protocol - click to create critical note";
            }
            else
            {
                tooltip = "Purple: Stroke protocol detected";
            }
        }
        else if (BackColor == DraftedBorderColor)
        {
            tooltip = "Green: Report is drafted";
        }
        else
        {
            tooltip = "Gray: Normal - no alerts";
        }

        _borderTooltip.SetToolTip(this, tooltip);
        _borderTooltip.SetToolTip(_contentLabel, tooltip);
    }

    /// <summary>
    /// Handle mouse clicks on content label.
    /// Left-click: paste clinical history to Mosaic (or create stroke note if enabled)
    /// Right-click: copy debug info to clipboard
    /// </summary>
    private void OnContentLabelMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            Services.Logger.Trace($"ClinicalHistory click: strokeDetected={_strokeDetected}, clickToCreate={_config.StrokeClickToCreateNote}, callbackSet={_onStrokeNoteClick != null}");

            // If stroke detected and click-to-create enabled, create note instead of paste
            if (_strokeDetected && _config.StrokeClickToCreateNote && _onStrokeNoteClick != null)
            {
                Services.Logger.Trace("ClinicalHistory: Invoking stroke note callback");
                bool created = _onStrokeNoteClick.Invoke();
                Services.Logger.Trace($"ClinicalHistory: Callback returned created={created}");
                if (!created)
                {
                    // Note already exists - show tooltip briefly
                    ShowAlreadyCreatedTooltip();
                }
                return;
            }

            // Default behavior: paste clinical history
            PasteClinicalHistoryToMosaic();
        }
        else if (e.Button == MouseButtons.Right)
        {
            CopyDebugInfoToClipboard();
        }
    }

    /// <summary>
    /// Show a tooltip indicating the critical note was already created.
    /// </summary>
    private void ShowAlreadyCreatedTooltip()
    {
        _borderTooltip.Show("Critical note already created for this study", this,
            Width / 2, Height / 2, 2000);
    }

    /// <summary>
    /// Paste the clinical history to Mosaic and return focus to previous window.
    /// </summary>
    private void PasteClinicalHistoryToMosaic()
    {
        var text = _contentLabel.Text;
        if (string.IsNullOrWhiteSpace(text) || text == "(No clinical history)")
        {
            ShowToast("No clinical history to paste");
            return;
        }

        // Format the text with leading newline for cleaner paste
        var formatted = $"\nClinical history: {text}.";

        // Run on background thread to avoid blocking UI
        System.Threading.Tasks.Task.Run(() =>
        {
            // Use paste lock to prevent race conditions with macro insertion
            // Both will execute in order - no skipping
            lock (ActionController.PasteLock)
            {
                try
                {
                    // Save current focus before we do anything (likely IntelliViewer)
                    NativeWindows.SavePreviousFocus();

                    // Set clipboard on UI thread (STA required for clipboard operations)
                    Logger.Trace($"Clinical history paste: setting clipboard to '{formatted.Substring(0, Math.Min(50, formatted.Length))}...'");
                    Invoke(() => Clipboard.SetText(formatted));
                    System.Threading.Thread.Sleep(50);

                    if (!NativeWindows.ActivateMosaicForcefully())
                    {
                        Invoke(() => ShowToast("Mosaic not found"));
                        return;
                    }

                    System.Threading.Thread.Sleep(200);
                    NativeWindows.SendHotkey("ctrl+v");
                    System.Threading.Thread.Sleep(100);

                    ActionController.LastPasteTime = DateTime.Now;

                    // Restore focus to previous window
                    NativeWindows.RestorePreviousFocus();

                    Invoke(() => ShowToast("Pasted to Mosaic"));
                }
                catch (Exception ex)
                {
                    Logger.Trace($"PasteClinicalHistoryToMosaic error: {ex.Message}");
                    Invoke(() => ShowToast("Paste failed"));
                }
            }
        });
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
    private void ShowCopiedToast() => ShowToast("Debug copied!");

    /// <summary>
    /// Show a brief toast with a custom message.
    /// </summary>
    private void ShowToast(string message)
    {
        var toast = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            ShowInTaskbar = false,
            TopMost = true,
            BackColor = Color.FromArgb(51, 51, 51),
            Size = new Size(140, 30),
            StartPosition = FormStartPosition.Manual,
            Location = new Point(Location.X + Width / 2 - 70, Location.Y - 35)
        };

        var label = new Label
        {
            Text = message,
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

        // Remove junk phrases and re-collapse any gaps left behind
        cleaned = Regex.Replace(cleaned, @"Other\s*\(Please Specify\)\s*;?\s*", " ", RegexOptions.IgnoreCase);
        cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();

        // Remove repeating phrases
        return RemoveRepeatingPhrases(cleaned);
    }

    /// <summary>
    /// Detect and remove repeating phrases in text.
    /// Handles cases like:
    /// - "fall fall fall fall" -> "fall"
    /// - "Chest Pain >18 Chest Pain >18 Chest Pain >18" -> "Chest Pain >18"
    /// - "A; B A; B B A; B" -> "A; B"
    /// </summary>
    private static string RemoveRepeatingPhrases(string text)
    {
        if (string.IsNullOrEmpty(text) || text.Length < 6)
            return text;

        string result = text;

        // FIRST: Check if whole string is just repetitions of a simple pattern
        // This catches "fall fall fall fall" -> "fall" before complex logic messes it up
        // Use (Length + 1) / 2 to handle odd-length strings like "fall fall" (9 chars, pattern "fall " is 5)
        for (int patternLen = 1; patternLen <= (result.Length + 1) / 2; patternLen++)
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

        // SECOND: Find and remove CONSECUTIVE duplicate phrases only
        // This handles cases like "Chest Pain >18 Chest Pain >18" but NOT "midepigastric pain... midepigastric region"
        // Only remove when the same phrase appears immediately after itself (with optional whitespace/punctuation between)
        bool foundDuplicate = true;
        while (foundDuplicate)
        {
            foundDuplicate = false;
            int maxPhraseLen = result.Length / 2;

            for (int phraseLen = maxPhraseLen; phraseLen >= 5; phraseLen--)
            {
                for (int startPos = 0; startPos <= result.Length - phraseLen * 2; startPos++)
                {
                    string phrase = result.Substring(startPos, phraseLen).Trim();
                    if (phrase.Length < 5) continue;

                    // Look for the same phrase immediately following (with optional whitespace/punctuation gap)
                    int afterFirstPhrase = startPos + phraseLen;

                    // Skip whitespace and minor punctuation between potential duplicates
                    while (afterFirstPhrase < result.Length &&
                           (char.IsWhiteSpace(result[afterFirstPhrase]) || result[afterFirstPhrase] == ';' || result[afterFirstPhrase] == ','))
                    {
                        afterFirstPhrase++;
                    }

                    // Check if the phrase repeats immediately after
                    if (afterFirstPhrase + phrase.Length <= result.Length)
                    {
                        string nextChunk = result.Substring(afterFirstPhrase, phrase.Length);
                        if (string.Equals(phrase, nextChunk.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            // Found consecutive duplicate - remove the second occurrence
                            result = result.Substring(0, startPos + phraseLen) + result.Substring(afterFirstPhrase + phrase.Length);
                            result = Regex.Replace(result, @"\s+", " ").Trim();
                            foundDuplicate = true;
                            break;
                        }
                    }
                }
                if (foundDuplicate) break;
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
        // Note: Don't include EXAM here - it causes false matches with "Reason for exam:" in clinical text
        var sectionHeaders = @"TECHNIQUE|FINDINGS|IMPRESSION|COMPARISON|PROCEDURE|CONCLUSION|RECOMMENDATION";

        // Find "CLINICAL HISTORY" section - look for content until next section header or end
        var match = Regex.Match(rawText,
            $@"CLINICAL HISTORY[:\s]*(.*?)(?=\b({sectionHeaders})\b|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!match.Success || match.Groups.Count < 2)
        {
            Services.Logger.Trace("ExtractClinicalHistory: Regex did not match");
            return null;
        }

        Services.Logger.Trace($"ExtractClinicalHistory: Regex matched, group1 length={match.Groups[1].Value.Length}");

        var content = match.Groups[1].Value.Trim();

        // Section exists but is empty - return empty string (not null)
        if (string.IsNullOrWhiteSpace(content))
            return string.Empty;

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

    /// <summary>
    /// Extract clinical history and return both pre-cleaned and cleaned versions.
    /// Used to detect if fixing was needed.
    /// </summary>
    public static (string? preCleaned, string? cleaned) ExtractClinicalHistoryWithFixInfo(string? rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText))
            return (null, null);

        if (!rawText.Contains("CLINICAL HISTORY", StringComparison.OrdinalIgnoreCase))
            return (null, null);

        // Note: Don't include EXAM here - it causes false matches with "Reason for exam:" in clinical text
        var sectionHeaders = @"TECHNIQUE|FINDINGS|IMPRESSION|COMPARISON|PROCEDURE|CONCLUSION|RECOMMENDATION";
        var match = Regex.Match(rawText,
            $@"CLINICAL HISTORY[:\s]*(.*?)(?=\b({sectionHeaders})\b|$)",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (!match.Success || match.Groups.Count < 2)
            return (null, null);

        var content = match.Groups[1].Value.Trim();

        // Section exists but is empty - return empty strings (not null)
        if (string.IsNullOrWhiteSpace(content))
            return (string.Empty, string.Empty);

        // Capture ALL lines BEFORE deduplication for pre-cleaned comparison
        var allLines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var allTrimmedLines = new System.Collections.Generic.List<string>();
        foreach (var line in allLines)
        {
            var trimmed = line.Trim();
            if (!string.IsNullOrWhiteSpace(trimmed))
                allTrimmedLines.Add(trimmed);
        }

        if (allTrimmedLines.Count == 0)
            return (string.Empty, string.Empty);

        // Pre-cleaned: BEFORE line dedup and phrase removal (just basic char cleanup)
        var preCleanedRaw = string.Join(" ", allTrimmedLines);
        var sb = new System.Text.StringBuilder();
        foreach (char c in preCleanedRaw)
        {
            if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) || c == '-' || c == '/' || c == '\'')
                sb.Append(c);
            else if (char.IsControl(c))
                sb.Append(' ');
        }
        var preCleaned = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();

        // Now deduplicate lines for the cleaned version
        var uniqueLines = new System.Collections.Generic.List<string>();
        foreach (var line in allTrimmedLines)
        {
            bool isDuplicate = false;
            foreach (var existing in uniqueLines)
            {
                if (string.Equals(existing, line, StringComparison.OrdinalIgnoreCase))
                {
                    isDuplicate = true;
                    break;
                }
            }
            if (!isDuplicate)
                uniqueLines.Add(line);
        }

        // Fully cleaned: with line dedup AND phrase removal
        var cleaned = CleanText(string.Join(" ", uniqueLines));

        return (preCleaned, cleaned);
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

    #region Gender Check

    /// <summary>
    /// Check report text for gender-specific terms that don't match the patient's gender.
    /// Returns a list of mismatched terms found in the report.
    /// </summary>
    public static List<string> CheckGenderMismatch(string? reportText, string? patientGender)
    {
        var mismatches = new List<string>();

        if (string.IsNullOrWhiteSpace(reportText) || string.IsNullOrWhiteSpace(patientGender))
            return mismatches;

        var reportLower = reportText.ToLowerInvariant();
        var genderUpper = patientGender.ToUpperInvariant();

        if (genderUpper == "MALE")
        {
            // Check for female-only terms in male patient's report
            foreach (var term in FemaleOnlyTerms)
            {
                // Use word boundary matching to avoid partial matches
                if (Regex.IsMatch(reportLower, $@"\b{Regex.Escape(term)}\b"))
                {
                    mismatches.Add(term);
                }
            }
        }
        else if (genderUpper == "FEMALE")
        {
            // Check for male-only terms in female patient's report
            foreach (var term in MaleOnlyTerms)
            {
                if (Regex.IsMatch(reportLower, $@"\b{Regex.Escape(term)}\b"))
                {
                    mismatches.Add(term);
                }
            }
        }

        return mismatches;
    }

    /// <summary>
    /// Set or clear the gender warning state.
    /// When active, replaces clinical history with a blinking red warning.
    /// </summary>
    public void SetGenderWarning(bool active, string? patientGender, List<string>? mismatchedTerms)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetGenderWarning(active, patientGender, mismatchedTerms));
            return;
        }

        if (active && mismatchedTerms != null && mismatchedTerms.Count > 0)
        {
            // Save current clinical history text if not already in warning mode
            if (!_genderWarningActive)
            {
                _savedClinicalHistoryText = _contentLabel.Text;
            }

            _genderWarningActive = true;
            _genderWarningText = $"GENDER MISMATCH!\nPatient: {patientGender}\nTerms: {string.Join(", ", mismatchedTerms)}";

            // Start blinking
            StartBlinking();
            UpdateBorderTooltip();
        }
        else
        {
            // Clear warning and restore clinical history
            if (_genderWarningActive)
            {
                _genderWarningActive = false;
                StopBlinking();

                // Restore label background to black
                _contentLabel.BackColor = Color.Black;

                // Restore current clinical history text (set by SetClinicalHistory before this call)
                // Use _currentDisplayedText which is kept up-to-date, not the old saved text
                if (!string.IsNullOrEmpty(_currentDisplayedText))
                {
                    _contentLabel.Text = _currentDisplayedText;
                    _contentLabel.ForeColor = _currentTextWasFixed ? FixedTextColor : NormalTextColor;
                }
                else if (!string.IsNullOrEmpty(_savedClinicalHistoryText))
                {
                    // Fallback to saved text if no current text
                    _contentLabel.Text = _savedClinicalHistoryText;
                    _contentLabel.ForeColor = NormalTextColor;
                }

                // Restore normal border (may be overridden by template mismatch or drafted state)
                if (!_templateMismatch)
                {
                    BackColor = NormalBorderColor;
                }
                UpdateBorderTooltip();

                _savedClinicalHistoryText = null;
            }
        }
    }

    private void StartBlinking()
    {
        if (_blinkTimer != null) return;

        _blinkState = true;
        UpdateBlinkDisplay();

        _blinkTimer = new System.Windows.Forms.Timer { Interval = 500 }; // Blink every 500ms
        _blinkTimer.Tick += (_, _) =>
        {
            _blinkState = !_blinkState;
            UpdateBlinkDisplay();
        };
        _blinkTimer.Start();
    }

    private void StopBlinking()
    {
        if (_blinkTimer != null)
        {
            _blinkTimer.Stop();
            _blinkTimer.Dispose();
            _blinkTimer = null;
        }
    }

    private void UpdateBlinkDisplay()
    {
        if (!_genderWarningActive) return;

        if (_blinkState)
        {
            // Bright red background, white text
            BackColor = Color.FromArgb(220, 0, 0);
            _contentLabel.BackColor = Color.FromArgb(180, 0, 0);
            _contentLabel.ForeColor = Color.White;
            _contentLabel.Text = _genderWarningText ?? "GENDER MISMATCH!";
        }
        else
        {
            // Darker red background
            BackColor = Color.FromArgb(120, 0, 0);
            _contentLabel.BackColor = Color.FromArgb(80, 0, 0);
            _contentLabel.ForeColor = Color.FromArgb(255, 200, 200);
            _contentLabel.Text = _genderWarningText ?? "GENDER MISMATCH!";
        }
    }

    /// <summary>
    /// Returns whether gender warning is currently active.
    /// </summary>
    public bool IsGenderWarningActive => _genderWarningActive;

    #endregion
}
