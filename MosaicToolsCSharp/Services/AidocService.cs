using System.Runtime.InteropServices;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.UIA3;

namespace MosaicTools.Services;

/// <summary>
/// A single Aidoc finding with positive/negative status.
/// </summary>
public record AidocFinding(string FindingType, bool IsPositive);

/// <summary>
/// Result of scraping the Aidoc shortcut widget.
/// </summary>
public record AidocResult(string PatientName, string Accession, string Department, List<AidocFinding> Findings);

/// <summary>
/// Scrapes the Aidoc shortcut widget (Electron app) for AI-detected findings.
/// Filters findings by relevance to the current study type.
/// </summary>
public class AidocService
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("gdi32.dll")]
    private static extern uint GetPixel(IntPtr hdc, int x, int y);

    private readonly AutomationService _automationService;

    public AidocService(AutomationService automationService)
    {
        _automationService = automationService;
    }

    // Track last logged state to avoid spamming logs every scrape cycle
    private string? _lastLoggedState;

    /// <summary>
    /// Check if a finding icon has a red/orange dot (positive finding indicator).
    /// Samples pixels across the icon looking for warm indicator colors.
    /// </summary>
    private static bool HasRedDot(System.Drawing.Rectangle iconRect, string? findingType = null)
    {
        IntPtr hdc = GetDC(IntPtr.Zero);
        try
        {
            // Sample a grid of points across the icon.
            // Requires Aidoc pulse animation to be disabled (No animation) to avoid
            // false positives from color bleed between adjacent finding icons.
            int w = iconRect.Width;
            int h = iconRect.Height;
            int orangeCount = 0;
            var sampledColors = new List<string>();

            for (int dy = 0; dy < h; dy += Math.Max(1, h / 6))
            {
                for (int dx = 0; dx < w; dx += Math.Max(1, w / 6))
                {
                    int sx = iconRect.Left + dx;
                    int sy = iconRect.Top + dy;
                    uint pixel = GetPixel(hdc, sx, sy);
                    if (pixel == 0xFFFFFFFF) continue; // CLR_INVALID

                    int r = (int)(pixel & 0xFF);
                    int g = (int)((pixel >> 8) & 0xFF);
                    int b = (int)((pixel >> 16) & 0xFF);

                    // Strict orange: R>200, G<130, B<80 (Aidoc positive dot is ~253,92,20)
                    if (r > 200 && g < 130 && b < 80)
                        orangeCount++;

                    // Collect non-background colors for diagnostic logging
                    if (r > 50 || g > 50 || b > 50) // skip near-black
                        if (!(r > 200 && g > 200 && b > 200)) // skip near-white
                            sampledColors.Add($"({r},{g},{b})");
                }
            }

            bool isPositive = orangeCount >= 2;

            // Log diagnostic pixel info
            if (sampledColors.Count > 0)
            {
                var uniqueColors = sampledColors.Distinct().Take(10);
                var tag = isPositive ? "POSITIVE" : "NEGATIVE";
                Logger.Trace($"Aidoc HasRedDot({findingType}): {tag} at [{iconRect.X},{iconRect.Y} {w}x{h}] orange={orangeCount} colors=[{string.Join(" ", uniqueColors)}]");
            }

            return isPositive;
        }
        finally
        {
            ReleaseDC(IntPtr.Zero, hdc);
        }
    }

    /// <summary>
    /// Scrape the Aidoc shortcut widget for current patient findings.
    /// Returns null if widget is not found, collapsed, or has no findings.
    /// Logs detailed element properties (Image Name/HelpText/Size) for debugging
    /// positive vs negative finding detection.
    /// </summary>
    public AidocResult? ScrapeShortcutWidget()
    {
        AutomationElement? widget = null, renderHost = null, group = null;
        try
        {
            var automation = _automationService.Automation;
            var desktop = _automationService.GetCachedDesktop();
            var cf = automation.ConditionFactory;

            widget = desktop.FindFirstChild(
                cf.ByName("aidoc-shortcut").And(cf.ByClassName("Chrome_WidgetWin_1")));

            if (widget == null)
                return null;

            renderHost = widget.FindFirstChild(cf.ByClassName("Chrome_RenderWidgetHostHWND"));
            if (renderHost == null)
                return null;

            group = renderHost.FindFirstChild(cf.ByControlType(FlaUI.Core.Definitions.ControlType.Group));
            if (group == null)
                return null;

            // Walk ALL children, building typed lists for logging and parsing
            var children = group.FindAllChildren();
            var parsed = new List<(string Type, string Text, System.Drawing.Rectangle Rect)>();
            var logEntries = new List<string>();

            try
            {
                foreach (var child in children)
                {
                    if (child.ControlType == FlaUI.Core.Definitions.ControlType.Text)
                    {
                        var name = child.Name ?? "";
                        parsed.Add(("Text", name, System.Drawing.Rectangle.Empty));
                        logEntries.Add($"Text:'{name}'");
                    }
                    else if (child.ControlType == FlaUI.Core.Definitions.ControlType.Image)
                    {
                        try
                        {
                            var rect = child.BoundingRectangle;
                            parsed.Add(("Image", "", rect));
                            logEntries.Add($"Image:{rect.Width}x{rect.Height}");
                        }
                        catch
                        {
                            parsed.Add(("Image", "", System.Drawing.Rectangle.Empty));
                            logEntries.Add("Image:(error)");
                        }
                    }
                    else
                    {
                        logEntries.Add($"{child.ControlType}:'{child.Name}'");
                    }
                }
            }
            finally
            {
                AutomationService.ReleaseElements(children);
            }

            // Log full element dump on changes
            var stateKey = string.Join("|", logEntries);
            if (stateKey != _lastLoggedState)
            {
                _lastLoggedState = stateKey;
                Logger.Trace($"Aidoc widget elements ({logEntries.Count}):");
                for (int i = 0; i < logEntries.Count; i++)
                    Logger.Trace($"  [{i}] {logEntries[i]}");
            }

            // Parse structure: [Image arrow][Image severity][Text name][Text accession][Text dept]
            // Optional: [Text status like 'Follow-up']
            // Then repeating: [Image findingIcon][Text findingType]
            // Ending with: [Text status]
            var textElements = parsed.Where(e => e.Type == "Text").ToList();
            if (textElements.Count < 4)
                return null;

            string patientName = textElements[0].Text;
            string accession = textElements[1].Text;
            string department = textElements[2].Text;

            // Extract all findings: each is an Image(w>=25) followed by a Text
            // Use pixel sampling on each icon to detect red dot (positive finding)
            var findings = new List<AidocFinding>();
            for (int i = 0; i < parsed.Count - 1; i++)
            {
                var el = parsed[i];
                var next = parsed[i + 1];
                if (el.Type == "Image" && el.Rect.Width >= 25 && next.Type == "Text"
                    && !string.IsNullOrWhiteSpace(next.Text))
                {
                    bool isPositive = el.Rect.Width > 0 && HasRedDot(el.Rect, next.Text);
                    findings.Add(new AidocFinding(next.Text, isPositive));
                    Logger.Trace($"Aidoc finding: {next.Text} positive={isPositive}");
                }
            }

            if (findings.Count == 0)
                return null;

            return new AidocResult(patientName, accession, department, findings);
        }
        catch (Exception ex)
        {
            Logger.Trace($"Aidoc scrape error: {ex.Message}");
            return null;
        }
        finally
        {
            AutomationService.ReleaseElement(group);
            AutomationService.ReleaseElement(renderHost);
            AutomationService.ReleaseElement(widget);
        }
    }

    // Study type keyword mappings for relevance filtering
    private static readonly Dictionary<string, string[]> StudyTypeFindings = new(StringComparer.OrdinalIgnoreCase)
    {
        // Head CT non-contrast
        ["HEAD_NONCON"] = new[] { "ICH", "MLS" },
        // CTA Head
        ["HEAD_CTA"] = new[] { "LVO", "M1-LVO", "VO", "BA" },
        // C-Spine
        ["CSPINE"] = new[] { "CSpFx" },
        // Chest CT
        ["CHEST"] = new[] { "PE", "Ptx", "PN", "MalETT", "RV/LV", "AD", "M-Aorta", "CAC", "RibFx" },
        // Abdomen/Pelvis CT
        ["ABDPELV"] = new[] { "FreeAir", "GIB", "M-AbdAo" },
        // Spine (non-cervical)
        ["SPINE"] = new[] { "VCFx" }
    };

    // Incidental findings relevant to chest/abdomen studies only
    private static readonly HashSet<string> ChestAbdIncidentalFindings = new(StringComparer.OrdinalIgnoreCase)
    {
        "IPE", "iptx"
    };

    /// <summary>
    /// Determine if a finding type is relevant to the current study description.
    /// Incidental findings (IPE, iptx) are always relevant.
    /// </summary>
    public static bool IsRelevantFinding(string findingType, string? studyDescription)
    {
        if (string.IsNullOrWhiteSpace(findingType))
            return false;

        if (string.IsNullOrWhiteSpace(studyDescription))
            return false;

        var desc = studyDescription.ToUpperInvariant();

        // Determine study type from description and check relevant findings
        string? studyType = ClassifyStudy(desc);
        if (studyType == null)
            return false;

        // Incidental findings (IPE, iptx) only relevant for chest/abdomen
        if (ChestAbdIncidentalFindings.Contains(findingType))
            return studyType is "CHEST" or "ABDPELV";

        if (StudyTypeFindings.TryGetValue(studyType, out var relevantFindings))
        {
            foreach (var f in relevantFindings)
            {
                if (string.Equals(f, findingType, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Classify a study description into a study type key.
    /// </summary>
    private static string? ClassifyStudy(string descUpper)
    {
        // CTA Head check must come before generic HEAD check
        if ((descUpper.Contains("CTA") || descUpper.Contains("ANGIO")) &&
            descUpper.Contains("HEAD"))
            return "HEAD_CTA";

        // Head CT non-contrast
        if (descUpper.Contains("HEAD") && !descUpper.Contains("CTA") && !descUpper.Contains("ANGIO"))
            return "HEAD_NONCON";

        // C-Spine (check before generic SPINE)
        if (descUpper.Contains("C-SPINE") || descUpper.Contains("CSPINE") || descUpper.Contains("CERVICAL SPINE"))
            return "CSPINE";

        // Chest CT
        if (descUpper.Contains("CHEST") || descUpper.Contains("THORAX") || descUpper.Contains("PE STUDY"))
            return "CHEST";

        // Abdomen/Pelvis CT
        if (descUpper.Contains("ABD") || descUpper.Contains("PELV"))
            return "ABDPELV";

        // Spine (non-cervical, checked after C-Spine)
        if (descUpper.Contains("SPINE"))
            return "SPINE";

        return null;
    }
}
