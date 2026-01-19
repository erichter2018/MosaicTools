using System.Threading;
using System.Text.RegularExpressions;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.UIA3;

namespace MosaicTools.Services;

/// <summary>
/// UI Automation service for Clario and Mosaic interaction.
/// Uses FlaUI which is a modern wrapper over Windows UI Automation.
/// Matches Python's pywinauto UIA backend logic.
/// </summary>
public class AutomationService : IDisposable
{
    private readonly UIA3Automation _automation;
    
    // Cached window for fast repeated scrapes
    private AutomationElement? _cachedSlimHubWindow;
    
    // Last scraped final report (public for external access)
    public string? LastFinalReport { get; private set; }

    // Last detected drafted state
    public bool LastDraftedState { get; private set; }

    // Last scraped study description (e.g., "CT ABDOMEN PELVIS WITHOUT IV CONTRAST")
    public string? LastDescription { get; private set; }

    // Last template name from final report (2nd line after EXAM:)
    public string? LastTemplateName { get; private set; }

    // Last scraped accession number
    public string? LastAccession { get; private set; }

    // Last scraped patient gender ("Male", "Female", or null if unknown)
    public string? LastPatientGender { get; private set; }

    // Debug flag
    private bool _hasLoggedDebugInfo = false;
    
    public AutomationService()
    {
        _automation = new UIA3Automation();
    }
    
    #region Clario Methods
    
    /// <summary>
    /// Find the Chrome window with Clario - Worklist.
    /// </summary>
    public AutomationElement? FindClarioWindow()
    {
        try
        {
            var desktop = _automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));
            
            foreach (var window in windows)
            {
                try
                {
                    var title = window.Name?.ToLowerInvariant() ?? "";
                    if (title.Contains("clario") && title.Contains("worklist"))
                    {
                        Logger.Trace($"Found Clario Window: '{window.Name}'");
                        return window;
                    }
                }
                catch { continue; }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Error finding Clario window: {ex.Message}");
        }
        
        return null;
    }
    
    /// <summary>
    /// Check if the Note Dialog is open and extract text from it.
    /// Fast path before doing full tree scan.
    /// </summary>
    public string? GetNoteDialogText(AutomationElement window)
    {
        try
        {
            var dialog = window.FindFirstDescendant(cf => 
                cf.ByAutomationId("content_patient_note_dialog_Main"));
            
            if (dialog == null) return null;
            
            // Found the dialog - look for note text field
            var noteFields = dialog.FindAllDescendants(cf => 
                cf.ByAutomationId("noteFieldMessage"));
            
            foreach (var field in noteFields)
            {
                var text = field.Name ?? "";
                if (!string.IsNullOrWhiteSpace(text))
                {
                    return text.Trim();
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Error checking for note dialog: {ex.Message}");
        }
        
        return null;
    }
    
    /// <summary>
    /// Exhaustive search for EXAM NOTE DataItems.
    /// Matches Python's get_exam_note_elements().
    /// </summary>
    public List<string> GetExamNoteElements(AutomationElement window, int timeoutSeconds = 15, Action<string>? statusCallback = null)
    {
        var results = new List<string>();
        var startTime = DateTime.Now;
        var lastToastTime = DateTime.Now;
        
        try
        {
            Logger.Trace($"Starting optimized Clario EXAM NOTE search ({timeoutSeconds}s timeout)...");
            
            while ((DateTime.Now - startTime).TotalSeconds < timeoutSeconds)
            {
                // Repeating toast every 5 seconds
                if ((DateTime.Now - lastToastTime).TotalSeconds >= 5)
                {
                    statusCallback?.Invoke("Still searching Clario...");
                    lastToastTime = DateTime.Now;
                }

                try
                {
                    // Lead with the broad DataItem scan which was proven to work in logs
                    var dataItems = window.FindAllDescendants(cf => 
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.DataItem));
                    
                    if (dataItems.Length > 0)
                    {
                        foreach (var item in dataItems)
                        {
                            var name = item.Name ?? "";
                            if (name.Contains("EXAM NOTE", StringComparison.OrdinalIgnoreCase))
                            {
                                results.Add(name);
                                if (results.Count >= 15) break;
                            }
                        }
                    }

                    if (results.Count > 0)
                    {
                        Logger.Trace($"SUCCESS: Found {results.Count} EXAM NOTE matches.");
                        return results;
                    }
                }
                catch (Exception ex)
                {
                    // Handle UIA_E_ELEMENTNOTAVAILABLE (0x80040201) and others
                    Logger.Trace($"Scrape loop error (retrying): {ex.Message}");
                    Thread.Sleep(200);
                }
                
                Thread.Sleep(1000); // Wait 1s between full tree scans to let UI stabilize
            }
            
            if (results.Count == 0)
            {
                Logger.Trace("TIMEOUT: No EXAM NOTE matches found.");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Outer search failed: {ex.Message}");
        }
        
        return results;
    }
    
    /// <summary>
    /// Perform the tiered Clario scrape.
    /// </summary>
    /// <param name="statusCallback">Optional callback for status updates</param>
    /// <param name="focusWindow">Whether to focus the Clario window (set false for background scrapes)</param>
    public string? PerformClarioScrape(Action<string>? statusCallback = null, bool focusWindow = true)
    {
        var window = FindClarioWindow();
        if (window == null)
        {
            Logger.Trace("Clario window NOT found.");
            return null;
        }

        // Only focus window if requested (skip for background scrapes to avoid stealing focus)
        if (focusWindow)
        {
            try
            {
                Logger.Trace("Focusing Clario window before scrape...");
                window.Focus();
                Thread.Sleep(500);
            }
            catch { }
        }
        
        // Fast path: Check if Note Dialog is open
        var noteDialogText = GetNoteDialogText(window);
        if (noteDialogText != null)
        {
            Logger.Trace("SUCCESS: Found open Note Dialog.");
            return noteDialogText;
        }
        
        // Normal path: Exhaustive search
        var notes = GetExamNoteElements(window, statusCallback: statusCallback);
        
        if (notes.Count > 0)
        {
            // Sort by timestamp (newest first), then by length as a tie-breaker
            var bestNote = notes
                .OrderByDescending(ExtractNoteTimestamp)
                .ThenByDescending(n => n.Length)
                .First();
                
            Logger.Trace($"SUCCESS: Returning newest match (len={bestNote.Length})");
            return bestNote;
        }
        
        return null;
    }

    /// <summary>
    /// Helper to extract timestamp for sorting.
    /// Matches MM/dd/yyyy h:mm tt
    /// </summary>
    private static DateTime ExtractNoteTimestamp(string text)
    {
        try
        {
            // Look for date/time pattern at the end of the string
            var match = Regex.Match(text, @"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2}\s*(?:AM|PM))", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                if (DateTime.TryParseExact($"{match.Groups[1].Value} {match.Groups[2].Value}", 
                    "MM/dd/yyyy h:mm tt", 
                    System.Globalization.CultureInfo.InvariantCulture, 
                    System.Globalization.DateTimeStyles.None, 
                    out var dt))
                {
                    return dt;
                }
            }
        }
        catch { }
        return DateTime.MinValue;
    }
    
    #endregion
    
    #region Mosaic Methods
    
    /// <summary>
    /// Find the Mosaic Info Hub or Mosaic Reporting window.
    /// </summary>
    public AutomationElement? FindMosaicWindow()
    {
        try
        {
            var desktop = _automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));
            
            foreach (var window in windows)
            {
                try
                {
                    var title = window.Name?.ToLowerInvariant() ?? "";
                    if (title.Contains("rvu counter") || title.Contains("test")) continue;
                    
                    if ((title.Contains("mosaic") && title.Contains("info hub")) ||
                        (title.Contains("mosaic") && title.Contains("reporting")))
                    {
                        return window;
                    }
                }
                catch { continue; }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Error finding Mosaic window: {ex.Message}");
        }
        
        return null;
    }
    
    /// <summary>
    /// Find the Mosaic editor element for pasting.
    /// </summary>
    public AutomationElement? FindMosaicEditor()
    {
        var window = FindMosaicWindow();
        if (window == null) return null;
        
        try
        {
            // Look for WebView2 container
            AutomationElement? searchRoot = null;
            
            try
            {
                searchRoot = window.FindFirstDescendant(cf => cf.ByAutomationId("webView"));
            }
            catch { }
            
            searchRoot ??= window;
            
            // Search for specific editor elements
            var descendants = searchRoot.FindAllDescendants();
            foreach (var elem in descendants)
            {
                try
                {
                    var className = elem.ClassName?.ToLowerInvariant() ?? "";
                    var name = elem.Name?.ToUpperInvariant() ?? "";
                    
                    // Match by tiptap/prosemirror class
                    if (className.Contains("tiptap") || className.Contains("prosemirror"))
                    {
                        return elem;
                    }
                    
                    // Match by ADDENDUM text
                    if (name.Contains("ADDENDUM:") && elem.ControlType != FlaUI.Core.Definitions.ControlType.Text)
                    {
                        return elem;
                    }
                }
                catch { continue; }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Error in FindMosaicEditor: {ex.Message}");
        }
        
        return null;
    }
    
    /// <summary>
    /// Get report text content from Mosaic.
    /// </summary>
    public string? GetReportTextContent()
    {
        var window = FindMosaicWindow();
        if (window == null) return null;
        
        try
        {
            var candidates = new List<(string text, int score)>();
            string[] keywords = { "TECHNIQUE", "CLINICAL HISTORY", "FINDINGS", "IMPRESSION" };
            
            var descendants = window.FindAllDescendants();
            foreach (var elem in descendants)
            {
                try
                {
                    var className = elem.ClassName ?? "";
                    if (className.Contains("ProseMirror"))
                    {
                        var text = elem.Name ?? "";
                        if (string.IsNullOrEmpty(text)) continue;
                        
                        int score = keywords.Count(kw => text.Contains(kw, StringComparison.OrdinalIgnoreCase));
                        candidates.Add((text, score));
                    }
                }
                catch { continue; }
            }
            
            if (candidates.Count > 0)
            {
                return candidates.OrderByDescending(c => c.score).First().text;
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Report scan error: {ex.Message}");
        }
        
        return null;
    }
    
    /// <summary>
    /// Fast scrape of the Final Report from Mosaic/SlimHub.
    /// Uses caching and targeted search for speed.
    /// </summary>
    public string? GetFinalReportFast(bool checkDraftedStatus = false)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        
        try
        {
            // Try to use cached window first
            if (_cachedSlimHubWindow != null)
            {
                try
                {
                    // Verify cached window is still valid
                    var title = _cachedSlimHubWindow.Name?.ToLowerInvariant() ?? "";
                    bool isValid = (title.Contains("mosaic") && title.Contains("info hub")) ||
                                   (title.Contains("mosaic") && title.Contains("reporting"));
                                   
                    if (!isValid)
                    {
                        _cachedSlimHubWindow = null;
                        Logger.Trace("GetFinalReportFast: Cached window invalid/closed.");
                    }
                }
                catch
                {
                    _cachedSlimHubWindow = null;
                }
            }
            
            // Find Mosaic window if not cached
            if (_cachedSlimHubWindow == null)
            {
                var desktop = _automation.GetDesktop();
                var windows = desktop.FindAllChildren(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));
                
                foreach (var window in windows)
                {
                    try
                    {
                        var title = window.Name?.ToLowerInvariant() ?? "";
                        if (title.Contains("rvu counter") || title.Contains("test")) continue;
                        
                        if ((title.Contains("mosaic") && title.Contains("info hub")) ||
                            (title.Contains("mosaic") && title.Contains("reporting")))
                        {
                            _cachedSlimHubWindow = window;
                            Logger.Trace($"Cached Mosaic window: {window.Name}");
                            break;
                        }
                    }
                    catch { continue; }
                }
            }
            
            if (_cachedSlimHubWindow == null)
            {
                Logger.Trace($"GetFinalReportFast: Mosaic window not found ({sw.ElapsedMilliseconds}ms)");
                return null;
            }

            // Single traversal: find DRAFTED status, Report document, Description, and Accession
            AutomationElement? reportDoc = null;
            LastDescription = null; // Reset
            LastTemplateName = null; // Reset - prevents stale template from previous study causing false mismatch
            LastAccession = null; // Reset - will be set if found, stays null if no study open
            LastPatientGender = null; // Reset - will be set if found

            // Extract accession: find "Current Study" text and the next text element after it
            try
            {
                var currentStudyElement = _cachedSlimHubWindow.FindFirstDescendant(cf =>
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text)
                    .And(cf.ByName("Current Study")));

                if (currentStudyElement != null)
                {
                    // Get parent and find all Text children
                    var parent = currentStudyElement.Parent;
                    if (parent != null)
                    {
                        var textElements = parent.FindAllChildren(cf =>
                            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text));

                        bool foundCurrentStudy = false;
                        foreach (var textEl in textElements)
                        {
                            if (foundCurrentStudy)
                            {
                                // This is the element after "Current Study" - should be accession
                                var accession = textEl.Name?.Trim();
                                if (!string.IsNullOrWhiteSpace(accession) && accession != "Current Study")
                                {
                                    LastAccession = accession;
                                    Logger.Trace($"Found Accession: {LastAccession}");
                                }
                                break;
                            }
                            if (textEl.Name == "Current Study")
                            {
                                foundCurrentStudy = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Accession extraction error: {ex.Message}");
            }

            // Extract patient gender: look for text matching "MALE, AGE" or "FEMALE, AGE"
            try
            {
                var textElements = _cachedSlimHubWindow.FindAllDescendants(cf =>
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text));

                foreach (var textEl in textElements)
                {
                    var text = textEl.Name?.ToUpperInvariant() ?? "";
                    if (text.StartsWith("MALE, AGE") || text.StartsWith("MALE,AGE"))
                    {
                        LastPatientGender = "Male";
                        Logger.Trace($"Found Patient Gender: Male");
                        break;
                    }
                    else if (text.StartsWith("FEMALE, AGE") || text.StartsWith("FEMALE,AGE"))
                    {
                        LastPatientGender = "Female";
                        Logger.Trace($"Found Patient Gender: Female");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Gender extraction error: {ex.Message}");
            }

            if (checkDraftedStatus)
            {
                // Search for DRAFTED, Report doc, and Description in one traversal using OR condition
                var draftedCondition = _cachedSlimHubWindow.ConditionFactory
                    .ByControlType(FlaUI.Core.Definitions.ControlType.Text)
                    .And(_cachedSlimHubWindow.ConditionFactory.ByName("DRAFTED"));
                var reportCondition = _cachedSlimHubWindow.ConditionFactory
                    .ByControlType(FlaUI.Core.Definitions.ControlType.Document)
                    .And(_cachedSlimHubWindow.ConditionFactory.ByName("Report", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                var descriptionCondition = _cachedSlimHubWindow.ConditionFactory
                    .ByControlType(FlaUI.Core.Definitions.ControlType.Text)
                    .And(_cachedSlimHubWindow.ConditionFactory.ByName("Description:", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                var combinedCondition = draftedCondition.Or(reportCondition).Or(descriptionCondition);

                var elements = _cachedSlimHubWindow.FindAllDescendants(combinedCondition);

                LastDraftedState = false;
                foreach (var el in elements)
                {
                    if (el.ControlType == FlaUI.Core.Definitions.ControlType.Text && el.Name == "DRAFTED")
                    {
                        LastDraftedState = true;
                    }
                    else if (el.ControlType == FlaUI.Core.Definitions.ControlType.Document &&
                             el.Name?.Contains("Report") == true && reportDoc == null)
                    {
                        reportDoc = el;
                    }
                    else if (el.ControlType == FlaUI.Core.Definitions.ControlType.Text &&
                             el.Name?.StartsWith("Description:", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        // Extract just the description text after "Description: "
                        LastDescription = el.Name.Substring("Description:".Length).Trim();
                        Logger.Trace($"Found Description: {LastDescription}");
                    }
                }
            }
            else
            {
                // Just find the Report document
                reportDoc = _cachedSlimHubWindow.FindFirstDescendant(cf =>
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Document)
                    .And(cf.ByName("Report", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));
            }

            if (reportDoc == null)
            {
                Logger.Trace($"GetFinalReportFast: Report Document not found ({sw.ElapsedMilliseconds}ms)");
                return null;
            }
            
            // Strategy from Python: Look for ProseMirror editor
            // This is robust for Tiptap editors used in Mosaic
            var flowDoc = reportDoc; // Start search from the document
            
            // Find all potential editors
            var candidates = flowDoc.FindAllDescendants(cf => 
                cf.ByClassName("ProseMirror", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));

            if (candidates.Length > 0)
            {
                Logger.Trace($"GetFinalReportFast: Found {candidates.Length} ProseMirror candidates");
                
                string bestText = "";
                int maxScore = -1;
                string[] keywords = { "TECHNIQUE", "CLINICAL HISTORY", "FINDINGS", "IMPRESSION", "EXAM" };

                foreach (var candidate in candidates)
                {
                    string text = candidate.Name ?? "";

                    // Try ValuePattern (common for edit controls)
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        var valPattern = candidate.Patterns.Value.PatternOrDefault;
                        if (valPattern != null) text = valPattern.Value.Value;
                    }

                    // Try TextPattern (rich edit controls)
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        var txtPattern = candidate.Patterns.Text.PatternOrDefault;
                        if (txtPattern != null) text = txtPattern.DocumentRange.GetText(-1);
                    }

                    if (string.IsNullOrWhiteSpace(text)) continue;

                    // MUST start with "EXAM:" to be the correct final report box
                    if (!text.TrimStart().StartsWith("EXAM:", StringComparison.OrdinalIgnoreCase))
                        continue;

                    int score = 0;
                    foreach (var kw in keywords)
                    {
                        if (text.Contains(kw, StringComparison.OrdinalIgnoreCase)) score++;
                    }

                    if (score > maxScore)
                    {
                        maxScore = score;
                        bestText = text;
                    }
                }

                if (!string.IsNullOrEmpty(bestText) && maxScore > 0)
                {
                    sw.Stop();
                    int lineCount = bestText.Split('\n').Length;
                    Logger.Trace($"GetFinalReportFast: ProseMirror SUCCESS in {sw.ElapsedMilliseconds}ms, {lineCount} lines, Score={maxScore}");
                    LastFinalReport = bestText;
                    LastTemplateName = ExtractTemplateName(bestText);
                    return bestText;
                }
            }
            
            Logger.Trace($"GetFinalReportFast: ProseMirror search failed (Candidates={candidates.Length}). Proceeding to fallback...");

            // FALLBACK: Sibling/Fragment search
            // This runs if ProseMirror elements weren't found OR didn't contain the report
            var examElement = reportDoc.FindFirstDescendant(cf => 
                    cf.ByName("EXAM:", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));

            if (examElement != null)
            {
                // The report text is fragmented across siblings.
                var container = examElement.Parent;
                if (container != null)
                {
                    var sb = new System.Text.StringBuilder();
                    var allDescendants = container.FindAllDescendants();
                    
                    foreach (var desc in allDescendants)
                    {
                        var t = desc.Name;
                        if (!string.IsNullOrWhiteSpace(t)) sb.AppendLine(t);
                    }
                    
                    var fullText = sb.ToString();
                    Logger.Trace($"GetFinalReportFast: Fallback Reconstruction found {fullText.Length} chars");
                    
                    if (fullText.Length > 50)
                    {
                        sw.Stop();
                        int lineCount = fullText.Split('\n').Length;
                        Logger.Trace($"GetFinalReportFast: Fallback SUCCESS in {sw.ElapsedMilliseconds}ms, {lineCount} lines");
                        LastFinalReport = fullText;
                        LastTemplateName = ExtractTemplateName(fullText);
                        return fullText;
                    }
                }
            }
            
            Logger.Trace($"GetFinalReportFast: Failed to find report content ({sw.ElapsedMilliseconds}ms)");

        }
        catch (Exception ex)
        {
            Logger.Trace($"GetFinalReportFast error: {ex.Message} ({sw.ElapsedMilliseconds}ms)");
            _cachedSlimHubWindow = null; // Invalidate cache on error
        }
        
        return null;
    }
    
    /// <summary>
    /// Check if Mosaic is currently recording using UIA (slow method).
    /// </summary>
    public bool IsDictationActiveUIA()
    {
        try
        {
            var desktop = _automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));
            
            foreach (var window in windows)
            {
                try
                {
                    var name = window.Name ?? "";
                    
                    // Direct window name match
                    if (name.Contains("Microphone recording"))
                        return true;
                    
                    // Check children if it's a Mosaic/SlimHub window
                    if (name.Contains("Mosaic") || name.Contains("SlimHub"))
                    {
                        var descendants = window.FindAllDescendants();
                        foreach (var child in descendants)
                        {
                            var childName = child.Name ?? "";
                            if (childName.Contains("Microphone recording") ||
                                childName.Contains("accessing your microphone"))
                            {
                                return true;
                            }
                        }
                    }
                }
                catch { continue; }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Error checking dictation state (UIA): {ex.Message}");
        }
        
        return false;
    }
    
    #endregion

    #region Template Matching

    /// <summary>
    /// Extract the template name from the final report (2nd line after EXAM:).
    /// </summary>
    public static string? ExtractTemplateName(string? reportText)
    {
        if (string.IsNullOrWhiteSpace(reportText)) return null;

        var lines = reportText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].Trim().StartsWith("EXAM:", StringComparison.OrdinalIgnoreCase))
            {
                // The template name is the next line after EXAM:
                if (i + 1 < lines.Length)
                {
                    return lines[i + 1].Trim();
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Common body parts to look for when matching templates.
    /// </summary>
    private static readonly string[] BodyParts = new[]
    {
        "HEAD", "BRAIN", "NECK", "CERVICAL", "C-SPINE", "CSPINE",
        "CHEST", "THORAX", "THORACIC", "T-SPINE", "TSPINE", "LUNG",
        "ABDOMEN", "ABDOMINAL", "PELVIS", "PELVIC", "LUMBAR", "L-SPINE", "LSPINE",
        "SPINE", "EXTREMITY", "UPPER EXTREMITY", "LOWER EXTREMITY",
        "ARM", "LEG", "SHOULDER", "HIP", "KNEE", "ANKLE", "WRIST", "ELBOW",
        "FOOT", "HAND", "FINGER", "TOE",
        "CARDIAC", "HEART", "CORONARY", "CTA", "MRA",
        "ANGIOGRAPHY", "ANGIOGRAM", "VENOGRAM",
        "PULMONARY VEINS", "PULMONARY ARTERIES", "PULMONARY EMBOLISM", "PE PROTOCOL",
        "AORTA", "AORTIC", "RUNOFF", "CAROTID",
        "SINUS", "ORBIT", "FACE", "FACIAL", "MAXILLOFACIAL", "TEMPORAL", "IAC",
        "RENAL", "KIDNEY", "UROGRAM", "ENTEROGRAPHY", "LIVER", "PANCREAS"
    };

    /// <summary>
    /// Extract body parts found in a text string.
    /// </summary>
    public static HashSet<string> ExtractBodyParts(string? text)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(text)) return result;

        var upperText = text.ToUpperInvariant();

        // Check for "CT ANGIOGRAPHY" and normalize to CTA
        if (upperText.Contains("CT ANGIOGRAPHY") || upperText.Contains("CT ANGIO"))
        {
            result.Add("CTA");
        }

        // Check for "MR ANGIOGRAPHY" and normalize to MRA
        if (upperText.Contains("MR ANGIOGRAPHY") || upperText.Contains("MR ANGIO"))
        {
            result.Add("MRA");
        }

        foreach (var part in BodyParts)
        {
            if (upperText.Contains(part))
            {
                // Normalize some common variations
                var normalized = part switch
                {
                    "ABDOMINAL" => "ABDOMEN",
                    "PELVIC" => "PELVIS",
                    "C-SPINE" or "CSPINE" => "CERVICAL",
                    "T-SPINE" or "TSPINE" => "THORACIC",
                    "L-SPINE" or "LSPINE" => "LUMBAR",
                    "THORAX" => "CHEST",
                    "LUNG" => "CHEST",
                    "BRAIN" => "HEAD", // Brain CT = Head CT
                    "ANGIOGRAPHY" or "ANGIOGRAM" => "CTA", // Normalize angio terms
                    "AORTIC" => "AORTA",
                    "FACIAL" or "MAXILLOFACIAL" => "FACE",
                    _ => part
                };
                result.Add(normalized);
            }
        }

        return result;
    }

    /// <summary>
    /// Check if the description and template body parts match.
    /// Returns true if they match exactly (no warning needed), false if any mismatch detected.
    /// </summary>
    public static bool DoBodyPartsMatch(string? description, string? templateName)
    {
        if (string.IsNullOrWhiteSpace(description) || string.IsNullOrWhiteSpace(templateName))
            return true; // Can't determine, assume OK

        var descParts = ExtractBodyParts(description);
        var templateParts = ExtractBodyParts(templateName);

        if (descParts.Count == 0 || templateParts.Count == 0)
            return true; // Can't determine, assume OK

        // Require complete agreement - both sets must be equal
        bool match = descParts.SetEquals(templateParts);

        if (match)
        {
            Logger.Trace($"Template MATCH: [{string.Join(", ", descParts)}]");
        }
        else
        {
            Logger.Trace($"Template MISMATCH: Description [{string.Join(", ", descParts)}] vs Template [{string.Join(", ", templateParts)}]");
        }

        return match;
    }

    #endregion

    public void Dispose()
    {
        _automation.Dispose();
    }
}
