using System.Threading;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
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

    // Last extracted Clario Priority (e.g., "STAT", "Stroke", "Routine")
    public string? LastClarioPriority { get; private set; }

    // Last extracted Clario Class (e.g., "Emergency", "Inpatient", "Outpatient")
    public string? LastClarioClass { get; private set; }

    // Whether the current study is detected as a Stroke study
    public bool IsStrokeStudy { get; private set; }

    // Last scraped patient name (title case, e.g., "Smith John")
    public string? LastPatientName { get; private set; }

    // Last scraped site code (e.g., "MLC", "UNM")
    public string? LastSiteCode { get; private set; }

    // Last scraped MRN (Medical Record Number) - required for XML file drop
    public string? LastMrn { get; private set; }

    
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

    #region Clario Priority/Class Extraction (Stroke Detection)

    /// <summary>
    /// Data extracted from a single UI element during tree traversal.
    /// </summary>
    private class ClarioElementData
    {
        public int Depth { get; set; }
        public string AutomationId { get; set; } = "";
        public string Name { get; set; } = "";
        public string Text { get; set; } = "";
    }

    /// <summary>
    /// Extracted Priority and Class data from Clario.
    /// </summary>
    public class ClarioPriorityData
    {
        public string Priority { get; set; } = "";
        public string Class { get; set; } = "";
        public string Accession { get; set; } = "";
    }

    /// <summary>
    /// Recursively collect UI elements from the Clario window up to a max depth.
    /// </summary>
    private List<ClarioElementData> GetClarioElementsRecursive(AutomationElement element, int maxDepth)
    {
        var elements = new List<ClarioElementData>();
        CollectClarioElements(element, 0, maxDepth, elements);
        return elements;
    }

    private void CollectClarioElements(AutomationElement element, int depth, int maxDepth, List<ClarioElementData> elements)
    {
        if (depth > maxDepth) return;

        try
        {
            string automationId = element.Properties.AutomationId.ValueOrDefault ?? "";
            string name = element.Name ?? "";
            string text = "";

            // Try to get text from Value pattern (for input fields)
            if (element.Patterns.Value.IsSupported)
            {
                text = element.Patterns.Value.Pattern.Value.ValueOrDefault ?? "";
            }

            if (!string.IsNullOrEmpty(automationId) ||
                !string.IsNullOrEmpty(name) ||
                !string.IsNullOrEmpty(text))
            {
                elements.Add(new ClarioElementData
                {
                    Depth = depth,
                    AutomationId = automationId.Trim(),
                    Name = name.Trim(),
                    Text = text.Trim()
                });
            }

            // Recurse into children
            foreach (var child in element.FindAllChildren())
            {
                CollectClarioElements(child, depth + 1, maxDepth, elements);
            }
        }
        catch { /* Element may become stale */ }
    }

    /// <summary>
    /// Extract Priority, Class, and Accession from collected elements.
    /// </summary>
    private ClarioPriorityData ExtractPriorityDataFromElements(List<ClarioElementData> elements)
    {
        var data = new ClarioPriorityData();

        for (int i = 0; i < elements.Count; i++)
        {
            var elem = elements[i];

            // PRIORITY: Look for automation ID or name containing "priority"
            if (string.IsNullOrEmpty(data.Priority))
            {
                if (elem.AutomationId.Contains("priority", StringComparison.OrdinalIgnoreCase) ||
                    (elem.Name.Contains("priority", StringComparison.OrdinalIgnoreCase) &&
                     elem.Name.Contains(":")))
                {
                    data.Priority = FindNextClarioValue(elements, i);
                }
            }

            // CLASS: Look for "class" but NOT "priority" (avoid "priority class")
            if (string.IsNullOrEmpty(data.Class))
            {
                if ((elem.AutomationId.Contains("class", StringComparison.OrdinalIgnoreCase) &&
                     !elem.AutomationId.Contains("priority", StringComparison.OrdinalIgnoreCase)) ||
                    (elem.Name.Contains("class", StringComparison.OrdinalIgnoreCase) &&
                     elem.Name.Contains(":") &&
                     !elem.Name.Contains("priority", StringComparison.OrdinalIgnoreCase)))
                {
                    data.Class = FindNextClarioValue(elements, i);
                }
            }

            // ACCESSION
            if (string.IsNullOrEmpty(data.Accession))
            {
                if (elem.AutomationId.Contains("accession", StringComparison.OrdinalIgnoreCase) ||
                    (elem.Name.Contains("accession", StringComparison.OrdinalIgnoreCase) &&
                     elem.Name.Contains(":")))
                {
                    data.Accession = FindNextClarioValue(elements, i);
                }
            }

            // Stop early if we found all three
            if (!string.IsNullOrEmpty(data.Priority) &&
                !string.IsNullOrEmpty(data.Class) &&
                !string.IsNullOrEmpty(data.Accession))
                break;
        }

        return data;
    }

    /// <summary>
    /// Find the next non-label value after a label element.
    /// </summary>
    private string FindNextClarioValue(List<ClarioElementData> elements, int startIndex)
    {
        string[] skipValues = { "priority", "class", "accession" };

        for (int j = startIndex + 1; j < Math.Min(startIndex + 10, elements.Count); j++)
        {
            var next = elements[j];

            // Check Name property first
            if (!string.IsNullOrEmpty(next.Name) &&
                !next.Name.Contains(":") &&
                !skipValues.Any(s => next.Name.Equals(s, StringComparison.OrdinalIgnoreCase)))
            {
                return next.Name;
            }

            // Fallback to Text property
            if (!string.IsNullOrEmpty(next.Text) &&
                !next.Text.Contains(":") &&
                !skipValues.Any(s => next.Text.Equals(s, StringComparison.OrdinalIgnoreCase)))
            {
                return next.Text;
            }
        }
        return "";
    }

    /// <summary>
    /// Extract Priority and Class from Clario window.
    /// Uses staggered depth search (12, 18, 25) to find data without over-traversing.
    /// </summary>
    /// <param name="targetAccession">Optional accession to verify match</param>
    /// <returns>Extracted data or null if not found</returns>
    public ClarioPriorityData? ExtractClarioPriorityAndClass(string? targetAccession = null)
    {
        var chromeWindow = FindClarioWindow();
        if (chromeWindow == null)
        {
            Logger.Trace("ExtractClarioPriorityAndClass: Clario window not found");
            return null;
        }

        var data = new ClarioPriorityData();
        int[] searchDepths = { 12, 18, 25 };  // Staggered search

        foreach (var maxDepth in searchDepths)
        {
            try
            {
                var elements = GetClarioElementsRecursive(chromeWindow, maxDepth);
                var extracted = ExtractPriorityDataFromElements(elements);

                // Merge newly found values
                if (string.IsNullOrEmpty(data.Priority)) data.Priority = extracted.Priority;
                if (string.IsNullOrEmpty(data.Class)) data.Class = extracted.Class;
                if (string.IsNullOrEmpty(data.Accession)) data.Accession = extracted.Accession;

                // Stop if we have Priority (minimum needed for stroke detection)
                if (!string.IsNullOrEmpty(data.Priority))
                {
                    Logger.Trace($"ExtractClarioPriorityAndClass: Found Priority='{data.Priority}', Class='{data.Class}', Accession='{data.Accession}' at depth {maxDepth}");
                    break;
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"ExtractClarioPriorityAndClass: Error at depth {maxDepth}: {ex.Message}");
            }
        }

        // Verify accession if provided
        if (targetAccession != null &&
            !string.IsNullOrEmpty(data.Accession) &&
            !data.Accession.Equals(targetAccession, StringComparison.OrdinalIgnoreCase))
        {
            Logger.Trace($"ExtractClarioPriorityAndClass: Accession mismatch - expected '{targetAccession}', got '{data.Accession}'");
            return null;
        }

        // Update stored values
        LastClarioPriority = data.Priority;
        LastClarioClass = data.Class;

        // Check for stroke
        IsStrokeStudy = CheckIfStrokeStudy(data.Priority, data.Class);

        if (string.IsNullOrEmpty(data.Priority) && string.IsNullOrEmpty(data.Class))
        {
            Logger.Trace("ExtractClarioPriorityAndClass: No Priority or Class found");
            return null;
        }

        return data;
    }

    /// <summary>
    /// Check if the study is a Stroke study based on Priority and Class.
    /// </summary>
    private bool CheckIfStrokeStudy(string? priority, string? patientClass)
    {
        // Check Priority field for "Stroke"
        if (!string.IsNullOrEmpty(priority) &&
            priority.Contains("Stroke", StringComparison.OrdinalIgnoreCase))
        {
            Logger.Trace($"CheckIfStrokeStudy: Detected STROKE in Priority field: '{priority}'");
            return true;
        }

        // Check Class field for "Stroke" (less common but possible)
        if (!string.IsNullOrEmpty(patientClass) &&
            patientClass.Contains("Stroke", StringComparison.OrdinalIgnoreCase))
        {
            Logger.Trace($"CheckIfStrokeStudy: Detected STROKE in Class field: '{patientClass}'");
            return true;
        }

        return false;
    }

    /// <summary>
    /// Clear the stroke detection state (call when study changes).
    /// </summary>
    public void ClearStrokeState()
    {
        LastClarioPriority = null;
        LastClarioClass = null;
        IsStrokeStudy = false;
    }

    #region Open Study via XML File Drop

    private const string XmlInFolder = @"C:\MModal\FluencyForImaging\Reporting\XML\IN";

    /// <summary>
    /// Opens a study in Mosaic by writing an OpenReport XML file to Fluency's monitored folder.
    /// This is more reliable than UI automation.
    /// </summary>
    /// <param name="accession">The accession number</param>
    /// <param name="mrn">The Medical Record Number (required)</param>
    /// <returns>True if XML was written successfully, false otherwise</returns>
    public bool OpenStudyInClario(string accession, string? mrn)
    {
        if (string.IsNullOrWhiteSpace(accession))
        {
            Logger.Trace("OpenStudyInClario: No accession provided");
            return false;
        }

        if (string.IsNullOrWhiteSpace(mrn))
        {
            Logger.Trace("OpenStudyInClario: No MRN provided - required for XML file drop");
            return false;
        }

        try
        {
            if (!Directory.Exists(XmlInFolder))
            {
                Logger.Trace($"OpenStudyInClario: XML folder not found: {XmlInFolder}");
                return false;
            }

            var timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            var pid = Environment.ProcessId;
            var filename = $"openreport{timestamp}.{pid}.xml";
            var filepath = Path.Combine(XmlInFolder, filename);

            var xml = $@"<Message>
  <Type>OpenReport</Type>
  <AccessionNumbers>
    <AccessionNumber>{accession}</AccessionNumber>
  </AccessionNumbers>
  <MedicalRecordNumber>{mrn}</MedicalRecordNumber>
</Message>";

            File.WriteAllText(filepath, xml);
            Logger.Trace($"OpenStudyInClario: Wrote XML to {filepath}");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Trace($"OpenStudyInClario: Error writing XML - {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Checks if the Fluency XML folder exists for opening studies.
    /// </summary>
    public bool IsXmlFolderAvailable() => Directory.Exists(XmlInFolder);

    #endregion

    /// <summary>
    /// Creates a Critical Results Communication Note in Clario.
    /// Steps: Click "Create" button → Click "Communication Note" → Click "Submit"
    /// </summary>
    public bool CreateCriticalCommunicationNote()
    {
        var clarioWindow = FindClarioWindow();
        if (clarioWindow == null)
        {
            Logger.Trace("CreateCriticalCommunicationNote: Clario window not found");
            return false;
        }

        try
        {
            // Step 1: Find and click "Create" button (targeted search)
            var createBtn = clarioWindow.FindFirstDescendant(cf =>
                cf.ByName("Create", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));

            if (createBtn == null)
            {
                Logger.Trace("CreateCriticalCommunicationNote: Create button not found");
                return false;
            }

            Logger.Trace($"CreateCriticalCommunicationNote: Found Create - Type={createBtn.ControlType}");
            ClickElement(createBtn, "Create");
            Thread.Sleep(150); // Brief wait for menu to appear

            // Step 2: Find and click "Communication Note" (targeted search)
            // Try Clario window first, then desktop (popup might be separate)
            var commNoteBtn = clarioWindow.FindFirstDescendant(cf =>
                cf.ByName("Communication Note"));

            if (commNoteBtn == null)
            {
                // Try desktop - popup might be a separate window
                var desktop = _automation.GetDesktop();
                commNoteBtn = desktop.FindFirstDescendant(cf =>
                    cf.ByName("Communication Note"));
            }

            if (commNoteBtn == null)
            {
                Logger.Trace("CreateCriticalCommunicationNote: Communication Note option not found");
                return false;
            }

            Logger.Trace($"CreateCriticalCommunicationNote: Found Communication Note - Type={commNoteBtn.ControlType}");
            ClickElement(commNoteBtn, "Communication Note");
            Thread.Sleep(200); // Wait for dialog to appear

            // Step 3: Find and click "Submit" button (targeted search with retry)
            FlaUI.Core.AutomationElements.AutomationElement? submitBtn = null;

            // Retry up to 3 times with short waits
            for (int retry = 0; retry < 3 && submitBtn == null; retry++)
            {
                if (retry > 0) Thread.Sleep(100);

                // Try finding Submit button directly
                submitBtn = clarioWindow.FindFirstDescendant(cf =>
                    cf.ByName("Submit").And(cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)));

                // If not found as Button, try other control types
                if (submitBtn == null)
                {
                    submitBtn = clarioWindow.FindFirstDescendant(cf =>
                        cf.ByName("Submit").And(cf.ByControlType(FlaUI.Core.Definitions.ControlType.Hyperlink)));
                }

                if (submitBtn == null)
                {
                    submitBtn = clarioWindow.FindFirstDescendant(cf =>
                        cf.ByName("Submit").And(cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text)));
                }
            }

            if (submitBtn == null)
            {
                Logger.Trace("CreateCriticalCommunicationNote: Submit button not found after retries");
                return false;
            }

            Logger.Trace($"CreateCriticalCommunicationNote: Found Submit - Type={submitBtn.ControlType}");
            ClickElement(submitBtn, "Submit");

            Logger.Trace("CreateCriticalCommunicationNote: SUCCESS - note created");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CreateCriticalCommunicationNote error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Helper to click an element using Invoke pattern or Click().
    /// </summary>
    private void ClickElement(FlaUI.Core.AutomationElements.AutomationElement element, string name)
    {
        var invokePattern = element.Patterns.Invoke.PatternOrDefault;
        if (invokePattern != null)
        {
            invokePattern.Invoke();
            Logger.Trace($"ClickElement: Clicked {name} via Invoke");
        }
        else
        {
            element.Click();
            Logger.Trace($"ClickElement: Clicked {name} via Click()");
        }
    }

    #endregion

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
    /// Click the "CREATE IMPRESSION" button in Mosaic.
    /// Returns true if successful.
    /// </summary>
    public bool ClickCreateImpression()
    {
        try
        {
            var window = FindMosaicWindow();
            if (window == null)
            {
                Logger.Trace("ClickCreateImpression: Mosaic window not found");
                return false;
            }

            // Find button by name "CREATE IMPRESSION"
            var button = window.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
                  .And(cf.ByName("CREATE IMPRESSION")));

            if (button == null)
            {
                Logger.Trace("ClickCreateImpression: Button not found");
                return false;
            }

            // Try to invoke it (preferred for buttons)
            var invokePattern = button.Patterns.Invoke.PatternOrDefault;
            if (invokePattern != null)
            {
                invokePattern.Invoke();
                Logger.Trace("ClickCreateImpression: Invoked via pattern");
                return true;
            }

            // Fallback: click at center
            var rect = button.BoundingRectangle;
            if (rect.Width > 0 && rect.Height > 0)
            {
                button.Click();
                Logger.Trace($"ClickCreateImpression: Clicked at {rect}");
                return true;
            }

            Logger.Trace("ClickCreateImpression: Button has no valid bounds");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Trace($"ClickCreateImpression error: {ex.Message}");
            return false;
        }
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
            LastPatientName = null; // Reset - will be set if found
            LastSiteCode = null; // Reset - will be set if found
            LastMrn = null; // Reset - will be set if found

            // Extract accession: find "Current Study" text and the next text element after it
            try
            {
                var currentStudyElement = _cachedSlimHubWindow.FindFirstDescendant(cf =>
                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text)
                    .And(cf.ByName("Current Study")));

                if (currentStudyElement != null)
                {
                    // Get parent and find all children (Text in 2.0.2, Button in 2.0.3)
                    var parent = currentStudyElement.Parent;
                    if (parent != null)
                    {
                        var allChildren = parent.FindAllChildren();

                        bool foundCurrentStudy = false;
                        foreach (var child in allChildren)
                        {
                            try
                            {
                                if (foundCurrentStudy)
                                {
                                    var accession = child.Name?.Trim();
                                    // Skip status words and empty elements
                                    if (string.IsNullOrWhiteSpace(accession) || accession == "Current Study" ||
                                        accession == "DRAFTED" || accession == "UNDRAFTED" || accession == "SIGNED")
                                    {
                                        continue; // Keep looking for accession
                                    }
                                    LastAccession = accession;
                                    Logger.Trace($"Found Accession: {LastAccession}");
                                    break;
                                }
                                if (child.Name == "Current Study")
                                {
                                    foundCurrentStudy = true;
                                }
                            }
                            catch { continue; }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Accession extraction error: {ex.Message}");
            }


            // Extract patient info from all descendants in a single traversal.
            // Handles both Mosaic 2.0.2 (combined Text "Description: XR CHEST") and
            // 2.0.3 (separate Text "Description:" + Button "XR CHEST") in one pass.
            // BEGIN Mosaic 2.0.2 compat: When 2.0.2 is retired, the pendingInfoField/Button
            // logic becomes the only path and the Text-based regex extraction can be removed.
            try
            {
                var allInfoElements = _cachedSlimHubWindow.FindAllDescendants();
                string? pendingInfoField = null; // Tracks label→button pairs for 2.0.3

                foreach (var el in allInfoElements)
                {
                    try
                    {
                        var ctrlType = el.ControlType;
                        var name = el.Name ?? "";

                        if (ctrlType == FlaUI.Core.Definitions.ControlType.Text)
                        {
                            var textUpper = name.ToUpperInvariant();

                            // Gender extraction
                            if (LastPatientGender == null)
                            {
                                if (textUpper.StartsWith("MALE, AGE") || textUpper.StartsWith("MALE,AGE"))
                                {
                                    LastPatientGender = "Male";
                                    Logger.Trace($"Found Patient Gender: Male");
                                }
                                else if (textUpper.StartsWith("FEMALE, AGE") || textUpper.StartsWith("FEMALE,AGE"))
                                {
                                    LastPatientGender = "Female";
                                    Logger.Trace($"Found Patient Gender: Female");
                                }
                            }

                            // Site code extraction: pattern "Site Code: XXX" (2.0.2: value in same element)
                            if (LastSiteCode == null)
                            {
                                var siteMatch = Regex.Match(name, @"Site\s*Code:\s*([A-Z]{2,5})", RegexOptions.IgnoreCase);
                                if (siteMatch.Success)
                                {
                                    LastSiteCode = siteMatch.Groups[1].Value.ToUpperInvariant();
                                    Logger.Trace($"Found Site Code: {LastSiteCode}");
                                }
                            }

                            // MRN extraction: pattern "MRN: XXX" (2.0.2: value in same element)
                            if (LastMrn == null)
                            {
                                var mrnMatch = Regex.Match(name, @"MRN:\s*([A-Z0-9]{5,20})", RegexOptions.IgnoreCase);
                                if (mrnMatch.Success)
                                {
                                    LastMrn = mrnMatch.Groups[1].Value.ToUpperInvariant();
                                    Logger.Trace($"Found MRN: {LastMrn}");
                                }
                            }

                            // Patient name extraction
                            if (LastPatientName == null && IsPatientNameCandidate(textUpper))
                            {
                                LastPatientName = ToTitleCase(textUpper);
                                Logger.Trace($"Found Patient Name: {LastPatientName}");
                            }

                            // Description extraction (2.0.2: "Description: CT ABDOMEN PELVIS...")
                            if (LastDescription == null && name.StartsWith("Description:", StringComparison.OrdinalIgnoreCase))
                            {
                                var desc = name.Substring("Description:".Length).Trim();
                                if (!string.IsNullOrEmpty(desc))
                                {
                                    LastDescription = desc;
                                    Logger.Trace($"Found Description: {LastDescription}");
                                }
                            }

                            // Mosaic 2.0.3: Track label-only Text elements for button-based extraction.
                            // In 2.0.3, "Description:" is a separate element from the value Button.
                            var trimmed = name.TrimEnd();
                            if (string.IsNullOrEmpty(LastDescription) && trimmed.Equals("Description:", StringComparison.OrdinalIgnoreCase))
                                pendingInfoField = "Description";
                            else if (LastMrn == null && trimmed.Equals("MRN:", StringComparison.OrdinalIgnoreCase))
                                pendingInfoField = "MRN";
                            else if (LastSiteCode == null && trimmed.Equals("Site Code:", StringComparison.OrdinalIgnoreCase))
                                pendingInfoField = "SiteCode";
                            else
                                pendingInfoField = null;
                        }
                        else if (ctrlType == FlaUI.Core.Definitions.ControlType.Button && pendingInfoField != null)
                        {
                            // Mosaic 2.0.3: Button value following a label Text element
                            var value = name.Trim();
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                switch (pendingInfoField)
                                {
                                    case "Description":
                                        if (string.IsNullOrEmpty(LastDescription))
                                        {
                                            LastDescription = value;
                                            Logger.Trace($"Found Description (v2.0.3): {value}");
                                        }
                                        break;
                                    case "MRN":
                                        if (LastMrn == null)
                                        {
                                            LastMrn = value.ToUpperInvariant();
                                            Logger.Trace($"Found MRN (v2.0.3): {LastMrn}");
                                        }
                                        break;
                                    case "SiteCode":
                                        if (LastSiteCode == null)
                                        {
                                            LastSiteCode = value.ToUpperInvariant();
                                            Logger.Trace($"Found Site Code (v2.0.3): {LastSiteCode}");
                                        }
                                        break;
                                }
                            }
                            pendingInfoField = null;
                        }
                        else
                        {
                            pendingInfoField = null; // Reset on non-Text/non-Button elements
                        }
                    }
                    catch { continue; }
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Patient info extraction error: {ex.Message}");
            }

            // END Mosaic 2.0.2 compat

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
                        var descVal = el.Name.Substring("Description:".Length).Trim();
                        if (!string.IsNullOrEmpty(descVal))  // Mosaic 2.0.3: label is separate from value
                        {
                            LastDescription = descVal;
                            Logger.Trace($"Found Description: {LastDescription}");
                        }
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

            // BEGIN Mosaic 2.0.2 compat: In 2.0.3, the Report Document is named
            // "Report {id} | RADPAIR" instead of just "Report". The substring search above
            // should match, but if it doesn't (e.g., cross-iframe accessibility issues),
            // fall back to finding any non-SlimHub Document element.
            // Remove this block when Mosaic 2.0.2 is fully retired.
            if (reportDoc == null)
            {
                try
                {
                    var documents = _cachedSlimHubWindow.FindAllDescendants(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Document));

                    foreach (var doc in documents)
                    {
                        try
                        {
                            var docName = doc.Name ?? "";
                            if (!docName.Equals("SlimHub", StringComparison.OrdinalIgnoreCase) &&
                                !string.IsNullOrWhiteSpace(docName))
                            {
                                reportDoc = doc;
                                Logger.Trace($"GetFinalReportFast: Found Report Document (v2.0.3 fallback): '{docName}'");
                                break;
                            }
                        }
                        catch { continue; }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Trace($"GetFinalReportFast: v2.0.3 Document search error: {ex.Message}");
                }
            }
            // END Mosaic 2.0.2 compat

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

                    // Mosaic 2.0.3: ProseMirror Name is empty; content is in child Text elements
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        try
                        {
                            var childTexts = candidate.FindAllDescendants(cf =>
                                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text));
                            if (childTexts.Length > 0)
                            {
                                var sb = new System.Text.StringBuilder();
                                foreach (var ct in childTexts)
                                {
                                    var childName = ct.Name;
                                    if (!string.IsNullOrWhiteSpace(childName))
                                        sb.AppendLine(childName);
                                }
                                text = sb.ToString();
                            }
                        }
                        catch { }
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

                // Mosaic 2.0.3 fallback: EXAM: heading may be absent (e.g., addendum view).
                // Pick the ProseMirror with the most report keywords, ignoring EXAM: requirement.
                if (maxScore <= 0)
                {
                    string bestFallback = "";
                    int bestFallbackScore = 0;
                    foreach (var candidate in candidates)
                    {
                        string text = candidate.Name ?? "";
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            try
                            {
                                var childTexts = candidate.FindAllDescendants(cf =>
                                    cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text));
                                if (childTexts.Length > 0)
                                {
                                    var sb = new System.Text.StringBuilder();
                                    foreach (var ct in childTexts)
                                    {
                                        var childName = ct.Name;
                                        if (!string.IsNullOrWhiteSpace(childName))
                                            sb.AppendLine(childName);
                                    }
                                    text = sb.ToString();
                                }
                            }
                            catch { continue; }
                        }
                        if (string.IsNullOrWhiteSpace(text) || text.Length < 50) continue;

                        int score = 0;
                        foreach (var kw in keywords)
                        {
                            if (text.Contains(kw, StringComparison.OrdinalIgnoreCase)) score++;
                        }
                        if (score > bestFallbackScore)
                        {
                            bestFallbackScore = score;
                            bestFallback = text;
                        }
                    }

                    if (!string.IsNullOrEmpty(bestFallback) && bestFallbackScore > 0)
                    {
                        sw.Stop();
                        int lineCount = bestFallback.Split('\n').Length;
                        Logger.Trace($"GetFinalReportFast: ProseMirror v2.0.3 fallback SUCCESS in {sw.ElapsedMilliseconds}ms, {lineCount} lines, Score={bestFallbackScore}");
                        LastFinalReport = bestFallback;
                        LastTemplateName = ExtractTemplateName(bestFallback);
                        return bestFallback;
                    }
                }
            }
            
            if (candidates.Length > 0)
            {
                try
                {
                    var sampleName = candidates[0].Name ?? "";
                    if (sampleName.Length > 80) sampleName = sampleName.Substring(0, 80);
                    Logger.Trace($"GetFinalReportFast: ProseMirror candidates found but no EXAM: match. First candidate: '{sampleName}'");
                }
                catch { }
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

    #region Discard Study Action

    /// <summary>
    /// Perform the full discard study flow:
    /// 1. Click ACTIONS button
    /// 2. Click "Discard Changes" menu item
    /// 3. Click "YES, DISCARD" button
    /// Returns true if successful.
    /// </summary>
    public bool ClickDiscardStudy()
    {
        try
        {
            var window = _cachedSlimHubWindow ?? FindMosaicWindow();
            if (window == null)
            {
                Logger.Trace("ClickDiscardStudy: Mosaic window not found");
                return false;
            }

            // Step 1: Click ACTIONS button
            var actionsButton = window.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
                  .And(cf.ByName("ACTIONS")));

            if (actionsButton == null)
            {
                Logger.Trace("ClickDiscardStudy: ACTIONS button not found");
                return false;
            }

            var invokePattern = actionsButton.Patterns.Invoke.PatternOrDefault;
            if (invokePattern != null)
            {
                invokePattern.Invoke();
                Logger.Trace("ClickDiscardStudy: Clicked ACTIONS button");
            }
            else
            {
                actionsButton.Click();
                Logger.Trace("ClickDiscardStudy: Clicked ACTIONS button (fallback)");
            }

            // Wait for menu to appear
            Thread.Sleep(400);

            // Step 2: Click "Discard Changes" menu item
            var discardMenuItem = window.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem)
                  .And(cf.ByName("Discard Changes")));

            if (discardMenuItem == null)
            {
                Logger.Trace("ClickDiscardStudy: Discard Changes menu item not found");
                return false;
            }

            invokePattern = discardMenuItem.Patterns.Invoke.PatternOrDefault;
            if (invokePattern != null)
            {
                invokePattern.Invoke();
                Logger.Trace("ClickDiscardStudy: Clicked Discard Changes menu item");
            }
            else
            {
                discardMenuItem.Click();
                Logger.Trace("ClickDiscardStudy: Clicked Discard Changes menu item (fallback)");
            }

            // Wait for dialog to appear
            Thread.Sleep(500);

            // Step 3: Click "YES, DISCARD" button
            var yesDiscardButton = window.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
                  .And(cf.ByName("YES, DISCARD")));

            if (yesDiscardButton == null)
            {
                Logger.Trace("ClickDiscardStudy: YES, DISCARD button not found");
                return false;
            }

            invokePattern = yesDiscardButton.Patterns.Invoke.PatternOrDefault;
            if (invokePattern != null)
            {
                invokePattern.Invoke();
                Logger.Trace("ClickDiscardStudy: Clicked YES, DISCARD button");
            }
            else
            {
                yesDiscardButton.Click();
                Logger.Trace("ClickDiscardStudy: Clicked YES, DISCARD button (fallback)");
            }

            Logger.Trace("ClickDiscardStudy: SUCCESS - study discarded");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Trace($"ClickDiscardStudy error: {ex.Message}");
            return false;
        }
    }

    #endregion

    #region Discard Dialog Detection

    /// <summary>
    /// Check if the "Confirm Discard Action?" dialog is currently visible.
    /// Used to detect when a study is being discarded (not signed).
    /// </summary>
    public bool IsDiscardDialogVisible()
    {
        try
        {
            // Use cached window if available, otherwise find it
            var window = _cachedSlimHubWindow ?? FindMosaicWindow();
            if (window == null) return false;

            // Look for the dialog window with name "Confirm Discard Action?"
            var discardDialog = window.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window)
                  .And(cf.ByName("Confirm Discard Action?")));

            if (discardDialog != null)
            {
                Logger.Trace("Discard dialog detected (by Window name)");
                return true;
            }

            // Fallback: look for the "YES, DISCARD" button
            var discardButton = window.FindFirstDescendant(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Button)
                  .And(cf.ByName("YES, DISCARD")));

            if (discardButton != null)
            {
                Logger.Trace("Discard dialog detected (by YES, DISCARD button)");
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            Logger.Trace($"IsDiscardDialogVisible error: {ex.Message}");
            return false;
        }
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
        "SPINE", "UPPER EXTREMITY", "LOWER EXTREMITY", // No generic "EXTREMITY" to avoid false matches
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

        // Normalize runoff studies - AORTA + RUNOFF is equivalent to ABDOMEN/PELVIS + LOWER EXTREMITY
        NormalizeRunoffStudy(descParts);
        NormalizeRunoffStudy(templateParts);

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

    /// <summary>
    /// Normalize runoff study body parts so equivalent descriptions match.
    /// AORTA + RUNOFF → ABDOMEN, PELVIS, LOWER EXTREMITY, RUNOFF (AORTA removed as it's implied)
    /// </summary>
    private static void NormalizeRunoffStudy(HashSet<string> parts)
    {
        if (!parts.Contains("RUNOFF")) return;

        // If AORTA is present with RUNOFF, expand to standard runoff parts
        // and remove AORTA since it's implied by abdomen/pelvis runoff
        if (parts.Contains("AORTA"))
        {
            parts.Remove("AORTA"); // AORTA is implied in runoff studies
            parts.Add("ABDOMEN");
            parts.Add("PELVIS");
            parts.Add("LOWER EXTREMITY");
        }

        // If ABDOMEN/PELVIS is present with RUNOFF, add LOWER EXTREMITY
        if ((parts.Contains("ABDOMEN") || parts.Contains("PELVIS")) && !parts.Contains("LOWER EXTREMITY"))
        {
            parts.Add("LOWER EXTREMITY");
        }

        // If LOWER EXTREMITY is present with RUNOFF, add ABDOMEN/PELVIS
        if (parts.Contains("LOWER EXTREMITY"))
        {
            parts.Add("ABDOMEN");
            parts.Add("PELVIS");
        }
    }

    #endregion

    #region Patient Name Extraction

    /// <summary>
    /// Common UI/medical terms to exclude from patient name detection.
    /// </summary>
    private static readonly HashSet<string> PatientNameExclusions = new(StringComparer.OrdinalIgnoreCase)
    {
        "CURRENT STUDY", "SITE CODE", "FINDINGS", "EXAM", "IMPRESSION", "TECHNIQUE",
        "CLINICAL HISTORY", "COMPARISON", "INDICATION", "CONTRAST", "ADDENDUM",
        "REPORT", "ACCESSION", "PATIENT", "PHYSICIAN", "RADIOLOGIST", "DOCTOR",
        "MALE", "FEMALE", "AGE", "DOB", "MRN", "STATUS", "PRIORITY", "CLASS",
        "EMERGENCY", "INPATIENT", "OUTPATIENT", "STAT", "ROUTINE", "STROKE",
        "DRAFTED", "SIGNED", "PENDING", "FINAL", "PRELIMINARY", "ACTIONS",
        "CT", "MRI", "MR", "XR", "US", "PET", "NM", "FLUORO", "MAMMOGRAM",
        "HEAD", "NECK", "CHEST", "ABDOMEN", "PELVIS", "SPINE", "BRAIN",
        "WITHOUT", "WITH", "CONTRAST", "IV", "ORAL", "PROTOCOL",
        "MY", "MACROS", "MY MACROS", "TEMPLATES", "MY TEMPLATES", "FAVORITES",
        "MY FAVORITES", "QUICK", "PICK", "PICK LIST", "TOOLS", "SETTINGS"
    };

    /// <summary>
    /// Check if text is a valid patient name candidate.
    /// Pattern: All-caps 2-4 word string (e.g., "SMITH JOHN" or "SMITH JOHN MICHAEL JR")
    /// </summary>
    private static bool IsPatientNameCandidate(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;

        text = text.Trim();

        // Length check: 5-50 characters
        if (text.Length < 5 || text.Length > 50) return false;

        // Split into words
        var words = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        // Must be 2-4 words
        if (words.Length < 2 || words.Length > 4) return false;

        // Check if excluded term
        if (PatientNameExclusions.Contains(text)) return false;

        // Check each word: must be all uppercase letters (with optional hyphens/apostrophes)
        foreach (var word in words)
        {
            // Skip suffixes like "JR", "SR", "II", "III"
            if (word.Length <= 3 && (word == "JR" || word == "SR" || word == "II" || word == "III" || word == "IV"))
                continue;

            // Must be at least 2 characters (for actual name parts)
            if (word.Length < 2) return false;

            // Check each character: must be A-Z, hyphen, or apostrophe
            foreach (var ch in word)
            {
                if (!((ch >= 'A' && ch <= 'Z') || ch == '-' || ch == '\''))
                    return false;
            }

            // Check if word is an excluded term
            if (PatientNameExclusions.Contains(word))
                return false;
        }

        return true;
    }

    /// <summary>
    /// Convert all-caps text to title case (e.g., "SMITH JOHN" → "Smith John")
    /// </summary>
    private static string ToTitleCase(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return text;

        var words = text.Split(' ');
        for (int i = 0; i < words.Length; i++)
        {
            if (words[i].Length > 0)
            {
                words[i] = char.ToUpper(words[i][0]) +
                           (words[i].Length > 1 ? words[i].Substring(1).ToLower() : "");
            }
        }
        return string.Join(" ", words);
    }

    #endregion

    public void Dispose()
    {
        _automation.Dispose();
    }
}
