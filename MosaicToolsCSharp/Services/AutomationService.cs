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
    public string? PerformClarioScrape(Action<string>? statusCallback = null)
    {
        var window = FindClarioWindow();
        if (window == null)
        {
            Logger.Trace("Clario window NOT found.");
            return null;
        }
        
        // Ensure window is focused
        try
        {
            Logger.Trace("Focusing Clario window before scrape...");
            window.Focus();
            Thread.Sleep(500);
        }
        catch { }
        
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
    
    public void Dispose()
    {
        _automation.Dispose();
    }
}
