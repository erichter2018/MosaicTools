using System.Text;
using System.Runtime.InteropServices;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;

class Program
{
    static StringBuilder _output = new StringBuilder();

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern IntPtr FindWindow(string? lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(IntPtr hWnd);

    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    static IntPtr _mosaicHwnd = IntPtr.Zero;

    [STAThread]
    static void Main(string[] args)
    {
        Console.WriteLine("Mosaic Diagnostic Tool v3");
        Console.WriteLine("=========================");
        Console.WriteLine();

        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            $"mosaic_diagnostic_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

        try
        {
            Log("Mosaic Diagnostic Tool v3");
            Log($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            Log($"Computer: {Environment.MachineName}");
            Log($"User: {Environment.UserName}");
            Log("=" + new string('=', 60));
            Log("");

            // First, find and activate Mosaic window using Win32
            Console.WriteLine("Finding Mosaic window...");
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd)) return true;

                var sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                var title = sb.ToString();

                if (title.Contains("Mosaic Reporting", StringComparison.OrdinalIgnoreCase) ||
                    title.Contains("Mosaic Info Hub", StringComparison.OrdinalIgnoreCase))
                {
                    _mosaicHwnd = hWnd;
                    return false; // stop enumeration
                }
                return true;
            }, IntPtr.Zero);

            if (_mosaicHwnd != IntPtr.Zero)
            {
                Console.WriteLine("Bringing Mosaic to foreground...");
                SetForegroundWindow(_mosaicHwnd);
                Thread.Sleep(1000); // Give it time to activate and render
                Console.WriteLine("Waiting for UI to settle...");
                Thread.Sleep(1000); // Extra wait for WebView2 content
            }

            using var automation = new UIA3Automation();

            // Find Mosaic window with FlaUI
            Console.WriteLine("Scanning with UI Automation...");
            var desktop = automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));

            AutomationElement? mosaicWindow = null;

            Log("ALL WINDOWS:");
            Log("-" + new string('-', 40));
            foreach (var window in windows)
            {
                try
                {
                    var title = window.Name ?? "";
                    if (!string.IsNullOrWhiteSpace(title))
                    {
                        Log($"  {title}");
                    }

                    var titleLower = title.ToLowerInvariant();
                    if ((titleLower.Contains("mosaic") && titleLower.Contains("info hub")) ||
                        (titleLower.Contains("mosaic") && titleLower.Contains("reporting")))
                    {
                        mosaicWindow = window;
                        Console.WriteLine($"Found Mosaic: {title}");
                    }
                }
                catch { }
            }
            Log("");

            if (mosaicWindow == null)
            {
                Log("ERROR: Mosaic window not found!");
                Log("Make sure Mosaic Reporting is open and has a study loaded.");
                Console.WriteLine("\nMosaic window not found!");
            }
            else
            {
                Log($"MOSAIC WINDOW: {mosaicWindow.Name}");
                Log("=" + new string('=', 60));
                Log("");

                // === RAW ELEMENT COUNT ===
                Log("RAW ELEMENT SCAN (no filtering):");
                try
                {
                    var allElements = mosaicWindow.FindAllDescendants();
                    Log($"  Total descendants (any type): {allElements.Length}");

                    // Count by control type
                    var typeCounts = new Dictionary<string, int>();
                    foreach (var el in allElements)
                    {
                        try
                        {
                            var typeName = el.ControlType.ToString();
                            if (!typeCounts.ContainsKey(typeName))
                                typeCounts[typeName] = 0;
                            typeCounts[typeName]++;
                        }
                        catch { }
                    }

                    Log("  Elements by type:");
                    foreach (var kv in typeCounts.OrderByDescending(x => x.Value))
                    {
                        Log($"    {kv.Key}: {kv.Value}");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  ERROR: {ex.Message}");
                }
                Log("");

                // === TEST 1: Find "Current Study" and next sibling (Accession) ===
                Log("TEST 1: Looking for 'Current Study' text element...");
                try
                {
                    var currentStudyElement = mosaicWindow.FindFirstDescendant(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text)
                        .And(cf.ByName("Current Study")));

                    if (currentStudyElement != null)
                    {
                        Log("  FOUND 'Current Study' element!");
                        var parent = currentStudyElement.Parent;
                        if (parent != null)
                        {
                            var siblings = parent.FindAllChildren(cf =>
                                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text));
                            Log($"  Parent has {siblings.Length} Text children:");
                            foreach (var sib in siblings)
                            {
                                Log($"    - \"{Truncate(sib.Name, 100)}\"");
                            }
                        }
                    }
                    else
                    {
                        Log("  NOT FOUND");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  ERROR: {ex.Message}");
                }
                Log("");

                // === TEST 2: Find Description element ===
                Log("TEST 2: Looking for 'Description:' text element...");
                try
                {
                    var descriptionCondition = mosaicWindow.ConditionFactory
                        .ByControlType(FlaUI.Core.Definitions.ControlType.Text)
                        .And(mosaicWindow.ConditionFactory.ByName("Description:", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));

                    var descElements = mosaicWindow.FindAllDescendants(descriptionCondition);
                    Log($"  Found {descElements.Length} elements matching 'Description:'");
                    foreach (var el in descElements)
                    {
                        Log($"    - \"{Truncate(el.Name, 200)}\"");
                    }
                    if (descElements.Length == 0)
                    {
                        Log("  NOT FOUND - This is why macros can't match study types!");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  ERROR: {ex.Message}");
                }
                Log("");

                // === TEST 3: Find Patient Gender ===
                Log("TEST 3: Looking for gender text (MALE/FEMALE, AGE)...");
                try
                {
                    var textElements = mosaicWindow.FindAllDescendants(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text));

                    Log($"  Found {textElements.Length} total Text elements");

                    bool foundGender = false;
                    foreach (var textEl in textElements)
                    {
                        try
                        {
                            var text = textEl.Name?.ToUpperInvariant() ?? "";
                            if (text.StartsWith("MALE, AGE") || text.StartsWith("MALE,AGE") ||
                                text.StartsWith("FEMALE, AGE") || text.StartsWith("FEMALE,AGE"))
                            {
                                Log($"  FOUND GENDER: \"{textEl.Name}\"");
                                foundGender = true;
                            }
                        }
                        catch { }
                    }
                    if (!foundGender)
                    {
                        Log("  NOT FOUND");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  ERROR: {ex.Message}");
                }
                Log("");

                // === TEST 4: Find DRAFTED status ===
                Log("TEST 4: Looking for 'DRAFTED' text element...");
                try
                {
                    var draftedElement = mosaicWindow.FindFirstDescendant(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text)
                        .And(cf.ByName("DRAFTED")));

                    if (draftedElement != null)
                    {
                        Log("  FOUND 'DRAFTED' - report is in drafted state");
                    }
                    else
                    {
                        Log("  NOT FOUND - report may not be drafted yet");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  ERROR: {ex.Message}");
                }
                Log("");

                // === TEST 5: Find Report Document ===
                Log("TEST 5: Looking for Report Document element...");
                try
                {
                    var reportDoc = mosaicWindow.FindFirstDescendant(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Document)
                        .And(cf.ByName("Report", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)));

                    if (reportDoc != null)
                    {
                        Log($"  FOUND Report Document: \"{Truncate(reportDoc.Name, 100)}\"");
                    }
                    else
                    {
                        Log("  NOT FOUND");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  ERROR: {ex.Message}");
                }
                Log("");

                // === DUMP ALL TEXT ELEMENTS ===
                Log("=" + new string('=', 60));
                Log("ALL TEXT ELEMENTS IN MOSAIC:");
                Log("-" + new string('-', 40));
                try
                {
                    var allText = mosaicWindow.FindAllDescendants(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Text));

                    Log($"Total: {allText.Length} Text elements");
                    Log("");

                    int count = 0;
                    foreach (var el in allText)
                    {
                        try
                        {
                            var name = el.Name ?? "";
                            if (string.IsNullOrWhiteSpace(name)) continue;

                            count++;
                            // Flag potentially interesting ones
                            string flag = "";
                            var upper = name.ToUpperInvariant();
                            if (upper.Contains("DESCRIPTION")) flag = " *** DESCRIPTION ***";
                            else if (upper.Contains("STUDY")) flag = " [study]";
                            else if (upper.Contains("EXAM")) flag = " [exam]";
                            else if (upper.Contains("CT ") || upper.Contains("MR ") || upper.Contains("US ") || upper.Contains("XR ")) flag = " [modality?]";
                            else if (upper.Contains("CHEST") || upper.Contains("ABDOMEN") || upper.Contains("HEAD") || upper.Contains("BRAIN")) flag = " [body part?]";

                            Log($"{count}. \"{Truncate(name, 200)}\"{flag}");

                            if (count > 300)
                            {
                                Log($"... (truncated, {allText.Length - count} more)");
                                break;
                            }
                        }
                        catch { }
                    }
                }
                catch (Exception ex)
                {
                    Log($"ERROR scanning text elements: {ex.Message}");
                }

                // === DUMP FIRST 50 ELEMENTS OF ANY TYPE ===
                Log("");
                Log("=" + new string('=', 60));
                Log("FIRST 50 ELEMENTS (any type) with names:");
                Log("-" + new string('-', 40));
                try
                {
                    var allElements = mosaicWindow.FindAllDescendants();
                    int count = 0;
                    foreach (var el in allElements)
                    {
                        try
                        {
                            var name = el.Name ?? "";
                            if (string.IsNullOrWhiteSpace(name)) continue;

                            count++;
                            Log($"{count}. [{el.ControlType}] \"{Truncate(name, 150)}\"");

                            if (count >= 50) break;
                        }
                        catch { }
                    }
                    if (count == 0)
                    {
                        Log("  (No named elements found!)");
                    }
                }
                catch (Exception ex)
                {
                    Log($"ERROR: {ex.Message}");
                }

                Console.WriteLine($"\nScan complete.");
            }

            // Write output
            File.WriteAllText(outputPath, _output.ToString());
            Console.WriteLine($"\nOutput saved to:\n{outputPath}");
            Console.WriteLine("\nPlease send this file to Erik for analysis.");
        }
        catch (Exception ex)
        {
            Log($"FATAL ERROR: {ex.Message}");
            Log(ex.StackTrace ?? "");
            Console.WriteLine($"\nFatal Error: {ex.Message}");

            try { File.WriteAllText(outputPath, _output.ToString()); }
            catch { }
        }

        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }

    static void Log(string message)
    {
        _output.AppendLine(message);
    }

    static string Truncate(string? text, int maxLen)
    {
        if (string.IsNullOrEmpty(text)) return "";
        text = text.Replace("\r", "\\r").Replace("\n", "\\n");
        if (text.Length > maxLen)
            return text.Substring(0, maxLen) + "...";
        return text;
    }
}
