using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace AidocInspector;

class Program
{
    static readonly StringBuilder _output = new();
    static int _maxDepth = 14;
    static int _elementCount = 0;
    static int _maxElements = 20000;

    [DllImport("user32.dll")]
    static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool IsWindow(IntPtr hWnd);

    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Log("=== Aidoc UI Automation Inspector v3 ===");
        Log($"Time: {DateTime.Now}");
        Log("");

        // --- Step 0: Find all Aidoc processes ---
        Log("--- Step 0: Aidoc processes ---");
        var aidocPids = new HashSet<uint>();
        foreach (var proc in Process.GetProcesses())
        {
            try
            {
                var procName = proc.ProcessName.ToLower();
                if (procName.Contains("aidoc"))
                {
                    aidocPids.Add((uint)proc.Id);
                    string module = "";
                    try { module = proc.MainModule?.FileName ?? ""; } catch { }
                    Log($"  PID={proc.Id} Name=\"{proc.ProcessName}\" Title=\"{TruncateStr(proc.MainWindowTitle, 80)}\" Module=\"{TruncateStr(module, 120)}\"");
                }
            }
            catch { }
        }
        // Remove our own PID
        aidocPids.Remove((uint)Environment.ProcessId);
        Log($"  Aidoc PIDs: [{string.Join(", ", aidocPids)}]");

        // --- Step 0.5: Use Win32 EnumWindows to find ALL windows from Aidoc PIDs ---
        Log("\n--- Step 0.5: Win32 EnumWindows for Aidoc PIDs ---");
        var aidocHwnds = new List<(IntPtr hwnd, string title, string className, uint pid, bool visible)>();
        EnumWindows((hwnd, param) =>
        {
            GetWindowThreadProcessId(hwnd, out uint pid);
            if (aidocPids.Contains(pid))
            {
                var titleSb = new StringBuilder(256);
                GetWindowText(hwnd, titleSb, 256);
                var classSb = new StringBuilder(256);
                GetClassName(hwnd, classSb, 256);
                bool visible = IsWindowVisible(hwnd);
                aidocHwnds.Add((hwnd, titleSb.ToString(), classSb.ToString(), pid, visible));
            }
            return true;
        }, IntPtr.Zero);

        foreach (var (hwnd, title, cls, pid, visible) in aidocHwnds)
        {
            Log($"  HWND=0x{hwnd:X} Title=\"{TruncateStr(title, 100)}\" Class=\"{cls}\" PID={pid} Visible={visible}");
        }
        Log($"  Total HWNDs: {aidocHwnds.Count}");

        using var automation = new UIA3Automation();
        var cf = automation.ConditionFactory;

        // --- Step 1: Use FlaUI to inspect each HWND ---
        Log("\n--- Step 1: FlaUI inspection of each Aidoc HWND ---");
        var inspectedWindows = new List<(AutomationElement el, string title, bool visible)>();

        foreach (var (hwnd, title, cls, pid, visible) in aidocHwnds)
        {
            try
            {
                var el = automation.FromHandle(hwnd);
                if (el != null)
                {
                    string name = "";
                    try { name = el.Name ?? ""; } catch { }
                    string elClass = "";
                    try { elClass = el.ClassName ?? ""; } catch { }
                    string bounds = SafeGetBounds(el);
                    string ctrlType = SafeGetControlType(el);

                    Log($"\n  HWND=0x{hwnd:X} FlaUI: [{ctrlType}] Name=\"{TruncateStr(name, 100)}\" Class=\"{elClass}\" {bounds} Visible={visible}");

                    // Count children
                    try
                    {
                        var children = el.FindAllChildren();
                        Log($"    Children: {children.Length}");
                    }
                    catch (Exception ex)
                    {
                        Log($"    Children error: {ex.Message}");
                    }

                    inspectedWindows.Add((el, title, visible));
                }
            }
            catch (Exception ex)
            {
                Log($"  HWND=0x{hwnd:X} FlaUI error: {ex.Message}");
            }
        }

        // --- Step 2: Deep walk the most interesting windows ---
        // Walk all visible windows, plus any with interesting titles
        Log($"\n{new string('=', 80)}");
        Log("DEEP TREE WALKS");
        Log(new string('=', 80));

        foreach (var (el, title, visible) in inspectedWindows)
        {
            _elementCount = 0;
            string name = "";
            try { name = el.Name ?? ""; } catch { }
            string bounds = SafeGetBounds(el);

            // Skip tiny/invisible windows unless they have a title
            bool isInteresting = visible || !string.IsNullOrEmpty(title) || !string.IsNullOrEmpty(name);
            if (!isInteresting) continue;

            Log($"\n--- Tree: \"{TruncateStr(name, 80)}\" (title=\"{TruncateStr(title, 80)}\") Visible={visible} {bounds} ---");
            WalkTree(el, 0);
            Log($"  Elements: {_elementCount}");
        }

        // --- Step 3: Text content scan for all windows ---
        Log($"\n{new string('=', 80)}");
        Log("TEXT CONTENT SCAN (all windows)");
        Log(new string('=', 80));

        foreach (var (el, title, visible) in inspectedWindows)
        {
            string name = "";
            try { name = el.Name ?? ""; } catch { }

            bool isInteresting = visible || !string.IsNullOrEmpty(title) || !string.IsNullOrEmpty(name);
            if (!isInteresting) continue;

            Log($"\n--- Text scan: \"{TruncateStr(name, 80)}\" ---");
            _elementCount = 0;
            ScanForText(el, 0);
        }

        SaveOutput();
    }

    static void WalkTree(AutomationElement element, int depth)
    {
        if (depth > _maxDepth || _elementCount > _maxElements) return;
        _elementCount++;

        string indent = new string(' ', depth * 2);
        Log($"{indent}{GetElementInfo(element)}");

        try
        {
            var children = element.FindAllChildren();
            foreach (var child in children)
                WalkTree(child, depth + 1);
        }
        catch (Exception ex)
        {
            Log($"{indent}  [Children error: {ex.Message}]");
        }
    }

    static void ScanForText(AutomationElement element, int depth)
    {
        if (depth > _maxDepth || _elementCount > _maxElements) return;
        _elementCount++;

        try
        {
            string name = "", automationId = "", controlType = "", className = "";
            try { name = element.Name ?? ""; } catch { }
            try { automationId = element.AutomationId ?? ""; } catch { }
            try { controlType = SafeGetControlType(element); } catch { }
            try { className = element.ClassName ?? ""; } catch { }

            string? valueStr = null, textStr = null, legacyVal = null, legacyName = null;

            try { if (element.Patterns.Value.IsSupported) valueStr = element.Patterns.Value.Pattern.Value.ValueOrDefault; } catch { }
            try { if (element.Patterns.Text.IsSupported) textStr = element.Patterns.Text.Pattern.DocumentRange?.GetText(4000); } catch { }
            try
            {
                if (element.Patterns.LegacyIAccessible.IsSupported)
                {
                    legacyVal = element.Patterns.LegacyIAccessible.Pattern.Value.ValueOrDefault;
                    legacyName = element.Patterns.LegacyIAccessible.Pattern.Name.ValueOrDefault;
                }
            }
            catch { }

            bool hasContent = (!string.IsNullOrWhiteSpace(name) && name.Length > 1)
                || !string.IsNullOrWhiteSpace(valueStr)
                || !string.IsNullOrWhiteSpace(textStr)
                || !string.IsNullOrWhiteSpace(legacyVal)
                || (!string.IsNullOrWhiteSpace(legacyName) && legacyName != name);

            if (hasContent)
            {
                string indent = new string(' ', depth * 2);
                Log($"{indent}[{controlType}] Class=\"{className}\" Id=\"{automationId}\"");
                if (!string.IsNullOrWhiteSpace(name) && name.Length > 1)
                    Log($"{indent}  Name: \"{TruncateStr(name, 500)}\"");
                if (!string.IsNullOrWhiteSpace(valueStr))
                    Log($"{indent}  Value: \"{TruncateStr(valueStr, 500)}\"");
                if (!string.IsNullOrWhiteSpace(textStr))
                    Log($"{indent}  TextPattern: \"{TruncateStr(textStr, 1000)}\"");
                if (!string.IsNullOrWhiteSpace(legacyVal))
                    Log($"{indent}  LegacyValue: \"{TruncateStr(legacyVal, 500)}\"");
                if (!string.IsNullOrWhiteSpace(legacyName) && legacyName != name)
                    Log($"{indent}  LegacyName: \"{TruncateStr(legacyName, 500)}\"");
            }

            var children = element.FindAllChildren();
            foreach (var child in children)
                ScanForText(child, depth + 1);
        }
        catch { }
    }

    static string GetElementInfo(AutomationElement el)
    {
        try
        {
            string controlType = SafeGetControlType(el);
            string name = "", automationId = "", className = "";
            try { name = el.Name ?? ""; } catch { }
            try { automationId = el.AutomationId ?? ""; } catch { }
            try { className = el.ClassName ?? ""; } catch { }
            string bounds = SafeGetBounds(el);

            var sb = new StringBuilder();
            sb.Append($"[{controlType}]");
            if (!string.IsNullOrEmpty(name)) sb.Append($" Name=\"{TruncateStr(name, 150)}\"");
            if (!string.IsNullOrEmpty(automationId)) sb.Append($" Id=\"{automationId}\"");
            if (!string.IsNullOrEmpty(className)) sb.Append($" Class=\"{className}\"");
            sb.Append($" {bounds}");

            var pats = new List<string>();
            try { if (el.Patterns.Value.IsSupported) pats.Add("Value"); } catch { }
            try { if (el.Patterns.Text.IsSupported) pats.Add("Text"); } catch { }
            try { if (el.Patterns.Invoke.IsSupported) pats.Add("Invoke"); } catch { }
            try { if (el.Patterns.SelectionItem.IsSupported) pats.Add("SelItem"); } catch { }
            try { if (el.Patterns.Selection.IsSupported) pats.Add("Selection"); } catch { }
            try { if (el.Patterns.Toggle.IsSupported) pats.Add("Toggle"); } catch { }
            try { if (el.Patterns.Scroll.IsSupported) pats.Add("Scroll"); } catch { }
            try { if (el.Patterns.ExpandCollapse.IsSupported) pats.Add("ExpandCollapse"); } catch { }
            try { if (el.Patterns.LegacyIAccessible.IsSupported) pats.Add("Legacy"); } catch { }
            try { if (el.Patterns.Grid.IsSupported) pats.Add("Grid"); } catch { }
            try { if (el.Patterns.GridItem.IsSupported) pats.Add("GridItem"); } catch { }
            try { if (el.Patterns.Table.IsSupported) pats.Add("Table"); } catch { }
            try { if (el.Patterns.TableItem.IsSupported) pats.Add("TableItem"); } catch { }
            try { if (el.Patterns.Window.IsSupported) pats.Add("Window"); } catch { }
            try { if (el.Patterns.Transform.IsSupported) pats.Add("Transform"); } catch { }

            if (pats.Count > 0) sb.Append($" [{string.Join(",", pats)}]");

            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"[ERROR: {ex.Message}]";
        }
    }

    static string SafeGetControlType(AutomationElement el)
    {
        try { return el.ControlType.ToString(); }
        catch { return "?"; }
    }

    static string SafeGetBounds(AutomationElement el)
    {
        try
        {
            var r = el.BoundingRectangle;
            return $"({r.X},{r.Y} {r.Width}x{r.Height})";
        }
        catch { return "(?)"; }
    }

    static string TruncateStr(string s, int max)
    {
        if (s == null) return "";
        if (s.Length <= max) return s;
        return s[..max] + "...";
    }

    static void Log(string msg)
    {
        Console.WriteLine(msg);
        _output.AppendLine(msg);
    }

    static void SaveOutput()
    {
        var path = @"C:\Users\erik.richter\Desktop\MosaicTools\AidocInspector\aidoc_inspection.txt";
        try
        {
            File.WriteAllText(path, _output.ToString());
            Console.WriteLine($"\nSaved to: {path}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\nFailed to save: {ex.Message}");
        }
    }
}
