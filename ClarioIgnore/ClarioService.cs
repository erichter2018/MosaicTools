using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace ClarioIgnore;

/// <summary>
/// Represents a single item in the Clario worklist.
/// </summary>
public class WorklistItem
{
    public string Procedure { get; set; } = "";
    public string Priority { get; set; } = "";
    public string Accession { get; set; } = "";
    public AutomationElement? RowElement { get; set; }

    public override string ToString() => $"{Procedure} ({Priority})";
}

/// <summary>
/// UI Automation service for reading Clario worklist and clicking skip buttons.
/// Uses FlaUI which is a modern wrapper over Windows UI Automation.
/// </summary>
public class ClarioService : IDisposable
{
    private readonly UIA3Automation _automation;

    public ClarioService()
    {
        _automation = new UIA3Automation();
    }

    /// <summary>
    /// Find the Chrome window with Clario - Worklist.
    /// </summary>
    public AutomationElement? FindClarioWindow()
    {
        try
        {
            var desktop = _automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));

            foreach (var window in windows)
            {
                try
                {
                    var title = window.Name?.ToLowerInvariant() ?? "";
                    if (title.Contains("clario") && title.Contains("worklist"))
                    {
                        Logger.Log($"Found Clario Window: '{window.Name}'");
                        return window;
                    }
                }
                catch { continue; }
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"Error finding Clario window: {ex.Message}");
        }

        return null;
    }

    /// <summary>
    /// Get all worklist items from the Clario window.
    /// Looks specifically for gridview-3019 which is the worklist grid.
    /// </summary>
    public List<WorklistItem> GetWorklistItems()
    {
        var items = new List<WorklistItem>();

        try
        {
            var window = FindClarioWindow();
            if (window == null)
            {
                Logger.Log("Clario window not found");
                return items;
            }

            // Find all elements - we'll filter for the worklist grid rows
            var allElements = window.FindAllDescendants();

            foreach (var elem in allElements)
            {
                try
                {
                    var autoId = elem.Properties.AutomationId.ValueOrDefault ?? "";

                    // Look for worklist rows: gridview-3019-record-ext-record-XXXX
                    // Skip headers: gridview-3019-hd-*
                    if (!autoId.StartsWith("gridview-3019-record-ext-record-"))
                        continue;

                    // Get all cells in this row
                    var cells = elem.FindAllDescendants();

                    // Collect non-empty cell values
                    var cellValues = new List<string>();
                    foreach (var cell in cells)
                    {
                        var cellName = cell.Name ?? "";
                        if (!string.IsNullOrWhiteSpace(cellName))
                            cellValues.Add(cellName.Trim());
                    }

                    // Field 4 (0-indexed) is the procedure
                    // Fields: 0=Priority, 1=Location, 2=Code, 3=Time, 4=Procedure, 5=Code, 6=Hospital, 7=Patient...
                    string procedure = cellValues.Count > 4 ? cellValues[4] : "";
                    string priority = cellValues.Count > 0 ? cellValues[0] : "";
                    string accession = cellValues.Count > 2 ? cellValues[2] : "";

                    if (!string.IsNullOrEmpty(procedure))
                    {
                        var worklistItem = new WorklistItem
                        {
                            Procedure = procedure,
                            Priority = priority,
                            Accession = accession,
                            RowElement = elem
                        };
                        items.Add(worklistItem);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"Error processing row: {ex.Message}");
                }
            }

            // Only return first N items (visible portion of worklist)
            var maxFiles = Configuration.Instance.MaxFilesToMonitor;
            if (items.Count > maxFiles)
            {
                items = items.Take(maxFiles).ToList();
            }

            Logger.Log($"Found {items.Count} worklist items (limited to {maxFiles})");
        }
        catch (Exception ex)
        {
            Logger.Log($"Error getting worklist items: {ex.Message}");
        }

        return items;
    }

    /// <summary>
    /// Check if a row is visible on screen (has valid coordinates).
    /// </summary>
    public bool IsRowVisible(WorklistItem item)
    {
        if (item.RowElement == null) return false;

        try
        {
            var rect = item.RowElement.BoundingRectangle;
            // Row must have positive coordinates and reasonable size
            return rect.X > 0 && rect.Y > 0 && rect.Width > 50 && rect.Height > 10;
        }
        catch
        {
            return false;
        }
    }

    // For coordinate-based clicking
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    // For screen capture without flicker
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [System.Runtime.InteropServices.DllImport("gdi32.dll")]
    private static extern uint GetPixel(IntPtr hdc, int nXPos, int nYPos);

    // For PostMessage click (no mouse movement)
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(POINT point);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr ChildWindowFromPointEx(IntPtr hWndParent, POINT pt, uint uFlags);

    [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string? lpszClass, string? lpszWindow);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

    private const uint GA_ROOT = 2; // Get root owner window

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    private const uint CWP_SKIPINVISIBLE = 0x0001;
    private const uint CWP_SKIPDISABLED = 0x0002;
    private const uint CWP_SKIPTRANSPARENT = 0x0004;

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct POINT { public int X; public int Y; }

    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;

    // PostMessage click constants
    private const uint WM_LBUTTONDOWN = 0x0201;
    private const uint WM_LBUTTONUP = 0x0202;
    private const int MK_LBUTTON = 0x0001;

    // Track the foreground window before clicking
    private IntPtr _savedForegroundWindow = IntPtr.Zero;

    /// <summary>
    /// Save the currently active window before a batch of clicks.
    /// </summary>
    public void SaveActiveWindow()
    {
        _savedForegroundWindow = GetForegroundWindow();
    }

    /// <summary>
    /// Restore focus to the previously active window after clicks.
    /// </summary>
    public void RestoreActiveWindow()
    {
        if (_savedForegroundWindow != IntPtr.Zero)
        {
            SetForegroundWindow(_savedForegroundWindow);
            _savedForegroundWindow = IntPtr.Zero;
        }
    }

    // Brightness threshold - icons brighter than this are considered "already skipped"
    // Inactive buttons ~72, active buttons ~125
    private const int SKIP_BRIGHTNESS_THRESHOLD = 100;

    /// <summary>
    /// Check if the skip button at the given coordinates is already active (bright).
    /// Returns true if already skipped (should NOT click), false if not skipped (safe to click).
    /// Uses direct Windows API to avoid cursor/screen flicker from GDI+ CopyFromScreen.
    /// </summary>
    private bool IsSkipButtonAlreadyActive(int x, int y)
    {
        IntPtr hdc = IntPtr.Zero;
        try
        {
            hdc = GetDC(IntPtr.Zero);
            if (hdc == IntPtr.Zero)
                return false;

            // Sample just 5 pixels (center + 4 corners) - fast and sufficient
            int offset = 4;
            int[][] samples = {
                new[] {0, 0},           // center
                new[] {-offset, -offset}, // top-left
                new[] {offset, -offset},  // top-right
                new[] {-offset, offset},  // bottom-left
                new[] {offset, offset}    // bottom-right
            };

            long totalBrightness = 0;
            int pixelCount = 0;

            foreach (var sample in samples)
            {
                uint pixel = GetPixel(hdc, x + sample[0], y + sample[1]);
                if (pixel != 0xFFFFFFFF)
                {
                    int r = (int)(pixel & 0xFF);
                    int g = (int)((pixel >> 8) & 0xFF);
                    int b = (int)((pixel >> 16) & 0xFF);
                    totalBrightness += (int)(0.299 * r + 0.587 * g + 0.114 * b);
                    pixelCount++;
                }
            }

            if (pixelCount == 0)
                return false;

            int avgBrightness = (int)(totalBrightness / pixelCount);
            bool isActive = avgBrightness > SKIP_BRIGHTNESS_THRESHOLD;

            Logger.Log($"Skip button at ({x},{y}): brightness={avgBrightness}, alreadyActive={isActive}");

            return isActive;
        }
        catch (Exception ex)
        {
            Logger.Log($"Error checking brightness: {ex.Message}");
            return false;
        }
        finally
        {
            if (hdc != IntPtr.Zero)
                ReleaseDC(IntPtr.Zero, hdc);
        }
    }

    /// <summary>
    /// Click at specific screen coordinates, saving and restoring mouse position.
    /// </summary>
    private void ClickAt(int x, int y)
    {
        // Save current mouse position
        GetCursorPos(out POINT originalPos);

        // Move, then wait for UI to register cursor position (most critical for web apps)
        SetCursorPos(x, y);
        Thread.Sleep(75);

        // Click
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
        Thread.Sleep(30);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        Thread.Sleep(50);

        // Restore mouse position
        SetCursorPos(originalPos.X, originalPos.Y);
    }

    /// <summary>
    /// Attempt to click using UI Automation patterns (no mouse movement).
    /// Returns true if pattern-based click succeeded, false to fall back to coordinate click.
    /// </summary>
    private bool TryPatternClick(AutomationElement element, string procedureName)
    {
        try
        {
            // Log available patterns for diagnostics (first few times only to avoid spam)
            var invokeSupported = element.Patterns.Invoke.IsSupported;
            var toggleSupported = element.Patterns.Toggle.IsSupported;
            var selectionItemSupported = element.Patterns.SelectionItem.IsSupported;

            Logger.Log($"Patterns for '{procedureName}': Invoke={invokeSupported}, Toggle={toggleSupported}, SelectionItem={selectionItemSupported}");

            // Try InvokePattern first (most common for buttons)
            if (invokeSupported)
            {
                Logger.Log($"Using InvokePattern for '{procedureName}' (no mouse movement)");
                element.Patterns.Invoke.Pattern.Invoke();
                return true;
            }

            // Try TogglePattern for toggle-style buttons
            if (toggleSupported)
            {
                Logger.Log($"Using TogglePattern for '{procedureName}' (no mouse movement)");
                element.Patterns.Toggle.Pattern.Toggle();
                return true;
            }

            // Try SelectionItemPattern for selectable items
            if (selectionItemSupported)
            {
                Logger.Log($"Using SelectionItemPattern for '{procedureName}' (no mouse movement)");
                element.Patterns.SelectionItem.Pattern.Select();
                return true;
            }

            Logger.Log($"No clickable patterns available for '{procedureName}'");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Log($"Pattern click failed for '{procedureName}': {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Attempt to click using PostMessage (no mouse movement).
    /// Sends WM_LBUTTONDOWN/UP directly to Chrome's render widget at the given screen coordinates.
    /// Returns true if messages were sent successfully.
    /// </summary>
    private bool TryPostMessageClick(int screenX, int screenY, string procedureName)
    {
        try
        {
            // Find window at screen coordinates
            var screenPt = new POINT { X = screenX, Y = screenY };
            IntPtr windowAtPoint = WindowFromPoint(screenPt);

            if (windowAtPoint == IntPtr.Zero)
            {
                Logger.Log($"PostMessage: No window found at ({screenX}, {screenY}) for '{procedureName}'");
                return false;
            }

            // Get the root/top-level window (WindowFromPoint may return a deeply nested child)
            IntPtr rootWindow = GetAncestor(windowAtPoint, GA_ROOT);
            if (rootWindow == IntPtr.Zero)
                rootWindow = windowAtPoint;

            // Find Chrome's render widget from the root window - this is where web content lives
            IntPtr renderWidget = FindChromeRenderWidget(rootWindow);
            IntPtr targetWindow = renderWidget != IntPtr.Zero ? renderWidget : windowAtPoint;

            // Log the render widget's screen position for debugging
            if (renderWidget != IntPtr.Zero && GetWindowRect(renderWidget, out RECT widgetRect))
            {
                Logger.Log($"RenderWidget rect: top={widgetRect.Top}, left={widgetRect.Left}, bottom={widgetRect.Bottom}, right={widgetRect.Right}");

                // Items near the bottom of the viewport have unreliable coordinates from UI Automation
                // Skip PostMessage for these - they'll fall back to regular mouse click
                int viewportHeight = widgetRect.Bottom - widgetRect.Top;
                int maxReliableY = widgetRect.Top + (int)(viewportHeight * 0.75); // Top 75% is reliable
                if (screenY > maxReliableY)
                {
                    Logger.Log($"PostMessage: screenY={screenY} exceeds reliable zone ({maxReliableY}), skipping PostMessage");
                    return false;
                }
            }

            // Convert screen coordinates to client coordinates using ScreenToClient
            // This properly accounts for title bar, menu bar, etc.
            var clientPt = new POINT { X = screenX, Y = screenY };
            if (!ScreenToClient(targetWindow, ref clientPt))
            {
                Logger.Log($"PostMessage: ScreenToClient failed for '{procedureName}'");
                return false;
            }

            int clientX = clientPt.X;
            int clientY = clientPt.Y;

            string widgetInfo = renderWidget != IntPtr.Zero ? "render widget" : "fallback window";
            Logger.Log($"PostMessage: screen=({screenX},{screenY}) -> client=({clientX},{clientY}) via {widgetInfo}, renderFound={renderWidget != IntPtr.Zero}");

            // Pack coordinates into lParam: (y << 16) | (x & 0xFFFF)
            IntPtr lParam = (IntPtr)((clientY << 16) | (clientX & 0xFFFF));
            IntPtr wParam = (IntPtr)MK_LBUTTON;

            // Send mouse down and up
            bool downResult = PostMessage(targetWindow, WM_LBUTTONDOWN, wParam, lParam);
            Thread.Sleep(30); // Brief delay between down and up
            bool upResult = PostMessage(targetWindow, WM_LBUTTONUP, IntPtr.Zero, lParam);

            if (downResult && upResult)
            {
                Logger.Log($"PostMessage click for '{procedureName}' at ({clientX}, {clientY}) - no mouse movement");
                return true;
            }
            else
            {
                Logger.Log($"PostMessage failed for '{procedureName}': down={downResult}, up={upResult}");
                return false;
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"PostMessage click failed for '{procedureName}': {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Find Chrome's render widget child window where web content is displayed.
    /// </summary>
    private IntPtr FindChromeRenderWidget(IntPtr chromeWindow)
    {
        // Chrome's render widget has class name "Chrome_RenderWidgetHostHWND"
        // First find the intermediate widget window
        IntPtr intermediateWidget = FindWindowEx(chromeWindow, IntPtr.Zero, "Chrome_RenderWidgetHostHWND", null);
        if (intermediateWidget != IntPtr.Zero)
        {
            Logger.Log($"FindChromeRenderWidget: Found at level 1");
            return intermediateWidget;
        }

        // Try searching one level deeper - Chrome sometimes nests windows
        IntPtr child = FindWindowEx(chromeWindow, IntPtr.Zero, null, null);
        int depth = 0;
        while (child != IntPtr.Zero && depth < 20)
        {
            IntPtr renderWidget = FindWindowEx(child, IntPtr.Zero, "Chrome_RenderWidgetHostHWND", null);
            if (renderWidget != IntPtr.Zero)
            {
                Logger.Log($"FindChromeRenderWidget: Found at level 2 (child {depth})");
                return renderWidget;
            }

            // Also check grandchildren
            IntPtr grandchild = FindWindowEx(child, IntPtr.Zero, null, null);
            while (grandchild != IntPtr.Zero)
            {
                renderWidget = FindWindowEx(grandchild, IntPtr.Zero, "Chrome_RenderWidgetHostHWND", null);
                if (renderWidget != IntPtr.Zero)
                {
                    Logger.Log($"FindChromeRenderWidget: Found at level 3");
                    return renderWidget;
                }
                grandchild = FindWindowEx(child, grandchild, null, null);
            }

            child = FindWindowEx(chromeWindow, child, null, null);
            depth++;
        }

        Logger.Log($"FindChromeRenderWidget: Not found after checking {depth} children");
        return IntPtr.Zero;
    }

    /// <summary>
    /// Click the skip button for a worklist row using coordinate-based clicking.
    /// The skip button is in the second column (first is "next", second is "skip").
    /// Checks if button is already active before clicking to avoid toggling it off.
    /// </summary>
    public bool ClickSkipButton(WorklistItem item)
    {
        if (item.RowElement == null) return false;

        try
        {
            // Get the row's bounding rectangle
            var rowRect = item.RowElement.BoundingRectangle;
            if (rowRect.IsEmpty || rowRect.Width <= 0)
            {
                Logger.Log($"Row has invalid bounds for '{item.Procedure}'");
                return false;
            }

            // Check if row is likely off-screen (scrolled out of view)
            // Off-screen rows often have very small height or unreasonable Y coordinates
            if (rowRect.Height < 15)
            {
                Logger.Log($"Row appears off-screen (height={rowRect.Height}) for '{item.Procedure}' - skipping");
                return false;
            }

            // Sanity check: Y coordinate should be reasonable (not negative, not way off screen)
            if (rowRect.Y < 0 || rowRect.Y > 3000)
            {
                Logger.Log($"Row has unreasonable Y coordinate ({rowRect.Y}) for '{item.Procedure}' - skipping");
                return false;
            }

            int clickX, clickY;

            // Find the second cell (first is "next", second is "skip")
            var cells = item.RowElement.FindAllChildren();
            AutomationElement? skipCell = null;

            if (cells.Length >= 2)
            {
                skipCell = cells[1]; // Second cell is the skip button
                var cellRect = skipCell.BoundingRectangle;

                if (!cellRect.IsEmpty && cellRect.Width > 0)
                {
                    clickX = (int)(cellRect.X + cellRect.Width / 2);
                    clickY = (int)(cellRect.Y + cellRect.Height / 2);
                }
                else
                {
                    // Fallback coordinates
                    clickX = (int)(rowRect.X + 33);
                    clickY = (int)(rowRect.Y + rowRect.Height / 2);
                }
            }
            else
            {
                // Fallback coordinates
                clickX = (int)(rowRect.X + 33);
                clickY = (int)(rowRect.Y + rowRect.Height / 2);
            }

            // Skip items beyond Y=800 - they have unreliable coordinates and will scroll up eventually
            const int MAX_RELIABLE_Y = 800;
            if (clickY > MAX_RELIABLE_Y)
            {
                Logger.Log($"Skipping '{item.Procedure}' at Y={clickY} (beyond {MAX_RELIABLE_Y}) - will scroll up later");
                return false;
            }

            // Check if button is already active (bright) - don't click if so
            if (IsSkipButtonAlreadyActive(clickX, clickY))
            {
                Logger.Log($"Skip button already active for '{item.Procedure}' - not clicking");
                return false; // Return false since we didn't click
            }

            // Try pattern-based clicking first (no mouse movement required)
            if (skipCell != null && TryPatternClick(skipCell, item.Procedure))
            {
                return true;
            }

            // Try PostMessage clicking (no mouse movement required)
            if (TryPostMessageClick(clickX, clickY, item.Procedure))
            {
                return true;
            }

            // Final fallback to coordinate-based clicking (moves mouse briefly)
            Logger.Log($"Falling back to mouse click at ({clickX}, {clickY}) for '{item.Procedure}'");
            ClickAt(clickX, clickY);

            return true;
        }
        catch (Exception ex)
        {
            Logger.Log($"Error clicking skip button: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Test mode: Log what would be skipped without actually clicking.
    /// </summary>
    public List<(WorklistItem Item, SkipRule Rule)> GetMatchingItems()
    {
        var matches = new List<(WorklistItem, SkipRule)>();
        var items = GetWorklistItems();

        foreach (var item in items)
        {
            var rule = Configuration.Instance.FindMatchingRule(item.Procedure, item.Priority);
            if (rule != null)
            {
                matches.Add((item, rule));
            }
        }

        return matches;
    }

    /// <summary>
    /// Diagnostic method to dump UI structure for debugging.
    /// Returns a string description of what was found.
    /// </summary>
    public string RunDiagnostic()
    {
        var result = new System.Text.StringBuilder();

        try
        {
            var desktop = _automation.GetDesktop();
            var allWindows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));

            AutomationElement? chromeWindow = null;
            foreach (var win in allWindows)
            {
                try
                {
                    var title = win.Name ?? "";
                    if (title.Contains("Clario", StringComparison.OrdinalIgnoreCase) &&
                        title.Contains("Chrome", StringComparison.OrdinalIgnoreCase))
                    {
                        chromeWindow = win;
                        break;
                    }
                }
                catch { }
            }

            if (chromeWindow == null)
            {
                result.AppendLine("ERROR: No Clario Chrome window found");
                return result.ToString();
            }

            result.AppendLine($"Using window: '{chromeWindow.Name}'");
            result.AppendLine();

            var allElements = chromeWindow.FindAllDescendants();
            result.AppendLine($"Total elements: {allElements.Length}");
            result.AppendLine();

            // Look for "Worklist" section header
            result.AppendLine("=== LOOKING FOR 'WORKLIST' SECTION ===");
            foreach (var elem in allElements)
            {
                try
                {
                    var name = elem.Name ?? "";
                    if (name.Equals("Worklist", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("Worklist", StringComparison.OrdinalIgnoreCase))
                    {
                        var controlType = elem.Properties.ControlType.ValueOrDefault;
                        var autoId = elem.Properties.AutomationId.ValueOrDefault ?? "";
                        result.AppendLine($"[{controlType}] Name='{name}', Id='{autoId}'");
                    }
                }
                catch { }
            }
            result.AppendLine();

            // Find the Auto-Next section and explore its structure
            result.AppendLine("=== FINDING AUTO-NEXT / WORKLIST SECTION ===");

            // Look for elements with IDs containing "worklist" but NOT "exam"
            foreach (var elem in allElements)
            {
                try
                {
                    var autoId = elem.Properties.AutomationId.ValueOrDefault ?? "";
                    if (autoId.Contains("worklist", StringComparison.OrdinalIgnoreCase) ||
                        autoId.Contains("queue", StringComparison.OrdinalIgnoreCase) ||
                        autoId.Contains("autonext", StringComparison.OrdinalIgnoreCase) ||
                        autoId.Contains("auto_next", StringComparison.OrdinalIgnoreCase) ||
                        autoId.Contains("auto-next", StringComparison.OrdinalIgnoreCase))
                    {
                        var controlType = elem.Properties.ControlType.ValueOrDefault;
                        var name = elem.Name ?? "";
                        result.AppendLine($"[{controlType}] Name='{name}' Id='{autoId}'");
                    }
                }
                catch { }
            }
            result.AppendLine();

            // Find the Auto-Next button and explore siblings/parent
            result.AppendLine("=== AUTO-NEXT BUTTON NEIGHBORHOOD ===");
            foreach (var elem in allElements)
            {
                try
                {
                    var name = elem.Name ?? "";
                    if (name == "Auto-Next")
                    {
                        result.AppendLine($"Found Auto-Next button");

                        // Go up and find siblings
                        var parent = elem.Parent;
                        for (int level = 1; level <= 5 && parent != null; level++)
                        {
                            var parentId = parent.Properties.AutomationId.ValueOrDefault ?? "";
                            var parentType = parent.Properties.ControlType.ValueOrDefault;
                            var children = parent.FindAllChildren();
                            result.AppendLine($"  Level {level} parent: [{parentType}] Id='{parentId}' - {children.Length} children");

                            // Look for Table or DataGrid children
                            var tables = parent.FindAllDescendants(cf => cf.ByControlType(ControlType.Table));
                            var grids = parent.FindAllDescendants(cf => cf.ByControlType(ControlType.DataGrid));
                            if (tables.Length > 0 || grids.Length > 0)
                            {
                                result.AppendLine($"    Found Tables: {tables.Length}, DataGrids: {grids.Length}");

                                foreach (var table in tables)
                                {
                                    var tableId = table.Properties.AutomationId.ValueOrDefault ?? "";
                                    if (!tableId.Contains("exam", StringComparison.OrdinalIgnoreCase))
                                    {
                                        result.AppendLine($"    Non-exam Table: Id='{tableId}'");

                                        // Get first few rows
                                        var rows = table.FindAllChildren();
                                        result.AppendLine($"      Rows/children: {rows.Length}");

                                        int rowNum = 0;
                                        foreach (var row in rows)
                                        {
                                            var rowId = row.Properties.AutomationId.ValueOrDefault ?? "";
                                            var rowName = row.Name ?? "";
                                            if (rowId.Contains("record"))
                                            {
                                                rowNum++;
                                                result.AppendLine($"      Row: Id='{rowId}'");

                                                // Get cells in this row
                                                var cells = row.FindAllDescendants(cf => cf.ByControlType(ControlType.DataItem));
                                                result.Append("        Cells: ");
                                                int cellNum = 0;
                                                foreach (var cell in cells)
                                                {
                                                    var cellName = cell.Name ?? "";
                                                    if (!string.IsNullOrWhiteSpace(cellName))
                                                    {
                                                        result.Append($"'{cellName}' | ");
                                                        cellNum++;
                                                        if (cellNum >= 6) break;
                                                    }
                                                }
                                                result.AppendLine();

                                                if (rowNum >= 3) break;
                                            }
                                        }
                                    }
                                }
                            }

                            parent = parent.Parent;
                        }
                        break;
                    }
                }
                catch { }
            }
            result.AppendLine();

            // Focus on gridview-3019 (the worklist) - examine row contents
            result.AppendLine("=== WORKLIST GRID (3019) ROW DETAILS ===");
            int rowCount = 0;
            foreach (var elem in allElements)
            {
                try
                {
                    var autoId = elem.Properties.AutomationId.ValueOrDefault ?? "";
                    // Look for worklist rows (not headers)
                    if (autoId.StartsWith("gridview-3019-record-ext-record-"))
                    {
                        rowCount++;
                        if (rowCount > 5) continue; // Only show first 5

                        result.AppendLine($"\n--- Row: {autoId} ---");

                        // Get all descendant cells
                        var cells = elem.FindAllDescendants();
                        result.AppendLine($"Total descendants: {cells.Length}");

                        // Show cell contents
                        result.Append("Cell values: ");
                        int cellNum = 0;
                        foreach (var cell in cells)
                        {
                            var cellName = cell.Name ?? "";
                            if (!string.IsNullOrWhiteSpace(cellName) && cellName.Length > 1)
                            {
                                result.Append($"'{cellName}' | ");
                                cellNum++;
                                if (cellNum >= 8) break;
                            }
                        }
                        result.AppendLine();

                        // Check for buttons IN this row
                        var rowButtons = elem.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
                        result.AppendLine($"Buttons in row: {rowButtons.Length}");
                        foreach (var btn in rowButtons)
                        {
                            var btnId = btn.Properties.AutomationId.ValueOrDefault ?? "";
                            var btnName = btn.Name ?? "";
                            result.AppendLine($"  Button: Name='{btnName}', Id='{btnId}'");
                        }
                    }
                }
                catch { }
            }
            result.AppendLine($"\nTotal worklist rows found: {rowCount}");
            result.AppendLine();

            // Now look at where empty buttons are and try to correlate with rows
            result.AppendLine("=== EMPTY BUTTONS - FULL ANCESTRY ===");
            int emptyBtnCount = 0;
            foreach (var elem in allElements)
            {
                try
                {
                    if (elem.Properties.ControlType.ValueOrDefault != ControlType.Button) continue;

                    var name = elem.Name ?? "";
                    var btnId = elem.Properties.AutomationId.ValueOrDefault ?? "";

                    if (string.IsNullOrWhiteSpace(name) || name == " ")
                    {
                        emptyBtnCount++;
                        if (emptyBtnCount > 8) continue;

                        result.AppendLine($"\nEmpty Button #{emptyBtnCount}: Id='{btnId}'");

                        // Trace full ancestry
                        var current = elem.Parent;
                        int level = 0;
                        while (current != null && level < 10)
                        {
                            level++;
                            var pId = current.Properties.AutomationId.ValueOrDefault ?? "";
                            var pType = current.Properties.ControlType.ValueOrDefault;

                            // Only show if it has an ID
                            if (!string.IsNullOrWhiteSpace(pId))
                            {
                                result.AppendLine($"  L{level}: [{pType}] Id='{pId}'");

                                // If this is a gridview row, note it
                                if (pId.Contains("gridview-") && pId.Contains("-record-"))
                                {
                                    result.AppendLine($"      *** THIS BUTTON IS IN A GRID ROW ***");
                                }
                            }

                            current = current.Parent;
                        }
                    }
                }
                catch { }
            }
            result.AppendLine($"\nTotal empty buttons: {emptyBtnCount}");

            // Look for any elements with "skip" in name or ID
            result.AppendLine("\n=== ELEMENTS WITH 'SKIP' IN NAME/ID ===");
            foreach (var elem in allElements)
            {
                try
                {
                    var name = elem.Name ?? "";
                    var autoId = elem.Properties.AutomationId.ValueOrDefault ?? "";
                    if (name.Contains("skip", StringComparison.OrdinalIgnoreCase) ||
                        autoId.Contains("skip", StringComparison.OrdinalIgnoreCase))
                    {
                        var controlType = elem.Properties.ControlType.ValueOrDefault;
                        result.AppendLine($"[{controlType}] Name='{name}', Id='{autoId}'");
                    }
                }
                catch { }
            }

            // Look at ALL element types in the first worklist row
            result.AppendLine("\n=== ALL ELEMENTS IN FIRST WORKLIST ROW ===");
            foreach (var elem in allElements)
            {
                try
                {
                    var autoId = elem.Properties.AutomationId.ValueOrDefault ?? "";
                    if (autoId == "gridview-3019-record-ext-record-6504" ||
                        autoId.StartsWith("gridview-3019-record-ext-record-"))
                    {
                        result.AppendLine($"Row: {autoId}");
                        var descendants = elem.FindAllDescendants();
                        foreach (var desc in descendants)
                        {
                            var descType = desc.Properties.ControlType.ValueOrDefault;
                            var descName = desc.Name ?? "";
                            var descId = desc.Properties.AutomationId.ValueOrDefault ?? "";
                            result.AppendLine($"  [{descType}] Name='{descName}', Id='{descId}'");
                        }
                        break; // Only show first row
                    }
                }
                catch { }
            }

            // Get bounding rectangles of worklist rows to find skip button position
            result.AppendLine("\n=== ROW BOUNDING RECTANGLES ===");
            result.AppendLine("(Skip button should be at left edge of each row)");
            int rectCount = 0;
            foreach (var elem in allElements)
            {
                try
                {
                    var autoId = elem.Properties.AutomationId.ValueOrDefault ?? "";
                    if (autoId.StartsWith("gridview-3019-record-ext-record-"))
                    {
                        rectCount++;
                        if (rectCount > 3) continue;

                        var rect = elem.BoundingRectangle;
                        result.AppendLine($"\nRow {rectCount}: {autoId}");
                        result.AppendLine($"  Bounds: X={rect.X}, Y={rect.Y}, W={rect.Width}, H={rect.Height}");

                        // Get first cell to compare
                        var cells = elem.FindAllChildren();
                        if (cells.Length > 0)
                        {
                            var firstCell = cells[0];
                            var cellRect = firstCell.BoundingRectangle;
                            result.AppendLine($"  First cell: X={cellRect.X}, Y={cellRect.Y}, W={cellRect.Width}");

                            // The skip button might be to the left of the first cell
                            var gapBeforeFirstCell = cellRect.X - rect.X;
                            result.AppendLine($"  Gap before first cell: {gapBeforeFirstCell}px");
                        }

                        // Get procedure name cell to extract text
                        foreach (var cell in elem.FindAllDescendants())
                        {
                            var cellName = cell.Name ?? "";
                            if (cellName.Contains("XR ") || cellName.Contains("CT ") ||
                                cellName.Contains("MR ") || cellName.Contains("US ") ||
                                cellName.Contains("VENOUS") || cellName.Contains("DOPPLER"))
                            {
                                result.AppendLine($"  Procedure: '{cellName}'");
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.AppendLine($"  Error: {ex.Message}");
                }
            }

            // Look at the header row to understand column structure
            result.AppendLine("\n=== WORKLIST HEADER ROW ===");
            foreach (var elem in allElements)
            {
                try
                {
                    var autoId = elem.Properties.AutomationId.ValueOrDefault ?? "";
                    if (autoId.Contains("gridview-3019") && autoId.Contains("-hd-"))
                    {
                        var name = elem.Name ?? "";
                        var rect = elem.BoundingRectangle;
                        result.AppendLine($"Header: '{name}' at X={rect.X}, Y={rect.Y}, W={rect.Width}");
                    }
                }
                catch { }
            }
        }
        catch (Exception ex)
        {
            result.AppendLine($"ERROR: {ex.Message}");
        }

        return result.ToString();
    }

    public void Dispose()
    {
        _automation?.Dispose();
    }
}
