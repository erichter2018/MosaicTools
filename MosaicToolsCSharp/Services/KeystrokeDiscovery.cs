using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;

namespace MosaicTools.Services;

/// <summary>
/// EXPERIMENTAL: Discovers undocumented keyboard shortcuts in Mosaic.
/// This is a dev-only tool - not for production use.
/// To remove: delete this file and the button in SettingsForm.cs
/// </summary>
public class KeystrokeDiscovery : IDisposable
{
    private readonly UIA3Automation _automation;
    private volatile bool _cancelRequested;
    private readonly Action<string> _statusCallback;

    // Dangerous keys to skip (would sign reports, close windows, etc.)
    private static readonly HashSet<string> SkipCombos = new(StringComparer.OrdinalIgnoreCase)
    {
        // Sign report
        "alt+f", "ctrl+alt+f",
        // Close window
        "ctrl+w", "alt+f4", "ctrl+q",
        // System keys
        "alt+tab", "ctrl+alt+delete",
        // Other potentially destructive
        "ctrl+alt+d", // might delete something
    };

    public KeystrokeDiscovery(Action<string> statusCallback)
    {
        _automation = new UIA3Automation();
        _statusCallback = statusCallback;
    }

    public void Cancel() => _cancelRequested = true;

    public void RunDiscovery(string outputPath)
    {
        _cancelRequested = false;
        var results = new List<DiscoveredShortcut>();

        // Define what to test
        var modifiers = new[] { "alt+", "ctrl+", "ctrl+shift+", "alt+shift+", "ctrl+alt+" };
        var letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        var numbers = "0123456789";
        var specialKeys = new[] { "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "f10", "f11", "f12" };

        var allCombos = new List<string>();

        // Build combo list
        foreach (var mod in modifiers)
        {
            foreach (var c in letters)
                allCombos.Add($"{mod}{c}");
            foreach (var c in numbers)
                allCombos.Add($"{mod}{c}");
            foreach (var key in specialKeys)
                allCombos.Add($"{mod}{key}");
        }

        // Filter out dangerous ones
        allCombos = allCombos.Where(c => !SkipCombos.Contains(c)).ToList();

        _statusCallback($"Testing {allCombos.Count} key combinations...");
        Logger.Trace($"KeystrokeDiscovery: Starting scan of {allCombos.Count} combinations");

        int tested = 0;
        int found = 0;

        foreach (var combo in allCombos)
        {
            if (_cancelRequested)
            {
                _statusCallback("Discovery cancelled.");
                break;
            }

            tested++;
            if (tested % 20 == 0)
            {
                _statusCallback($"Testing... {tested}/{allCombos.Count} ({found} found)");
            }

            try
            {
                var result = TestKeystroke(combo);
                if (result != null)
                {
                    results.Add(result);
                    found++;
                    Logger.Trace($"FOUND: {combo} -> {result.Effect}");
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Error testing {combo}: {ex.Message}");
            }

            // Small delay between tests
            Thread.Sleep(50);
        }

        // Write results
        WriteResults(outputPath, results);

        _statusCallback($"Done! Found {found} shortcuts. Results in: {outputPath}");
        Logger.Trace($"KeystrokeDiscovery: Complete. {found} shortcuts found.");
    }

    private DiscoveredShortcut? TestKeystroke(string combo)
    {
        // 1. Focus Mosaic
        if (!NativeWindows.ActivateMosaicForcefully())
        {
            Logger.Trace("KeystrokeDiscovery: Could not activate Mosaic");
            return null;
        }
        Thread.Sleep(100);

        // 2. Capture before state
        var before = CaptureState();

        // 3. Send the keystroke
        NativeWindows.SendHotkey(combo);
        Thread.Sleep(400); // Wait for UI to respond

        // 4. Capture after state
        var after = CaptureState();

        // 5. Compare states
        var changes = CompareStates(before, after);

        // 6. Reset - press Escape to dismiss any dialogs
        NativeWindows.SendHotkey("escape");
        Thread.Sleep(100);

        // 7. If focus was lost, try to restore
        if (after.FocusedWindowTitle != before.FocusedWindowTitle &&
            !after.FocusedWindowTitle.Contains("Mosaic", StringComparison.OrdinalIgnoreCase))
        {
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);
        }

        if (changes.Count > 0)
        {
            return new DiscoveredShortcut
            {
                Combo = combo,
                Effect = string.Join("; ", changes),
                NewWindowTitle = after.NewWindowTitles.FirstOrDefault()
            };
        }

        return null;
    }

    private StateSnapshot CaptureState()
    {
        var state = new StateSnapshot();

        try
        {
            // Get foreground window title
            var fgHwnd = NativeWindows.GetForegroundWindow();
            state.FocusedWindowTitle = NativeWindows.GetWindowTitle(fgHwnd);

            // Get all window titles related to Mosaic
            var desktop = _automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));

            foreach (var win in windows)
            {
                try
                {
                    var title = win.Name ?? "";
                    if (title.Contains("Mosaic", StringComparison.OrdinalIgnoreCase) ||
                        title.Contains("SlimHub", StringComparison.OrdinalIgnoreCase))
                    {
                        state.WindowTitles.Add(title);
                    }
                }
                catch { }
            }

            // Get clipboard
            try
            {
                state.ClipboardText = ClipboardService.GetText() ?? "";
            }
            catch { }

            // Get focused element in Mosaic
            try
            {
                var mosaicWin = FindMosaicWindow();
                if (mosaicWin != null)
                {
                    var focused = mosaicWin.FindFirstDescendant(cf =>
                        cf.ByControlType(FlaUI.Core.Definitions.ControlType.Document));
                    state.FocusedElementName = focused?.Name ?? "";
                }
            }
            catch { }
        }
        catch (Exception ex)
        {
            Logger.Trace($"CaptureState error: {ex.Message}");
        }

        return state;
    }

    private AutomationElement? FindMosaicWindow()
    {
        try
        {
            var desktop = _automation.GetDesktop();
            var windows = desktop.FindAllChildren(cf =>
                cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));

            foreach (var window in windows)
            {
                var title = window.Name?.ToLowerInvariant() ?? "";
                if ((title.Contains("mosaic") && title.Contains("info hub")) ||
                    (title.Contains("mosaic") && title.Contains("reporting")))
                {
                    return window;
                }
            }
        }
        catch { }
        return null;
    }

    private List<string> CompareStates(StateSnapshot before, StateSnapshot after)
    {
        var changes = new List<string>();

        // New windows appeared?
        var newWindows = after.WindowTitles.Except(before.WindowTitles).ToList();
        if (newWindows.Count > 0)
        {
            after.NewWindowTitles = newWindows;
            changes.Add($"New window: {newWindows.First()}");
        }

        // Focus changed to different window?
        if (after.FocusedWindowTitle != before.FocusedWindowTitle &&
            !string.IsNullOrEmpty(after.FocusedWindowTitle))
        {
            // Only report if it's something meaningful (not just focus shift)
            if (!after.FocusedWindowTitle.Contains("Mosaic", StringComparison.OrdinalIgnoreCase) ||
                after.FocusedWindowTitle != before.FocusedWindowTitle)
            {
                changes.Add($"Focus -> {after.FocusedWindowTitle}");
            }
        }

        // Clipboard changed?
        if (after.ClipboardText != before.ClipboardText &&
            !string.IsNullOrEmpty(after.ClipboardText))
        {
            var preview = after.ClipboardText.Length > 50
                ? after.ClipboardText.Substring(0, 50) + "..."
                : after.ClipboardText;
            preview = preview.Replace("\r", "").Replace("\n", " ");
            changes.Add($"Clipboard changed: {preview}");
        }

        return changes;
    }

    private void WriteResults(string outputPath, List<DiscoveredShortcut> results)
    {
        var lines = new List<string>
        {
            "# Mosaic Keyboard Shortcut Discovery Results",
            $"# Generated: {DateTime.Now}",
            $"# Total found: {results.Count}",
            "",
            "Shortcut | Effect | New Window",
            "---------|--------|------------"
        };

        foreach (var r in results.OrderBy(r => r.Combo))
        {
            lines.Add($"{r.Combo} | {r.Effect} | {r.NewWindowTitle ?? ""}");
        }

        // Also add a simple list
        lines.Add("");
        lines.Add("# Simple list:");
        foreach (var r in results.OrderBy(r => r.Combo))
        {
            lines.Add($"{r.Combo}: {r.Effect}");
        }

        File.WriteAllLines(outputPath, lines);
    }

    public void Dispose()
    {
        _automation.Dispose();
    }

    private class StateSnapshot
    {
        public string FocusedWindowTitle { get; set; } = "";
        public List<string> WindowTitles { get; set; } = new();
        public List<string> NewWindowTitles { get; set; } = new();
        public string ClipboardText { get; set; } = "";
        public string FocusedElementName { get; set; } = "";
    }

    private class DiscoveredShortcut
    {
        public string Combo { get; set; } = "";
        public string Effect { get; set; } = "";
        public string? NewWindowTitle { get; set; }
    }
}
