using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace MosaicTools.Services;

/// <summary>
/// Native Windows API interop for window manipulation.
/// Matches Python's win32gui/ctypes usage.
/// </summary>
public static class NativeWindows
{
    #region Window Management
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool BringWindowToTop(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll")]
    public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    
    [DllImport("kernel32.dll")]
    public static extern uint GetCurrentThreadId();
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);
    
    [DllImport("user32.dll")]
    public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);
    
    public const int SW_RESTORE = 9;
    public const int SW_SHOWMINIMIZED = 2;
    
    [StructLayout(LayoutKind.Sequential)]
    public struct WINDOWPLACEMENT
    {
        public int length;
        public int flags;
        public int showCmd;
        public System.Drawing.Point ptMinPosition;
        public System.Drawing.Point ptMaxPosition;
        public System.Drawing.Rectangle rcNormalPosition;
    }
    
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    
    #endregion
    
    #region Keyboard Simulation
    
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    
    public const int VK_MENU = 0x12; // Alt
    public const int VK_CONTROL = 0x11;
    public const int VK_SHIFT = 0x10;
    public const int VK_END = 0x23;
    public const int VK_NEXT = 0x22; // Page Down
    public const int KEYEVENTF_KEYUP = 0x0002;
    
    // Virtual key codes for common keys
    public static byte GetVirtualKeyCode(char c)
    {
        return (byte)char.ToUpper(c);
    }
    
    /// <summary>
    /// Send Alt+Key combination using raw keybd_event.
    /// </summary>
    public static void SendAltKey(char key)
    {
        byte vk = GetVirtualKeyCode(key);
        
        keybd_event(VK_MENU, 0, 0, UIntPtr.Zero);
        Thread.Sleep(20); // Delay for modifier registration
        keybd_event(vk, 0, 0, UIntPtr.Zero);
        Thread.Sleep(20); // Hold key
        keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        Thread.Sleep(10);
        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }

    /// <summary>
    /// Forcibly release all modifier keys.
    /// Matches Python's release of ctrl, shift, alt.
    /// </summary>
    public static void KeyUpModifiers()
    {
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }

    /// <summary>
    /// Send a hotkey combination (e.g., "ctrl+shift+r") using keybd_event.
    /// </summary>
    public static void SendHotkey(string hotkey)
    {
        var parts = hotkey.ToLower().Split('+');
        var modifiers = new List<byte>();
        byte mainKey = 0;

        foreach (var part in parts)
        {
            var p = part.Trim();
            if (p == "ctrl") modifiers.Add(VK_CONTROL);
            else if (p == "shift") modifiers.Add(VK_SHIFT);
            else if (p == "alt") modifiers.Add(VK_MENU);
            else if (p.Length == 1) mainKey = GetVirtualKeyCode(p[0]);
            else if (p.StartsWith("f") && int.TryParse(p.Substring(1), out var f)) mainKey = (byte)(0x6F + f);
        }

        foreach (var mod in modifiers)
        {
            keybd_event(mod, 0, 0, UIntPtr.Zero);
            Thread.Sleep(5);
        }

        if (mainKey != 0)
        {
            keybd_event(mainKey, 0, 0, UIntPtr.Zero);
            Thread.Sleep(10);
            keybd_event(mainKey, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            Thread.Sleep(5);
        }

        for (int i = modifiers.Count - 1; i >= 0; i--)
        {
            keybd_event(modifiers[i], 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            Thread.Sleep(5);
        }
    }
    
    #endregion
    
    #region Registry Access (for Dictation State)
    
    /// <summary>
    /// Check if microphone is in use via Windows registry.
    /// Matches Python's is_dictation_active_registry().
    /// </summary>
    public static bool? IsMicrophoneActiveFromRegistry()
    {
        try
        {
            string[] targetExes = { "msedgewebview2.exe", "webviewhost.exe", "mosaic" };
            string[] micRoots = {
                @"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged",
                @"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"
            };
            
            bool seenTarget = false;
            
            foreach (var rootPath in micRoots)
            {
                using var rootKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(rootPath);
                if (rootKey == null) continue;
                
                foreach (var subkeyName in rootKey.GetSubKeyNames())
                {
                    bool isTarget = false;
                    foreach (var exe in targetExes)
                    {
                        if (subkeyName.Contains(exe, StringComparison.OrdinalIgnoreCase))
                        {
                            isTarget = true;
                            break;
                        }
                    }
                    
                    if (!isTarget) continue;
                    
                    seenTarget = true;
                    
                    using var subkey = rootKey.OpenSubKey(subkeyName);
                    if (subkey == null) continue;
                    
                    var startVal = subkey.GetValue("LastUsedTimeStart");
                    var stopVal = subkey.GetValue("LastUsedTimeStop");
                    
                    if (startVal is long start && stopVal is long stop)
                    {
                        // Active if start != 0 and stop == 0
                        if (start != 0 && stop == 0)
                        {
                            return true;
                        }
                    }
                }
            }
            
            return seenTarget ? false : null;
        }
        catch (Exception ex)
        {
            Logger.Trace($"Registry dictation check failed: {ex.Message}");
            return null;
        }
    }
    
    #endregion
    
    #region Window Search Helpers
    
    /// <summary>
    /// Get window title text.
    /// </summary>
    public static string GetWindowTitle(IntPtr hWnd)
    {
        var sb = new System.Text.StringBuilder(256);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
    
    /// <summary>
    /// Find window by title containing keywords.
    /// </summary>
    public static IntPtr FindWindowByTitle(string[] mandatory, string[]? optional = null)
    {
        IntPtr found = IntPtr.Zero;
        
        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd)) return true;
            
            var title = GetWindowTitle(hWnd).ToLowerInvariant();
            
            // All mandatory keywords must be present
            foreach (var kw in mandatory)
            {
                if (!title.Contains(kw.ToLowerInvariant()))
                    return true; // Continue searching
            }
            
            // At least one optional keyword must be present (if specified)
            if (optional != null && optional.Length > 0)
            {
                bool hasOptional = false;
                foreach (var kw in optional)
                {
                    if (title.Contains(kw.ToLowerInvariant()))
                    {
                        hasOptional = true;
                        break;
                    }
                }
                if (!hasOptional) return true;
            }
            
            found = hWnd;
            return false; // Stop searching
        }, IntPtr.Zero);
        
        return found;
    }
    
    /// <summary>
    /// Activate a window (bring to foreground).
    /// Matches Python's _activate_window_by_title().
    /// </summary>
    public static bool ActivateWindow(IntPtr hWnd, int timeoutMs = 2000)
    {
        if (hWnd == IntPtr.Zero) return false;
        
        try
        {
            // Early exit if already active
            if (GetForegroundWindow() == hWnd) return true;
            
            // Get thread IDs
            uint currentThread = GetCurrentThreadId();
            uint targetThread = GetWindowThreadProcessId(hWnd, out _);
            
            // Attach thread input
            bool attached = false;
            if (currentThread != targetThread)
            {
                attached = AttachThreadInput(currentThread, targetThread, true);
            }
            
            try
            {
                // Check if minimized
                var placement = new WINDOWPLACEMENT { length = Marshal.SizeOf<WINDOWPLACEMENT>() };
                GetWindowPlacement(hWnd, ref placement);
                
                if (placement.showCmd == SW_SHOWMINIMIZED)
                {
                    ShowWindow(hWnd, SW_RESTORE);
                }
                
                BringWindowToTop(hWnd);
                SetForegroundWindow(hWnd);
            }
            finally
            {
                if (attached)
                {
                    AttachThreadInput(currentThread, targetThread, false);
                }
            }
            
            // Check if it worked
            Thread.Sleep(50);
            if (GetForegroundWindow() == hWnd) return true;
            
            // Fallback: Alt key trick (like AHK)
            keybd_event(VK_MENU, 0, 0, UIntPtr.Zero);
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            Thread.Sleep(20);
            SetForegroundWindow(hWnd);
            
            // Wait for activation
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (GetForegroundWindow() == hWnd) return true;
                Thread.Sleep(20);
            }
            
            return GetForegroundWindow() == hWnd;
        }
        catch (Exception ex)
        {
            Logger.Trace($"Window activation failed: {ex.Message}");
            return false;
        }
    }
    
    /// <summary>
    /// Forcefully activate Mosaic using aggressive techniques.
    /// Matches Python's _activate_mosaic_forcefully().
    /// </summary>
    public static bool ActivateMosaicForcefully()
    {
        var hWnd = FindWindowByTitle(new[] { "mosaic" }, new[] { "reporting", "info hub" });
        if (hWnd == IntPtr.Zero) return false;

        for (int i = 0; i < 3; i++) // Try up to 3 times
        {
            if (GetForegroundWindow() == hWnd) return true;

            try
            {
                uint currentThread = GetCurrentThreadId();
                uint targetThread = GetWindowThreadProcessId(hWnd, out _);

                AttachThreadInput(currentThread, targetThread, true);
                SwitchToThisWindow(hWnd, true);
                BringWindowToTop(hWnd);
                SetForegroundWindow(hWnd);
                AttachThreadInput(currentThread, targetThread, false);

                // Double-tap with gentle method
                Thread.Sleep(100); // Increased from 50ms
                ActivateWindow(hWnd);

                // Verification loop (200ms)
                var sw = Stopwatch.StartNew();
                while (sw.ElapsedMilliseconds < 200)
                {
                    if (GetForegroundWindow() == hWnd) return true;
                    Thread.Sleep(20);
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Force activate attempt {i+1} failed: {ex.Message}");
            }
            Thread.Sleep(100); // Wait before retry
        }

        return GetForegroundWindow() == hWnd;
    }
    
    #endregion
    
    #region Focus Restoration
    
    private static IntPtr _previousFocusHwnd = IntPtr.Zero;
    
    /// <summary>
    /// Save the current foreground window for later restoration.
    /// Skips saving if the current window is Mosaic itself.
    /// </summary>
    public static void SavePreviousFocus()
    {
        var foreground = GetForegroundWindow();
        if (foreground == IntPtr.Zero) 
        {
            _previousFocusHwnd = IntPtr.Zero;
            return;
        }
        
        var title = GetWindowTitle(foreground);
        // Don't save if it's already Mosaic
        if (title.Contains("Mosaic", StringComparison.OrdinalIgnoreCase))
        {
            _previousFocusHwnd = IntPtr.Zero;
            return;
        }
        
        _previousFocusHwnd = foreground;
        Logger.Trace($"Focus saved: {title}");
    }
    
    /// <summary>
    /// Restore focus to the previously saved window.
    /// </summary>
    public static void RestorePreviousFocus(int delayMs = 50)
    {
        if (_previousFocusHwnd == IntPtr.Zero) return;
        if (!IsWindow(_previousFocusHwnd))
        {
            _previousFocusHwnd = IntPtr.Zero;
            return;
        }
        
        Thread.Sleep(delayMs);
        SetForegroundWindow(_previousFocusHwnd);
        Logger.Trace($"Focus restored: {GetWindowTitle(_previousFocusHwnd)}");
        _previousFocusHwnd = IntPtr.Zero;
    }
    
    #endregion
    
    #region UI Automation Check (Legacy/Fallback)
    
    /// <summary>
    /// Check if microphone is active by scanning the UI tree for the recording icon.
    /// Matches Python's is_dictation_active_uia().
    /// </summary>
    public static bool IsMicrophoneActiveFromUia()
    {
        try
        {
            // Note: FlaUI is relatively slow for global scans.
            using var automation = new FlaUI.UIA3.UIA3Automation();
            var desktop = automation.GetDesktop();
            
            // 1. Broad search for any window with "Microphone recording" in title
            var micWindows = desktop.FindAllChildren(cf => cf.ByName("Microphone recording", FlaUI.Core.Definitions.PropertyConditionFlags.IgnoreCase));
            if (micWindows.Length > 0) return true;
            
            // 2. Search specific windows (Mosaic/SlimHub) for indicators
            var windows = desktop.FindAllChildren(cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window));
            foreach (var win in windows)
            {
                var title = win.Name;
                if (title.Contains("Mosaic", StringComparison.OrdinalIgnoreCase) || 
                    title.Contains("SlimHub", StringComparison.OrdinalIgnoreCase) ||
                    title.Contains("InteleViewer", StringComparison.OrdinalIgnoreCase))
                {
                    // Check for common indicator names/labels
                    var indicator = win.FindFirstDescendant(cf => 
                        cf.ByName("Microphone recording", FlaUI.Core.Definitions.PropertyConditionFlags.IgnoreCase)
                        .Or(cf.ByName("accessing your microphone", FlaUI.Core.Definitions.PropertyConditionFlags.IgnoreCase)));
                    
                    if (indicator != null) return true;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"UIA dictation check failed: {ex.Message}");
        }
        return false;
    }
    
    #endregion
    
    #region Custom Windows Messages

    public const int WM_TRIGGER_SCRAPE = 0x0401;
    public const int WM_TRIGGER_DEBUG = 0x0402;
    public const int WM_TRIGGER_BEEP = 0x0403;
    public const int WM_TRIGGER_SHOW_REPORT = 0x0404;
    public const int WM_TRIGGER_CAPTURE_SERIES = 0x0405;
    public const int WM_TRIGGER_GET_PRIOR = 0x0406;
    public const int WM_TRIGGER_TOGGLE_RECORD = 0x0407;
    public const int WM_TRIGGER_PROCESS_REPORT = 0x0408;
    public const int WM_TRIGGER_SIGN_REPORT = 0x0409;
    public const int WM_TRIGGER_OPEN_SETTINGS = 0x040A;
    public const int WM_TRIGGER_CREATE_IMPRESSION = 0x040B;
    public const int WM_TRIGGER_DISCARD_STUDY = 0x040C;
    public const int WM_TRIGGER_CHECK_UPDATES = 0x040D;

    #endregion

    #region WM_COPYDATA IPC (for RVUCounter integration)

    public const int WM_COPYDATA = 0x004A;

    // Message types for RVUCounter
    public const int MSG_STUDY_SIGNED = 1;           // Study was signed via MosaicTools
    public const int MSG_STUDY_CLOSED_UNSIGNED = 2;  // Study closed without being signed

    [StructLayout(LayoutKind.Sequential)]
    public struct COPYDATASTRUCT
    {
        public IntPtr dwData;    // Message type ID
        public int cbData;       // Data length in bytes
        public IntPtr lpData;    // Pointer to data
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindow(string? lpClassName, string lpWindowName);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, ref COPYDATASTRUCT lParam);

    /// <summary>
    /// Send a string message to RVUCounter via WM_COPYDATA.
    /// </summary>
    /// <param name="messageType">MSG_STUDY_SIGNED or MSG_STUDY_CLOSED_UNSIGNED</param>
    /// <param name="accession">The accession number</param>
    /// <returns>True if message was sent successfully</returns>
    public static bool SendToRvuCounter(int messageType, string accession)
    {
        try
        {
            // Find the RVUCounter window by title
            var rvuHandle = FindWindow(null, "RVUCounter");
            if (rvuHandle == IntPtr.Zero)
            {
                Logger.Trace($"SendToRvuCounter: RVUCounter window not found");
                return false;
            }

            // Prepare the string data
            var bytes = System.Text.Encoding.Unicode.GetBytes(accession + '\0');
            var dataPtr = Marshal.AllocHGlobal(bytes.Length);

            try
            {
                Marshal.Copy(bytes, 0, dataPtr, bytes.Length);

                var copyData = new COPYDATASTRUCT
                {
                    dwData = new IntPtr(messageType),
                    cbData = bytes.Length,
                    lpData = dataPtr
                };

                SendMessage(rvuHandle, WM_COPYDATA, IntPtr.Zero, ref copyData);

                var msgName = messageType == MSG_STUDY_SIGNED ? "SIGNED" : "CLOSED_UNSIGNED";
                Logger.Trace($"SendToRvuCounter: Sent {msgName} for '{accession}'");
                return true;
            }
            finally
            {
                Marshal.FreeHGlobal(dataPtr);
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"SendToRvuCounter error: {ex.Message}");
            return false;
        }
    }

    #endregion

    #region Window Utilities

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_SHOWWINDOW = 0x0040;

    /// <summary>
    /// Forcibly set a window to be TopMost.
    /// </summary>
    public static void ForceTopMost(IntPtr hWnd)
    {
        SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    }

    #endregion

    #region Windows Startup

    private const string StartupRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
    private const string AppName = "MosaicTools";

    /// <summary>
    /// Set or remove the app from Windows startup.
    /// </summary>
    public static void SetRunAtStartup(bool enable)
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(StartupRegistryKey, true);
            if (key == null)
            {
                Logger.Trace("SetRunAtStartup: Could not open registry key");
                return;
            }

            if (enable)
            {
                var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                if (!string.IsNullOrEmpty(exePath))
                {
                    key.SetValue(AppName, $"\"{exePath}\"");
                    Logger.Trace($"SetRunAtStartup: Enabled - {exePath}");
                }
            }
            else
            {
                key.DeleteValue(AppName, false);
                Logger.Trace("SetRunAtStartup: Disabled");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"SetRunAtStartup error: {ex.Message}");
        }
    }

    /// <summary>
    /// Check if the app is set to run at Windows startup.
    /// </summary>
    public static bool GetRunAtStartup()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(StartupRegistryKey, false);
            if (key == null) return false;

            var value = key.GetValue(AppName);
            return value != null;
        }
        catch (Exception ex)
        {
            Logger.Trace($"GetRunAtStartup error: {ex.Message}");
            return false;
        }
    }

    #endregion
}
