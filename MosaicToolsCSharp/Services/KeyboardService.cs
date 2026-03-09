using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MosaicTools.Services;

/// <summary>
/// Global keyboard hook service using Win32 WH_KEYBOARD_LL.
/// Hook runs on a dedicated thread with its own message pump for maximum reliability.
/// </summary>
public class KeyboardService : IDisposable
{
    // Win32 interop
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int WM_QUIT = 0x0012;

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    [DllImport("user32.dll")]
    private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    [DllImport("user32.dll")]
    private static extern bool PostThreadMessage(uint idThread, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref MSG lpMsg);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref MSG lpMsg);

    // Session change notification
    [DllImport("wtsapi32.dll")]
    private static extern bool WTSRegisterSessionNotification(IntPtr hWnd, int dwFlags);

    [DllImport("wtsapi32.dll")]
    private static extern bool WTSUnRegisterSessionNotification(IntPtr hWnd);

    [StructLayout(LayoutKind.Sequential)]
    private struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public uint vkCode;
        public uint scanCode;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public System.Drawing.Point pt;
    }

    private const int VK_LCONTROL = 0xA2, VK_RCONTROL = 0xA3;
    private const int VK_LSHIFT = 0xA0, VK_RSHIFT = 0xA1;
    private const int VK_LMENU = 0xA4, VK_RMENU = 0xA5;
    private const int VK_LWIN = 0x5B, VK_RWIN = 0x5C;

    // Custom message ID for reinstall request
    private const uint WM_USER_REINSTALL = 0x0400 + 1;
    private const int WM_WTSSESSION_CHANGE = 0x02B1;
    private const int WTS_SESSION_UNLOCK = 0x8;
    private const int NOTIFY_FOR_THIS_SESSION = 0;

    private IntPtr _hookId = IntPtr.Zero;
    private LowLevelKeyboardProc? _hookProc; // prevent GC of delegate
    private readonly ConcurrentDictionary<string, Action> _hotkeyActions = new();
    private readonly ConcurrentDictionary<string, Action> _directHotkeyActions = new();
    private volatile bool _isRecording = false;
    private long _lastHookCallbackTicks = DateTime.UtcNow.Ticks;
    private System.Threading.Timer? _hookHealthTimer;
    private Action<string>? _recordCallback;
    private readonly ConcurrentDictionary<string, DateTime> _lastTriggers = new();

    // Dedicated hook thread
    private Thread? _hookThread;
    private uint _hookThreadId;
    private volatile bool _hookThreadRunning;
    private readonly ManualResetEventSlim _hookThreadReady = new(false);

    // Health check backoff: prevents constant reinstall when user is mouse-only (no keyboard)
    // GetLastInputInfo includes mouse movement, so idle=0 doesn't mean keyboard is active.
    private int _reinstallsSinceLastCallback;  // consecutive reinstalls without a callback in between
    private int _healthCheckSkipsRemaining;    // backoff counter — skip this many checks before trying again

    // UI thread synchronization (kept for session unlock listener)
    public System.Windows.Forms.Control? UiSyncTarget { get; set; }

    public void Start()
    {
        if (_hookThreadRunning) return;

        LogLowLevelHooksTimeout();
        StartHookThread();

        // Health timer: check every 5s, reinstall if no callbacks for 15s while user is active
        _hookHealthTimer ??= new System.Threading.Timer(CheckHookHealth, null, 5_000, 5_000);
    }

    /// <summary>
    /// Log the LowLevelHooksTimeout registry value at startup for diagnostics.
    /// </summary>
    private static void LogLowLevelHooksTimeout()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"Control Panel\Desktop");
            var val = key?.GetValue("LowLevelHooksTimeout");
            Logger.Trace(val != null
                ? $"Keyboard hook: LowLevelHooksTimeout registry = {val}ms"
                : "Keyboard hook: LowLevelHooksTimeout not set (default ~300ms)");
        }
        catch { }
    }

    private void StartHookThread()
    {
        _hookThreadReady.Reset();
        _hookThread = new Thread(HookThreadProc)
        {
            Name = "KeyboardHookThread",
            IsBackground = true
        };
        _hookThread.SetApartmentState(ApartmentState.STA);
        _hookThread.Start();

        // Wait up to 3s for hook thread to be ready
        if (!_hookThreadReady.Wait(3000))
        {
            Logger.Trace("Keyboard hook: WARNING - hook thread did not become ready in 3s");
        }
    }

    private void HookThreadProc()
    {
        _hookThreadId = GetCurrentThreadId();
        _hookThreadRunning = true;

        try
        {
            _hookProc = HookCallback;
            using var curProcess = Process.GetCurrentProcess();
            using var curModule = curProcess.MainModule!;
            _hookId = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, GetModuleHandle(curModule.ModuleName), 0);

            if (_hookId == IntPtr.Zero)
            {
                var error = Marshal.GetLastWin32Error();
                Logger.Trace($"Keyboard hook failed: SetWindowsHookEx error={error}");
                _hookThreadRunning = false;
                _hookThreadReady.Set();
                return;
            }

            Logger.Trace("Keyboard hook started on dedicated thread");
            Interlocked.Exchange(ref _lastHookCallbackTicks, DateTime.UtcNow.Ticks);
            _hookThreadReady.Set();

            // Set up session unlock listener on UI thread
            SetupSessionNotification();

            // Message pump — keeps the hook alive
            while (GetMessage(out var msg, IntPtr.Zero, 0, 0) > 0)
            {
                if (msg.message == WM_USER_REINSTALL)
                {
                    // Reinstall request from health timer
                    ReinstallHookOnThread();
                    continue;
                }
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Keyboard hook thread error: {ex.Message}");
        }
        finally
        {
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }
            _hookThreadRunning = false;
            _hookThreadReady.Set(); // unblock anyone waiting
        }
    }

    /// <summary>
    /// Reinstall the hook on the hook thread (called from within the message pump).
    /// </summary>
    private void ReinstallHookOnThread()
    {
        try
        {
            if (_hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
            }

            _hookProc = HookCallback;
            using var curProcess = Process.GetCurrentProcess();
            using var curModule = curProcess.MainModule!;
            _hookId = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, GetModuleHandle(curModule.ModuleName), 0);

            if (_hookId == IntPtr.Zero)
            {
                var error = Marshal.GetLastWin32Error();
                Logger.Trace($"Keyboard hook reinstall failed: error={error}");
            }
            else
            {
                Interlocked.Exchange(ref _lastHookCallbackTicks, DateTime.UtcNow.Ticks);
                Logger.Trace("Keyboard hook reinstalled successfully");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Keyboard hook reinstall error: {ex.Message}");
        }
    }

    /// <summary>
    /// Set up session unlock detection. When the workstation unlocks, proactively reinstall the hook.
    /// </summary>
    private void SetupSessionNotification()
    {
        try
        {
            var syncTarget = UiSyncTarget;
            if (syncTarget == null || !syncTarget.IsHandleCreated || syncTarget.IsDisposed) return;

            syncTarget.BeginInvoke(() =>
            {
                try
                {
                    Microsoft.Win32.SystemEvents.SessionSwitch += OnSessionSwitch;
                    Logger.Trace("Keyboard hook: session unlock listener registered");
                }
                catch (Exception ex)
                {
                    Logger.Trace($"Session notification setup error: {ex.Message}");
                }
            });
        }
        catch (Exception ex)
        {
            Logger.Trace($"Session notification setup error: {ex.Message}");
        }
    }

    private void OnSessionSwitch(object? sender, Microsoft.Win32.SessionSwitchEventArgs e)
    {
        if (e.Reason == Microsoft.Win32.SessionSwitchReason.SessionUnlock)
        {
            Logger.Trace("Session unlocked — proactively reinstalling keyboard hook");
            _reinstallsSinceLastCallback = 0;
            _healthCheckSkipsRemaining = 0;
            RequestReinstall();
        }
    }

    /// <summary>
    /// Request hook reinstall by posting a message to the hook thread's message pump.
    /// Thread-safe — can be called from any thread.
    /// </summary>
    private void RequestReinstall()
    {
        if (_hookThreadRunning && _hookThreadId != 0)
        {
            PostThreadMessage(_hookThreadId, WM_USER_REINSTALL, IntPtr.Zero, IntPtr.Zero);
        }
    }

    public void Stop()
    {
        if (_hookThreadRunning && _hookThreadId != 0)
        {
            // Post WM_QUIT to exit the message pump
            PostThreadMessage(_hookThreadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            _hookThread?.Join(2000);
        }

        if (_hookId != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hookId);
            _hookId = IntPtr.Zero;
        }
        _hookThreadRunning = false;
    }

    public void RegisterHotkey(string hotkey, Action action)
    {
        if (string.IsNullOrWhiteSpace(hotkey)) return;
        var normalized = NormalizeHotkey(hotkey);
        _hotkeyActions[normalized] = action;
        Logger.Trace($"Registered hotkey: {normalized}");
    }

    public void RegisterDirectHotkey(string hotkey, Action action)
    {
        if (string.IsNullOrWhiteSpace(hotkey)) return;
        var normalized = NormalizeHotkey(hotkey);
        _directHotkeyActions[normalized] = action;
        Logger.Trace($"Registered direct hotkey: {normalized}");
    }

    public void ClearHotkeys()
    {
        _hotkeyActions.Clear();
        _directHotkeyActions.Clear();
    }

    public bool IsWinKeyHeld()
    {
        return IsKeyDown(VK_LWIN) || IsKeyDown(VK_RWIN);
    }

    public void StartRecording(Action<string> callback)
    {
        _isRecording = true;
        _recordCallback = callback;
    }

    /// <summary>
    /// Returns seconds since last user input (keyboard OR mouse).
    /// </summary>
    private static double GetIdleSeconds()
    {
        var info = new LASTINPUTINFO { cbSize = (uint)Marshal.SizeOf<LASTINPUTINFO>() };
        if (!GetLastInputInfo(ref info)) return 0;
        return (Environment.TickCount - (int)info.dwTime) / 1000.0;
    }

    private void CheckHookHealth(object? state)
    {
        try
        {
            if (!_hookThreadRunning)
            {
                // Hook thread died — restart it
                Logger.Trace("Hook health: hook thread not running — restarting");
                _reinstallsSinceLastCallback = 0;
                _healthCheckSkipsRemaining = 0;
                StartHookThread();
                return;
            }

            if (_hotkeyActions.Count == 0 && _directHotkeyActions.Count == 0) return;

            // Backoff: skip checks after consecutive reinstalls without any callbacks.
            // This prevents constant reinstall spam when user is mouse/mic-only (no keyboard).
            // GetLastInputInfo counts mouse movement as activity, so idle<15s doesn't mean keyboard is in use.
            if (_healthCheckSkipsRemaining > 0)
            {
                _healthCheckSkipsRemaining--;
                return;
            }

            var secondsSinceCallback = (DateTime.UtcNow.Ticks - Interlocked.Read(ref _lastHookCallbackTicks))
                                        / (double)TimeSpan.TicksPerSecond;
            var idleSeconds = GetIdleSeconds();

            // Reinstall if: user has been active (input within 15s) but hook hasn't fired for 15s
            if (secondsSinceCallback > 15 && idleSeconds < 15)
            {
                _reinstallsSinceLastCallback++;

                if (_reinstallsSinceLastCallback <= 2)
                {
                    // First 2 reinstalls: try immediately (hook might genuinely be dead)
                    Logger.Trace($"Hook health: User active (idle {idleSeconds:F0}s) but no hook callbacks for {secondsSinceCallback:F0}s — reinstall #{_reinstallsSinceLastCallback}");
                    RequestReinstall();
                }
                else
                {
                    // 3+ reinstalls with no callback: user is probably mouse/mic-only.
                    // Back off exponentially: skip 12/24/48/... checks (60s/120s/240s/... at 5s interval), capped at 5 min.
                    var skipCount = Math.Min(60, 12 * (1 << (_reinstallsSinceLastCallback - 3)));
                    _healthCheckSkipsRemaining = skipCount;
                    Logger.Trace($"Hook health: {_reinstallsSinceLastCallback} reinstalls without callback — backing off {skipCount * 5}s (user likely mouse/mic-only)");
                    RequestReinstall();
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Hook health check error: {ex.Message}");
        }
    }

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        Interlocked.Exchange(ref _lastHookCallbackTicks, DateTime.UtcNow.Ticks);
        _reinstallsSinceLastCallback = 0;
        _healthCheckSkipsRemaining = 0;

        if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
        {
            try
            {
                var hookStruct = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                var vk = (int)hookStruct.vkCode;

                if (IsModifierVK(vk))
                    return CallNextHookEx(_hookId, nCode, wParam, lParam);

                bool ctrl = IsKeyDown(VK_LCONTROL) || IsKeyDown(VK_RCONTROL);
                bool shift = IsKeyDown(VK_LSHIFT) || IsKeyDown(VK_RSHIFT);
                bool alt = IsKeyDown(VK_LMENU) || IsKeyDown(VK_RMENU);
                bool win = IsKeyDown(VK_LWIN) || IsKeyDown(VK_RWIN);

                var keyName = VKToName(vk);
                if (string.IsNullOrEmpty(keyName))
                    return CallNextHookEx(_hookId, nCode, wParam, lParam);

                var hotkeyStr = BuildHotkeyString(keyName, ctrl, shift, alt, win);

                // Recording mode (settings UI hotkey capture)
                if (_isRecording && _recordCallback != null)
                {
                    _isRecording = false;
                    var cb = _recordCallback;
                    _recordCallback = null;
                    cb(hotkeyStr);
                    return CallNextHookEx(_hookId, nCode, wParam, lParam);
                }

                var now = DateTime.UtcNow;

                // Direct hotkeys (execute immediately on hook thread)
                if (_directHotkeyActions.TryGetValue(hotkeyStr, out var directAction))
                {
                    if (_lastTriggers.TryGetValue(hotkeyStr, out var lastDirect) && (now - lastDirect).TotalMilliseconds < 100)
                        return CallNextHookEx(_hookId, nCode, wParam, lParam);
                    _lastTriggers[hotkeyStr] = now;
                    Logger.Trace($"Direct hotkey triggered: {hotkeyStr}");
                    var action2 = directAction;
                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        try { action2(); }
                        catch (Exception ex) { Logger.Trace($"Direct hotkey error: {ex.Message}"); }
                    });
                    return CallNextHookEx(_hookId, nCode, wParam, lParam);
                }

                // Normal hotkeys (run on ThreadPool)
                if (_hotkeyActions.TryGetValue(hotkeyStr, out var action))
                {
                    if (_lastTriggers.TryGetValue(hotkeyStr, out var lastNormal) && (now - lastNormal).TotalMilliseconds < 300)
                        return CallNextHookEx(_hookId, nCode, wParam, lParam);
                    _lastTriggers[hotkeyStr] = now;
                    Logger.Trace($"Hotkey triggered: {hotkeyStr}");
                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        try { action(); }
                        catch (Exception ex) { Logger.Trace($"Hotkey action error: {ex.Message}"); }
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Hook callback error: {ex.Message}");
            }
        }

        return CallNextHookEx(_hookId, nCode, wParam, lParam);
    }

    private static bool IsKeyDown(int vk) => (GetAsyncKeyState(vk) & 0x8000) != 0;

    private static bool IsModifierVK(int vk) =>
        vk is VK_LCONTROL or VK_RCONTROL or VK_LSHIFT or VK_RSHIFT
           or VK_LMENU or VK_RMENU or VK_LWIN or VK_RWIN;

    private static string BuildHotkeyString(string keyName, bool ctrl, bool shift, bool alt, bool win)
    {
        var parts = new List<string>();
        if (ctrl) parts.Add("ctrl");
        if (shift) parts.Add("shift");
        if (alt) parts.Add("alt");
        if (win) parts.Add("win");
        parts.Add(keyName);
        return string.Join("+", parts);
    }

    private static string VKToName(int vk) => vk switch
    {
        >= 0x41 and <= 0x5A => ((char)(vk + 32)).ToString(), // A-Z → a-z
        >= 0x30 and <= 0x39 => ((char)vk).ToString(),        // 0-9
        >= 0x60 and <= 0x69 => ((char)('0' + vk - 0x60)).ToString(), // NumPad 0-9
        >= 0x70 and <= 0x7B => $"f{vk - 0x6F}",              // F1-F12
        0x20 => "space", 0x0D => "enter", 0x1B => "escape",
        0x08 => "backspace", 0x09 => "tab", 0x14 => "capslock",
        0x2C => "printscreen", 0x91 => "scrolllock", 0x13 => "pause",
        0x2D => "insert", 0x2E => "delete", 0x24 => "home", 0x23 => "end",
        0x21 => "pageup", 0x22 => "pagedown",
        0x26 => "up", 0x28 => "down", 0x25 => "left", 0x27 => "right",
        0xBD => "-", 0xBB => "=", 0xDB => "[", 0xDD => "]", 0xDC => "\\",
        0xBA => ";", 0xDE => "'", 0xBC => ",", 0xBE => ".", 0xBF => "/",
        0xC0 => "`",
        _ => ""
    };

    private static string NormalizeHotkey(string hotkey)
    {
        var parts = hotkey.ToLowerInvariant().Split('+', StringSplitOptions.RemoveEmptyEntries);
        var mods = new List<string>();
        var keys = new List<string>();

        foreach (var p in parts)
        {
            var trimmed = p.Trim();
            if (trimmed is "ctrl" or "control") mods.Add("ctrl");
            else if (trimmed is "shift") mods.Add("shift");
            else if (trimmed is "alt") mods.Add("alt");
            else if (trimmed is "win" or "windows" or "meta") mods.Add("win");
            else keys.Add(trimmed);
        }

        mods.Sort((a, b) =>
        {
            var order = new[] { "ctrl", "shift", "alt", "win" };
            return Array.IndexOf(order, a).CompareTo(Array.IndexOf(order, b));
        });

        return string.Join("+", mods.Concat(keys));
    }

    public void Dispose()
    {
        try { Microsoft.Win32.SystemEvents.SessionSwitch -= OnSessionSwitch; } catch { }
        _hookHealthTimer?.Dispose();
        _hookHealthTimer = null;
        Stop();
        _hookThreadReady.Dispose();
    }
}
