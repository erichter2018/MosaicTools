using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MosaicTools.Services;

/// <summary>
/// Global keyboard hook service using Win32 WH_KEYBOARD_LL.
/// </summary>
public class KeyboardService : IDisposable
{
    // Win32 interop
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_SYSKEYDOWN = 0x0104;

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

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public uint vkCode;
        public uint scanCode;
        public uint flags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    private const int VK_LCONTROL = 0xA2, VK_RCONTROL = 0xA3;
    private const int VK_LSHIFT = 0xA0, VK_RSHIFT = 0xA1;
    private const int VK_LMENU = 0xA4, VK_RMENU = 0xA5;
    private const int VK_LWIN = 0x5B, VK_RWIN = 0x5C;

    private IntPtr _hookId = IntPtr.Zero;
    private LowLevelKeyboardProc? _hookProc; // prevent GC of delegate
    private readonly ConcurrentDictionary<string, Action> _hotkeyActions = new();
    private readonly ConcurrentDictionary<string, Action> _directHotkeyActions = new();
    private volatile bool _isRecording = false;
    private Action<string>? _recordCallback;
    private readonly ConcurrentDictionary<string, DateTime> _lastTriggers = new();

    public event Action<string>? HotkeyTriggered;

    public void Start()
    {
        if (_hookId != IntPtr.Zero) return;

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
            }
            else
            {
                Logger.Trace("Keyboard hook started (Win32)");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Keyboard hook failed: {ex.Message}");
        }
    }

    public void Stop()
    {
        if (_hookId != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hookId);
            _hookId = IntPtr.Zero;
        }
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

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
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
                    try { directAction(); }
                    catch (Exception ex) { Logger.Trace($"Direct hotkey error: {ex.Message}"); }
                    return CallNextHookEx(_hookId, nCode, wParam, lParam);
                }

                // Normal hotkeys (run on ThreadPool)
                if (_hotkeyActions.TryGetValue(hotkeyStr, out var action))
                {
                    if (_lastTriggers.TryGetValue(hotkeyStr, out var lastNormal) && (now - lastNormal).TotalMilliseconds < 300)
                        return CallNextHookEx(_hookId, nCode, wParam, lParam);
                    _lastTriggers[hotkeyStr] = now;
                    Logger.Trace($"Hotkey triggered: {hotkeyStr}");
                    HotkeyTriggered?.Invoke(hotkeyStr);
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
        Stop();
    }
}
