using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using SharpHook;
using SharpHook.Native;
using System.Collections.Concurrent;

namespace MosaicTools.Services;

/// <summary>
/// Global keyboard hook service using SharpHook.
/// Matches Python's keyboard library usage.
/// </summary>
public class KeyboardService : IDisposable
{
    private TaskPoolGlobalHook? _hook;
    private readonly ConcurrentDictionary<string, Action> _hotkeyActions = new();
    private readonly HashSet<KeyCode> _pressedModifiers = new();
    private volatile bool _isRecording = false;
    private Action<string>? _recordCallback;
    
    public event Action<string>? HotkeyTriggered;
    
    public void Start()
    {
        if (_hook != null) return;
        
        try
        {
            _hook = new TaskPoolGlobalHook();
            _hook.KeyPressed += OnKeyPressed;
            _hook.KeyReleased += OnKeyReleased;
            _hook.RunAsync();
            
            Logger.Trace("Keyboard hook started");
        }
        catch (Exception ex)
        {
            Logger.Trace($"Keyboard hook failed: {ex.Message}");
        }
    }
    
    public void Stop()
    {
        _hook?.Dispose();
        _hook = null;
    }
    
    /// <summary>
    /// Register a hotkey with an action.
    /// Format: "ctrl+shift+f1" or "alt+g"
    /// </summary>
    public void RegisterHotkey(string hotkey, Action action)
    {
        if (string.IsNullOrWhiteSpace(hotkey)) return;
        
        var normalized = NormalizeHotkey(hotkey);
        _hotkeyActions[normalized] = action;
        Logger.Trace($"Registered hotkey: {normalized}");
    }
    
    /// <summary>
    /// Unregister all hotkeys.
    /// </summary>
    public void ClearHotkeys()
    {
        _hotkeyActions.Clear();
    }
    
    /// <summary>
    /// Start recording a hotkey (for settings UI).
    /// </summary>
    public void StartRecording(Action<string> callback)
    {
        _isRecording = true;
        _recordCallback = callback;
        _pressedModifiers.Clear();
    }
    
    private void OnKeyPressed(object? sender, KeyboardHookEventArgs e)
    {
        try
        {
            var code = e.Data.KeyCode;
            var mask = e.RawEvent.Mask;
            
            // Track modifiers (keep manual tracking as fallback for recording mode)
            if (IsModifier(code))
            {
                _pressedModifiers.Add(code);
                return;
            }
            
            // Recording mode
            if (_isRecording && _recordCallback != null)
            {
                var recorded = BuildHotkeyString(code, mask);
                _isRecording = false;
                _recordCallback(recorded);
                _recordCallback = null;
                return;
            }
            
            // Check registered hotkeys
            var current = BuildHotkeyString(code, mask);
            if (_hotkeyActions.TryGetValue(current, out var action))
            {
                Logger.Trace($"Hotkey triggered: {current}");
                HotkeyTriggered?.Invoke(current);
                
                // Run action on thread pool to not block hook
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    try
                    {
                        action();
                    }
                    catch (Exception ex)
                    {
                        Logger.Trace($"Hotkey action error: {ex.Message}");
                    }
                });
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Key pressed handler error: {ex.Message}");
        }
    }
    
    private void OnKeyReleased(object? sender, KeyboardHookEventArgs e)
    {
        if (IsModifier(e.Data.KeyCode))
        {
            _pressedModifiers.Remove(e.Data.KeyCode);
        }
    }
    
    private bool IsModifier(KeyCode code)
    {
        return code is KeyCode.VcLeftControl or KeyCode.VcRightControl
            or KeyCode.VcLeftShift or KeyCode.VcRightShift
            or KeyCode.VcLeftAlt or KeyCode.VcRightAlt
            or KeyCode.VcLeftMeta or KeyCode.VcRightMeta;
    }
    
    private string BuildHotkeyString(KeyCode mainKey, ModifierMask mask)
    {
        var parts = new List<string>();
        
        // Add modifiers in order: ctrl, shift, alt, win
        bool ctrl = mask.HasFlag(ModifierMask.LeftCtrl) || mask.HasFlag(ModifierMask.RightCtrl) || 
                   _pressedModifiers.Contains(KeyCode.VcLeftControl) || _pressedModifiers.Contains(KeyCode.VcRightControl);
        bool shift = mask.HasFlag(ModifierMask.LeftShift) || mask.HasFlag(ModifierMask.RightShift) ||
                    _pressedModifiers.Contains(KeyCode.VcLeftShift) || _pressedModifiers.Contains(KeyCode.VcRightShift);
        bool alt = mask.HasFlag(ModifierMask.LeftAlt) || mask.HasFlag(ModifierMask.RightAlt) ||
                  _pressedModifiers.Contains(KeyCode.VcLeftAlt) || _pressedModifiers.Contains(KeyCode.VcRightAlt);
        bool win = mask.HasFlag(ModifierMask.LeftMeta) || mask.HasFlag(ModifierMask.RightMeta) ||
                  _pressedModifiers.Contains(KeyCode.VcLeftMeta) || _pressedModifiers.Contains(KeyCode.VcRightMeta);

        if (ctrl) parts.Add("ctrl");
        if (shift) parts.Add("shift");
        if (alt) parts.Add("alt");
        if (win) parts.Add("win");
        
        // Add main key
        var keyName = KeyCodeToName(mainKey);
        if (!string.IsNullOrEmpty(keyName))
            parts.Add(keyName);
        
        return string.Join("+", parts);
    }
    
    private static string KeyCodeToName(KeyCode code)
    {
        // SharpHook (libuiohook) KeyCodes are not contiguous ASCII/VirtualKey offsets.
        // Mapping common ones explicitly.
        return code switch
        {
            KeyCode.VcA => "a", KeyCode.VcB => "b", KeyCode.VcC => "c", KeyCode.VcD => "d",
            KeyCode.VcE => "e", KeyCode.VcF => "f", KeyCode.VcG => "g", KeyCode.VcH => "h",
            KeyCode.VcI => "i", KeyCode.VcJ => "j", KeyCode.VcK => "k", KeyCode.VcL => "l",
            KeyCode.VcM => "m", KeyCode.VcN => "n", KeyCode.VcO => "o", KeyCode.VcP => "p",
            KeyCode.VcQ => "q", KeyCode.VcR => "r", KeyCode.VcS => "s", KeyCode.VcT => "t",
            KeyCode.VcU => "u", KeyCode.VcV => "v", KeyCode.VcW => "w", KeyCode.VcX => "x",
            KeyCode.VcY => "y", KeyCode.VcZ => "z",
            KeyCode.Vc1 => "1", KeyCode.Vc2 => "2", KeyCode.Vc3 => "3", KeyCode.Vc4 => "4",
            KeyCode.Vc5 => "5", KeyCode.Vc6 => "6", KeyCode.Vc7 => "7", KeyCode.Vc8 => "8",
            KeyCode.Vc9 => "9", KeyCode.Vc0 => "0",
            KeyCode.VcF1 => "f1", KeyCode.VcF2 => "f2", KeyCode.VcF3 => "f3", KeyCode.VcF4 => "f4",
            KeyCode.VcF5 => "f5", KeyCode.VcF6 => "f6", KeyCode.VcF7 => "f7", KeyCode.VcF8 => "f8",
            KeyCode.VcF9 => "f9", KeyCode.VcF10 => "f10", KeyCode.VcF11 => "f11", KeyCode.VcF12 => "f12",
            KeyCode.VcSpace => "space",
            KeyCode.VcEnter => "enter",
            KeyCode.VcEscape => "escape",
            KeyCode.VcBackspace => "backspace",
            KeyCode.VcTab => "tab",
            KeyCode.VcCapsLock => "capslock",
            KeyCode.VcPrintScreen => "printscreen",
            KeyCode.VcScrollLock => "scrolllock",
            KeyCode.VcPause => "pause",
            KeyCode.VcInsert => "insert",
            KeyCode.VcDelete => "delete",
            KeyCode.VcHome => "home",
            KeyCode.VcEnd => "end",
            KeyCode.VcPageUp => "pageup",
            KeyCode.VcPageDown => "pagedown",
            KeyCode.VcUp => "up",
            KeyCode.VcDown => "down",
            KeyCode.VcLeft => "left",
            KeyCode.VcRight => "right",
            KeyCode.VcMinus => "-",
            KeyCode.VcEquals => "=",
            KeyCode.VcOpenBracket => "[",
            KeyCode.VcCloseBracket => "]",
            KeyCode.VcBackslash => "\\",
            KeyCode.VcSemicolon => ";",
            KeyCode.VcQuote => "'",
            KeyCode.VcComma => ",",
            KeyCode.VcPeriod => ".",
            KeyCode.VcSlash => "/",
            KeyCode.VcBackQuote => "`",
            _ => ""
        };
    }
    
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
        
        // Sort modifiers
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
