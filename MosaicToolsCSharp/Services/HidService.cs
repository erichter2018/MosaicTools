using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using HidSharp;

namespace MosaicTools.Services;

/// <summary>
/// PowerMic and SpeechMike HID device communication service.
/// Reads button presses from the vendor-specific HID interface.
/// </summary>
public class HidService : IDisposable
{
    // Nuance / Dictaphone Vendor IDs + Philips (0x0911)
    private static readonly int[] VendorIds = { 0x0554, 0x0558, 0x0911 };

    // Device type constants
    public const string DeviceAuto = "Auto";
    public const string DevicePowerMic = "PowerMic";
    public const string DeviceSpeechMike = "SpeechMike";

    // Button byte patterns [byte0, byte1, byte2] for PowerMic
    public static readonly Dictionary<string, byte[]> ButtonDefinitions = new()
    {
        ["Left Button"] = new byte[] { 0x00, 0x80, 0x00 },
        ["Right Button"] = new byte[] { 0x00, 0x00, 0x02 },
        ["T Button"] = new byte[] { 0x00, 0x01, 0x00 },
        ["Record Button"] = new byte[] { 0x00, 0x04, 0x00 },
        ["Skip Back"] = new byte[] { 0x00, 0x02, 0x00 },
        ["Skip Forward"] = new byte[] { 0x00, 0x08, 0x00 },
        ["Rewind"] = new byte[] { 0x00, 0x10, 0x00 },
        ["Fast Forward"] = new byte[] { 0x00, 0x20, 0x00 },
        ["Stop/Play"] = new byte[] { 0x00, 0x40, 0x00 },
        ["Checkmark"] = new byte[] { 0x00, 0x00, 0x01 }
    };

    // PowerMic buttons
    public static readonly string[] PowerMicButtons =
    {
        "Left Button", "Right Button", "Record Button", "T Button",
        "Skip Back", "Skip Forward", "Rewind", "Fast Forward",
        "Stop/Play", "Checkmark", "Scan Button"
    };

    // SpeechMike buttons (Event mode - use SpeechControl to configure front mouse buttons as hotkeys)
    public static readonly string[] SpeechMikeButtons =
    {
        "Record", "Play/Pause", "Fast Forward", "Rewind",
        "EoL", "Ins/Ovr", "-i-",
        "Index Finger", "F1", "F2", "F3", "F4"
    };

    // Legacy - returns all possible buttons for backwards compatibility
    public static readonly string[] AllButtons =
    {
        "Left Button", "Right Button", "Record Button", "T Button",
        "Skip Back", "Skip Forward", "Rewind", "Fast Forward",
        "Stop/Play", "Checkmark", "Scan Button",
        "Record", "Play/Pause", "EoL", "Ins/Ovr", "-i-",
        "Index Finger", "F1", "F2", "F3", "F4"
    };

    private HidDevice? _device;
    private HidStream? _stream;
    private Thread? _listenerThread;
    private volatile bool _running;
    private bool _isPhilips;
    private string _preferredDevice = DeviceAuto;

    // Debounce for SpeechMike phantom events from physical movement
    private string? _lastPhilipsButton;
    private DateTime _lastPhilipsButtonTime = DateTime.MinValue;
    private const int PhilipsDebounceMs = 300;

    public event Action<string>? ButtonPressed;
    public event Action<string>? DeviceConnected;
    public event Action<bool>? RecordButtonStateChanged; // For PTT: true=down, false=up

    public bool IsConnected => _stream != null;
    public string? ConnectedDeviceName { get; private set; }
    public bool IsPhilipsConnected => _isPhilips && IsConnected;

    /// <summary>
    /// Set the preferred device type. Pass DeviceAuto to connect to any available device.
    /// </summary>
    public void SetPreferredDevice(string deviceType)
    {
        _preferredDevice = deviceType;
        // If connected to wrong device, disconnect to force reconnect
        if (_stream != null && !string.IsNullOrEmpty(ConnectedDeviceName))
        {
            bool shouldDisconnect = deviceType switch
            {
                DevicePowerMic => _isPhilips,
                DeviceSpeechMike => !_isPhilips,
                _ => false
            };
            if (shouldDisconnect)
            {
                Logger.Trace($"Disconnecting {ConnectedDeviceName} to switch to preferred: {deviceType}");
                Disconnect();
            }
        }
    }

    /// <summary>
    /// Get list of available dictation microphones (PowerMic, SpeechMike).
    /// </summary>
    public static List<(string Name, string Type)> GetAvailableDevices()
    {
        var result = new List<(string Name, string Type)>();
        try
        {
            var devices = DeviceList.Local.GetHidDevices();
            foreach (var device in devices)
            {
                if (Array.IndexOf(VendorIds, device.VendorID) < 0) continue;
                bool isPhilips = device.VendorID == 0x0911;
                var name = device.GetProductName() ?? (isPhilips ? "SpeechMike" : "PowerMic");
                var type = isPhilips ? DeviceSpeechMike : DevicePowerMic;
                // Avoid duplicates (same type)
                if (!result.Exists(r => r.Type == type))
                    result.Add((name, type));
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"GetAvailableDevices error: {ex.Message}");
        }
        return result;
    }

    public void Start()
    {
        if (_running) return;

        _running = true;
        _listenerThread = new Thread(ListenerLoop) { IsBackground = true };
        _listenerThread.Start();
    }

    public void Stop()
    {
        _running = false;
        _listenerThread?.Join(1000);
        Disconnect();
    }

    private void ListenerLoop()
    {
        bool lastRecordDown = false;
        byte[] lastData = new byte[4];
        ushort lastPhilipsButtons = 0;

        while (_running)
        {
            try
            {
                // Connection logic
                if (_stream == null)
                {
                    if (!TryConnect())
                    {
                        Thread.Sleep(5000);
                        continue;
                    }
                }

                // Read data (non-blocking with timeout)
                var buffer = new byte[64];
                int bytesRead = 0;

                try
                {
                    bytesRead = _stream!.Read(buffer, 0, buffer.Length);
                }
                catch (TimeoutException)
                {
                    Thread.Sleep(10);
                    continue;
                }
                catch
                {
                    Disconnect();
                    continue;
                }

                if (_isPhilips)
                {
                    // Philips SpeechMike: reports via vendor-specific interface (0xFFA0)
                    // HidSharp adds 1-byte report ID prefix, so actual data is offset by 1
                    // Byte 8 (index 8) = F1/F2/F3/F4/Index finger buttons
                    // Byte 9 (index 9) = Record/FF/Rewind/EoL/Ins/Ovr/-i-

                    if (bytesRead < 10)
                    {
                        Logger.Trace($"SpeechMike: Short read ({bytesRead} bytes), skipping");
                        continue;
                    }

                    byte btn7 = buffer[8];  // F buttons (user's byte 7)
                    byte btn8 = buffer[9];  // Record/EoL etc (user's byte 8)
                    ushort btnCombined = (ushort)((btn7 << 8) | btn8);

                    // Debounce
                    if (btnCombined == lastPhilipsButtons)
                    {
                        Thread.Sleep(10);
                        continue;
                    }

                    ushort prevButtons = lastPhilipsButtons;
                    lastPhilipsButtons = btnCombined;

                    // Ignore compound events (multiple bits set) - these are status updates, not button presses
                    // Must check BEFORE Record state tracking to avoid phantom PTT releases
                    int bitsSet = System.Numerics.BitOperations.PopCount(btn7) + System.Numerics.BitOperations.PopCount(btn8);
                    if (bitsSet > 1)
                    {
                        Logger.Trace($"SpeechMike: Ignoring compound event (btn7=0x{btn7:X2}, btn8=0x{btn8:X2}, {bitsSet} bits set)");
                        // Reset so the next real button press is treated as fresh
                        lastPhilipsButtons = 0;
                        continue;
                    }

                    // Track Record button state for PTT (after compound filter)
                    bool recordDown = (btn8 & 0x01) != 0;
                    if (recordDown != lastRecordDown)
                    {
                        RecordButtonStateChanged?.Invoke(recordDown);
                        lastRecordDown = recordDown;
                    }

                    // Button release - don't fire any action
                    if (btnCombined == 0x0000) continue;

                    // Only fire on fresh button press
                    if (prevButtons != 0x0000) continue;

                    // Match SpeechMike buttons
                    string? matchedButton = null;

                    // Byte 8 buttons (exact match since we already filtered compound events)
                    if (btn8 == 0x01) matchedButton = "Record";
                    else if (btn8 == 0x04) matchedButton = "Play/Pause";
                    else if (btn8 == 0x08) matchedButton = "Fast Forward";
                    else if (btn8 == 0x10) matchedButton = "Rewind";
                    else if (btn8 == 0x20) matchedButton = "EoL";
                    else if (btn8 == 0x40) matchedButton = "Ins/Ovr";
                    else if (btn8 == 0x80) matchedButton = "-i-";

                    // Byte 7 buttons
                    else if (btn7 == 0x02) matchedButton = "F1";
                    else if (btn7 == 0x04) matchedButton = "F2";
                    else if (btn7 == 0x08) matchedButton = "F3";
                    else if (btn7 == 0x10) matchedButton = "F4";
                    else if (btn7 == 0x20) matchedButton = "Index Finger";

                    if (matchedButton != null)
                    {
                        // Debounce: ignore same button firing again within 300ms (phantom from movement)
                        var now = DateTime.UtcNow;
                        if (matchedButton == _lastPhilipsButton &&
                            (now - _lastPhilipsButtonTime).TotalMilliseconds < PhilipsDebounceMs)
                        {
                            Logger.Trace($"SpeechMike DEBOUNCED: {matchedButton} ({(now - _lastPhilipsButtonTime).TotalMilliseconds:F0}ms since last)");
                            continue;
                        }
                        _lastPhilipsButton = matchedButton;
                        _lastPhilipsButtonTime = now;

                        Logger.Trace($"SpeechMike button: {matchedButton} (btn7=0x{btn7:X2}, btn8=0x{btn8:X2}, prev=0x{prevButtons:X4})");
                        ButtonPressed?.Invoke(matchedButton);
                    }
                }
                else
                {
                    // PowerMic path
                    if (bytesRead < 3) continue;

                    var btnData = new byte[] {
                        buffer[0],
                        buffer[1],
                        buffer[2],
                        bytesRead > 3 ? buffer[3] : (byte)0
                    };

                    // Debounce
                    if (btnData[0] == lastData[0] && btnData[1] == lastData[1] &&
                        btnData[2] == lastData[2] && btnData[3] == lastData[3])
                    {
                        Thread.Sleep(10);
                        continue;
                    }
                    Array.Copy(btnData, lastData, 4);

                    // Track Record Button state for PTT
                    bool recordDown = (btnData[2] & 0x04) != 0;
                    if (recordDown != lastRecordDown)
                    {
                        RecordButtonStateChanged?.Invoke(recordDown);
                        lastRecordDown = recordDown;
                    }

                    // Only fire if button pressed
                    if (btnData[1] == 0 && btnData[2] == 0 && btnData[3] == 0)
                        continue;

                    // Match button pattern
                    string? matchedButton = null;

                    // Byte 3: Right Button and Checkmark
                    if ((btnData[3] & 0x02) != 0) matchedButton = "Right Button";
                    else if ((btnData[3] & 0x40) != 0) matchedButton = "Right Button";
                    else if ((btnData[3] & 0x01) != 0) matchedButton = "Checkmark";

                    // Byte 2: Core navigation and record
                    else if ((btnData[2] & 0x80) != 0) matchedButton = "Left Button";
                    else if ((btnData[2] & 0x20) != 0) matchedButton = "Fast Forward";
                    else if ((btnData[2] & 0x10) != 0) matchedButton = "Rewind";
                    else if ((btnData[2] & 0x08) != 0) matchedButton = "Skip Forward";
                    else if ((btnData[2] & 0x04) != 0) matchedButton = "Record Button";
                    else if ((btnData[2] & 0x02) != 0) matchedButton = "Skip Back";
                    else if ((btnData[2] & 0x01) != 0) matchedButton = "T Button";
                    else if ((btnData[2] & 0x40) != 0) matchedButton = "Stop/Play";

                    // Byte 1: Support Tab key
                    else if ((btnData[1] & 0x01) != 0) matchedButton = "T Button";

                    if (matchedButton != null)
                    {
                        Logger.Trace($"PowerMic button: {matchedButton}");
                        ButtonPressed?.Invoke(matchedButton);
                    }
                }

                Thread.Sleep(10);
            }
            catch (Exception ex)
            {
                Logger.Trace($"HID listener error: {ex.Message}");
                Thread.Sleep(2000);
            }
        }
    }

    private bool TryConnect()
    {
        try
        {
            var devices = DeviceList.Local.GetHidDevices();

            foreach (var device in devices)
            {
                if (Array.IndexOf(VendorIds, device.VendorID) < 0) continue;

                bool isPhilips = device.VendorID == 0x0911;

                // Check if this device matches the preference
                if (_preferredDevice == DevicePowerMic && isPhilips) continue;
                if (_preferredDevice == DeviceSpeechMike && !isPhilips) continue;


                try
                {
                    var maxInputLen = device.GetMaxInputReportLength();
                    Logger.Trace($"Trying HID device: VID=0x{device.VendorID:X4}, PID=0x{device.ProductID:X4}, MaxInput={maxInputLen}");

                    // SpeechMike needs at least 10 bytes for button data (bytes 8-9)
                    if (isPhilips && maxInputLen < 10)
                    {
                        Logger.Trace($"Skipping SpeechMike interface (MaxInput={maxInputLen} too small, need 10+)");
                        continue;
                    }

                    var stream = device.Open();
                    stream.ReadTimeout = 100;

                    _device = device;
                    _stream = stream;
                    _isPhilips = isPhilips;

                    var name = device.GetProductName() ?? (_isPhilips ? "SpeechMike" : "PowerMic");
                    ConnectedDeviceName = name;
                    Logger.Trace($"Connected to {name} (VID=0x{device.VendorID:X4}, PID=0x{device.ProductID:X4})");
                    DeviceConnected?.Invoke($"Connected to {name}");

                    return true;
                }
                catch (Exception ex)
                {
                    Logger.Trace($"HID Connection Failed (VID=0x{device.VendorID:X4}, PID=0x{device.ProductID:X4}): {ex.Message}");
                }
            }

        }
        catch (Exception ex)
        {
            Logger.Trace($"HID enumerate error: {ex.Message}");
        }

        return false;
    }

    private void Disconnect()
    {
        _stream?.Dispose();
        _stream = null;
        _device = null;
        _isPhilips = false;
        ConnectedDeviceName = null;
    }

    public void Dispose()
    {
        Stop();
    }
}
