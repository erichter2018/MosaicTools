using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using HidSharp;

namespace MosaicTools.Services;

/// <summary>
/// PowerMic HID device communication service.
/// Matches Python's background_listener() and HID logic.
/// </summary>
public class HidService : IDisposable
{
    // Nuance / Dictaphone Vendor IDs
    private static readonly int[] VendorIds = { 0x0554, 0x0558 };
    
    // Button byte patterns [byte0, byte1, byte2]
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
    
    public static readonly string[] AllButtons = 
    {
        "Left Button", "Right Button", "Record Button", "T Button",
        "Skip Back", "Skip Forward", "Rewind", "Fast Forward", 
        "Stop/Play", "Checkmark", "Scan Button", "F1", "F2", "F3"
    };
    
    private HidDevice? _device;
    private HidStream? _stream;
    private Thread? _listenerThread;
    private volatile bool _running;
    private DateTime _lastSyncTime = DateTime.Now;
    
    public event Action<string>? ButtonPressed;
    public event Action<string>? DeviceConnected;
    public event Action<bool>? RecordButtonStateChanged; // For PTT: true=down, false=up
    
    public bool IsConnected => _stream != null;
    
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
                    // Device disconnected
                    Disconnect();
                    continue;
                }
                
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
                
                // Track Record Button state for PTT (bit 0x04 in byte 2 - shifted by ReportID)
                bool recordDown = (btnData[2] & 0x04) != 0;
                if (recordDown != lastRecordDown)
                {
                    Logger.Trace($"PTT State: {(recordDown ? "DOWN" : "UP")}");
                    RecordButtonStateChanged?.Invoke(recordDown);
                    lastRecordDown = recordDown;
                }
                
                // Only fire action if at least one button bit is set in ANY potential data byte (1, 2, or 3)
                if (btnData[1] == 0 && btnData[2] == 0 && btnData[3] == 0)
                {
                    continue;
                }
                
                // Match button pattern using bitwise priority across ALL data bytes
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
                
                // Byte 1: Support Tab and Function keys
                else if ((btnData[1] & 0x01) != 0) matchedButton = "T Button";
                else if ((btnData[1] & 0x10) != 0) matchedButton = "F1";
                else if ((btnData[1] & 0x20) != 0) matchedButton = "F2";
                else if ((btnData[1] & 0x40) != 0) matchedButton = "F3";
                
                if (matchedButton != null)
                {
                    Logger.Trace($"HID Button Match: {matchedButton} (Data: {BitConverter.ToString(btnData)})");
                    ButtonPressed?.Invoke(matchedButton);
                }
                else
                {
                   // Log any non-zero pattern that didn't match
                   Logger.Trace($"HID Activity (Unknown Map): (Len: {bytesRead}, Data: {BitConverter.ToString(buffer, 0, bytesRead)})");
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
                
                try
                {
                    Logger.Trace($"Attempting HID connect: {device.GetProductName() ?? "PowerMic"} at {device.DevicePath}");
                    
                    var stream = device.Open();
                    stream.ReadTimeout = 100; // Non-blocking with short timeout
                    
                    _device = device;
                    _stream = stream;
                    
                    var name = device.GetProductName() ?? "PowerMic";
                    Logger.Trace($"HID Connection Successful: {name}");
                    DeviceConnected?.Invoke($"Connected to {name}");
                    
                    return true;
                }
                catch (Exception ex)
                {
                    Logger.Trace($"HID Connection Failed: {ex.Message}");
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
    }
    
    public void Dispose()
    {
        Stop();
    }
}
