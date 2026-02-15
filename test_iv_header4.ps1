Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinApi4 {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr FindWindow(string className, string windowName);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    public const byte VK_CONTROL = 0x11;
    public const byte VK_A = 0x41;
    public const byte VK_C = 0x43;
    public const uint KEYEVENTF_KEYUP = 0x02;
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;
}
"@

[System.Windows.Forms.Clipboard]::Clear()

$hwnd = [WinApi4]::FindWindow("SunAwtDialog", "Header Info")
if ($hwnd -eq [IntPtr]::Zero) {
    Write-Host "Header Info dialog not found!"
    exit 1
}

# Get dialog rect
$rect = New-Object WinApi4+RECT
[WinApi4]::GetWindowRect($hwnd, [ref]$rect)
Write-Host "Dialog rect: L=$($rect.Left) T=$($rect.Top) R=$($rect.Right) B=$($rect.Bottom)"
$centerX = [int](($rect.Left + $rect.Right) / 2)
$centerY = [int](($rect.Top + $rect.Bottom) / 2)
Write-Host "Center: $centerX, $centerY"

# Activate and click in center of dialog
[WinApi4]::SetForegroundWindow($hwnd)
Start-Sleep -Milliseconds 300
[WinApi4]::SetCursorPos($centerX, $centerY)
Start-Sleep -Milliseconds 100
[WinApi4]::mouse_event([WinApi4]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 50
[WinApi4]::mouse_event([WinApi4]::MOUSEEVENTF_LEFTUP, 0, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 300

# Ctrl+A
Write-Host "Ctrl+A..."
[WinApi4]::keybd_event([WinApi4]::VK_CONTROL, 0, 0, [UIntPtr]::Zero)
[WinApi4]::keybd_event([WinApi4]::VK_A, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 50
[WinApi4]::keybd_event([WinApi4]::VK_A, 0, [WinApi4]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
[WinApi4]::keybd_event([WinApi4]::VK_CONTROL, 0, [WinApi4]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 300

# Ctrl+C
Write-Host "Ctrl+C..."
[WinApi4]::keybd_event([WinApi4]::VK_CONTROL, 0, 0, [UIntPtr]::Zero)
[WinApi4]::keybd_event([WinApi4]::VK_C, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 50
[WinApi4]::keybd_event([WinApi4]::VK_C, 0, [WinApi4]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
[WinApi4]::keybd_event([WinApi4]::VK_CONTROL, 0, [WinApi4]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 500

$clip = [System.Windows.Forms.Clipboard]::GetText()
if ($clip) {
    Write-Host "`nGot $($clip.Length) chars. Key fields:"
    foreach ($line in $clip -split "`n") {
        if ($line -match "\(0020,0011\)|\(0020,0013\)|\(0008,0050\)|\(0010,0010\)|\(0008,0060\)|\(0008,103e\)") {
            Write-Host "  $($line.Trim())"
        }
    }
} else {
    Write-Host "`nStill empty. The text area may need different interaction."
}
