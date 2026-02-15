Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinApi3 {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr FindWindow(string className, string windowName);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    public const byte VK_CONTROL = 0x11;
    public const byte VK_A = 0x41;
    public const byte VK_C = 0x43;
    public const uint KEYEVENTF_KEYUP = 0x02;
}
"@

[System.Windows.Forms.Clipboard]::Clear()
Start-Sleep -Milliseconds 200

$hwnd = [WinApi3]::FindWindow("SunAwtDialog", "Header Info")
if ($hwnd -eq [IntPtr]::Zero) {
    Write-Host "Header Info dialog not found!"
    exit 1
}

Write-Host "Found Header Info dialog at 0x$($hwnd.ToString('X'))"
[WinApi3]::SetForegroundWindow($hwnd)
Start-Sleep -Milliseconds 800

# Verify it's foreground
$fg = [WinApi3]::GetForegroundWindow()
$sb = New-Object System.Text.StringBuilder 512
[WinApi3]::GetWindowText($fg, $sb, 512)
Write-Host "Foreground window: $($sb.ToString())"

# Try keybd_event instead of SendKeys
Write-Host "Sending Ctrl+A via keybd_event..."
[WinApi3]::keybd_event([WinApi3]::VK_CONTROL, 0, 0, [UIntPtr]::Zero)
[WinApi3]::keybd_event([WinApi3]::VK_A, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 50
[WinApi3]::keybd_event([WinApi3]::VK_A, 0, [WinApi3]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
[WinApi3]::keybd_event([WinApi3]::VK_CONTROL, 0, [WinApi3]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 300

Write-Host "Sending Ctrl+C via keybd_event..."
[WinApi3]::keybd_event([WinApi3]::VK_CONTROL, 0, 0, [UIntPtr]::Zero)
[WinApi3]::keybd_event([WinApi3]::VK_C, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 50
[WinApi3]::keybd_event([WinApi3]::VK_C, 0, [WinApi3]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
[WinApi3]::keybd_event([WinApi3]::VK_CONTROL, 0, [WinApi3]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 500

$clip = [System.Windows.Forms.Clipboard]::GetText()
if ($clip) {
    Write-Host "`nGot $($clip.Length) chars"
    foreach ($line in $clip -split "`n") {
        if ($line -match "\(0020,0011\)|\(0020,0013\)|\(0008,0050\)|\(0010,0010\)|\(0008,0060\)|\(0008,103e\)") {
            Write-Host $line.Trim()
        }
    }
} else {
    Write-Host "`nClipboard still empty. Checking all clipboard formats..."
    $dataObj = [System.Windows.Forms.Clipboard]::GetDataObject()
    if ($dataObj) {
        $formats = $dataObj.GetFormats()
        Write-Host "Available formats: $($formats -join ', ')"
    } else {
        Write-Host "No data object on clipboard"
    }
}
