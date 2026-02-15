Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinApi2 {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr FindWindow(string className, string windowName);
}
"@

# Clear clipboard first
[System.Windows.Forms.Clipboard]::Clear()
Start-Sleep -Milliseconds 200

$hwnd = [WinApi2]::FindWindow("SunAwtDialog", "Header Info")
if ($hwnd -ne [IntPtr]::Zero) {
    Write-Host "Found Header Info dialog, activating..."
    [WinApi2]::SetForegroundWindow($hwnd)
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.SendKeys]::SendWait("^(a)")
    Start-Sleep -Milliseconds 300
    [System.Windows.Forms.SendKeys]::SendWait("^(c)")
    Start-Sleep -Milliseconds 500

    $clip = [System.Windows.Forms.Clipboard]::GetText()
    if ($clip) {
        Write-Host "Got $($clip.Length) chars"
        Write-Host ""
        # Show key fields only
        foreach ($line in $clip -split "`n") {
            if ($line -match "\(0020,0011\)|\(0020,0013\)|\(0008,0050\)|\(0010,0010\)|\(0008,0060\)|\(0008,103e\)|\(0008,1030\)|\(0020,0010\)") {
                Write-Host $line.Trim()
            }
        }
    } else {
        Write-Host "Clipboard empty!"
    }
} else {
    Write-Host "Header Info dialog not found!"
}
