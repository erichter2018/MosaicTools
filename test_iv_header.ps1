# Try to read the Header Info dialog contents
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

$automation = [System.Windows.Automation.AutomationElement]
$desktop = $automation::RootElement

# Find the Header Info dialog
$nameCondition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::NameProperty, "Header Info"
)
$dialog = $desktop.FindFirst([System.Windows.Automation.TreeScope]::Children, $nameCondition)

if ($dialog) {
    Write-Host "Found Header Info dialog!"
    Write-Host "  ClassName: $($dialog.Current.ClassName)"
    Write-Host "  FrameworkId: $($dialog.Current.FrameworkId)"
    Write-Host "  BoundingRect: $($dialog.Current.BoundingRectangle)"
    Write-Host ""

    # Try to enumerate children
    $allCondition = [System.Windows.Automation.Condition]::TrueCondition
    $children = $dialog.FindAll([System.Windows.Automation.TreeScope]::Descendants, $allCondition)
    Write-Host "Total descendant elements: $($children.Count)"
    Write-Host ""

    foreach ($child in $children) {
        try {
            $name = $child.Current.Name
            $ctrlType = $child.Current.ControlType.ProgrammaticName
            $autoId = $child.Current.AutomationId
            $className = $child.Current.ClassName

            $line = "  [$ctrlType] Class='$className'"
            if ($autoId) { $line += " AutoId='$autoId'" }
            if ($name) {
                $truncName = if ($name.Length -gt 200) { $name.Substring(0, 200) + "..." } else { $name }
                $line += " Name='$truncName'"
            }

            # Try to get Value pattern
            try {
                $valuePattern = $child.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
                if ($valuePattern) {
                    $val = $valuePattern.Current.Value
                    if ($val) {
                        $truncVal = if ($val.Length -gt 200) { $val.Substring(0, 200) + "..." } else { $val }
                        $line += " Value='$truncVal'"
                    }
                }
            } catch { }

            # Try to get Text pattern
            try {
                $textPattern = $child.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern)
                if ($textPattern) {
                    $text = $textPattern.DocumentRange.GetText(500)
                    if ($text) { $line += " Text='$text'" }
                }
            } catch { }

            # Try LegacyIAccessible
            try {
                $legacy = $child.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern)
                if ($legacy) {
                    $legVal = $legacy.Current.Value
                    $legName = $legacy.Current.Name
                    $legDesc = $legacy.Current.Description
                    if ($legVal) { $line += " LegacyValue='$legVal'" }
                    if ($legDesc) { $line += " LegacyDesc='$legDesc'" }
                }
            } catch { }

            Write-Host $line
        } catch {
            Write-Host "  [Error reading element]"
        }
    }
} else {
    Write-Host "Header Info dialog not found via UIA"
}

# Also try clipboard approach - activate dialog, Ctrl+A, Ctrl+C
Write-Host "`n=== Attempting clipboard capture ==="
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinApi {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr FindWindow(string className, string windowName);
}
"@

$hwnd = [WinApi]::FindWindow("SunAwtDialog", "Header Info")
if ($hwnd -ne [IntPtr]::Zero) {
    Write-Host "Activating Header Info dialog..."
    [WinApi]::SetForegroundWindow($hwnd)
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.SendKeys]::SendWait("^(a)")
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.SendKeys]::SendWait("^(c)")
    Start-Sleep -Milliseconds 500

    $clip = [System.Windows.Forms.Clipboard]::GetText()
    if ($clip) {
        Write-Host "Clipboard content ($($clip.Length) chars):"
        Write-Host $clip
    } else {
        Write-Host "Clipboard empty after Ctrl+A, Ctrl+C"
    }
}
