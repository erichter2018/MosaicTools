Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinFind {
    [DllImport("user32.dll")] public static extern IntPtr FindWindow(string className, string windowName);
}
"@

$hwnd = [WinFind]::FindWindow("SunAwtDialog", "Header Info")
Write-Host "HWND: 0x$($hwnd.ToString('X'))"

if ($hwnd -eq [IntPtr]::Zero) {
    Write-Host "Not found!"
    exit 1
}

# Get UIA element from HWND
$element = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
Write-Host "UIA Element: Name='$($element.Current.Name)' Class='$($element.Current.ClassName)' Framework='$($element.Current.FrameworkId)'"
Write-Host ""

$allCondition = [System.Windows.Automation.Condition]::TrueCondition

function Dump-Element {
    param($el, [int]$depth = 0, [int]$maxDepth = 10)
    if ($depth -gt $maxDepth) { return }

    $indent = "  " * $depth
    try {
        $name = $el.Current.Name
        $ctrlType = $el.Current.ControlType.ProgrammaticName
        $className = $el.Current.ClassName
        $autoId = $el.Current.AutomationId
        $rect = $el.Current.BoundingRectangle

        $line = "${indent}[$ctrlType] Class='$className'"
        if ($autoId) { $line += " AutoId='$autoId'" }
        $line += " Rect=($([int]$rect.X),$([int]$rect.Y),$([int]$rect.Width)x$([int]$rect.Height))"
        if ($name) {
            $tn = if ($name.Length -gt 120) { $name.Substring(0, 120) + "..." } else { $name }
            $line += " Name='$tn'"
        }

        # Value
        try {
            $vp = $el.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
            $val = $vp.Current.Value
            if ($val) {
                $tv = if ($val.Length -gt 200) { $val.Substring(0, 200) + "..." } else { $val }
                $line += " Value='$tv'"
            }
        } catch {}

        # LegacyIAccessible
        try {
            $lip = $el.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern)
            $lv = $lip.Current.Value
            $ldesc = $lip.Current.Description
            $lrole = $lip.Current.Role
            if ($lv) { $line += " LegVal='$($lv.Substring(0,[Math]::Min(200,$lv.Length)))'" }
            if ($ldesc) { $line += " LegDesc='$ldesc'" }
            $line += " Role=$lrole"
        } catch {}

        Write-Host $line
    } catch {
        Write-Host "${indent}[ERROR reading element: $_]"
    }

    try {
        $children = $el.FindAll([System.Windows.Automation.TreeScope]::Children, $allCondition)
        Write-Host "${indent}  ($($children.Count) children)"
        foreach ($child in $children) {
            Dump-Element -el $child -depth ($depth+1) -maxDepth $maxDepth
        }
    } catch {
        Write-Host "${indent}  (error enumerating children: $_)"
    }
}

Dump-Element -el $element -depth 0 -maxDepth 10
