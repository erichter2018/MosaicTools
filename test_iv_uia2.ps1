Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$automation = [System.Windows.Automation.AutomationElement]
$desktop = $automation::RootElement

# Find Header Info by class since name search failed before
$allCondition = [System.Windows.Automation.Condition]::TrueCondition
$windows = $desktop.FindAll([System.Windows.Automation.TreeScope]::Children, $allCondition)

$dialog = $null
foreach ($w in $windows) {
    try {
        if ($w.Current.Name -eq "Header Info") {
            $dialog = $w
            break
        }
    } catch {}
}

if (-not $dialog) {
    # Try by class
    foreach ($w in $windows) {
        try {
            if ($w.Current.ClassName -eq "SunAwtDialog" -and $w.Current.Name -match "Header") {
                $dialog = $w
                break
            }
        } catch {}
    }
}

if (-not $dialog) {
    Write-Host "Header Info not found!"
    exit 1
}

Write-Host "Found: Name='$($dialog.Current.Name)' Class='$($dialog.Current.ClassName)' Framework='$($dialog.Current.FrameworkId)'"
Write-Host ""

# Deep recursive dump
function Dump-Element {
    param($element, [int]$depth = 0, [int]$maxDepth = 8)
    if ($depth -gt $maxDepth) { return }

    $indent = "  " * $depth
    $name = $element.Current.Name
    $ctrlType = $element.Current.ControlType.ProgrammaticName
    $className = $element.Current.ClassName
    $autoId = $element.Current.AutomationId

    $line = "${indent}[$ctrlType] Class='$className'"
    if ($autoId) { $line += " AutoId='$autoId'" }
    if ($name) {
        $truncName = if ($name.Length -gt 150) { $name.Substring(0, 150) + "..." } else { $name }
        $line += " Name='$truncName'"
    }

    # Try Value pattern
    try {
        $vp = $element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
        if ($vp) {
            $val = $vp.Current.Value
            if ($val) {
                $tv = if ($val.Length -gt 150) { $val.Substring(0, 150) + "..." } else { $val }
                $line += " Value='$tv'"
            }
        }
    } catch {}

    # Try LegacyIAccessible
    try {
        $lip = $element.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern)
        if ($lip) {
            $lv = $lip.Current.Value
            $ln = $lip.Current.Name
            if ($lv -and $lv -ne $name) { $line += " LegVal='$($lv.Substring(0, [Math]::Min(150,$lv.Length)))'" }
        }
    } catch {}

    Write-Host $line

    try {
        $children = $element.FindAll([System.Windows.Automation.TreeScope]::Children, $allCondition)
        foreach ($child in $children) {
            try { Dump-Element -element $child -depth ($depth+1) -maxDepth $maxDepth } catch {}
        }
    } catch {}
}

Dump-Element -element $dialog -depth 0 -maxDepth 8
