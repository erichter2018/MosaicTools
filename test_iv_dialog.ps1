# Find all InteleViewer windows including the new dialog
Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public class WinEnum2 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);

    public static List<string> results = new List<string>();

    public static void EnumAll(uint targetPid) {
        EnumWindows((hWnd, lParam) => {
            uint pid;
            GetWindowThreadProcessId(hWnd, out pid);
            if (pid == targetPid) {
                AddWindow(hWnd, 0);
                EnumChildWindows(hWnd, (child, lp) => {
                    AddWindow(child, 1);
                    EnumChildWindows(child, (grandchild, lp2) => {
                        AddWindow(grandchild, 2);
                        return true;
                    }, IntPtr.Zero);
                    return true;
                }, IntPtr.Zero);
            }
            return true;
        }, IntPtr.Zero);
    }

    static void AddWindow(IntPtr hWnd, int depth) {
        var sb = new StringBuilder(512);
        GetWindowText(hWnd, sb, 512);
        var title = sb.ToString();
        sb.Clear();
        GetClassName(hWnd, sb, 512);
        var className = sb.ToString();
        bool visible = IsWindowVisible(hWnd);

        var indent = new string(' ', depth * 2);
        results.Add(string.Format("{0}[{1}] Class='{2}' Vis={3} Title='{4}'",
            indent, hWnd.ToInt64(), className, visible,
            title.Length > 120 ? title.Substring(0, 120) + "..." : title));
    }
}
"@

$ivProc = Get-Process -Name "InteleViewer" | Select-Object -First 1
[WinEnum2]::EnumAll([uint32]$ivProc.Id)
foreach ($r in [WinEnum2]::results) { Write-Host $r }
