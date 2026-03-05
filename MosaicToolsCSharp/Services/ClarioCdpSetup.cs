using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace MosaicTools.Services;

/// <summary>
/// Automates Chrome CDP setup for Clario: creates a directory junction so the CDP profile
/// shares the real Chrome profile (cookies, settings, logins), and creates a desktop shortcut
/// that launches Chrome with --remote-debugging-port=9224.
///
/// Chrome 136+ blocks --remote-debugging-port on the default User Data directory,
/// but a junction with a different path pointing to the same directory bypasses this check.
/// </summary>
public static class ClarioCdpSetup
{
    private const int CdpPort = 9224;
    private const string JunctionName = "ClarioCDP";

    private static readonly string ChromePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        "Google", "Chrome", "Application", "chrome.exe");

    private static readonly string ChromePathX86 = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
        "Google", "Chrome", "Application", "chrome.exe");

    private static readonly string ChromeUserData = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Google", "Chrome", "User Data");

    private static readonly string JunctionPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Google", "Chrome", JunctionName);

    /// <summary>Check if the junction is already set up correctly.</summary>
    public static bool IsJunctionConfigured()
    {
        try
        {
            var info = new DirectoryInfo(JunctionPath);
            return info.Exists
                   && (info.Attributes & FileAttributes.ReparsePoint) != 0;
        }
        catch { return false; }
    }

    /// <summary>Find Chrome executable.</summary>
    public static string? FindChrome()
    {
        if (File.Exists(ChromePath)) return ChromePath;
        if (File.Exists(ChromePathX86)) return ChromePathX86;
        return null;
    }

    /// <summary>
    /// Create the directory junction and desktop shortcut.
    /// Returns (success, message).
    /// </summary>
    public static (bool Success, string Message) Setup(string clarioUrl)
    {
        var chrome = FindChrome();
        if (chrome == null)
            return (false, "Chrome not found in Program Files.");

        if (!Directory.Exists(ChromeUserData))
            return (false, "Chrome User Data directory not found. Launch Chrome at least once first.");

        // Create junction: ClarioCDP → User Data
        try
        {
            if (Directory.Exists(JunctionPath))
            {
                var info = new DirectoryInfo(JunctionPath);
                if ((info.Attributes & FileAttributes.ReparsePoint) != 0)
                {
                    // Already a junction — leave it
                }
                else
                {
                    // Real directory — remove and replace with junction
                    Directory.Delete(JunctionPath, true);
                    CreateJunction(JunctionPath, ChromeUserData);
                }
            }
            else
            {
                CreateJunction(JunctionPath, ChromeUserData);
            }
        }
        catch (Exception ex)
        {
            return (false, $"Failed to create directory junction: {ex.Message}");
        }

        // Create desktop shortcut
        try
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var shortcutPath = Path.Combine(desktop, "Clario CDP.lnk");

            var args = $"--remote-debugging-port={CdpPort} --remote-allow-origins=* "
                     + $"--user-data-dir=\"{JunctionPath}\" /new-window {clarioUrl}";

            CreateShortcut(shortcutPath, chrome, args,
                Path.GetDirectoryName(chrome)!,
                FindClarioIcon());

            return (true, $"Setup complete.\n\n"
                        + $"• Junction: {JunctionPath} → User Data\n"
                        + $"• Shortcut: {shortcutPath}\n\n"
                        + "Close all Chrome windows, then launch Clario from the new shortcut.\n"
                        + "Your Chrome settings, cookies, and logins will be preserved.");
        }
        catch (Exception ex)
        {
            return (false, $"Junction created but shortcut failed: {ex.Message}");
        }
    }

    private static string? FindClarioIcon()
    {
        var iconPath = @"C:\RP_Source\Icons\Clario - Login.ico";
        return File.Exists(iconPath) ? iconPath : null;
    }

    private static void CreateJunction(string junctionPath, string targetPath)
    {
        // Use mklink /J via cmd — works without admin rights for same-drive junctions
        var psi = new ProcessStartInfo("cmd.exe",
            $"/c mklink /J \"{junctionPath}\" \"{targetPath}\"")
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        var proc = Process.Start(psi)!;
        proc.WaitForExit(5000);
        if (proc.ExitCode != 0)
        {
            var err = proc.StandardError.ReadToEnd();
            throw new IOException($"mklink /J failed (exit {proc.ExitCode}): {err}");
        }
    }

    private static void CreateShortcut(string shortcutPath, string targetPath, string arguments,
        string workingDir, string? iconPath)
    {
        // Use WScript.Shell COM for .lnk creation
        var shellType = Type.GetTypeFromProgID("WScript.Shell")!;
        dynamic shell = Activator.CreateInstance(shellType)!;
        try
        {
            var shortcut = shell.CreateShortcut(shortcutPath);
            try
            {
                shortcut.TargetPath = targetPath;
                shortcut.Arguments = arguments;
                shortcut.WorkingDirectory = workingDir;
                if (iconPath != null)
                    shortcut.IconLocation = iconPath + ",0";
                shortcut.Save();
            }
            finally
            {
                Marshal.FinalReleaseComObject(shortcut);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(shell);
        }
    }
}
