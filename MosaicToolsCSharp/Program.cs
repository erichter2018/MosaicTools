using System;
using System.Threading;
using System.Windows.Forms;
using MosaicTools.UI;
using MosaicTools.Services;

namespace MosaicTools;

/// <summary>
/// Entry point for MosaicTools application.
/// </summary>
static class Program
{
    private static Mutex? _mutex;
    
    [STAThread]
    static void Main(string[] args)
    {
        // Parse command line arguments
        foreach (var arg in args)
        {
            if (arg.Equals("-headless", StringComparison.OrdinalIgnoreCase) ||
                arg.Equals("--headless", StringComparison.OrdinalIgnoreCase))
            {
                App.IsHeadless = true;
            }
        }
        
        // Single instance enforcement
        const string mutexName = "MosaicTools_SingleInstance_Mutex";
        _mutex = new Mutex(true, mutexName, out bool createdNew);
        
        if (!createdNew)
        {
            // Another instance is running - could send message to activate it
            MessageBox.Show("MosaicTools is already running.", "MosaicTools",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Normalize executable name (ensures we're always MosaicTools.exe)
        if (NormalizeExecutableName())
        {
            _mutex?.ReleaseMutex();
            _mutex?.Dispose();
            return; // Exit - new instance launching as MosaicTools.exe
        }

        try
        {
            // Log startup
            Logger.Trace("--- NEW SESSION ---");
            Logger.Trace("App Init Started");
            
            // Enable visual styles
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            
            // Initialize configuration
            var config = Configuration.Load();
            Logger.Trace($"Configuration loaded for: {config.DoctorName}");
            
            // Create and run main form
            using var mainForm = new MainForm(config);
            Application.Run(mainForm);
        }
        catch (Exception ex)
        {
            Logger.Trace($"TOP LEVEL CRASH: {ex}");
            
            // Write crash log
            try
            {
                var crashPath = Path.Combine(AppContext.BaseDirectory, "mosaic_crash_log.txt");
                File.WriteAllText(crashPath, $"CRASH AT STARTUP: {ex}\n\n{ex.StackTrace}");
            }
            catch { }
            
            MessageBox.Show($"Application failed to start:\n\n{ex.Message}\n\nCheck mosaic_crash_log.txt for details.",
                "MosaicTools Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _mutex?.ReleaseMutex();
            _mutex?.Dispose();
        }
    }

    /// <summary>
    /// Ensures the executable is named MosaicTools.exe.
    /// If renamed (e.g., "MosaicTools (1).exe"), renames self and restarts.
    /// Also cleans up duplicate executables.
    /// </summary>
    /// <returns>True if app should exit (restart in progress)</returns>
    private static bool NormalizeExecutableName()
    {
        var exePath = Application.ExecutablePath;
        var exeDir = Path.GetDirectoryName(exePath) ?? ".";
        var currentName = Path.GetFileName(exePath);
        var targetPath = Path.Combine(exeDir, "MosaicTools.exe");

        // Already correct name?
        if (currentName.Equals("MosaicTools.exe", StringComparison.OrdinalIgnoreCase))
        {
            CleanupDuplicateExecutables(exeDir);
            return false;
        }

        try
        {
            // If MosaicTools.exe already exists, rename it to _old
            if (File.Exists(targetPath))
            {
                var oldPath = Path.Combine(exeDir, "MosaicTools_old.exe");
                if (File.Exists(oldPath))
                    File.Delete(oldPath);
                File.Move(targetPath, oldPath);
            }

            // Rename self to MosaicTools.exe
            File.Move(exePath, targetPath);

            // Restart from the correct path
            System.Diagnostics.Process.Start(targetPath);
            return true;
        }
        catch
        {
            // If rename fails, continue running with current name
            return false;
        }
    }

    /// <summary>
    /// Delete duplicate MosaicTools executables (e.g., "MosaicTools (1).exe").
    /// </summary>
    private static void CleanupDuplicateExecutables(string exeDir)
    {
        try
        {
            foreach (var file in Directory.GetFiles(exeDir, "MosaicTools*.exe"))
            {
                var name = Path.GetFileName(file);

                // Keep the main exe
                if (name.Equals("MosaicTools.exe", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Keep _new.exe (might be mid-update)
                if (name.Equals("MosaicTools_new.exe", StringComparison.OrdinalIgnoreCase))
                    continue;

                try { File.Delete(file); }
                catch { /* ignore locked files */ }
            }
        }
        catch { }
    }
}
