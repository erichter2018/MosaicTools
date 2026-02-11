using System;
using System.Text;
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
        // Register CodePages encoding provider (required for RichTextBox RTF in self-contained deployment)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Install global exception handlers FIRST, before anything else
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += OnThreadException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;

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
            // Log startup with version info
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            Logger.Trace("========================================");
            Logger.Trace($"=== NEW SESSION - v{version} ===");
            Logger.Trace("========================================");
            Logger.Trace($"App Init Started (Headless: {App.IsHeadless})");
            
            // Enable dark mode for scrollbars, menus, etc (must be before visual styles)
            NativeWindows.EnableDarkMode();

            // Enable visual styles (SetHighDpiMode must be called first per MS docs)
            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
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

            // Rename self to MosaicTools.exe -- rollback if this fails
            try
            {
                File.Move(exePath, targetPath);
            }
            catch
            {
                // Rollback: restore original MosaicTools.exe so the app isn't left broken
                var oldPath = Path.Combine(exeDir, "MosaicTools_old.exe");
                if (File.Exists(oldPath) && !File.Exists(targetPath))
                {
                    try { File.Move(oldPath, targetPath); } catch { }
                }
                return false;
            }

            // Restart from the correct path - preserve command line arguments (especially -headless)
            var args = Environment.GetCommandLineArgs();
            var argsToPass = args.Length > 1 ? string.Join(" ", args.Skip(1).Select(a => a.Contains(' ') ? $"\"{ a}\"" : a)) : "";

            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = targetPath,
                Arguments = argsToPass,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(startInfo);
            return true;
        }
        catch (Exception ex)
        {
            Logger.Trace($"NormalizeExecutableName failed: {ex.Message}");
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

    /// <summary>
    /// Handle exceptions on the UI thread.
    /// </summary>
    private static void OnThreadException(object sender, ThreadExceptionEventArgs e)
    {
        LogCrash("UI THREAD EXCEPTION", e.Exception);
    }

    /// <summary>
    /// Handle exceptions on background threads.
    /// </summary>
    private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        var ex = e.ExceptionObject as Exception;
        LogCrash($"UNHANDLED EXCEPTION (Terminating: {e.IsTerminating})", ex);

        if (e.IsTerminating)
        {
            try
            {
                MessageBox.Show(
                    $"MosaicTools encountered an unexpected error and needs to close.\n\n{ex?.Message}\n\nDetails saved to mosaic_crash_log.txt",
                    "MosaicTools Crash", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }
    }

    /// <summary>
    /// Log crash details to both the trace log and a dedicated crash log.
    /// </summary>
    private static void LogCrash(string context, Exception? ex)
    {
        var message = $"{context}: {ex?.GetType().Name}: {ex?.Message}";
        var fullDetails = $"{context}\n{ex}";

        // Log to trace file
        try
        {
            Logger.Trace("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
            Logger.Trace(message);
            Logger.Trace(ex?.StackTrace ?? "(no stack trace)");
            Logger.Trace("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
        }
        catch { }

        // Also write to dedicated crash log for easy finding
        try
        {
            var crashPath = Path.Combine(AppContext.BaseDirectory, "mosaic_crash_log.txt");
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            File.AppendAllText(crashPath, $"\n\n=== CRASH AT {timestamp} ===\n{fullDetails}\n");
        }
        catch { }
    }
}
