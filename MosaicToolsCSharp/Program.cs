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
}
