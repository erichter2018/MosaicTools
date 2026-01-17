using System;
using System.IO;

namespace MosaicTools.Services;

/// <summary>
/// Simple trace logger matching Python's log_trace() function.
/// </summary>
public static class Logger
{
    private static readonly string LogPath = Path.Combine(AppContext.BaseDirectory, "mosaic_setup_trace.txt");
    private static readonly object _lock = new();
    
    public static void Trace(string message)
    {
        try
        {
            lock (_lock)
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var line = $"{timestamp}: {message}\n";
                File.AppendAllText(LogPath, line);
            }
        }
        catch
        {
            // Silently ignore logging failures
        }
    }
}
