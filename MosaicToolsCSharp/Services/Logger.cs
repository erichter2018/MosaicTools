using System;
using System.IO;

namespace MosaicTools.Services;

/// <summary>
/// Simple trace logger matching Python's log_trace() function.
/// Caps file size at 1MB using FIFO (removes old entries first).
/// </summary>
public static class Logger
{
    private static readonly string LogPath = Path.Combine(AppContext.BaseDirectory, "mosaic_setup_trace.txt");
    private static readonly object _lock = new();
    private const long MaxFileSize = 1024 * 1024; // 1MB

    public static void Trace(string message)
    {
        try
        {
            lock (_lock)
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var line = $"{timestamp}: {message}\n";

                // Check if file needs trimming
                if (File.Exists(LogPath))
                {
                    var fileInfo = new FileInfo(LogPath);
                    if (fileInfo.Length > MaxFileSize)
                    {
                        TrimLogFile();
                    }
                }

                File.AppendAllText(LogPath, line);
            }
        }
        catch
        {
            // Silently ignore logging failures
        }
    }

    private static void TrimLogFile()
    {
        try
        {
            // Read all lines, keep the last 50% to leave room for new entries
            var lines = File.ReadAllLines(LogPath);
            var keepCount = lines.Length / 2;
            var linesToKeep = lines[^keepCount..]; // Keep last half
            File.WriteAllLines(LogPath, linesToKeep);
        }
        catch
        {
            // If trimming fails, just delete and start fresh
            try { File.Delete(LogPath); } catch { }
        }
    }
}
