using System;
using System.IO;

namespace MosaicTools.Services;

/// <summary>
/// Simple trace logger matching Python's log_trace() function.
/// Caps file size at 1MB using FIFO (removes old entries first).
/// </summary>
public static class Logger
{
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MosaicTools", "mosaic_tools_log.txt");
    public static string LogFilePath => LogPath;
    private static readonly object _lock = new();
    private const long MaxFileSize = 1024 * 1024; // 1MB

    public static void Trace(string message)
    {
        try
        {
            lock (_lock)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);

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
            var fi = new FileInfo(LogPath);
            if (fi.Length < 1024 * 1024) return; // Only trim at 1MB

            // Read bytes and find newline near the midpoint
            var bytes = File.ReadAllBytes(LogPath);
            int mid = bytes.Length / 2;
            // Find next newline after midpoint
            while (mid < bytes.Length && bytes[mid] != (byte)'\n') mid++;
            if (mid < bytes.Length) mid++; // Skip past the newline

            // Write second half
            using var fs = new FileStream(LogPath, FileMode.Create, FileAccess.Write);
            fs.Write(bytes, mid, bytes.Length - mid);
        }
        catch
        {
            // If trimming fails, just delete and start fresh
            try { File.Delete(LogPath); } catch { }
        }
    }
}
