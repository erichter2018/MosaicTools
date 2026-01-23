using System;
using System.IO;

namespace ClarioIgnore;

public static class Logger
{
    private static readonly string LogFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ClarioIgnore");

    private static readonly string LogPath = Path.Combine(LogFolder, "clarioignore_log.txt");

    private static readonly object _lock = new();

    public static void Log(string message)
    {
        try
        {
            lock (_lock)
            {
                Directory.CreateDirectory(LogFolder);

                // Keep log file under 1MB by truncating
                if (File.Exists(LogPath))
                {
                    var info = new FileInfo(LogPath);
                    if (info.Length > 1024 * 1024)
                    {
                        File.WriteAllText(LogPath, "--- Log truncated ---\n");
                    }
                }

                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                File.AppendAllText(LogPath, $"[{timestamp}] {message}\n");
            }
        }
        catch
        {
            // Silently fail - logging should never crash the app
        }
    }
}
