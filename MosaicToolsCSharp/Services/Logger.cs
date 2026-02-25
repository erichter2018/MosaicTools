using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;

namespace MosaicTools.Services;

/// <summary>
/// Async trace logger. All writes go to a lock-free queue and are flushed
/// by a dedicated background thread, so Logger.Trace() never blocks callers.
/// This is critical for the keyboard hook callback — any blocking there
/// causes Windows to remove the hook (system-wide input lag/death).
/// </summary>
public static class Logger
{
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MosaicTools", "mosaic_tools_log.txt");
    public static string LogFilePath => LogPath;
    private const long MaxFileSize = 1024 * 1024; // 1MB

    private static readonly ConcurrentQueue<string> _queue = new();
    private static readonly ManualResetEventSlim _signal = new(false);
    private static readonly Thread _writerThread;
    private static volatile bool _stopping;

    static Logger()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
        _writerThread = new Thread(WriterLoop)
        {
            IsBackground = true,
            Name = "LogWriter",
            Priority = ThreadPriority.BelowNormal
        };
        _writerThread.Start();
    }

    /// <summary>
    /// Queue a log message. Returns immediately — never blocks the caller.
    /// </summary>
    public static void Trace(string message)
    {
        try
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            _queue.Enqueue($"{timestamp}: {message}\n");
            _signal.Set(); // Wake writer thread
        }
        catch
        {
            // Silently ignore — logging must never crash the app
        }
    }

    /// <summary>
    /// Flush pending log messages and stop the writer thread.
    /// Call during application shutdown.
    /// </summary>
    public static void Shutdown()
    {
        _stopping = true;
        _signal.Set();
        _writerThread.Join(3000); // Wait up to 3s for final flush
    }

    private static void WriterLoop()
    {
        while (!_stopping)
        {
            _signal.Wait(1000); // Wake on signal or every 1s
            _signal.Reset();
            FlushQueue();
        }
        FlushQueue(); // Final flush on shutdown
    }

    private static void FlushQueue()
    {
        try
        {
            // Drain the queue into a single batch write
            using var sw = new StreamWriter(LogPath, append: true);
            int count = 0;
            while (_queue.TryDequeue(out var line))
            {
                sw.Write(line);
                count++;
                if (count > 1000) break; // Batch cap to prevent unbounded writes
            }

            if (count > 0)
                sw.Flush();
        }
        catch
        {
            // If write fails, drain the queue to prevent unbounded growth
            while (_queue.Count > 10000 && _queue.TryDequeue(out _)) { }
        }

        // Trim check (outside the StreamWriter to avoid file contention)
        try
        {
            if (File.Exists(LogPath))
            {
                var fi = new FileInfo(LogPath);
                if (fi.Length > MaxFileSize)
                    TrimLogFile();
            }
        }
        catch { }
    }

    private static void TrimLogFile()
    {
        try
        {
            var bytes = File.ReadAllBytes(LogPath);
            int mid = bytes.Length / 2;
            while (mid < bytes.Length && bytes[mid] != (byte)'\n') mid++;
            if (mid < bytes.Length) mid++;

            using var fs = new FileStream(LogPath, FileMode.Create, FileAccess.Write);
            fs.Write(bytes, mid, bytes.Length - mid);
        }
        catch
        {
            try { File.Delete(LogPath); } catch { }
        }
    }
}
