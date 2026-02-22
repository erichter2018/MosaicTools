using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

internal static class Program
{
    private const uint PROCESS_SUSPEND_RESUME = 0x0800;
    private const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(uint desiredAccess, [MarshalAs(UnmanagedType.Bool)] bool inheritHandle, int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr hObject);

    [DllImport("ntdll.dll")]
    private static extern int NtSuspendProcess(IntPtr processHandle);

    [DllImport("ntdll.dll")]
    private static extern int NtResumeProcess(IntPtr processHandle);

    private static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            PrintUsage();
            return 1;
        }

        string action = args[0].Trim().ToLowerInvariant();

        if (action == "trap")
        {
            if (args.Length < 2)
            {
                PrintUsage();
                return 1;
            }

            string target = args[1].Trim();
            string folder = args.Length >= 3 ? args[2] : @"C:\MModal\FluencyForImaging\Reporting\XML\IN";
            string pattern = args.Length >= 4 ? args[3] : "settoken*.xml";
            return TrapAndSuspend(target, folder, pattern);
        }

        if (args.Length < 2)
        {
            PrintUsage();
            return 1;
        }

        string target2 = args[1].Trim();

        if (action is not ("suspend" or "resume"))
        {
            Console.WriteLine("First argument must be 'suspend', 'resume', or 'trap'.");
            PrintUsage();
            return 1;
        }

        return ApplyToTargets(action, target2);
    }

    private static int TrapAndSuspend(string target, string folder, string pattern)
    {
        if (!Directory.Exists(folder))
        {
            Console.WriteLine($"Folder not found: {folder}");
            return 2;
        }

        Console.WriteLine($"Watching: {folder}");
        Console.WriteLine($"Pattern:  {pattern}");
        Console.WriteLine($"Target:   {target}");
        Console.WriteLine("Speak the macro phrase now. This tool will suspend the target when a matching file appears.");
        Console.WriteLine("Press Ctrl+C to cancel.");

        using var hit = new AutoResetEvent(false);
        string? matchedPath = null;
        string? matchedEvent = null;

        using var fsw = new FileSystemWatcher(folder, pattern)
        {
            IncludeSubdirectories = false,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime | NotifyFilters.LastWrite,
            InternalBufferSize = 64 * 1024,
            EnableRaisingEvents = true
        };

        void OnHit(string kind, string path)
        {
            if (Interlocked.CompareExchange(ref matchedPath, path, null) == null)
            {
                matchedEvent = kind;
                hit.Set();
            }
        }

        fsw.Created += (_, e) => OnHit("Created", e.FullPath);
        fsw.Changed += (_, e) => OnHit("Changed", e.FullPath);
        fsw.Renamed += (_, e) => OnHit("Renamed", e.FullPath);

        // Poll fallback in case the watcher misses a very fast create+delete cycle.
        using var cts = new CancellationTokenSource();
        var pollTask = Task.Run(async () =>
        {
            while (!cts.Token.IsCancellationRequested)
            {
                try
                {
                    var file = Directory.EnumerateFiles(folder, pattern).FirstOrDefault();
                    if (!string.IsNullOrEmpty(file))
                    {
                        OnHit("Poll", file);
                        return;
                    }
                }
                catch { }

                try { await Task.Delay(5, cts.Token); } catch { }
            }
        });

        hit.WaitOne();
        cts.Cancel();
        try { pollTask.Wait(200); } catch { }

        Console.WriteLine($"Matched ({matchedEvent}): {matchedPath}");
        var rc = ApplyToTargets("suspend", target);

        if (rc == 0)
        {
            Console.WriteLine();
            Console.WriteLine("If suspend succeeded, immediately inspect:");
            Console.WriteLine($"  {folder}");
            try
            {
                foreach (var fi in new DirectoryInfo(folder).GetFiles("*.xml").OrderByDescending(f => f.LastWriteTime).Take(20))
                    Console.WriteLine($"  {fi.LastWriteTime:HH:mm:ss.fff}  {fi.Name}  ({fi.Length} bytes)");
            }
            catch { }
        }

        return rc;
    }

    private static int ApplyToTargets(string action, string target)
    {
        var targets = ResolveTargets(target);
        if (targets.Count == 0)
        {
            Console.WriteLine($"No process found matching '{target}'.");
            return 2;
        }

        Console.WriteLine($"{action.ToUpperInvariant()} {targets.Count} process(es):");
        foreach (var p in targets.OrderBy(p => p.Id))
        {
            Console.WriteLine($"  PID {p.Id,-6} {p.ProcessName}");
            try
            {
                Apply(action, p.Id);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"    ERROR: {ex.Message}");
            }
            finally
            {
                p.Dispose();
            }
        }

        Console.WriteLine("Done.");
        return 0;
    }

    private static void Apply(string action, int pid)
    {
        IntPtr handle = OpenProcess(PROCESS_SUSPEND_RESUME | PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
        if (handle == IntPtr.Zero)
            throw new Win32Exception(Marshal.GetLastWin32Error(), $"OpenProcess failed for PID {pid}");

        try
        {
            int status = action == "suspend" ? NtSuspendProcess(handle) : NtResumeProcess(handle);
            if (status != 0)
                throw new InvalidOperationException($"Nt{(action == "suspend" ? "Suspend" : "Resume")}Process returned NTSTATUS 0x{status:X8}");
        }
        finally
        {
            CloseHandle(handle);
        }
    }

    private static List<Process> ResolveTargets(string target)
    {
        var result = new List<Process>();
        var seen = new HashSet<int>();

        if (int.TryParse(target, out int pid))
        {
            try
            {
                var p = Process.GetProcessById(pid);
                result.Add(p);
                seen.Add(p.Id);
            }
            catch { }
            return result;
        }

        string normalized = target.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
            ? Path.GetFileNameWithoutExtension(target)
            : target;

        foreach (var p in Process.GetProcessesByName(normalized))
        {
            if (seen.Add(p.Id)) result.Add(p);
            else p.Dispose();
        }

        return result;
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  ProcFreeze suspend <processName|pid>");
        Console.WriteLine("  ProcFreeze resume  <processName|pid>");
        Console.WriteLine("  ProcFreeze trap    <processName|pid> [folder] [pattern]");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  ProcFreeze suspend MosaicInfoHub");
        Console.WriteLine("  ProcFreeze resume MosaicInfoHub");
        Console.WriteLine("  ProcFreeze trap MosaicInfoHub");
        Console.WriteLine("  ProcFreeze trap MosaicInfoHub C:\\MModal\\FluencyForImaging\\Reporting\\XML\\IN settoken*.xml");
    }
}
