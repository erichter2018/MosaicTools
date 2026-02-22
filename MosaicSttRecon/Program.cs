using System.Diagnostics;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;
        Console.WriteLine("Mosaic STT Recon (Guided)");
        Console.WriteLine("=========================");
        Console.WriteLine();
        Console.WriteLine("This tool helps compare a normal dictated phrase vs a Mosaic voice macro phrase.");
        Console.WriteLine("It captures local process/network snapshots and writes a report to your Desktop.");
        Console.WriteLine();

        try
        {
            var session = new ReconSession();
            session.Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("ERROR: " + ex.Message);
            Console.WriteLine(ex);
        }

        Console.WriteLine();
        Console.Write("Press Enter to exit...");
        Console.ReadLine();
    }
}

internal sealed class ReconSession
{
    private readonly StringBuilder _report = new();

    public void Run()
    {
        var started = DateTime.Now;
        Log("Mosaic STT Recon");
        Log($"Generated: {started:yyyy-MM-dd HH:mm:ss}");
        Log($"Machine: {Environment.MachineName}");
        Log($"User: {Environment.UserName}");
        Log($"OS: {Environment.OSVersion}");
        Log("");

        var processTree = ProcessTreeSnapshot.Capture();
        var mosaicPids = FindMosaicPids(processTree);
        var mosaicProcs = processTree.Processes
            .Where(p => mosaicPids.Contains(p.Pid))
            .OrderBy(p => p.Pid)
            .ToList();

        Console.WriteLine("Scanning for Mosaic-related processes...");
        if (mosaicProcs.Count == 0)
        {
            Console.WriteLine("No Mosaic-related process was found.");
            Console.WriteLine("Open Mosaic (with a study loaded if possible), then rerun this tool.");
            Log("No Mosaic-related process found.");
            WriteReport();
            return;
        }

        Console.WriteLine($"Found {mosaicProcs.Count} Mosaic-related process(es).");
        Console.WriteLine();
        Console.WriteLine("Included processes (Mosaic window and child processes):");
        foreach (var p in mosaicProcs)
        {
            Console.WriteLine($"  PID {p.Pid,-6} {p.Name,-24} Parent {p.ParentPid,-6}  \"{Trim(p.MainWindowTitle, 70)}\"");
        }
        Console.WriteLine();

        Log("Mosaic-related processes");
        Log("-----------------------");
        foreach (var p in mosaicProcs)
            Log($"PID={p.Pid} Parent={p.ParentPid} Name={p.Name} Title=\"{p.MainWindowTitle}\"");
        Log("");

        WriteModuleHints(mosaicProcs);

        Console.WriteLine("How this test works:");
        Console.WriteLine("1) You press Enter to start a capture window.");
        Console.WriteLine("2) Immediately speak the requested phrase in Mosaic.");
        Console.WriteLine("3) The tool samples for 10 seconds.");
        Console.WriteLine();
        Console.WriteLine("Use one normal phrase first (not a macro), then a phrase that triggers a Mosaic voice macro.");
        Console.WriteLine();

        WaitForEnter("Phase 1: Press Enter, then immediately dictate a normal phrase (non-macro).");
        var normalPhase = CapturePhase("normal_phrase", mosaicPids, TimeSpan.FromSeconds(10));

        WaitForEnter("Phase 2: Press Enter, then immediately say a Mosaic voice macro phrase.");
        var macroPhase = CapturePhase("macro_phrase", mosaicPids, TimeSpan.FromSeconds(10));

        LogPhase(normalPhase);
        LogPhase(macroPhase);
        LogDiff(normalPhase, macroPhase);
        LogInterpretation(normalPhase, macroPhase);

        WriteReport();
    }

    private static HashSet<int> FindMosaicPids(ProcessTreeSnapshot snapshot)
    {
        var seeds = new HashSet<int>();
        foreach (var p in snapshot.Processes)
        {
            bool nameMatch = p.Name.Contains("mosaic", StringComparison.OrdinalIgnoreCase);
            bool titleMatch = p.MainWindowTitle.Contains("mosaic", StringComparison.OrdinalIgnoreCase);
            if (nameMatch || titleMatch)
                seeds.Add(p.Pid);
        }

        if (seeds.Count == 0)
            return seeds;

        var byParent = snapshot.Processes.GroupBy(p => p.ParentPid).ToDictionary(g => g.Key, g => g.Select(x => x.Pid).ToList());
        var queue = new Queue<int>(seeds);
        while (queue.Count > 0)
        {
            int pid = queue.Dequeue();
            if (!byParent.TryGetValue(pid, out var children)) continue;
            foreach (var childPid in children)
            {
                if (seeds.Add(childPid))
                    queue.Enqueue(childPid);
            }
        }

        return seeds;
    }

    private void WriteModuleHints(List<ProcInfo> mosaicProcs)
    {
        Log("Interesting loaded modules (best effort)");
        Log("----------------------------------------");

        string[] hints =
        {
            "speech", "dict", "nuance", "dragon", "vosk", "whisper", "audio", "micro", "webview",
            "grpc", "ws2_32", "winhttp", "websocket"
        };

        foreach (var p in mosaicProcs)
        {
            try
            {
                using var proc = Process.GetProcessById(p.Pid);
                var names = new List<string>();
                foreach (ProcessModule module in proc.Modules)
                {
                    var name = module.ModuleName ?? "";
                    if (hints.Any(h => name.Contains(h, StringComparison.OrdinalIgnoreCase)))
                        names.Add(name);
                }

                names = names.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
                Log($"PID {p.Pid} {p.Name}: {(names.Count == 0 ? "(no hint modules matched)" : string.Join(", ", names))}");
            }
            catch (Exception ex)
            {
                Log($"PID {p.Pid} {p.Name}: module enumeration failed ({ex.Message})");
            }
        }

        Log("");
    }

    private PhaseCapture CapturePhase(string label, HashSet<int> mosaicPids, TimeSpan duration)
    {
        Console.WriteLine($"Capturing '{label}' for {duration.TotalSeconds:0} seconds...");
        var phase = new PhaseCapture(label, DateTime.Now, duration);
        var sw = Stopwatch.StartNew();

        while (sw.Elapsed < duration)
        {
            phase.SampleCount++;

            var procSnapshot = ProcessTreeSnapshot.Capture();
            foreach (var p in procSnapshot.Processes.Where(p => mosaicPids.Contains(p.Pid)))
            {
                phase.Processes[p.Pid] = p;
            }

            foreach (var conn in NetstatSnapshot.GetTcpConnections().Where(c => mosaicPids.Contains(c.Pid)))
                phase.TcpConnections.Add(conn);

            foreach (var pipe in NamedPipeSnapshot.GetPipeNames())
                phase.PipeNames.Add(pipe);

            Thread.Sleep(1000);
        }

        phase.EndTime = DateTime.Now;
        Console.WriteLine($"Completed '{label}'. Captured {phase.SampleCount} samples.");
        Console.WriteLine();
        return phase;
    }

    private void LogPhase(PhaseCapture phase)
    {
        Log($"Phase: {phase.Label}");
        Log(new string('-', 7 + phase.Label.Length));
        Log($"Start: {phase.StartTime:HH:mm:ss.fff}");
        Log($"End:   {phase.EndTime:HH:mm:ss.fff}");
        Log($"Samples: {phase.SampleCount}");
        Log("");

        Log("Observed TCP connections (Mosaic-related PIDs)");
        if (phase.TcpConnections.Count == 0)
        {
            Log("  (none observed)");
        }
        else
        {
            foreach (var c in phase.TcpConnections.OrderBy(c => c.Pid).ThenBy(c => c.Local).ThenBy(c => c.Remote))
                Log($"  PID {c.Pid,-6} {c.State,-12} {c.Local,-24} -> {c.Remote}");
        }
        Log("");

        Log("Named pipes snapshot (system-wide, no PID mapping)");
        var suspiciousPipes = phase.PipeNames
            .Where(IsSuspiciousPipeName)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (suspiciousPipes.Count == 0)
        {
            Log($"  Suspicious names: none (total pipes seen: {phase.PipeNames.Count})");
        }
        else
        {
            Log($"  Suspicious names ({suspiciousPipes.Count}, total pipes seen {phase.PipeNames.Count}):");
            foreach (var p in suspiciousPipes)
                Log($"    {p}");
        }
        Log("");
    }

    private void LogDiff(PhaseCapture normal, PhaseCapture macro)
    {
        Log("Differences (macro_phrase vs normal_phrase)");
        Log("------------------------------------------");

        var extraTcp = macro.TcpConnections.Except(normal.TcpConnections).OrderBy(c => c.Pid).ThenBy(c => c.Local).ToList();
        var missingTcp = normal.TcpConnections.Except(macro.TcpConnections).OrderBy(c => c.Pid).ThenBy(c => c.Local).ToList();

        Log("TCP only seen during macro phrase");
        if (extraTcp.Count == 0) Log("  (none)");
        foreach (var c in extraTcp)
            Log($"  PID {c.Pid,-6} {c.State,-12} {c.Local,-24} -> {c.Remote}");
        Log("");

        Log("TCP only seen during normal phrase");
        if (missingTcp.Count == 0) Log("  (none)");
        foreach (var c in missingTcp)
            Log($"  PID {c.Pid,-6} {c.State,-12} {c.Local,-24} -> {c.Remote}");
        Log("");

        var extraPipes = macro.PipeNames.Except(normal.PipeNames, StringComparer.OrdinalIgnoreCase)
            .Where(IsSuspiciousPipeName).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
        var missingPipes = normal.PipeNames.Except(macro.PipeNames, StringComparer.OrdinalIgnoreCase)
            .Where(IsSuspiciousPipeName).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();

        Log("Suspicious pipe names only seen during macro phrase");
        if (extraPipes.Count == 0) Log("  (none)");
        foreach (var p in extraPipes) Log($"  {p}");
        Log("");

        Log("Suspicious pipe names only seen during normal phrase");
        if (missingPipes.Count == 0) Log("  (none)");
        foreach (var p in missingPipes) Log($"  {p}");
        Log("");
    }

    private void LogInterpretation(PhaseCapture normal, PhaseCapture macro)
    {
        var extraLocalTcp = macro.TcpConnections.Except(normal.TcpConnections)
            .Any(c => c.Local.Contains("127.0.0.1") || c.Remote.Contains("127.0.0.1") || c.Local.Contains("[::1]") || c.Remote.Contains("[::1]"));
        var anyRemoteTcp = macro.TcpConnections.Any(c => !c.Remote.StartsWith("127.0.0.1") && !c.Remote.StartsWith("[::1]") && !c.Remote.StartsWith("0.0.0.0"));
        var suspiciousPipeDelta = macro.PipeNames.Except(normal.PipeNames, StringComparer.OrdinalIgnoreCase).Any(IsSuspiciousPipeName);

        Log("Interpretation Hints");
        Log("--------------------");
        Log("This tool is a coarse recon step. It cannot inspect encrypted payloads or internal function calls.");
        if (extraLocalTcp)
            Log("- Macro phrase showed a localhost connection difference. That suggests local IPC (service/helper) may be involved.");
        if (suspiciousPipeDelta)
            Log("- Macro phrase showed suspicious named pipe differences. That suggests a local Windows IPC channel may exist.");
        if (anyRemoteTcp)
            Log("- Mosaic-related processes maintained remote TCP connections during capture. If behavior still differs, the command may be encoded inside existing traffic (e.g., WebSocket message payload).");
        if (!extraLocalTcp && !suspiciousPipeDelta)
            Log("- No obvious endpoint/pipe difference was captured. The macro decision may be happening inside an existing long-lived connection or in-process parser.");
        Log("");
        Log("Next step recommendation:");
        Log("- If you see localhost/pipe clues: use ProcMon/API Monitor focused on those names.");
        Log("- If you only see stable remote connections: capture traffic (Fiddler/Wireshark) and compare normal vs macro phrase messages.");
        Log("- If no observable transport difference: inspect Mosaic process modules/assets or use API Monitor/Frida hooks.");
        Log("");
    }

    private void WaitForEnter(string prompt)
    {
        Console.WriteLine(prompt);
        Console.ReadLine();
    }

    private void WriteReport()
    {
        string fileName = $"mosaic_stt_recon_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
        string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string desktopPath = Path.Combine(desktop, fileName);

        try
        {
            File.WriteAllText(desktopPath, _report.ToString());
            Console.WriteLine($"Report written to: {desktopPath}");
            return;
        }
        catch (UnauthorizedAccessException)
        {
            // Sandbox/test environments may block Desktop writes. Fall back to repo-local output.
        }
        catch (IOException)
        {
            // Fall back below.
        }

        string fallbackDir = Path.Combine(Environment.CurrentDirectory, "MosaicSttReconReports");
        Directory.CreateDirectory(fallbackDir);
        string fallbackPath = Path.Combine(fallbackDir, fileName);
        File.WriteAllText(fallbackPath, _report.ToString());
        Console.WriteLine($"Desktop write unavailable. Report written to: {fallbackPath}");
    }

    private void Log(string line) => _report.AppendLine(line);

    private static string Trim(string text, int max)
    {
        if (string.IsNullOrEmpty(text)) return text;
        return text.Length <= max ? text : text[..(max - 3)] + "...";
    }

    private static bool IsSuspiciousPipeName(string pipeName)
    {
        string[] hints = { "mosaic", "speech", "dict", "nuance", "dragon", "audio", "webview", "grpc", "stt" };
        return hints.Any(h => pipeName.Contains(h, StringComparison.OrdinalIgnoreCase));
    }
}

internal sealed class PhaseCapture
{
    public string Label { get; }
    public DateTime StartTime { get; }
    public DateTime EndTime { get; set; }
    public TimeSpan Duration { get; }
    public int SampleCount { get; set; }
    public Dictionary<int, ProcInfo> Processes { get; } = new();
    public HashSet<TcpConn> TcpConnections { get; } = new();
    public HashSet<string> PipeNames { get; } = new(StringComparer.OrdinalIgnoreCase);

    public PhaseCapture(string label, DateTime startTime, TimeSpan duration)
    {
        Label = label;
        StartTime = startTime;
        EndTime = startTime;
        Duration = duration;
    }
}

internal readonly record struct ProcInfo(int Pid, int ParentPid, string Name, string MainWindowTitle);

internal sealed class ProcessTreeSnapshot
{
    public List<ProcInfo> Processes { get; } = new();

    public static ProcessTreeSnapshot Capture()
    {
        var parentMap = Toolhelp.GetParentPidMap();
        var snap = new ProcessTreeSnapshot();

        foreach (var proc in Process.GetProcesses())
        {
            try
            {
                int pid = proc.Id;
                int parentPid = parentMap.TryGetValue(pid, out var pp) ? pp : -1;
                string name = Safe(() => proc.ProcessName) ?? "";
                string title = Safe(() => proc.MainWindowTitle) ?? "";
                snap.Processes.Add(new ProcInfo(pid, parentPid, name, title));
            }
            catch
            {
            }
            finally
            {
                try { proc.Dispose(); } catch { }
            }
        }

        return snap;
    }

    private static string? Safe(Func<string> getter)
    {
        try { return getter(); } catch { return null; }
    }
}

internal static class Toolhelp
{
    private const uint TH32CS_SNAPPROCESS = 0x00000002;
    private static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct PROCESSENTRY32
    {
        public uint dwSize;
        public uint cntUsage;
        public uint th32ProcessID;
        public nuint th32DefaultHeapID;
        public uint th32ModuleID;
        public uint cntThreads;
        public uint th32ParentProcessID;
        public int pcPriClassBase;
        public uint dwFlags;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szExeFile;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, uint th32ProcessID);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool Process32First(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool Process32Next(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr hObject);

    public static Dictionary<int, int> GetParentPidMap()
    {
        var map = new Dictionary<int, int>();
        IntPtr snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (snapshot == INVALID_HANDLE_VALUE)
            return map;

        try
        {
            var pe = new PROCESSENTRY32 { dwSize = (uint)Marshal.SizeOf<PROCESSENTRY32>() };
            if (!Process32First(snapshot, ref pe))
                return map;

            do
            {
                map[(int)pe.th32ProcessID] = (int)pe.th32ParentProcessID;
            }
            while (Process32Next(snapshot, ref pe));
        }
        finally
        {
            CloseHandle(snapshot);
        }

        return map;
    }
}

internal readonly record struct TcpConn(int Pid, string Local, string Remote, string State);

internal static class NetstatSnapshot
{
    public static List<TcpConn> GetTcpConnections()
    {
        var list = new List<TcpConn>();
        try
        {
            using var proc = new Process();
            proc.StartInfo = new ProcessStartInfo
            {
                FileName = "netstat.exe",
                Arguments = "-ano -p TCP",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            proc.Start();
            string output = proc.StandardOutput.ReadToEnd();
            _ = proc.StandardError.ReadToEnd();
            proc.WaitForExit(5000);

            using var reader = new StringReader(output);
            while (reader.ReadLine() is { } line)
            {
                line = line.Trim();
                if (!line.StartsWith("TCP", StringComparison.OrdinalIgnoreCase)) continue;
                var parts = SplitColumns(line);
                if (parts.Length < 5) continue;

                string local = NormalizeEndpoint(parts[1]);
                string remote = NormalizeEndpoint(parts[2]);
                string state = parts[3];
                if (!int.TryParse(parts[4], out int pid)) continue;

                list.Add(new TcpConn(pid, local, remote, state));
            }
        }
        catch
        {
        }

        return list;
    }

    private static string[] SplitColumns(string line) => line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

    private static string NormalizeEndpoint(string endpoint)
    {
        endpoint = endpoint.Trim();
        if (endpoint == "*:*") return endpoint;

        if (endpoint.StartsWith("[") && endpoint.Contains("]:"))
            return endpoint;

        int lastColon = endpoint.LastIndexOf(':');
        if (lastColon <= 0 || lastColon == endpoint.Length - 1)
            return endpoint;

        string host = endpoint[..lastColon];
        string port = endpoint[(lastColon + 1)..];
        if (host == "0.0.0.0" || host == "::")
            return $"{host}:{port}";

        if (IPAddress.TryParse(host, out var ip))
            return ip.ToString() + ":" + port;

        return endpoint;
    }
}

internal static class NamedPipeSnapshot
{
    public static IEnumerable<string> GetPipeNames()
    {
        try
        {
            return Directory.EnumerateFileSystemEntries("\\\\.\\pipe\\")
                .Select(path => path.StartsWith("\\\\.\\pipe\\", StringComparison.OrdinalIgnoreCase) ? path[9..] : path)
                .ToArray();
        }
        catch
        {
            return Array.Empty<string>();
        }
    }
}



