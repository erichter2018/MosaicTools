using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;

namespace MosaicTools.Services;

/// <summary>
/// Connectivity state for a monitored server.
/// </summary>
public enum ConnectivityState
{
    Unknown,
    Good,      // < 100ms
    Slow,      // 100-500ms
    Degraded,  // > 500ms
    Offline    // Connection failed
}

/// <summary>
/// Status information for a monitored server.
/// </summary>
public class ServerStatus
{
    public string Name { get; set; } = "";
    public ConnectivityState State { get; set; } = ConnectivityState.Unknown;
    public double CurrentLatencyMs { get; set; }
    public double MinLatencyMs { get; set; } = double.MaxValue;
    public double MaxLatencyMs { get; set; }
    public double AvgLatencyMs { get; set; }
    public int SuccessCount { get; set; }
    public int FailCount { get; set; }
    public double PacketLossPercent => SuccessCount + FailCount == 0 ? 0
        : (double)FailCount / (SuccessCount + FailCount) * 100;
    public DateTime LastCheck { get; set; }
    public DateTime? LastSuccess { get; set; }
    public string? ErrorMessage { get; set; }
}

/// <summary>
/// Service for monitoring network connectivity to critical systems.
/// </summary>
public class ConnectivityService : IDisposable
{
    private readonly Configuration _config;
    private System.Threading.Timer? _checkTimer;
    private readonly Dictionary<string, ServerStatus> _statuses = new();
    private readonly Dictionary<string, List<double>> _latencyHistory = new(); // rolling history for avg
    private const int LatencyHistorySize = 10;
    private int _isChecking; // int for Interlocked atomicity
    private bool _disposed;

    /// <summary>
    /// Event fired when any server status changes.
    /// </summary>
    public event Action? StatusChanged;

    /// <summary>
    /// Current status of all monitored servers.
    /// </summary>
    public IReadOnlyDictionary<string, ServerStatus> Statuses => _statuses;

    public ConnectivityService(Configuration config)
    {
        _config = config;
        InitializeStatuses();
    }

    private void InitializeStatuses()
    {
        _statuses.Clear();
        _latencyHistory.Clear();

        foreach (var server in _config.ConnectivityServers)
        {
            _statuses[server.Name] = new ServerStatus { Name = server.Name };
            _latencyHistory[server.Name] = new List<double>();
        }
    }

    /// <summary>
    /// Start the connectivity monitoring service.
    /// </summary>
    public void Start()
    {
        if (!_config.ConnectivityMonitorEnabled)
        {
            Logger.Trace("ConnectivityService: Monitoring disabled, not starting");
            return;
        }

        Logger.Trace($"ConnectivityService: Starting with {_config.ConnectivityCheckIntervalSeconds}s interval");

        // Reinitialize statuses in case servers changed
        InitializeStatuses();

        // Run initial check
        _ = CheckNowAsync();

        // Start timer for periodic checks
        var intervalMs = _config.ConnectivityCheckIntervalSeconds * 1000;
        _checkTimer = new System.Threading.Timer(OnTimerTick, null, intervalMs, intervalMs);
    }

    /// <summary>
    /// Stop the connectivity monitoring service.
    /// </summary>
    public void Stop()
    {
        Logger.Trace("ConnectivityService: Stopping");
        _checkTimer?.Change(Timeout.Infinite, Timeout.Infinite);
        _checkTimer?.Dispose();
        _checkTimer = null;
    }

    /// <summary>
    /// Restart the service with current configuration.
    /// </summary>
    public void Restart()
    {
        Stop();
        Start();
    }

    private void OnTimerTick(object? state)
    {
        _ = CheckNowAsync();
    }

    /// <summary>
    /// Perform an immediate connectivity check on all servers.
    /// </summary>
    public async Task CheckNowAsync()
    {
        if (Interlocked.CompareExchange(ref _isChecking, 1, 0) != 0)
        {
            Logger.Trace("ConnectivityService: Check already in progress, skipping");
            return;
        }

        try
        {
            var enabledServers = _config.ConnectivityServers.Where(s => s.Enabled && !string.IsNullOrWhiteSpace(s.Host)).ToList();

            if (enabledServers.Count == 0)
            {
                Logger.Trace("ConnectivityService: No enabled servers with hosts configured");
                return;
            }

            Logger.Trace($"ConnectivityService: Checking {enabledServers.Count} servers");

            // Check all servers in parallel
            var tasks = enabledServers.Select(async server =>
            {
                var (ok, ms, err) = await TestServerAsync(server);
                UpdateStatus(server.Name, ok, ms, err);
            });

            await Task.WhenAll(tasks);

            // Notify listeners
            StatusChanged?.Invoke();
        }
        catch (Exception ex)
        {
            Logger.Trace($"ConnectivityService: Check error: {ex.Message}");
        }
        finally
        {
            Interlocked.Exchange(ref _isChecking, 0);
        }
    }

    /// <summary>
    /// Test a single server synchronously and return result (for Settings "Test" button).
    /// </summary>
    public async Task<(bool ok, double ms, string? err)> TestSingleServerAsync(ServerConfig server)
    {
        return await TestServerAsync(server);
    }

    private async Task<(bool ok, double ms, string? err)> TestServerAsync(ServerConfig server)
    {
        try
        {
            if (server.Port > 0)
            {
                return await TestTcpAsync(server.Host, server.Port, _config.ConnectivityTimeoutMs);
            }
            else
            {
                return await TestPingAsync(server.Host, _config.ConnectivityTimeoutMs);
            }
        }
        catch (Exception ex)
        {
            return (false, 0, ex.Message);
        }
    }

    private async Task<(bool ok, double ms, string? err)> TestTcpAsync(string host, int port, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            using var client = new TcpClient();
            using var cts = new CancellationTokenSource(timeoutMs);
            await client.ConnectAsync(host, port, cts.Token);
            sw.Stop();
            return (true, sw.ElapsedMilliseconds, null);
        }
        catch (OperationCanceledException)
        {
            return (false, timeoutMs, "Connection timeout");
        }
        catch (SocketException ex)
        {
            return (false, 0, ex.SocketErrorCode.ToString());
        }
        catch (Exception ex)
        {
            return (false, 0, ex.Message);
        }
    }

    private async Task<(bool ok, double ms, string? err)> TestPingAsync(string host, int timeoutMs)
    {
        try
        {
            using var ping = new Ping();
            var reply = await ping.SendPingAsync(host, timeoutMs);
            if (reply.Status == IPStatus.Success)
            {
                return (true, reply.RoundtripTime, null);
            }
            return (false, 0, reply.Status.ToString());
        }
        catch (PingException ex)
        {
            return (false, 0, ex.InnerException?.Message ?? ex.Message);
        }
        catch (Exception ex)
        {
            return (false, 0, ex.Message);
        }
    }

    private void UpdateStatus(string serverName, bool success, double latencyMs, string? errorMessage)
    {
        if (!_statuses.ContainsKey(serverName))
        {
            _statuses[serverName] = new ServerStatus { Name = serverName };
            _latencyHistory[serverName] = new List<double>();
        }

        var status = _statuses[serverName];
        status.LastCheck = DateTime.Now;
        status.ErrorMessage = errorMessage;

        if (success)
        {
            status.SuccessCount++;
            status.CurrentLatencyMs = latencyMs;
            status.LastSuccess = DateTime.Now;

            // Update min/max
            if (latencyMs < status.MinLatencyMs)
                status.MinLatencyMs = latencyMs;
            if (latencyMs > status.MaxLatencyMs)
                status.MaxLatencyMs = latencyMs;

            // Update rolling average
            var history = _latencyHistory[serverName];
            history.Add(latencyMs);
            if (history.Count > LatencyHistorySize)
                history.RemoveAt(0);
            status.AvgLatencyMs = history.Average();

            // Determine state based on latency and packet loss
            status.State = DetermineState(latencyMs, status.PacketLossPercent);
        }
        else
        {
            status.FailCount++;
            status.CurrentLatencyMs = 0;
            status.State = ConnectivityState.Offline;
        }
    }

    private ConnectivityState DetermineState(double latencyMs, double packetLoss)
    {
        // If significant packet loss, degrade state
        if (packetLoss > 10)
            return ConnectivityState.Degraded;

        // State based on latency
        if (latencyMs < 100)
            return ConnectivityState.Good;
        if (latencyMs < 500)
            return ConnectivityState.Slow;
        return ConnectivityState.Degraded;
    }

    /// <summary>
    /// Get the status for a specific server by name.
    /// </summary>
    public ServerStatus? GetStatus(string serverName)
    {
        return _statuses.TryGetValue(serverName, out var status) ? status : null;
    }

    /// <summary>
    /// Get all enabled server statuses.
    /// </summary>
    public IEnumerable<ServerStatus> GetEnabledStatuses()
    {
        var enabledNames = _config.ConnectivityServers
            .Where(s => s.Enabled && !string.IsNullOrWhiteSpace(s.Host))
            .Select(s => s.Name)
            .ToHashSet();

        return _statuses.Values.Where(s => enabledNames.Contains(s.Name));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Stop();
    }
}
