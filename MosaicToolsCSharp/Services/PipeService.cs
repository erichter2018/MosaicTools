using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MosaicTools.Services;

// --- Message models ---

/// <summary>
/// Study data broadcast from MosaicTools to RVUCounter on each scrape tick (sent only when changed).
/// </summary>
public record StudyDataMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("accession")] string? Accession,
    [property: JsonPropertyName("description")] string? Description,
    [property: JsonPropertyName("templateName")] string? TemplateName,
    [property: JsonPropertyName("patientName")] string? PatientName,
    [property: JsonPropertyName("patientGender")] string? PatientGender,
    [property: JsonPropertyName("mrn")] string? Mrn,
    [property: JsonPropertyName("siteCode")] string? SiteCode,
    [property: JsonPropertyName("clarioPriority")] string? ClarioPriority,
    [property: JsonPropertyName("clarioClass")] string? ClarioClass,
    [property: JsonPropertyName("drafted")] bool Drafted,
    [property: JsonPropertyName("hasCritical")] bool HasCritical,
    [property: JsonPropertyName("timestamp")] string Timestamp
);

/// <summary>
/// Study lifecycle event from MosaicTools to RVUCounter (sent immediately).
/// </summary>
public record StudyEventMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("eventType")] string EventType,
    [property: JsonPropertyName("accession")] string? Accession,
    [property: JsonPropertyName("hasCritical")] bool HasCritical
);

/// <summary>
/// Shift info from RVUCounter to MosaicTools.
/// </summary>
public record ShiftInfoMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("totalRvu")] double TotalRvu,
    [property: JsonPropertyName("recordCount")] int RecordCount,
    [property: JsonPropertyName("shiftStart")] string? ShiftStart,
    [property: JsonPropertyName("isShiftActive")] bool IsShiftActive,
    [property: JsonPropertyName("currentHourRvu")] double? CurrentHourRvu = null,
    [property: JsonPropertyName("priorHourRvu")] double? PriorHourRvu = null,
    [property: JsonPropertyName("estimatedTotalRvu")] double? EstimatedTotalRvu = null
);

/// <summary>
/// Distraction alert from RVUCounter — study has been open too long.
/// </summary>
public record DistractionAlertMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("studyType")] string? StudyType,
    [property: JsonPropertyName("elapsedSeconds")] double ElapsedSeconds,
    [property: JsonPropertyName("expectedSeconds")] double ExpectedSeconds,
    [property: JsonPropertyName("alertLevel")] int AlertLevel
);

/// <summary>
/// Named pipe IPC server that broadcasts study data to RVUCounter and receives shift RVU totals back.
/// Wire protocol: each message is [4 bytes uint32 LE length][N bytes UTF-8 JSON].
/// </summary>
public class PipeService : IDisposable
{
    private const string PipeName = "MosaicToolsPipe";

    private NamedPipeServerStream? _pipe;
    private CancellationTokenSource _cts = new();
    private readonly object _writeLock = new();
    private readonly object _shiftLock = new();

    private StudyDataMessage? _lastSentStudyData;
    private ShiftInfoMessage? _latestShiftInfo;
    private volatile bool _isConnected;

    public bool IsConnected => _isConnected;

    public ShiftInfoMessage? LatestShiftInfo
    {
        get { lock (_shiftLock) return _latestShiftInfo; }
    }

    /// <summary>
    /// Raised when new shift info is received from the pipe client.
    /// </summary>
    public event Action? ShiftInfoUpdated;

    /// <summary>
    /// Raised when a distraction alert is received from RVUCounter.
    /// </summary>
    public event Action<DistractionAlertMessage>? DistractionAlertReceived;

    public void Start()
    {
        Task.Run(ServerLoopAsync);
    }

    private async Task ServerLoopAsync()
    {
        while (!_cts.IsCancellationRequested)
        {
            try
            {
                _pipe = new NamedPipeServerStream(
                    PipeName,
                    PipeDirection.InOut,
                    1,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous);

                Logger.Trace("PipeService: Waiting for connection...");
                await _pipe.WaitForConnectionAsync(_cts.Token);
                _isConnected = true;
                Logger.Trace("PipeService: Client connected");

                // Re-send last study data to newly connected client
                if (_lastSentStudyData != null)
                {
                    WriteMessage(_lastSentStudyData);
                }

                await ReadLoopAsync(_pipe, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                Logger.Trace($"PipeService: Connection error: {ex.Message}");
            }
            finally
            {
                _isConnected = false;
                try { _pipe?.Dispose(); } catch { }
                _pipe = null;
            }

            // Brief delay before re-listening
            if (!_cts.IsCancellationRequested)
            {
                try { await Task.Delay(500, _cts.Token); } catch (OperationCanceledException) { break; }
            }
        }
    }

    private async Task ReadLoopAsync(NamedPipeServerStream pipe, CancellationToken ct)
    {
        while (!ct.IsCancellationRequested && pipe.IsConnected)
        {
            // Read 4-byte length prefix
            var lengthBuf = new byte[4];
            var bytesRead = await ReadExactAsync(pipe, lengthBuf, ct);
            if (bytesRead < 4)
            {
                Logger.Trace("PipeService: Client disconnected (incomplete length)");
                return;
            }

            var length = (int)BitConverter.ToUInt32(lengthBuf, 0);
            if (length <= 0 || length > 1_048_576) // 1MB sanity limit
            {
                Logger.Trace($"PipeService: Invalid message length {length}, disconnecting");
                return;
            }

            var payloadBuf = new byte[length];
            bytesRead = await ReadExactAsync(pipe, payloadBuf, ct);
            if (bytesRead < length)
            {
                Logger.Trace("PipeService: Client disconnected (incomplete payload)");
                return;
            }

            var json = Encoding.UTF8.GetString(payloadBuf);
            DispatchMessage(json);
        }
    }

    private void DispatchMessage(string json)
    {
        try
        {
            // Peek at the type field to determine which model to deserialize
            using var doc = JsonDocument.Parse(json);
            var type = doc.RootElement.GetProperty("type").GetString();

            if (type == "shift_info")
            {
                var msg = JsonSerializer.Deserialize<ShiftInfoMessage>(json);
                if (msg != null)
                {
                    lock (_shiftLock)
                    {
                        _latestShiftInfo = msg;
                    }
                    Logger.Trace($"PipeService: Received shift_info: rvu={msg.TotalRvu:F1}, records={msg.RecordCount}, active={msg.IsShiftActive}, curHr={msg.CurrentHourRvu?.ToString("F1") ?? "null"}, prevHr={msg.PriorHourRvu?.ToString("F1") ?? "null"}, est={msg.EstimatedTotalRvu?.ToString("F1") ?? "null"}");
                    ShiftInfoUpdated?.Invoke();
                }
            }
            else if (type == "distraction_alert")
            {
                var msg = JsonSerializer.Deserialize<DistractionAlertMessage>(json);
                if (msg != null)
                {
                    Logger.Trace($"PipeService: Received distraction_alert: level={msg.AlertLevel}, study={msg.StudyType}, elapsed={msg.ElapsedSeconds:F0}s");
                    DistractionAlertReceived?.Invoke(msg);
                }
            }
            else
            {
                Logger.Trace($"PipeService: Unknown message type '{type}'");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"PipeService: Error parsing message: {ex.Message}");
        }
    }

    /// <summary>
    /// Send study data to the pipe client. Only sends if data has changed since last send.
    /// </summary>
    public void SendStudyData(StudyDataMessage msg)
    {
        if (!_isConnected) return;
        if (msg == _lastSentStudyData) return; // Record equality check

        _lastSentStudyData = msg;
        WriteMessage(msg);
    }

    /// <summary>
    /// Send a study lifecycle event immediately.
    /// </summary>
    public void SendStudyEvent(StudyEventMessage msg)
    {
        if (!_isConnected) return;
        WriteMessage(msg);
    }

    private void WriteMessage(object msg)
    {
        lock (_writeLock)
        {
            try
            {
                if (_pipe == null || !_pipe.IsConnected) return;

                var json = JsonSerializer.Serialize(msg, msg.GetType());
                var payload = Encoding.UTF8.GetBytes(json);
                var lengthPrefix = BitConverter.GetBytes((uint)payload.Length);

                _pipe.Write(lengthPrefix, 0, 4);
                _pipe.Write(payload, 0, payload.Length);
                _pipe.Flush();
            }
            catch (IOException ex)
            {
                Logger.Trace($"PipeService: Write error: {ex.Message}");
                _isConnected = false;
            }
        }
    }

    /// <summary>
    /// Read exactly count bytes from stream. Returns actual bytes read (less than count means disconnect).
    /// </summary>
    private static async Task<int> ReadExactAsync(Stream stream, byte[] buffer, CancellationToken ct)
    {
        int offset = 0;
        while (offset < buffer.Length)
        {
            var read = await stream.ReadAsync(buffer, offset, buffer.Length - offset, ct);
            if (read == 0) return offset; // Disconnected
            offset += read;
        }
        return offset;
    }

    public void Dispose()
    {
        _cts.Cancel();
        _cts.Dispose();
        try { _pipe?.Dispose(); } catch { }
        _pipe = null;
    }
}
