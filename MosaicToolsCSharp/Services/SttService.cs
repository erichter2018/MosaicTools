// [CustomSTT] Orchestrator for audio capture + STT provider + cost tracking
using NAudio.Wave;

namespace MosaicTools.Services;

/// <summary>
/// Manages PowerMic audio capture, STT provider lifecycle, and cost tracking.
/// </summary>
public class SttService : IDisposable
{
    private readonly Configuration _config;
    private ISttProvider? _provider;
    private WaveInEvent? _waveIn;
    private int _selectedDeviceIndex = -1;
    private volatile bool _recording;
    private System.Threading.Timer? _keepAliveTimer;

    // Cost tracking
    private double _totalDurationSeconds;
    private decimal _sessionCost;

    public bool IsRecording => _recording;
    public bool IsConnected => _provider?.IsConnected ?? false;
    public decimal SessionCost => _sessionCost;
    public string ProviderName => _provider?.Name ?? "None";

    public event Action<SttResult>? TranscriptionReceived;
    public event Action<decimal>? CostUpdated;
    public event Action<string>? StatusChanged;
    public event Action<bool>? RecordingStateChanged;
    public event Action<string>? ErrorOccurred;

    public SttService(Configuration config)
    {
        _config = config;
    }

    /// <summary>
    /// Initialize provider and find audio device. Call once on startup.
    /// Returns error message or null on success.
    /// </summary>
    public string? Initialize()
    {
        // Find PowerMic audio device
        _selectedDeviceIndex = FindAudioDevice();
        if (_selectedDeviceIndex < 0)
        {
            var msg = "PowerMic audio device not found. Check Settings.";
            Logger.Trace($"SttService: {msg}");
            return msg;
        }

        // Create provider
        _provider = CreateProvider();
        if (_provider == null)
        {
            return "STT provider not configured. Check Settings.";
        }

        _provider.TranscriptionReceived += OnTranscriptionReceived;
        _provider.ErrorOccurred += OnProviderError;
        _provider.ConnectionStateChanged += OnConnectionStateChanged;

        // KeepAlive timer (disabled until first recording starts/stops)
        _keepAliveTimer = new System.Threading.Timer(async _ =>
        {
            if (_provider?.IsConnected == true && !_recording)
                await _provider.SendKeepAliveAsync();
        }, null, Timeout.Infinite, Timeout.Infinite);

        Logger.Trace($"SttService: Initialized (device={_selectedDeviceIndex}, provider={_provider.Name})");
        return null;
    }

    /// <summary>
    /// Start recording audio and streaming to provider.
    /// Connects to provider if not already connected.
    /// </summary>
    public async Task StartRecordingAsync()
    {
        if (_recording) return;
        if (_provider == null || _selectedDeviceIndex < 0) return;

        // Connect if needed
        if (!_provider.IsConnected)
        {
            StatusChanged?.Invoke("Connecting...");
            var connected = await _provider.ConnectAsync();
            if (!connected)
            {
                StatusChanged?.Invoke("Connection failed");
                return;
            }
        }

        // Stop keepalive timer (we're actively sending audio now)
        _keepAliveTimer?.Change(Timeout.Infinite, Timeout.Infinite);

        // Start audio capture
        try
        {
            _waveIn = new WaveInEvent
            {
                DeviceNumber = _selectedDeviceIndex,
                WaveFormat = new WaveFormat(16000, 16, 1),
                BufferMilliseconds = 100
            };
            _waveIn.DataAvailable += OnAudioDataAvailable;
            _waveIn.RecordingStopped += OnRecordingStopped;
            _waveIn.StartRecording();

            _recording = true;
            RecordingStateChanged?.Invoke(true);
            StatusChanged?.Invoke("Recording...");
            Logger.Trace("SttService: Recording started");
        }
        catch (Exception ex)
        {
            Logger.Trace($"SttService: Failed to start recording: {ex.Message}");
            ErrorOccurred?.Invoke($"Audio capture failed: {ex.Message}");
            StatusChanged?.Invoke("Audio error");
        }
    }

    /// <summary>
    /// Stop recording and finalize pending transcripts.
    /// Keeps WebSocket connection alive for next PTT press.
    /// </summary>
    public async Task StopRecordingAsync()
    {
        if (!_recording) return;

        _recording = false;
        RecordingStateChanged?.Invoke(false);

        try
        {
            _waveIn?.StopRecording();
        }
        catch (Exception ex)
        {
            Logger.Trace($"SttService: StopRecording error: {ex.Message}");
        }

        // Send Finalize to flush pending transcripts
        if (_provider?.IsConnected == true)
        {
            await _provider.FinalizeAsync();
        }

        // Start keepalive timer to keep WebSocket open between PTT presses
        _keepAliveTimer?.Change(8000, 8000);

        StatusChanged?.Invoke("Stopped");
        Logger.Trace("SttService: Recording stopped");
    }

    /// <summary>
    /// Fully disconnect from provider (e.g., on settings change or shutdown).
    /// </summary>
    public async Task DisconnectAsync()
    {
        if (_recording)
        {
            await StopRecordingAsync();
        }

        _keepAliveTimer?.Change(Timeout.Infinite, Timeout.Infinite);

        if (_provider != null)
        {
            await _provider.DisconnectAsync();
        }

        StatusChanged?.Invoke("Disconnected");
    }

    /// <summary>
    /// Reset session cost counter.
    /// </summary>
    public void ResetCost()
    {
        _totalDurationSeconds = 0;
        _sessionCost = 0;
        CostUpdated?.Invoke(0);
    }

    private void OnAudioDataAvailable(object? sender, WaveInEventArgs e)
    {
        if (!_recording || _provider == null) return;
        _provider.SendAudio(e.Buffer, 0, e.BytesRecorded);
    }

    private void OnRecordingStopped(object? sender, StoppedEventArgs e)
    {
        if (e.Exception != null)
        {
            Logger.Trace($"SttService: Recording stopped with error: {e.Exception.Message}");
            ErrorOccurred?.Invoke($"Audio device error: {e.Exception.Message}");
            StatusChanged?.Invoke("Audio error");
        }

        _waveIn?.Dispose();
        _waveIn = null;
    }

    private void OnTranscriptionReceived(SttResult result)
    {
        // Track cost from final results
        if (result.IsFinal && result.Duration > 0)
        {
            _totalDurationSeconds += result.Duration;
            _sessionCost = (decimal)(_totalDurationSeconds / 60.0) * (_provider?.CostPerMinute ?? 0);
            Logger.Trace($"SttService: Cost update - duration={result.Duration:F2}s, total={_totalDurationSeconds:F1}s, cost=${_sessionCost:F4}");
            CostUpdated?.Invoke(_sessionCost);
        }

        TranscriptionReceived?.Invoke(result);
    }

    private void OnProviderError(string error)
    {
        Logger.Trace($"SttService: Provider error: {error}");
        ErrorOccurred?.Invoke(error);
    }

    private void OnConnectionStateChanged(bool connected)
    {
        if (!connected && _recording)
        {
            // Connection dropped while recording
            _recording = false;
            RecordingStateChanged?.Invoke(false);
            try { _waveIn?.StopRecording(); } catch { }
            StatusChanged?.Invoke("Disconnected");

            // Attempt reconnection
            _ = Task.Run(ReconnectAsync);
        }
    }

    private async Task ReconnectAsync()
    {
        for (int attempt = 1; attempt <= 3; attempt++)
        {
            StatusChanged?.Invoke($"Reconnecting ({attempt}/3)...");
            Logger.Trace($"SttService: Reconnect attempt {attempt}");

            await Task.Delay(1000 * attempt); // Exponential backoff

            if (_provider == null) break;
            var success = await _provider.ConnectAsync();
            if (success)
            {
                StatusChanged?.Invoke("Reconnected");
                Logger.Trace("SttService: Reconnected successfully");
                return;
            }
        }

        StatusChanged?.Invoke("Reconnection failed");
        ErrorOccurred?.Invoke("Could not reconnect to STT service.");
        Logger.Trace("SttService: Reconnection failed after 3 attempts");
    }

    private int FindAudioDevice()
    {
        var configuredName = _config.SttAudioDeviceName;
        int deviceCount = WaveInEvent.DeviceCount;

        // If user configured a specific device, look for exact match
        if (!string.IsNullOrEmpty(configuredName))
        {
            for (int i = 0; i < deviceCount; i++)
            {
                var caps = WaveInEvent.GetCapabilities(i);
                if (caps.ProductName.Contains(configuredName, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Trace($"SttService: Found configured device '{caps.ProductName}' at index {i}");
                    return i;
                }
            }
            Logger.Trace($"SttService: Configured device '{configuredName}' not found, falling back to auto-detect");
        }

        // Auto-detect PowerMic by known name fragments
        string[] micKeywords = { "PowerMic", "Dictaphone", "Nuance", "SpeechMike", "Philips" };
        for (int i = 0; i < deviceCount; i++)
        {
            var caps = WaveInEvent.GetCapabilities(i);
            foreach (var keyword in micKeywords)
            {
                if (caps.ProductName.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Trace($"SttService: Auto-detected '{caps.ProductName}' at index {i}");
                    return i;
                }
            }
        }

        // Log available devices for debugging
        for (int i = 0; i < deviceCount; i++)
        {
            var caps = WaveInEvent.GetCapabilities(i);
            Logger.Trace($"SttService: Available device [{i}]: '{caps.ProductName}'");
        }

        return -1;
    }

    private ISttProvider? CreateProvider()
    {
        if (string.IsNullOrEmpty(_config.SttApiKey))
        {
            Logger.Trace("SttService: No API key configured");
            return null;
        }

        return _config.SttProvider switch
        {
            "deepgram" => new DeepgramProvider(_config.SttApiKey, _config.SttModel, _config.SttAutoPunctuate),
            _ => new DeepgramProvider(_config.SttApiKey, _config.SttModel, _config.SttAutoPunctuate)
        };
    }

    /// <summary>
    /// Get list of available audio input devices for settings UI.
    /// </summary>
    public static List<string> GetAudioDevices()
    {
        var devices = new List<string>();
        for (int i = 0; i < WaveInEvent.DeviceCount; i++)
        {
            var caps = WaveInEvent.GetCapabilities(i);
            devices.Add(caps.ProductName);
        }
        return devices;
    }

    public void Dispose()
    {
        _keepAliveTimer?.Dispose();
        try { _waveIn?.StopRecording(); } catch { }
        _waveIn?.Dispose();
        _provider?.Dispose();
    }
}
