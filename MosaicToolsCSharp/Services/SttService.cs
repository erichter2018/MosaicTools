// [CustomSTT] Orchestrator for audio capture + STT provider
using NAudio.Wave;

namespace MosaicTools.Services;

/// <summary>
/// Manages PowerMic audio capture and STT provider lifecycle.
/// </summary>
public class SttService : IDisposable
{
    private readonly Configuration _config;
    private ISttProvider? _provider;
    private WaveInEvent? _waveIn;
    private int _selectedDeviceIndex = -1;
    private volatile bool _recording;
    private System.Threading.Timer? _keepAliveTimer;

    public bool IsRecording => _recording;
    public bool IsConnected => _provider?.IsConnected ?? false;
    public string ProviderName => _provider?.Name ?? "None";

    public event Action<SttResult>? TranscriptionReceived;
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

        // Keep _recording true briefly so OnAudioDataAvailable continues sending
        // audio to the provider — captures the tail end of the last spoken word.
        RecordingStateChanged?.Invoke(false);
        await Task.Delay(300); // Speech tail buffer

        _recording = false;

        // Capture and null the reference to prevent races with OnRecordingStopped/OnConnectionStateChanged
        var waveIn = _waveIn;
        _waveIn = null;

        if (waveIn != null)
        {
            try
            {
                waveIn.RecordingStopped -= OnRecordingStopped;
                waveIn.StopRecording();
                // Let final NAudio buffers drain through DataAvailable
                await Task.Delay(100);
                waveIn.DataAvailable -= OnAudioDataAvailable;
            }
            catch (Exception ex)
            {
                Logger.Trace($"SttService: StopRecording error: {ex.Message}");
            }
            finally
            {
                try { waveIn.Dispose(); } catch { }
            }
        }

        // Send Finalize to flush pending transcripts (provider waits for final response)
        if (_provider?.IsConnected == true)
        {
            await _provider.FinalizeAsync();

            // Disconnect after finalize to prevent turn state bleeding between PTT presses.
            // Providers like AssemblyAI accumulate turn context across a persistent WebSocket,
            // causing text from the previous session to appear when recording restarts.
            // StartRecordingAsync will reconnect automatically on next PTT press.
            await _provider.DisconnectAsync();
        }

        // No keepalive needed — we disconnect between PTT presses now
        _keepAliveTimer?.Change(Timeout.Infinite, Timeout.Infinite);

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
        // Don't dispose _waveIn here — disposal is handled by StopRecordingAsync/Dispose
        // to avoid racing with the native StopRecording call that triggered this callback.
    }

    private void OnTranscriptionReceived(SttResult result)
    {
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

            var waveIn = _waveIn;
            _waveIn = null;
            if (waveIn != null)
            {
                try
                {
                    waveIn.DataAvailable -= OnAudioDataAvailable;
                    waveIn.RecordingStopped -= OnRecordingStopped;
                    waveIn.StopRecording();
                }
                catch { }
                finally
                {
                    try { waveIn.Dispose(); } catch { }
                }
            }

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
        // Check that the active provider has credentials configured
        var hasKey = _config.SttProvider switch
        {
            "deepgram" => !string.IsNullOrEmpty(_config.SttApiKey),
            "speechmatics" => !string.IsNullOrEmpty(_config.SttSpeechmaticsApiKey),
            "assemblyai" => !string.IsNullOrEmpty(_config.SttAssemblyAIApiKey),
            "corti" => !string.IsNullOrEmpty(_config.SttCortiClientId),
            _ => !string.IsNullOrEmpty(_config.SttApiKey)
        };
        if (!hasKey)
        {
            Logger.Trace("SttService: No API key configured");
            return null;
        }

        return _config.SttProvider switch
        {
            "deepgram" => new DeepgramProvider(_config.SttApiKey, _config.SttModel, _config.SttAutoPunctuate),
            "assemblyai" => new AssemblyAIProvider(_config.SttAssemblyAIApiKey, _config.SttAutoPunctuate),
            "corti" => new CortiProvider(_config.SttCortiClientId, _config.SttCortiClientSecret, _config.SttCortiEnvironment, _config.SttAutoPunctuate),
            "speechmatics" => new SpeechmaticsProvider(_config.SttSpeechmaticsApiKey, _config.SttSpeechmaticsRegion, _config.SttAutoPunctuate),
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

        var waveIn = _waveIn;
        _waveIn = null;
        if (waveIn != null)
        {
            try
            {
                waveIn.DataAvailable -= OnAudioDataAvailable;
                waveIn.RecordingStopped -= OnRecordingStopped;
                waveIn.StopRecording();
            }
            catch { }
            finally
            {
                try { waveIn.Dispose(); } catch { }
            }
        }

        // Dispose provider on a background thread to avoid blocking the UI thread
        // (provider Dispose calls DisconnectAsync synchronously which can take seconds)
        var provider = _provider;
        _provider = null;
        if (provider != null)
            Task.Run(() => { try { provider.Dispose(); } catch { } });
    }
}
