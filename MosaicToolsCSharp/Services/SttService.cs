// [CustomSTT] Orchestrator for audio capture + STT provider
using NAudio.Wave;

namespace MosaicTools.Services;

/// <summary>
/// Manages PowerMic audio capture and STT provider lifecycle.
/// Thin orchestrator — all timing, flush, and disconnect logic lives in providers.
/// </summary>
public class SttService : IDisposable
{
    private readonly Configuration _config;
    private readonly string? _keytermOverride;
    private ISttProvider? _provider;
    private WaveInEvent? _waveIn;
    private int _selectedDeviceIndex = -1;
    private volatile bool _recording;
    private volatile bool _stopping; // Suppress reconnect during intentional stop

    public bool IsRecording => _recording;
    public bool IsConnected => _provider?.IsConnected ?? false;
    public string ProviderName => _provider?.Name ?? "None";

    public event Action<SttResult>? TranscriptionReceived;
    public event Action<string>? StatusChanged;
    public event Action<bool>? RecordingStateChanged;
    public event Action<string>? ErrorOccurred;

    public SttService(Configuration config, string? keytermOverride = null)
    {
        _config = config;
        _keytermOverride = keytermOverride;
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

        // Start session (provider connects if needed)
        if (!_provider.IsConnected)
        {
            StatusChanged?.Invoke("Connecting...");
            var connected = await _provider.StartSessionAsync();
            if (!connected)
            {
                StatusChanged?.Invoke("Connection failed");
                return;
            }
        }

        // Start audio capture using provider's declared format
        try
        {
            var fmt = _provider.AudioFormat;
            _waveIn = new WaveInEvent
            {
                DeviceNumber = _selectedDeviceIndex,
                WaveFormat = new WaveFormat(fmt.SampleRate, fmt.BitsPerSample, fmt.Channels),
                BufferMilliseconds = fmt.BufferMs
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
    /// Stop recording. Provider owns the full stop pipeline —
    /// tail buffer, flush, and disconnect decisions.
    /// Audio keeps flowing during EndSessionAsync so the provider
    /// captures trailing speech.
    /// </summary>
    public async Task StopRecordingAsync()
    {
        if (!_recording) return;
        _stopping = true; // Suppress OnConnectionStateChanged reconnect

        // Update UI immediately
        RecordingStateChanged?.Invoke(false);

        // Provider runs its full stop pipeline (tail buffer, flush, optional disconnect).
        // _recording stays true so OnAudioDataAvailable keeps sending audio during this.
        try
        {
            if (_provider is { IsConnected: true })
                await _provider.EndSessionAsync();
        }
        catch (Exception ex)
        {
            Logger.Trace($"SttService: EndSession error: {ex.Message}");
        }

        // Now stop audio capture
        _recording = false;
        _stopping = false;

        // Capture and null the reference to prevent races with OnRecordingStopped/OnConnectionStateChanged
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
            catch (Exception ex)
            {
                Logger.Trace($"SttService: StopRecording error: {ex.Message}");
            }
            finally
            {
                try { waveIn.Dispose(); } catch { }
            }
        }

        StatusChanged?.Invoke("Stopped");
        Logger.Trace("SttService: Recording stopped");

        // Pre-connect for next PTT press so the user doesn't wait on reconnection.
        // Providers that disconnect each session (AssemblyAI, Deepgram) benefit most.
        // Providers still connected (Corti) will no-op in StartSessionAsync.
        var provider = _provider;
        if (provider is { IsConnected: false })
        {
            _ = Task.Run(async () =>
            {
                try { await provider.StartSessionAsync(); }
                catch (Exception ex) { Logger.Trace($"SttService: Pre-connect failed: {ex.Message}"); }
            });
        }
    }

    /// <summary>
    /// Fully shut down provider (e.g., on settings change or app exit).
    /// </summary>
    public async Task DisconnectAsync()
    {
        if (_recording)
        {
            await StopRecordingAsync();
        }

        if (_provider != null)
        {
            await _provider.ShutdownAsync();
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
        if (!connected && _recording && !_stopping)
        {
            // Connection dropped unexpectedly while recording
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
            var success = await _provider.StartSessionAsync();
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

        var keyterms = _keytermOverride ?? _config.SttDeepgramKeyterms;
        return _config.SttProvider switch
        {
            "deepgram" => new DeepgramProvider(_config.SttApiKey, _config.SttModel, _config.SttAutoPunctuate, keyterms),
            "assemblyai" => new AssemblyAIProvider(_config.SttAssemblyAIApiKey, _config.SttAutoPunctuate),
            "corti" => new CortiProvider(_config.SttCortiClientId, _config.SttCortiClientSecret, _config.SttCortiEnvironment, _config.SttAutoPunctuate),
            "speechmatics" => new SpeechmaticsProvider(_config.SttSpeechmaticsApiKey, _config.SttSpeechmaticsRegion, _config.SttAutoPunctuate),
            _ => new DeepgramProvider(_config.SttApiKey, _config.SttModel, _config.SttAutoPunctuate, keyterms)
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
        // (provider Dispose calls ShutdownAsync synchronously which can take seconds)
        var provider = _provider;
        _provider = null;
        if (provider != null)
            Task.Run(() => { try { provider.Dispose(); } catch { } });
    }
}
