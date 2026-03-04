// [CustomSTT] Orchestrator for audio capture + STT provider(s)
using NAudio.Wave;

namespace MosaicTools.Services;

/// <summary>
/// Manages PowerMic audio capture and STT provider lifecycle.
/// Supports single-provider mode (unchanged from original) and ensemble mode
/// (fan-out to 3 providers with merger).
/// </summary>
public class SttService : IDisposable
{
    private readonly Configuration _config;
    private readonly string? _keytermOverride;
    private readonly Dictionary<string, string>? _perProviderKeyterms;

    // Single-provider mode (unchanged path)
    private ISttProvider? _provider;

    // Ensemble mode
    private ISttProvider? _primaryProvider;     // Deepgram (always)
    private ISttProvider? _secondaryProvider1;  // AssemblyAI
    private ISttProvider? _secondaryProvider2;  // Speechmatics
    private SttEnsembleMerger? _merger;
    private bool _ensembleMode;

    private WaveInEvent? _waveIn;
    private int _selectedDeviceIndex = -1;
    private volatile bool _recording;
    private volatile bool _stopping;

    public bool IsRecording => _recording;
    public bool IsConnected => _ensembleMode
        ? (_primaryProvider?.IsConnected ?? false)
        : (_provider?.IsConnected ?? false);
    public string ProviderName => _ensembleMode
        ? "Ensemble (DG+AAI+SM)"
        : (_provider?.Name ?? "None");

    public event Action<SttResult>? TranscriptionReceived;
    /// <summary>
    /// Fires for every provider's final result (before merge). Used for per-provider keyterm learning.
    /// </summary>
    public event Action<SttResult>? RawProviderFinalReceived;
    /// <summary>
    /// Fires after each ensemble merge with live stats snapshot. Only in ensemble mode.
    /// </summary>
    public event Action<EnsembleStats>? EnsembleStatsUpdated;
    public event Action<string>? StatusChanged;
    public event Action<bool>? RecordingStateChanged;
    public event Action<string>? ErrorOccurred;

    public SttService(Configuration config, string? keytermOverride = null,
        Dictionary<string, string>? perProviderKeyterms = null)
    {
        _config = config;
        _keytermOverride = keytermOverride;
        _perProviderKeyterms = perProviderKeyterms;
    }

    public string? Initialize()
    {
        _selectedDeviceIndex = FindAudioDevice();
        if (_selectedDeviceIndex < 0)
        {
            var msg = "PowerMic audio device not found. Check Settings.";
            Logger.Trace($"SttService: {msg}");
            return msg;
        }

        // Decide: ensemble or single-provider
        if (_config.SttEnsembleEnabled && CanRunEnsemble())
        {
            return InitializeEnsemble();
        }
        else
        {
            // Single-provider mode — unchanged path
            _ensembleMode = false;
            _provider = CreateProvider();
            if (_provider == null)
                return "STT provider not configured. Check Settings.";

            _provider.TranscriptionReceived += OnTranscriptionReceived;
            _provider.ErrorOccurred += OnProviderError;
            _provider.ConnectionStateChanged += OnConnectionStateChanged;

            Logger.Trace($"SttService: Initialized single-provider (device={_selectedDeviceIndex}, provider={_provider.Name})");
            return null;
        }
    }

    private bool CanRunEnsemble()
    {
        return !string.IsNullOrEmpty(_config.SttApiKey) &&
               !string.IsNullOrEmpty(_config.SttAssemblyAIApiKey) &&
               !string.IsNullOrEmpty(_config.SttSpeechmaticsApiKey);
    }

    private string? InitializeEnsemble()
    {
        _ensembleMode = true;

        var dgKeyterms = GetProviderKeyterms("deepgram");
        var aaiKeyterms = GetProviderKeyterms("assemblyai");
        var smKeyterms = GetProviderKeyterms("speechmatics");

        _primaryProvider = new DeepgramProvider(_config.SttApiKey, _config.SttModel, _config.SttAutoPunctuate, dgKeyterms);
        _secondaryProvider1 = new AssemblyAIProvider(_config.SttAssemblyAIApiKey, _config.SttAutoPunctuate, aaiKeyterms);
        _secondaryProvider2 = new SpeechmaticsProvider(_config.SttSpeechmaticsApiKey, _config.SttSpeechmaticsRegion, _config.SttAutoPunctuate, smKeyterms);

        _merger = new SttEnsembleMerger(
            _config,
            _config.SttEnsembleWaitMs,
            _config.SttEnsembleConfidenceThreshold);

        _merger.MergedResultReady += result =>
        {
            TranscriptionReceived?.Invoke(result);
        };
        _merger.StatsUpdated += stats =>
        {
            EnsembleStatsUpdated?.Invoke(stats);
        };

        // Primary (Deepgram): interims go to display, finals go to merger + raw event
        _primaryProvider.TranscriptionReceived += result =>
        {
            var tagged = result with { ProviderName = "deepgram" };
            if (!result.IsFinal)
            {
                TranscriptionReceived?.Invoke(tagged); // Interims for live display
            }
            else
            {
                RawProviderFinalReceived?.Invoke(tagged);
                _merger?.SubmitResult(tagged);
            }
        };
        _primaryProvider.ErrorOccurred += OnProviderError;
        _primaryProvider.ConnectionStateChanged += OnConnectionStateChanged;

        // Secondary providers: only finals, errors logged silently
        _secondaryProvider1.TranscriptionReceived += result =>
        {
            if (!result.IsFinal) return;
            var tagged = result with { ProviderName = "assemblyai" };
            RawProviderFinalReceived?.Invoke(tagged);
            _merger?.SubmitResult(tagged);
        };
        _secondaryProvider1.ErrorOccurred += err => Logger.Trace($"SttService [AAI secondary]: {err}");
        _secondaryProvider1.ConnectionStateChanged += connected =>
        {
            if (!connected) Logger.Trace("SttService: AAI secondary disconnected");
        };

        _secondaryProvider2.TranscriptionReceived += result =>
        {
            if (!result.IsFinal) return;
            var tagged = result with { ProviderName = "speechmatics" };
            RawProviderFinalReceived?.Invoke(tagged);
            _merger?.SubmitResult(tagged);
        };
        _secondaryProvider2.ErrorOccurred += err => Logger.Trace($"SttService [SM secondary]: {err}");
        _secondaryProvider2.ConnectionStateChanged += connected =>
        {
            if (!connected) Logger.Trace("SttService: SM secondary disconnected");
        };

        Logger.Trace($"SttService: Initialized ensemble mode (device={_selectedDeviceIndex})");
        return null;
    }

    private string GetProviderKeyterms(string providerName)
    {
        if (_perProviderKeyterms != null && _perProviderKeyterms.TryGetValue(providerName, out var kt))
            return kt;
        return _keytermOverride ?? _config.SttDeepgramKeyterms;
    }

    public async Task StartRecordingAsync()
    {
        if (_recording) return;

        if (_ensembleMode)
        {
            if (_primaryProvider == null) return;

            // Primary must connect
            if (!_primaryProvider.IsConnected)
            {
                StatusChanged?.Invoke("Connecting...");
                var connected = await _primaryProvider.StartSessionAsync();
                if (!connected)
                {
                    StatusChanged?.Invoke("Connection failed");
                    return;
                }
            }

            // Secondaries connect in parallel (failures tolerated)
            var s1 = _secondaryProvider1;
            var s2 = _secondaryProvider2;
            _ = Task.Run(async () =>
            {
                try { if (s1 is { IsConnected: false }) await s1.StartSessionAsync(); }
                catch (Exception ex) { Logger.Trace($"SttService: AAI connect failed: {ex.Message}"); }
            });
            _ = Task.Run(async () =>
            {
                try { if (s2 is { IsConnected: false }) await s2.StartSessionAsync(); }
                catch (Exception ex) { Logger.Trace($"SttService: SM connect failed: {ex.Message}"); }
            });
        }
        else
        {
            if (_provider == null || _selectedDeviceIndex < 0) return;

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
        }

        // Start audio capture
        try
        {
            var fmt = _ensembleMode ? _primaryProvider!.AudioFormat : _provider!.AudioFormat;
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

    public async Task StopRecordingAsync()
    {
        if (!_recording) return;
        _stopping = true;

        RecordingStateChanged?.Invoke(false);

        if (_ensembleMode)
        {
            // End primary first
            try
            {
                if (_primaryProvider is { IsConnected: true })
                    await _primaryProvider.EndSessionAsync();
            }
            catch (Exception ex) { Logger.Trace($"SttService: Primary EndSession error: {ex.Message}"); }

            // End secondaries in parallel with timeout
            var s1 = _secondaryProvider1;
            var s2 = _secondaryProvider2;
            var t1 = Task.Run(async () =>
            {
                try { if (s1 is { IsConnected: true }) await s1.EndSessionAsync(); } catch { }
            });
            var t2 = Task.Run(async () =>
            {
                try { if (s2 is { IsConnected: true }) await s2.EndSessionAsync(); } catch { }
            });
            await Task.WhenAny(Task.WhenAll(t1, t2), Task.Delay(2000));
        }
        else
        {
            try
            {
                if (_provider is { IsConnected: true })
                    await _provider.EndSessionAsync();
            }
            catch (Exception ex)
            {
                Logger.Trace($"SttService: EndSession error: {ex.Message}");
            }
        }

        _recording = false;
        _stopping = false;

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

        // Pre-connect for next PTT press
        if (_ensembleMode)
        {
            var p = _primaryProvider;
            var s1 = _secondaryProvider1;
            var s2 = _secondaryProvider2;
            _ = Task.Run(async () =>
            {
                try { if (p is { IsConnected: false }) await p.StartSessionAsync(); } catch { }
                try { if (s1 is { IsConnected: false }) await s1.StartSessionAsync(); } catch { }
                try { if (s2 is { IsConnected: false }) await s2.StartSessionAsync(); } catch { }
            });
        }
        else
        {
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
    }

    public async Task DisconnectAsync()
    {
        if (_recording)
            await StopRecordingAsync();

        if (_ensembleMode)
        {
            if (_primaryProvider != null) await _primaryProvider.ShutdownAsync();
            if (_secondaryProvider1 != null) await _secondaryProvider1.ShutdownAsync();
            if (_secondaryProvider2 != null) await _secondaryProvider2.ShutdownAsync();
        }
        else
        {
            if (_provider != null) await _provider.ShutdownAsync();
        }

        StatusChanged?.Invoke("Disconnected");
    }

    private void OnAudioDataAvailable(object? sender, WaveInEventArgs e)
    {
        if (!_recording) return;

        if (_ensembleMode)
        {
            // Fan out same audio to all providers
            _primaryProvider?.SendAudio(e.Buffer, 0, e.BytesRecorded);
            _secondaryProvider1?.SendAudio(e.Buffer, 0, e.BytesRecorded);
            _secondaryProvider2?.SendAudio(e.Buffer, 0, e.BytesRecorded);
        }
        else
        {
            _provider?.SendAudio(e.Buffer, 0, e.BytesRecorded);
        }
    }

    private void OnRecordingStopped(object? sender, StoppedEventArgs e)
    {
        if (e.Exception != null)
        {
            Logger.Trace($"SttService: Recording stopped with error: {e.Exception.Message}");
            ErrorOccurred?.Invoke($"Audio device error: {e.Exception.Message}");
            StatusChanged?.Invoke("Audio error");
        }
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
            _ = Task.Run(ReconnectAsync);
        }
    }

    private async Task ReconnectAsync()
    {
        for (int attempt = 1; attempt <= 3; attempt++)
        {
            StatusChanged?.Invoke($"Reconnecting ({attempt}/3)...");
            Logger.Trace($"SttService: Reconnect attempt {attempt}");

            await Task.Delay(1000 * attempt);

            if (_ensembleMode)
            {
                if (_primaryProvider == null) break;
                var success = await _primaryProvider.StartSessionAsync();
                if (success)
                {
                    StatusChanged?.Invoke("Reconnected");
                    Logger.Trace("SttService: Reconnected successfully");
                    // Reconnect secondaries in background
                    var s1 = _secondaryProvider1;
                    var s2 = _secondaryProvider2;
                    _ = Task.Run(async () =>
                    {
                        try { if (s1 is { IsConnected: false }) await s1.StartSessionAsync(); } catch { }
                        try { if (s2 is { IsConnected: false }) await s2.StartSessionAsync(); } catch { }
                    });
                    return;
                }
            }
            else
            {
                if (_provider == null) break;
                var success = await _provider.StartSessionAsync();
                if (success)
                {
                    StatusChanged?.Invoke("Reconnected");
                    Logger.Trace("SttService: Reconnected successfully");
                    return;
                }
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

        for (int i = 0; i < deviceCount; i++)
        {
            var caps = WaveInEvent.GetCapabilities(i);
            Logger.Trace($"SttService: Available device [{i}]: '{caps.ProductName}'");
        }

        return -1;
    }

    private ISttProvider? CreateProvider()
    {
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
            "assemblyai" => new AssemblyAIProvider(_config.SttAssemblyAIApiKey, _config.SttAutoPunctuate, keyterms),
            "corti" => new CortiProvider(_config.SttCortiClientId, _config.SttCortiClientSecret, _config.SttCortiEnvironment, _config.SttAutoPunctuate),
            "speechmatics" => new SpeechmaticsProvider(_config.SttSpeechmaticsApiKey, _config.SttSpeechmaticsRegion, _config.SttAutoPunctuate, keyterms),
            _ => new DeepgramProvider(_config.SttApiKey, _config.SttModel, _config.SttAutoPunctuate, keyterms)
        };
    }

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

        if (_ensembleMode)
        {
            var p = _primaryProvider; _primaryProvider = null;
            var s1 = _secondaryProvider1; _secondaryProvider1 = null;
            var s2 = _secondaryProvider2; _secondaryProvider2 = null;
            Task.Run(() =>
            {
                try { p?.Dispose(); } catch { }
                try { s1?.Dispose(); } catch { }
                try { s2?.Dispose(); } catch { }
            });
        }
        else
        {
            var provider = _provider;
            _provider = null;
            if (provider != null)
                Task.Run(() => { try { provider.Dispose(); } catch { } });
        }
    }
}
