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
    private ISttProvider? _primaryProvider;     // Anchor/driver (configurable, default Deepgram)
    private ISttProvider? _secondaryProvider1;  // Configurable (default: Soniox)
    private ISttProvider? _secondaryProvider2;  // Configurable (default: Speechmatics)
    private string _s1Name = "";               // Provider name for secondary 1
    private string _s2Name = "";               // Provider name for secondary 2
    private SttEnsembleMerger? _merger;
    private bool _ensembleMode;

    private WaveInEvent? _waveIn;
    private int _selectedDeviceIndex = -1;
    private volatile bool _recording;
    private volatile bool _stopping;

    public bool IsRecording => _recording;
    public bool IsStopping => _stopping;
    public bool IsConnected => _ensembleMode
        ? (_primaryProvider?.IsConnected ?? false)
        : (_provider?.IsConnected ?? false);
    public string ProviderName => _ensembleMode
        ? "Ensemble (DG" +
          (_s1Name != "none" ? $"+{ShortName(_s1Name)}" : "") +
          (_s2Name != "none" ? $"+{ShortName(_s2Name)}" : "") + ")"
        : (_provider?.Name ?? "None");

    /// <summary>Returns the two secondary provider names when in ensemble mode.</summary>
    public (string s1, string s2) EnsembleSecondaries => (_s1Name, _s2Name);

    /// <summary>Exposes the merger for external validation counter updates.</summary>
    internal SttEnsembleMerger? Merger => _merger;

    private static string ShortName(string provider) => provider switch
    {
        "soniox" => "SNX",
        "speechmatics" => "SM",
        "assemblyai" => "AAI",
        "elevenlabs" => "EL",
        "smallestai" => "SAI",
        _ => provider.ToUpperInvariant()[..Math.Min(3, provider.Length)]
    };

    public event Action<SttResult>? TranscriptionReceived;
    /// <summary>
    /// Fires for every provider's final result (before merge). Used for per-provider keyterm learning.
    /// </summary>
    public event Action<SttResult>? RawProviderFinalReceived;
    /// <summary>
    /// Fires after each ensemble merge with live stats snapshot. Only in ensemble mode.
    /// </summary>
    public event Action<EnsembleStats>? EnsembleStatsUpdated;
    /// <summary>
    /// Fires after an ensemble merge that produced corrections, for validation against signed reports.
    /// </summary>
    public event Action<List<CorrectionRecord>>? EnsembleCorrectionsEmitted;
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

    /// <summary>Returns the maximum keyterm capacity for each provider.</summary>
    public static int GetKeytermLimit(string provider) => provider switch
    {
        "deepgram" => 100,      // keyterm URL params
        "assemblyai" => 1000,   // U3 Pro keyterms_prompt
        "speechmatics" => 1000, // additional_vocab array
        "soniox" => 500,        // context.terms
        "smallestai" => 100,    // keywords param
        "elevenlabs" => 0,      // no keyterm support
        _ => 100
    };

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
        // Need anchor provider + at least one secondary that's not "none" with a valid key
        var anchor = _config.SttEnsembleAnchor;
        if (!HasKeyForProvider(anchor)) return false;
        var s1 = _config.SttEnsembleSecondary1;
        var s2 = _config.SttEnsembleSecondary2;
        return (s1 != "none" && HasKeyForProvider(s1)) ||
               (s2 != "none" && HasKeyForProvider(s2));
    }

    private bool HasKeyForProvider(string provider) => provider switch
    {
        "deepgram" => !string.IsNullOrEmpty(_config.SttApiKey),
        "soniox" => !string.IsNullOrEmpty(_config.SttSonioxApiKey),
        "speechmatics" => !string.IsNullOrEmpty(_config.SttSpeechmaticsApiKey),
        "assemblyai" => !string.IsNullOrEmpty(_config.SttAssemblyAIApiKey),
        "elevenlabs" => !string.IsNullOrEmpty(_config.SttElevenLabsApiKey),
        "smallestai" => !string.IsNullOrEmpty(_config.SttSmallestAiApiKey),
        "none" => true,
        _ => false
    };

    // Providers should punctuate if either global auto-punctuate OR final-report-only punctuation is on
    private bool EffectivePunctuate => _config.SttAutoPunctuate || _config.SttAutoPunctuateFinalReport;

    private ISttProvider CreateEnsembleSecondary(string provider, string keyterms)
    {
        var punctuate = EffectivePunctuate;
        return provider switch
        {
            "deepgram" => new DeepgramProvider(_config.SttApiKey, _config.SttModel, punctuate, keyterms),
            "soniox" => new SonioxProvider(_config.SttSonioxApiKey, punctuate, keyterms),
            "speechmatics" => new SpeechmaticsProvider(_config.SttSpeechmaticsApiKey, _config.SttSpeechmaticsRegion, punctuate, keyterms),
            "assemblyai" => new AssemblyAIProvider(_config.SttAssemblyAIApiKey, punctuate, keyterms),
            "elevenlabs" => new ElevenLabsProvider(_config.SttElevenLabsApiKey, punctuate),
            "smallestai" => new SmallestAiProvider(_config.SttSmallestAiApiKey, punctuate, keyterms),
            _ => throw new ArgumentException($"Unknown ensemble provider: {provider}")
        };
    }

    private string _anchorName = "deepgram";

    private ISttProvider CreateAnchorProvider(string provider, string keyterms)
    {
        var punctuate = EffectivePunctuate;
        return provider switch
        {
            "deepgram" => new DeepgramProvider(_config.SttApiKey, _config.SttModel, punctuate, keyterms),
            "soniox" => new SonioxProvider(_config.SttSonioxApiKey, punctuate, keyterms),
            "speechmatics" => new SpeechmaticsProvider(_config.SttSpeechmaticsApiKey, _config.SttSpeechmaticsRegion, punctuate, keyterms),
            "assemblyai" => new AssemblyAIProvider(_config.SttAssemblyAIApiKey, punctuate, keyterms),
            "elevenlabs" => new ElevenLabsProvider(_config.SttElevenLabsApiKey, punctuate),
            "smallestai" => new SmallestAiProvider(_config.SttSmallestAiApiKey, punctuate, keyterms),
            _ => new DeepgramProvider(_config.SttApiKey, _config.SttModel, punctuate, keyterms)
        };
    }

    private string? InitializeEnsemble()
    {
        _ensembleMode = true;
        _anchorName = _config.SttEnsembleAnchor;
        _s1Name = _config.SttEnsembleSecondary1;
        _s2Name = _config.SttEnsembleSecondary2;

        var anchorKeyterms = GetProviderKeyterms(_anchorName);
        _primaryProvider = CreateAnchorProvider(_anchorName, anchorKeyterms);

        // Create secondary providers (skip "none")
        if (_s1Name != "none" && HasKeyForProvider(_s1Name))
        {
            _secondaryProvider1 = CreateEnsembleSecondary(_s1Name, GetProviderKeyterms(_s1Name));
        }
        if (_s2Name != "none" && HasKeyForProvider(_s2Name))
        {
            _secondaryProvider2 = CreateEnsembleSecondary(_s2Name, GetProviderKeyterms(_s2Name));
        }

        _merger = new SttEnsembleMerger(
            _config,
            _config.SttEnsembleWaitMs,
            _config.SttEnsembleConfidenceThreshold,
            _anchorName, _s1Name, _s2Name);

        _merger.MergedResultReady += result =>
        {
            TranscriptionReceived?.Invoke(result);
        };
        _merger.StatsUpdated += stats =>
        {
            EnsembleStatsUpdated?.Invoke(stats);
        };
        _merger.CorrectionsEmitted += corrections =>
        {
            EnsembleCorrectionsEmitted?.Invoke(corrections);
        };

        // Wire "Clear All-Time Stats" button callback to reset live merger counters
        _config.OnClearAllEnsembleStats = () => _merger.ResetAllStats();

        // Primary (anchor): interims go to display, finals go to merger + raw event
        var anchorShort = ShortName(_anchorName);
        _primaryProvider.TranscriptionReceived += result =>
        {
            var tagged = result with { ProviderName = _anchorName };
            if (!result.IsFinal)
            {
                TranscriptionReceived?.Invoke(tagged); // Interims for live display
            }
            else
            {
                var wc = result.Words?.Length ?? 0;
                var lowConf = result.Words?.Count(w => w.Confidence < _config.SttEnsembleConfidenceThreshold) ?? 0;
                var preview = result.Transcript.Length > 60 ? result.Transcript[..60] + "..." : result.Transcript;
                Logger.Trace($"Ensemble [{anchorShort}] final: {wc} words ({lowConf} low-conf), \"{preview}\"");
                if (result.Words != null)
                {
                    var wordDetails = string.Join(", ", result.Words.Select(w =>
                        $"{w.Text}({w.Confidence:F2})"));
                    Logger.Trace($"Ensemble [{anchorShort}] words: {wordDetails}");
                }
                RawProviderFinalReceived?.Invoke(tagged);
                _merger?.SubmitResult(tagged);
            }
        };
        _primaryProvider.ErrorOccurred += OnProviderError;
        _primaryProvider.ConnectionStateChanged += OnConnectionStateChanged;

        // Secondary providers: only finals, errors logged silently
        if (_secondaryProvider1 != null) WireSecondary(_secondaryProvider1, _s1Name);
        if (_secondaryProvider2 != null) WireSecondary(_secondaryProvider2, _s2Name);

        Logger.Trace($"SttService: Initialized ensemble mode (device={_selectedDeviceIndex})");
        return null;
    }

    private void WireSecondary(ISttProvider provider, string name)
    {
        var shortName = ShortName(name);
        provider.TranscriptionReceived += result =>
        {
            if (!result.IsFinal) return;
            var tagged = result with { ProviderName = name };
            var wc = result.Words?.Length ?? 0;
            var preview = result.Transcript.Length > 60 ? result.Transcript[..60] + "..." : result.Transcript;
            Logger.Trace($"Ensemble [{shortName}] final: {wc} words, \"{preview}\"");
            if (result.Words != null)
            {
                var wordDetails = string.Join(", ", result.Words.Select(w =>
                    $"{w.Text}({w.Confidence:F2})"));
                Logger.Trace($"Ensemble [{shortName}] words: {wordDetails}");
            }
            RawProviderFinalReceived?.Invoke(tagged);
            _merger?.SubmitResult(tagged);
        };
        provider.ErrorOccurred += err => Logger.Trace($"SttService [{shortName} secondary]: {err}");
        provider.ConnectionStateChanged += connected =>
        {
            if (!connected) Logger.Trace($"SttService: {name} secondary disconnected");
        };
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
                catch (Exception ex) { Logger.Trace($"SttService: {_s1Name} connect failed: {ex.Message}"); }
            });
            _ = Task.Run(async () =>
            {
                try { if (s2 is { IsConnected: false }) await s2.StartSessionAsync(); }
                catch (Exception ex) { Logger.Trace($"SttService: {_s2Name} connect failed: {ex.Message}"); }
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
            // End primary first (essential — triggers final transcript)
            try
            {
                if (_primaryProvider is { IsConnected: true })
                    await _primaryProvider.EndSessionAsync();
            }
            catch (Exception ex) { Logger.Trace($"SttService: Primary EndSession error: {ex.Message}"); }

            // End secondaries in background (non-blocking) — their cleanup isn't needed
            // before the next recording can start, and waiting up to 2s blocks the action thread.
            var s1 = _secondaryProvider1;
            var s2 = _secondaryProvider2;
            _ = Task.Run(async () =>
            {
                var t1 = Task.Run(async () =>
                {
                    try { if (s1 is { IsConnected: true }) await s1.EndSessionAsync(); } catch { }
                });
                var t2 = Task.Run(async () =>
                {
                    try { if (s2 is { IsConnected: true }) await s2.EndSessionAsync(); } catch { }
                });
                await Task.WhenAny(Task.WhenAll(t1, t2), Task.Delay(2000));
            });
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

        // Pre-connect all providers in parallel for next PTT press.
        // Sequential pre-connect caused AAI's 1500ms rate-limit delay to block SM connect.
        if (_ensembleMode)
        {
            var p = _primaryProvider;
            var s1 = _secondaryProvider1;
            var s2 = _secondaryProvider2;
            _ = Task.Run(async () =>
            {
                try { if (p is { IsConnected: false }) await p.StartSessionAsync(); } catch { }
            });
            _ = Task.Run(async () =>
            {
                try { if (s1 is { IsConnected: false }) await s1.StartSessionAsync(); } catch { }
            });
            _ = Task.Run(async () =>
            {
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
            "soniox" => !string.IsNullOrEmpty(_config.SttSonioxApiKey),
            "elevenlabs" => !string.IsNullOrEmpty(_config.SttElevenLabsApiKey),
            "smallestai" => !string.IsNullOrEmpty(_config.SttSmallestAiApiKey),
            "corti" => !string.IsNullOrEmpty(_config.SttCortiClientId),
            _ => !string.IsNullOrEmpty(_config.SttApiKey)
        };
        if (!hasKey)
        {
            Logger.Trace("SttService: No API key configured");
            return null;
        }

        var keyterms = _keytermOverride ?? _config.SttDeepgramKeyterms;
        var punctuate = EffectivePunctuate;
        return _config.SttProvider switch
        {
            "deepgram" => new DeepgramProvider(_config.SttApiKey, _config.SttModel, punctuate, keyterms),
            "assemblyai" => new AssemblyAIProvider(_config.SttAssemblyAIApiKey, punctuate, keyterms),
            "soniox" => new SonioxProvider(_config.SttSonioxApiKey, punctuate, keyterms),
            "elevenlabs" => new ElevenLabsProvider(_config.SttElevenLabsApiKey, punctuate),
            "smallestai" => new SmallestAiProvider(_config.SttSmallestAiApiKey, punctuate, keyterms),
            "corti" => new CortiProvider(_config.SttCortiClientId, _config.SttCortiClientSecret, _config.SttCortiEnvironment, punctuate),
            "speechmatics" => new SpeechmaticsProvider(_config.SttSpeechmaticsApiKey, _config.SttSpeechmaticsRegion, punctuate, keyterms),
            _ => new DeepgramProvider(_config.SttApiKey, _config.SttModel, punctuate, keyterms)
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
            // Flush study stats to all-time on shutdown (crash safety)
            _merger?.FlushAlltimeStats();

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
