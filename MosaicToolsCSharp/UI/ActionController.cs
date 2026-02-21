using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MosaicTools.Services;
using MosaicTools.UI;

namespace MosaicTools.UI;

/// <summary>
/// Central action controller that coordinates all services.
/// Matches Python's action dispatch in MosaicToolsApp.
/// </summary>
public struct ActionRequest
{
    public string Action { get; set; }
    public string Source { get; set; }
}

public class ActionController : IDisposable
{
    private readonly Configuration _config;
    private readonly MainForm _mainForm;
    
    // Services
    private readonly HidService _hidService;
    private readonly KeyboardService _keyboardService;
    private readonly AutomationService _automationService;
    private readonly IMosaicReader _mosaicReader;
    private readonly IMosaicCommander _mosaicCommander;
    private NoteFormatter _noteFormatter;
    private GetPriorService _getPriorService;
    private readonly OcrService _ocrService;
    private readonly PipeService _pipeService;
    private readonly TemplateDatabase _templateDatabase;
    private readonly AidocService _aidocService;
    private readonly RadAiService? _radAiService;  // [RadAI] — null if RadAI not installed
    private readonly RecoMdService _recoMdService;
    private SttService? _sttService;  // [CustomSTT] — null if Custom STT not enabled
    private bool _sttDirectPasteActive; // [CustomSTT] Track whether direct paste is active
    private IntPtr _sttPasteTargetWindow; // [CustomSTT] Window that had focus when recording started
    private readonly object _directPasteLock = new(); // [CustomSTT] Serialize direct paste operations

    public PipeService PipeService => _pipeService;

    /// <summary>
    /// Safe BeginInvoke that silently no-ops if the MainForm handle isn't created.
    /// Prevents crash loops when timers fire during form creation/destruction.
    /// </summary>
    private void InvokeUI(Action action)
    {
        if (_mainForm.IsHandleCreated && !_mainForm.IsDisposed)
            _mainForm.BeginInvoke(action);
    }

    // Batched UI updates: collects multiple UI actions and executes them in a single
    // BeginInvoke call. This reduces WM_USER message spam on the UI thread, preventing
    // the low-level keyboard hook from being silently removed by Windows when the
    // message pump can't service hook callbacks fast enough (LowLevelHooksTimeout).
    private readonly List<Action> _uiBatch = new();

    private void BatchUI(Action action) => _uiBatch.Add(action);

    private void FlushUI()
    {
        if (_uiBatch.Count == 0) return;
        var actions = _uiBatch.ToArray();
        _uiBatch.Clear();
        InvokeUI(() => { foreach (var a in actions) a(); });
    }

    // Action Queue (Must be STA for Clipboard and SendKeys)
    private readonly ConcurrentQueue<ActionRequest> _actionQueue = new();
    private readonly AutoResetEvent _actionEvent = new(false);
    private readonly Thread _actionThread;
    
    // State
    private volatile bool _dictationActive = false;
    private volatile bool _isUserActive = false;
    private int _scrapeRunning = 0; // Reentrancy guard for scrape timer
    private long _scrapeStartedTicks; // Watchdog: when current scrape started (Interlocked)
    private volatile bool _stopThread = false;

    // Impression search state
    private bool _searchingForImpression = false;
    private bool _impressionFromProcessReport = false; // True if opened by Process Report (stays until Sign)
    private DateTime _impressionSearchStartTime; // When we started searching - used for initial delay

    // Impression delete state
    private string? _pendingImpressionReplaceText;
    private volatile bool _impressionDeletePending;
    private int NormalScrapeIntervalMs => _config.ScrapeIntervalSeconds * 1000;
    private int _fastScrapeIntervalMs = 1000;
    private int _studyLoadScrapeIntervalMs = 500;
    private const int IdleScrapeIntervalMs = 10_000; // Slow polling when no study is open
    private const int DeepIdleScrapeIntervalMs = 30_000; // Very slow polling after extended idle (reduces UIA load)
    private int _consecutiveIdleScrapes = 0; // Counts scrapes with no report content
    private int _scrapesSinceLastGc = 0; // Counter for periodic GC to release FlaUI COM wrappers
    private int _scrapeHeartbeatCount = 0; // Heartbeat counter for diagnostic logging

    // PTT (Push-to-talk) state (int for Interlocked atomicity)
    private int _pttBusy = 0;
    private DateTime _lastSyncTime = DateTime.Now;
    private readonly System.Threading.Timer? _syncTimer;
    private System.Threading.Timer? _scrapeTimer;
    private DateTime _lastManualToggleTime = DateTime.MinValue;
    private int _consecutiveInactiveCount = 0; // For indicator debounce

    // Accession tracking - only track non-empty accessions
    private string? _lastNonEmptyAccession;

    // Accession flap debounce - prevent false "study closed" events
    // When accession goes empty, we defer the close and wait for it to stabilize
    private volatile string? _pendingCloseAccession;   // The accession that went empty
    private int _pendingCloseTickCount;        // How many scrape ticks it's been empty

    /// <summary>
    /// Returns true if a study is currently open (has a non-empty accession).
    /// </summary>
    public bool IsStudyOpen => !string.IsNullOrEmpty(_lastNonEmptyAccession);

    // RVUCounter integration - track whether current accession was signed
    private bool _currentAccessionSigned = false;

    // Report changes tracking - baseline report for diff highlighting
    private string? _baselineReport;
    private bool _processReportPressedForCurrentAccession = false;
    private bool _draftedAutoProcessDetected = false; // Set true when drafted study's report changes from baseline
    private bool _autoShowReportDoneForAccession = false; // Only auto-show report overlay once per accession
    private bool _needsBaselineCapture = false; // Set true on new study, cleared when baseline captured
    private bool _baselineIsFromTemplateDb = false; // True when baseline came from template DB fallback (section-only diffing)
    private int _baselineCaptureAttempts = 0; // Tick counter for template DB fallback timing
    private bool _templateRecordedForStudy = false; // Only record template once per study, before Process Report
    private string? _lastPopupReportText; // Track what's currently displayed in popup for change detection
    private DateTime _staleSetTime; // When SetStaleState(true) was last triggered

    // RVUCounter integration - track if discard dialog was shown for current accession
    // If dialog was shown while on accession X, and then X changes, X was discarded
    private bool _discardDialogShownForCurrentAccession = false;

    // Pending macro insertion - wait for clinical history to be visible before inserting
    private string? _pendingMacroAccession;
    private string? _pendingMacroDescription;

    // Alert state tracking for alerts-only mode
    private bool _templateMismatchActive = false;
    private bool _genderMismatchActive = false;
    private bool _strokeDetectedActive = false;
    private bool _pendingClarioPriorityRetry = false;
    private bool _clinicalHistoryVisible = false;

    // Aidoc state tracking
    private string? _lastAidocFindings; // comma-joined relevant finding types
    private bool _lastAidocRelevant = false;
    private readonly HashSet<string> _aidocConfirmedNegative = new(StringComparer.OrdinalIgnoreCase); // "once negative, stay negative" per study

    // RecoMD state tracking — continuous send on every scrape tick
    private string? _recoMdOpenedForAccession; // accession currently opened in RecoMD
    private string? _lastRecoMdSentText;       // last text sent (to detect changes for logging)

    // Critical note tracking - session-scoped set of accessions that already have critical notes
    private readonly ConcurrentDictionary<string, byte> _criticalNoteCreatedForAccessions = new();

    // Session-wide tracking to prevent duplicate pastes on study reopen
    // Use ConcurrentDictionary for thread safety (scrape timer reads, STA thread writes)
    private readonly ConcurrentDictionary<string, byte> _macrosInsertedForAccessions = new();
    private readonly ConcurrentDictionary<string, byte> _clinicalHistoryFixedForAccessions = new();

    // Ignore Inpatient Drafted - tracks completion of auto-insertions for Ctrl+A
    private bool _macrosCompleteForCurrentAccession = false;
    private bool _autoFixCompleteForCurrentAccession = false;
    private bool _ctrlASentForCurrentAccession = false;

    // Critical studies tracker - studies where critical notes were placed this session
    private readonly List<CriticalStudyEntry> _criticalStudies = new();
    private readonly object _criticalStudiesLock = new();

    /// <summary>
    /// List of studies where critical notes were placed this session.
    /// </summary>
    public IReadOnlyList<CriticalStudyEntry> CriticalStudies
    {
        get { lock (_criticalStudiesLock) return _criticalStudies.ToList(); }
    }

    /// <summary>
    /// Automation service for UI automation tasks (exposed for critical studies popup).
    /// </summary>
    public AutomationService Automation => _automationService;

    /// <summary>
    /// Event raised when the critical studies list changes.
    /// </summary>
    public event Action? CriticalStudiesChanged;

    // Window/Level cycle state for InteleViewer
    private int _windowLevelCycleIndex = 0;

    // Shared paste lock to prevent concurrent clipboard operations
    public static readonly object PasteLock = new();
    public static DateTime LastPasteTime = DateTime.MinValue;

    // Guard to skip duplicate RecoMD paste actions queued by impatient button presses
    private volatile bool _recoMdPasteInProgress;
    // RecoMD send throttle — only send every Nth scrape tick to reduce chatter
    private int _recoMdSendTickCounter;
    
    public ActionController(Configuration config, MainForm mainForm, RecoMdService recoMdService)
    {
        _config = config;
        _mainForm = mainForm;
        _recoMdService = recoMdService;

        _hidService = new HidService();
        _keyboardService = new KeyboardService();
        _keyboardService.UiSyncTarget = mainForm;
        _automationService = new AutomationService();
        _mosaicReader = _automationService;
        _mosaicCommander = _automationService;
        _noteFormatter = new NoteFormatter(config.DoctorName, config.CriticalFindingsTemplate, config.TargetTimezone);
        _getPriorService = new GetPriorService(_config.ComparisonTemplate);
        _ocrService = new OcrService();
        _pipeService = new PipeService();
        _pipeService.Start();
        _templateDatabase = new TemplateDatabase();
        _aidocService = new AidocService(_automationService.Automation);
        _radAiService = RadAiService.TryCreate();  // [RadAI]

        _actionThread = new Thread(ActionLoop) { IsBackground = true };
        _actionThread.SetApartmentState(ApartmentState.STA);
        _actionThread.Start();

        // 500ms heartbeat for registry sync (matches MicIndicator polling rate)
        _syncTimer = new System.Threading.Timer(OnSyncTimerCallback, null, 500, 500);
    }

    private void ActionLoop()
    {
        while (!_stopThread)
        {
            _actionEvent.WaitOne(500);
            if (_stopThread) break;

            while (_actionQueue.TryDequeue(out var req))
            {
                _isUserActive = true;
                
                // Save previous focus before activating Mosaic
                NativeWindows.SavePreviousFocus();
                
                try
                {
                    ExecuteAction(req);
                }
                catch (Exception ex)
                {
                    Logger.Trace($"Action loop error ({req.Action}): {ex.Message}");
                    InvokeUI(() => _mainForm.ShowStatusToast($"Error: {ex.Message}"));
                }
                finally
                {
                    _isUserActive = false;
                    
                    // Restore focus after action completes
                    NativeWindows.RestorePreviousFocus(50);
                    
                    InvokeUI(() => _mainForm.EnsureWindowsOnTop());
                }
            }
        }
    }
    
    public void Start()
    {
        // Wire up HID events (always active, including headless)
        _hidService.ButtonPressed += OnMicButtonPressed;
        _hidService.RecordButtonStateChanged += OnRecordButtonStateChanged;
        _hidService.DeviceConnected += msg =>
            InvokeUI(() => _mainForm.ShowStatusToast(msg));

        _hidService.SetPreferredDevice(_config.PreferredMicrophone);
        _hidService.Start();

        // Register hotkeys
        RegisterHotkeys();
        _keyboardService.Start();

        // Start background dictation sync (skip in headless mode)
        if (!App.IsHeadless && _config.IndicatorEnabled)
        {
            StartDictationSync();
        }
        
        // Start Mosaic scrape timer (always on)
        StartMosaicScrapeTimer();

        // Load template database and prune stale entries
        if (_config.TemplateDatabaseEnabled)
        {
            _templateDatabase.Load();
            _templateDatabase.Cleanup();
        }

        // [CustomSTT] Initialize custom STT service
        if (_config.CustomSttEnabled)
        {
            InitializeSttService();
        }

        Logger.Trace($"ActionController started (Headless={App.IsHeadless})");
    }
    
    public void Stop()
    {
        _hidService.Stop();
        _keyboardService.Stop();
    }

    public void RefreshServices()
    {
        Logger.Trace("Refreshing ActionController services...");
        // Recreate services that depend on config
        _noteFormatter = new NoteFormatter(_config.DoctorName, _config.CriticalFindingsTemplate, _config.TargetTimezone);
        _getPriorService = new GetPriorService(_config.ComparisonTemplate);
        _hidService.SetPreferredDevice(_config.PreferredMicrophone);
        RegisterHotkeys();
        _keyboardService.Start(); // ensure hook is running (idempotent)

        // Restart scraper to pick up any interval changes
        ToggleMosaicScraper(true);

        // [CustomSTT] Re-initialize STT service on settings change
        // Always reinitialize when enabled — provider, model, or key may have changed
        if (_config.CustomSttEnabled)
        {
            if (_sttService != null)
            {
                try { _sttService.Dispose(); } catch { } // Dispose handles disconnect internally
                _sttService = null;
                Logger.Trace("SttService: Disposed old instance for reinit");
            }
            InitializeSttService();
        }
        else if (_sttService != null)
        {
            try { _sttService.Dispose(); } catch { }
            _sttService = null;
            Logger.Trace("SttService: Disabled and disposed");
        }
    }

    // [CustomSTT] Initialize the custom STT service
    private void InitializeSttService()
    {
        _sttService = new SttService(_config);
        var error = _sttService.Initialize();
        if (error != null)
        {
            InvokeUI(() => _mainForm.ShowStatusToast(error));
            Logger.Trace($"SttService: Init error: {error}");
            _sttService.Dispose();
            _sttService = null;
            return;
        }

        _sttService.TranscriptionReceived += result =>
        {
            // Always update TranscriptionForm
            InvokeUI(() => _mainForm.OnSttTranscriptionReceived(result));

            // [CustomSTT] Direct paste: final results go straight into Mosaic's transcript box.
            // Must run on STA thread (clipboard requires it). Use a dedicated STA thread that
            // stays alive through the entire activate → focus → paste sequence.
            if (result.IsFinal && !string.IsNullOrEmpty(result.Transcript) && _sttDirectPasteActive)
            {
                // Client-side spoken punctuation → symbols (replaces Deepgram dictation mode
                // so we can keep "colon" as a word for medical context).
                var transcript = ApplySpokenPunctuation(result.Transcript).Trim();
                transcript = ApplyRadiologyCleanup(transcript);
                if (transcript.Length > 0 && !(transcript.Length > 1 && char.IsDigit(transcript[1])))
                    transcript = char.ToLower(transcript[0]) + transcript[1..];
                var textToPaste = " " + transcript;
                var svc = _automationService;
                var t = new Thread(() =>
                {
                    lock (_directPasteLock)
                    {
                        NativeWindows.ActivateMosaicForcefully();
                        Thread.Sleep(100);
                        // FocusTranscriptBox disabled: Mosaic has two text boxes (transcript
                        // and final report). Skipping explicit focus lets the paste go to
                        // whichever box the user last clicked, which is the desired behavior.
                        // Re-enable if we ever need to force transcript-only pasting again.
                        // svc.FocusTranscriptBox();
                        // Thread.Sleep(100);

                        InsertTextToFocusedEditor(textToPaste);
                        Thread.Sleep(50);
                        Logger.Trace($"CustomSTT: Direct paste ({textToPaste.Length} chars): \"{(textToPaste.Length > 40 ? textToPaste[..40] + "..." : textToPaste)}\"");

                        // Clear the indicator immediately after successful paste
                        InvokeUI(() => _mainForm.ClearTranscriptionForm());

                        // Restore focus to the app the user was in before paste
                        var restoreHwnd = _sttPasteTargetWindow;
                        if (restoreHwnd != IntPtr.Zero && NativeWindows.IsWindow(restoreHwnd))
                        {
                            Thread.Sleep(50);
                            NativeWindows.ActivateWindow(restoreHwnd, 500);
                        }
                    }
                });
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
            }
        };
        _sttService.RecordingStateChanged += recording =>
            InvokeUI(() => _mainForm.UpdateSttRecordingState(recording));
        _sttService.StatusChanged += status =>
            InvokeUI(() => _mainForm.UpdateSttStatus(status));
        _sttService.ErrorOccurred += error =>
            InvokeUI(() => _mainForm.ShowStatusToast(error));

        Logger.Trace("SttService: Initialized successfully");
    }

    public void RefreshFloatingToolbar() =>
        InvokeUI(() => _mainForm.RefreshFloatingToolbar(_config.FloatingButtons));

    /// <summary>
    /// Get the name of the currently connected microphone, or null if not connected.
    /// </summary>
    public string? GetConnectedMicrophoneName() => _hidService.ConnectedDeviceName;
    
    private void RegisterHotkeys()
    {
        _keyboardService.ClearHotkeys();
        
        foreach (var (action, mapping) in _config.ActionMappings)
        {
            if (!string.IsNullOrEmpty(mapping.Hotkey))
            {
                var actionName = action; // Capture for closure

                // CycleWindowLevel uses direct hotkey for instant response (no ThreadPool)
                if (actionName == Actions.CycleWindowLevel)
                {
                    _keyboardService.RegisterDirectHotkey(mapping.Hotkey, () => PerformCycleWindowLevel());
                }
                else
                {
                    _keyboardService.RegisterHotkey(mapping.Hotkey, () => TriggerAction(actionName, "Hotkey"));
                }
            }
        }
    }
    
    /// <summary>
    /// Get the action mappings for the currently connected microphone.
    /// </summary>
    private Dictionary<string, ActionMapping> GetCurrentMicMappings()
    {
        return _hidService.IsPhilipsConnected ? _config.SpeechMikeActionMappings : _config.ActionMappings;
    }

    private void OnMicButtonPressed(string button)
    {
        // If PTT is on or Custom STT is on, the Record Button is handled by OnRecordButtonStateChanged
        // and should not trigger its mapped action (usually System Beep or Toggle)
        // "Record Button" = PowerMic, "Record" = SpeechMike
        if ((_config.DeadManSwitch || _config.CustomSttEnabled) && (button == "Record Button" || button == "Record"))  // [CustomSTT]
        {
            return;
        }

        // Find action for this button using the correct mappings for connected device
        var mappings = GetCurrentMicMappings();
        foreach (var (action, mapping) in mappings)
        {
            if (mapping.MicButton == button)
            {
                // Skip CycleWindowLevel in headless mode (requires InteleViewer integration)
                if (action == Actions.CycleWindowLevel && App.IsHeadless)
                {
                    return;
                }

                // CycleWindowLevel bypasses the queue for instant response
                if (action == Actions.CycleWindowLevel)
                {
                    PerformCycleWindowLevel();
                }
                else
                {
                    TriggerAction(action, button);
                }
                return;
            }
        }
    }
    
    private void OnRecordButtonStateChanged(bool isDown)
    {
        // [CustomSTT] When Custom STT is enabled without PTT, toggle recording on press
        if (_config.CustomSttEnabled && _sttService != null && !_config.DeadManSwitch)
        {
            if (isDown)
            {
                ThreadPool.QueueUserWorkItem(_ => PerformToggleRecord(null, sendKey: true));
            }
            return;
        }

        // 1. Dead Man's Switch (Push-to-Talk) Active Logic
        if (_config.DeadManSwitch)
        {
            if (isDown)
            {
                if (Interlocked.CompareExchange(ref _pttBusy, 1, 0) == 0 && !_dictationActive)
                {
                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        try { PerformToggleRecord(true, sendKey: true); }
                        finally { Interlocked.Exchange(ref _pttBusy, 0); }
                    });
                }
                else
                {
                    Interlocked.CompareExchange(ref _pttBusy, 0, 1); // Reset if condition not met
                }
            }
            else
            {
                if (Interlocked.CompareExchange(ref _pttBusy, 1, 0) == 0 && _dictationActive)
                {
                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        try {
                            Thread.Sleep(50);
                            PerformToggleRecord(false, sendKey: true);
                        }
                        finally { Interlocked.Exchange(ref _pttBusy, 0); }
                    });
                }
                else
                {
                    Interlocked.CompareExchange(ref _pttBusy, 0, 1); // Reset if condition not met
                }
            }
            return;
        }
        
        // 2. Passive Monitoring - Disabled to prevent conflict with mapped actions.
        // State sync is handled by manual actions or the background reality check.
    }

    
    public void TriggerAction(string action, string source = "Manual")
    {
        Logger.Trace($"TriggerAction Queued: {action} (Source: {source})");

        // Wake up scrape timer if it's in deep idle — user activity means a study may be open soon
        if (_consecutiveIdleScrapes >= 10)
        {
            _consecutiveIdleScrapes = 0;
            RestartScrapeTimer(NormalScrapeIntervalMs);
            Logger.Trace("Scrape wake-up: user action, resuming normal interval");
        }

        _actionQueue.Enqueue(new ActionRequest { Action = action, Source = source });
        _actionEvent.Set();
    }

    private void ExecuteAction(ActionRequest req)
    {
        Logger.Trace($"Executing Action: {req.Action} (Source: {req.Source})");

        // GLOBAL FIX: If triggered by a Hotkey, release physically held modifiers
        // to prevent interference with emitted keystrokes.
        // Skip for CycleWindowLevel - needs to be instant, handles its own modifiers
        if (req.Source == "Hotkey" && req.Action != Actions.CycleWindowLevel)
        {
            NativeWindows.KeyUpModifiers();
            Thread.Sleep(50);
        }

        switch (req.Action)
        {
            case Actions.SystemBeep:
                PerformBeep();
                break;
            case Actions.GetPrior:
                PerformGetPrior();
                break;
            case Actions.CriticalFindings:
                // Win key held = debug mode
                bool debugMode = _keyboardService.IsWinKeyHeld();
                PerformCriticalFindings(debugMode);
                break;
            case Actions.ShowReport:
                PerformShowReport();
                break;
            case Actions.CaptureSeries:
                PerformCaptureSeries();
                break;
            case Actions.ToggleRecord:
                PerformToggleRecord(null, sendKey: true);
                break;
            case Actions.ProcessReport:
                PerformProcessReport(req.Source);
                break;
            case Actions.SignReport:
                PerformSignReport(req.Source);
                break;
            case Actions.CreateImpression:
                // Skip Forward (PowerMic) and -i- (SpeechMike) are hardcoded in Mosaic to trigger Create Impression
                // Only invoke via UI automation if triggered by other sources
                if (req.Source != "Skip Forward" && req.Source != "-i-")
                {
                    PerformCreateImpression();
                }
                break;
            case Actions.DiscardStudy:
                PerformDiscardStudy();
                break;
            case Actions.ShowPickLists:
                PerformShowPickLists();
                break;
            case Actions.CycleWindowLevel:
                PerformCycleWindowLevel();
                break;
            case Actions.CreateCriticalNote:
                PerformCreateCriticalNote();
                break;
            case Actions.RadAiImpression:  // [RadAI]
                PerformRadAiImpression();
                break;
            case Actions.TriggerRecoMd:
                PerformRecoMd();
                break;
            case Actions.PasteRecoMd:
                PerformPasteRecoMd();
                break;
            case "__InsertMacros__":
                PerformInsertMacros();
                break;
            case "__InsertPickListText__":
                PerformInsertPickListText();
                break;
            case "__RadAiInsert__":  // [RadAI]
                PerformRadAiInsert();
                break;
            case "__ReplaceImpression__":
                PerformReplaceImpression();
                break;
        }
    }
    
    #region Action Implementations
    
    private volatile int _syncCheckToken = 0;
    private string PrepareTextForPaste(string text)
    {
        if (_config.SeparatePastedItems && !text.StartsWith("\n"))
            return "\n" + text;
        return text;
    }

    private void InsertTextToFocusedEditor(string text)
    {
        if (_config.ExperimentalUseSendInputInsert)
        {
            var ok = NativeWindows.SendUnicodeText(text);
            if (!ok)
                Logger.Trace($"InsertTextToFocusedEditor: SendInput failed ({text.Length} chars)");
            return;
        }

        ClipboardService.SetText(text);
        Thread.Sleep(50);
        NativeWindows.SendHotkey("ctrl+v");
    }

    public bool IsAddendumOpen()
    {
        // Primary: check the flag set during ProseMirror scanning (works even when
        // LastFinalReport is stale due to the U+FFFC fallback logic)
        if (_mosaicReader.IsAddendumDetected)
            return true;

        // Fallback: check LastFinalReport text directly
        var report = _mosaicReader.LastFinalReport;
        return !string.IsNullOrEmpty(report) && report.TrimStart().StartsWith("Addendum", StringComparison.OrdinalIgnoreCase);
    }

    private void PerformBeep()
    {
        // Protect from Registry Sync overwriting state during toggle
        _lastManualToggleTime = DateTime.Now;

        // 1. Determine STARTING state using registry truth when available
        bool? realState = NativeWindows.IsMicrophoneActiveFromRegistry();
        bool currentState = realState ?? _dictationActive;
        bool startingActive = !currentState;

        // Phantom detection: if toggling from ON→OFF, wait briefly and verify mic actually stopped
        if (currentState == true && !startingActive)
        {
            Thread.Sleep(300);
            bool? verified = NativeWindows.IsMicrophoneActiveFromRegistry();
            if (verified == true)
            {
                // Mic still active — this was a phantom button press. Suppress.
                Logger.Trace("PerformBeep: Phantom detected (mic still active after 300ms). Suppressing.");
                return;
            }
        }

        // 2. Play Feedback Beep FIRST for snappiness (Python matches this order)
        bool shouldPlay = startingActive ? _config.StartBeepEnabled : _config.StopBeepEnabled;
        if (shouldPlay)
        {
            int freq = startingActive ? 1000 : 500;
            double vol = startingActive ? _config.StartBeepVolume : _config.StopBeepVolume;
            
            if (startingActive)
            {
                // Start beep: Delayed by dictation_pause_ms
                int startDelay = Math.Max(0, _config.DictationPauseMs);
                Logger.Trace($"Beep: START (High Pitch, {startDelay}ms delay)");
                ThreadPool.QueueUserWorkItem(_ => {
                    Thread.Sleep(startDelay);
                    AudioService.PlayBeep(1000, 200, vol);
                });
            }
            else
            {
                // Stop beep: Instant
                Logger.Trace("Beep: STOP (Low Pitch, Instant)");
                AudioService.PlayBeepAsync(500, 200, vol);
            }
        }

        // 3. Update internal state and UI
        _dictationActive = startingActive;
        InvokeUI(() => _mainForm.UpdateIndicatorState(_dictationActive));
        Logger.Trace($"System Beep: State toggled to {(_dictationActive ? "ON" : "OFF")}");

        // 4. Reality Check (Python Parity)
        // We only correct if reality says TRUE but we are FALSE.
        // We wait for the system to settle before checking.
        int currentToken = Interlocked.Increment(ref _syncCheckToken);
        int syncDelayMs = Math.Max(1500, (int)(_config.DictationPauseMs * 2.5));

        ThreadPool.QueueUserWorkItem(_ =>
        {
            Thread.Sleep(syncDelayMs);
            if (_syncCheckToken != currentToken) return;

            // Check registry + UIA fallback
            bool? isRealActive = NativeWindows.IsMicrophoneActiveFromRegistry();
            if (isRealActive == null) isRealActive = NativeWindows.IsMicrophoneActiveFromUia();

            if (isRealActive == true && !_dictationActive)
            {
                Logger.Trace("SYNC FIX: System is recording, but app state was OFF. Syncing to ON.");
                _dictationActive = true;
                InvokeUI(() => _mainForm.UpdateIndicatorState(_dictationActive));
            }
            else if (isRealActive == false && _dictationActive)
            {
                // Optionally sync to OFF if it's been a while? Python doesn't do this to avoid false negatives.
                Logger.Trace("SYNC INFO: Registry says OFF, but app state is ON. (Python: Ignoring to avoid False Negatives)");
            }
        });
    }
    
    private void OnSyncTimerCallback(object? state)
    {
        try
        {
            bool? registryActive = NativeWindows.IsMicrophoneActiveFromRegistry();
            if (!registryActive.HasValue) return;

            bool active = registryActive.Value;

            // 1. Instant ON: If it's active, reset count and update UI immediately
            if (active)
            {
                _consecutiveInactiveCount = 0;
                InvokeUI(() =>
                {
                    // Safety net: if user is dictating but indicator was hidden (e.g. scrape
                    // failed to detect study re-open), force-show it.
                    if (_config.IndicatorEnabled && _config.HideIndicatorWhenNoStudy)
                        _mainForm.EnsureIndicatorVisible();
                    _mainForm.UpdateIndicatorState(true);
                });
            }
            else
            {
                // 2. Sticky OFF: Require 3 consecutive inactive checks (~750ms) to turn off
                _consecutiveInactiveCount++;
                if (_consecutiveInactiveCount >= 3)
                {
                    InvokeUI(() => _mainForm.UpdateIndicatorState(false));
                }
            }

            // 3. Update internal logical state (with 500ms lockout for manual toggles)
            if ((DateTime.Now - _lastManualToggleTime).TotalMilliseconds > 500)
            {
                if (active != _dictationActive)
                {
                    // Only sync logical state once the debounce has settled for OFF
                    if (active || _consecutiveInactiveCount >= 3)
                    {
                        Logger.Trace($"Registry Sync: Internal state updated to {active}");
                        _dictationActive = active;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"SyncTimer error: {ex.Message}");
        }
    }

    private void PerformToggleRecord(bool? desiredState = null, bool sendKey = true)
    {
        Logger.Trace($"PerformToggleRecord (desired={desiredState}, sendKey={sendKey})");

        _lastManualToggleTime = DateTime.Now;

        // [CustomSTT] Route to SttService when Custom STT is enabled
        if (_config.CustomSttEnabled && _sttService != null)
        {
            PerformToggleRecordStt(desiredState);
            return;
        }

        // 1. Python-style early exit: If we specify a state and are already there, just beeps
        if (desiredState.HasValue)
        {
            bool currentReal = NativeWindows.IsMicrophoneActiveFromRegistry() ?? _dictationActive;
            if (currentReal == desiredState.Value)
            {
                Logger.Trace($"ToggleRecord: Already in desired state ({desiredState}). skipping key, playing beep.");
                
                bool shouldBeep = desiredState.Value ? _config.StartBeepEnabled : _config.StopBeepEnabled;
                if (shouldBeep)
                {
                    int freq = desiredState.Value ? 1000 : 500;
                    double vol = desiredState.Value ? _config.StartBeepVolume : _config.StopBeepVolume;
                    
                    // Python uses 200ms for these beeps
                    if (desiredState.Value) // START
                    {
                        int delay = Math.Max(0, _config.DictationPauseMs);
                        ThreadPool.QueueUserWorkItem(_ => {
                            Thread.Sleep(delay);
                            AudioService.PlayBeep(freq, 200, vol);
                        });
                    }
                    else // STOP
                    {
                        AudioService.PlayBeepAsync(freq, 200, vol);
                    }
                }
                
                _dictationActive = currentReal;
                return;
            }
        }

        // 2. Perform Toggle (Alt+R) matching Python's activation rules
        if (sendKey)
        {
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);
            NativeWindows.SendAltKey('R');
        }
        
        // 3. Update internal state (Syncing with Python's predictive model)
        if (desiredState.HasValue)
        {
            _dictationActive = desiredState.Value;
        }
        else
        {
            // Python: self.dictation_active = not self.dictation_active
            bool currentReal = NativeWindows.IsMicrophoneActiveFromRegistry() ?? _dictationActive;
            _dictationActive = !currentReal;
        }
        
        // Removed optimistic UI update: light now only turns on when registry confirms recording state.

        // 4. Play beep feedback strictly matching Python
        bool shouldPlay = _dictationActive ? _config.StartBeepEnabled : _config.StopBeepEnabled;
        if (shouldPlay)
        {
            int freq = _dictationActive ? 1000 : 500;
            double vol = _dictationActive ? _config.StartBeepVolume : _config.StopBeepVolume;
            
            if (_dictationActive)
            {
                // Start beep: Delayed by dictation_pause_ms
                int delay = Math.Max(0, _config.DictationPauseMs);
                ThreadPool.QueueUserWorkItem(_ => {
                    Thread.Sleep(delay);
                    AudioService.PlayBeep(freq, 200, vol);
                });
            }
            else
            {
                // Stop beep: Instant
                AudioService.PlayBeepAsync(freq, 200, vol);
            }
        }

        Logger.Trace($"Dictation toggle initiated. State: {_dictationActive}");
    }

    // [CustomSTT] Toggle recording via SttService instead of Mosaic's built-in dictation
    private void PerformToggleRecordStt(bool? desiredState)
    {
        bool isRecording = _sttService!.IsRecording;

        // If we have a desired state and are already there, just beep
        if (desiredState.HasValue && desiredState.Value == isRecording)
        {
            Logger.Trace($"CustomSTT: Already in desired state ({desiredState}). Playing beep.");
            bool shouldBeep = desiredState.Value ? _config.SttStartBeepEnabled : _config.SttStopBeepEnabled;
            if (shouldBeep)
            {
                int freq = desiredState.Value ? 1000 : 500;
                double vol = desiredState.Value ? _config.SttStartBeepVolume : _config.SttStopBeepVolume;
                AudioService.PlayBeepAsync(freq, 200, vol);
            }
            return;
        }

        bool newState = desiredState ?? !isRecording;

        if (newState)
        {
            // Start recording
            _sttPasteTargetWindow = NativeWindows.GetForegroundWindow(); // [CustomSTT] Save target window before showing TranscriptionForm
            _sttDirectPasteActive = true; // [CustomSTT] Enable direct paste to foreground window
            if (_config.SttShowIndicator) // [CustomSTT] Only show indicator if enabled in settings
                InvokeUI(() => _mainForm.ShowTranscriptionForm());
            Task.Run(async () => await _sttService!.StartRecordingAsync()).Wait(2000);
            _dictationActive = true;

            if (_config.SttStartBeepEnabled)
            {
                AudioService.PlayBeepAsync(1000, 200, _config.SttStartBeepVolume);
            }
        }
        else
        {
            // Stop recording — StopRecordingAsync now waits for finalize + disconnects
            Task.Run(async () => await _sttService!.StopRecordingAsync()).Wait(4000);
            _dictationActive = false;

            if (_config.SttStopBeepEnabled)
            {
                AudioService.PlayBeepAsync(500, 200, _config.SttStopBeepVolume);
            }

            // Small delay for any final paste to complete, then clean up UI
            Thread.Sleep(500);
            _sttDirectPasteActive = false;
            InvokeUI(() => _mainForm.HideTranscriptionForm());
        }

        Logger.Trace($"CustomSTT: Recording toggled to {newState}");
    }

    private void PerformProcessReport(string source = "Manual")
    {
        Logger.Trace($"Process Report (Source: {source})");

        // Mark that Process Report was pressed for this accession (for diff highlighting)
        _processReportPressedForCurrentAccession = true;

        // Safety: Release all modifiers before starting automated sequence
        NativeWindows.KeyUpModifiers();
        Thread.Sleep(50);

        // [CustomSTT] Custom STT path: text is already in Mosaic via direct paste. Just stop + Alt+P.
        if (_config.CustomSttEnabled && _sttService != null)
        {
            // Stop recording if active, wait for final transcripts to be pasted
            if (_sttService.IsRecording)
            {
                Logger.Trace("Process Report (CustomSTT): Stopping recording...");
                Task.Run(async () => await _sttService.StopRecordingAsync()).Wait(4000);
                _dictationActive = false;
                if (_config.SttStopBeepEnabled)
                    AudioService.PlayBeepAsync(500, 200, _config.SttStopBeepVolume);
            }

            // Small delay for any final paste to complete
            Thread.Sleep(500);
            _sttDirectPasteActive = false;

            // Hide the transcription indicator
            InvokeUI(() => _mainForm.HideTranscriptionForm());

            // Send Alt+P — text is already in Mosaic via direct paste
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);
            NativeWindows.SendAltKey('P');
        }
        else
        {
            // Standard (non-CustomSTT) path
            bool dictationWasActive = _dictationActive;

            // 1. Auto-stop dictation if enabled
            if (_config.AutoStopDictation && _dictationActive)
            {
                Logger.Trace("Process Report: Auto-stopping dictation...");
                PerformToggleRecord(false, sendKey: true); // Explicitly ensure we STOP
                Thread.Sleep(200);
            }

            // 2. Conditional Alt+P logic for hardcoded mic buttons
            // PowerMic: Skip Back is hardcoded to Process Report
            // SpeechMike: Ins/Ovr is hardcoded to Process Report
            bool isHardcodedProcessButton = (source == "Skip Back") || (source == "Ins/Ovr");
            var currentMappings = GetCurrentMicMappings();
            bool isButtonMappedToProcess =
                currentMappings.GetValueOrDefault(Actions.ProcessReport)?.MicButton == "Skip Back" ||
                currentMappings.GetValueOrDefault(Actions.ProcessReport)?.MicButton == "Ins/Ovr";

            if (isHardcodedProcessButton && isButtonMappedToProcess)
            {
                // If the hardware button is pressed, and it's mapped to Process Report
                if (dictationWasActive)
                {
                    // If dictation was ON, the hardware button might fail to process the report.
                    // We send it manually AFTER stopping dictation.
                    Logger.Trace($"Process Report: Dictation was ON + {source}. Sending Alt+P manually.");
                    NativeWindows.ActivateMosaicForcefully();
                    Thread.Sleep(100);
                    NativeWindows.SendAltKey('P');
                }
                else
                {
                    // Hardware handles it when dictation is OFF.
                    Logger.Trace($"Process Report: Dictation was OFF + {source}. Skipping redundant Alt+P.");
                }
            }
            else
            {
                // Standard behavior for Hotkeys, Toolbar, or non-hardcoded buttons
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(100);
                NativeWindows.SendAltKey('P');
            }
        }
        
        // Scroll down if enabled (3 rapid Page Down presses)
        // Scroll down if enabled (Smart Scroll)
        if (_config.ScrollToBottomOnProcess)
        {
            int pageDowns = 0;
            
            {
                string? report = _mosaicReader.LastFinalReport;
                if (!string.IsNullOrEmpty(report))
                {
                    int lines = report.Split('\n').Length;
                    
                    if (lines >= _config.ScrollThreshold3) pageDowns = 3;
                    else if (lines >= _config.ScrollThreshold2) pageDowns = 2;
                    else if (lines >= _config.ScrollThreshold1) pageDowns = 1;
                    
                    Logger.Trace($"Smart Scroll: Report has {lines} lines -> sending {pageDowns} PgDn(s).");
                    
                    if (_config.ShowLineCountToast)
                    {
                        string pgDnText = pageDowns == 1 ? "1 PgDn" : $"{pageDowns} PgDns";
                        _mainForm.ShowStatusToast($"Smart Scroll: {lines} lines -> {pgDnText}");
                    }
                }
                else
                {
                    Logger.Trace("Smart Scroll: No report scraped yet. Skipping scroll.");
                }
            }
            
            if (pageDowns > 0)
            {
                Thread.Sleep(50);
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(50);
                for (int i = 0; i < pageDowns; i++)
                {
                    NativeWindows.keybd_event(NativeWindows.VK_NEXT, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(10);
                    NativeWindows.keybd_event(NativeWindows.VK_NEXT, 0, NativeWindows.KEYEVENTF_KEYUP, UIntPtr.Zero);
                    Thread.Sleep(30);
                }
                Logger.Trace($"Process Report: Sent {pageDowns} Page Down keys to scroll down");
            }
        }

        // Start impression search if enabled
        if (_config.ShowImpression)
        {
            StartImpressionSearch();
        }

        // Auto-show report overlay if enabled (only first time per accession)
        // Skip if popup is already open (don't trigger toggle/cycle logic)
        if (_config.ShowReportAfterProcess && !_autoShowReportDoneForAccession)
        {
            _autoShowReportDoneForAccession = true;
            var popupRef = _currentReportPopup;
            bool popupAlreadyOpen = popupRef != null && !popupRef.IsDisposed && popupRef.Visible;
            if (popupAlreadyOpen)
            {
                Logger.Trace("Process Report: Skipping auto-show (popup already open), marking as stale");
                InvokeUI(() => { if (popupRef != null && !popupRef.IsDisposed) popupRef.SetStaleState(true); });
            }
            else
            {
                Logger.Trace("Process Report: Auto-showing report overlay (first time for accession)");
                PerformShowReport();
            }
        }
        else
        {
            var popupRef2 = _currentReportPopup;
            if (popupRef2 != null && !popupRef2.IsDisposed && popupRef2.Visible)
            {
                // Popup is open but auto-show is disabled or already done - still mark as stale during processing
                Logger.Trace("Process Report: Marking popup as stale during processing");
                InvokeUI(() => { if (popupRef2 != null && !popupRef2.IsDisposed) popupRef2.SetStaleState(true); });
            }
        }

        // Auto-create critical note for stroke cases if enabled
        if (_config.StrokeAutoCreateNote && _strokeDetectedActive)
        {
            var accession = _mosaicReader.LastAccession;
            if (!HasCriticalNoteForAccession(accession))
            {
                Logger.Trace($"Process Report: Auto-creating critical note for stroke case {accession}");
                CreateStrokeCriticalNote(accession);
            }
        }

        // Auto-send to RecoMD if enabled — kick-start immediately, then scrape timer continues
        if (_config.RecoMdEnabled && _config.RecoMdAutoOnProcess)
        {
            var immediateText = _mosaicReader.LastFinalReport;
            if (!string.IsNullOrEmpty(immediateText))
            {
                var recoText = RecoMdService.CleanReportText(immediateText);
                var recoAcc = _mosaicReader.LastAccession;
                var recoDesc = _mosaicReader.LastDescription;
                var recoName = _mosaicReader.LastPatientName;
                var recoGender = _mosaicReader.LastPatientGender;
                var recoMrn = _mosaicReader.LastMrn;
                var recoAge = _mosaicReader.LastPatientAge ?? 0;
                Logger.Trace("RecoMD: Immediate send after Process Report");
                _recoMdOpenedForAccession = recoAcc;
                _lastRecoMdSentText = recoText;
                Task.Run(async () =>
                {
                    if (string.IsNullOrEmpty(recoAcc)) return;
                    await _recoMdService.OpenReportAsync(recoAcc,
                        recoDesc, recoName, recoGender, recoMrn, recoAge);
                    await _recoMdService.SendReportTextAsync(recoAcc, recoText);
                });
            }
            BringRecoMdToFront();
        }

        // [RadAI] Auto-generate and insert RadAI impression on Process Report
        // Instead of polling in a Task.Run (which races with Mosaic's editor rebuild),
        // we set a flag and let the scrape timer trigger RadAI when the report stabilizes.
        // After Alt+P, Mosaic rebuilds the ProseMirror editor for 15-25 seconds during which
        // GetFinalReportFast can't update LastFinalReport (U+FFFC filtering). The scrape timer
        // naturally detects when the editor is ready again.
        if (_config.RadAiAutoOnProcess && _radAiService != null)
        {
            Logger.Trace("RadAI: Scheduling auto-insert after Process Report");
            _radAiPreProcessReport = _mosaicReader.LastFinalReport ?? "";
            _radAiAutoInsertRequestTime = DateTime.Now;
            _radAiAutoInsertPending = true;
        }

        // Popup will auto-update via scrape timer when report changes
    }

    private void StartImpressionSearch()
    {
        Logger.Trace("Starting impression search (fast scrape mode)...");
        _searchingForImpression = true;
        _impressionFromProcessReport = true; // Mark as manually triggered - stays until Sign Report
        _impressionSearchStartTime = DateTime.Now; // Track when we started - wait 2s before showing

        // Show the impression window with waiting message
        InvokeUI(() => _mainForm.ShowImpressionWindow());

        // Switch to fast scrape rate (1 second)
        RestartScrapeTimer(_fastScrapeIntervalMs);
    }

    private void OnImpressionFound(string impression)
    {
        Logger.Trace("Impression found! Switching to slow scrape mode.");
        _searchingForImpression = false;

        // Update the impression window with content
        InvokeUI(() => _mainForm.UpdateImpression(impression));

        // Revert to configured scrape rate
        RestartScrapeTimer(NormalScrapeIntervalMs);
    }

    private void RestartScrapeTimer(int intervalMs)
    {
        var timer = _scrapeTimer;
        try { timer?.Change(intervalMs, intervalMs); }
        catch (ObjectDisposedException) { return; }
        Logger.Trace($"Scrape timer interval changed to {intervalMs}ms");
    }

    private void PerformSignReport(string source = "Manual")
    {
        Logger.Trace($"Sign Report (Source: {source})");

        // Close report popup if open
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible)
        {
            InvokeUI(() =>
            {
                try { _currentReportPopup?.Close(); } catch { }
                _currentReportPopup = null;
                _lastPopupReportText = null;
            });
        }
        // [RadAI] Cancel pending auto-insert and close popup/overlay on sign
        _radAiAutoInsertPending = false;
        if (_currentRadAiPopup != null && !_currentRadAiPopup.IsDisposed)
        {
            _pendingRadAiImpressionItems = null;
            InvokeUI(() =>
            {
                try { _currentRadAiPopup?.Close(); } catch { }
                _currentRadAiPopup = null;
            });
        }
        if (_currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed)
        {
            _pendingRadAiImpressionItems = null;
            InvokeUI(() =>
            {
                try { _currentRadAiOverlay?.Close(); } catch { }
                _currentRadAiOverlay = null;
            });
        }

        // Mark current accession as signed for RVUCounter integration
        // This flag will be used when accession changes to send the appropriate notification
        _currentAccessionSigned = true;
        Logger.Trace($"RVUCounter: Marked accession '{_lastNonEmptyAccession}' as signed");

        // [CustomSTT] When Custom STT is enabled, always send Alt+F (Mosaic doesn't have the PowerMic)
        if (_config.CustomSttEnabled)
        {
            Logger.Trace("Sign Report (CustomSTT): Always sending Alt+F");
            NativeWindows.KeyUpModifiers();
            Thread.Sleep(50);
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);
            NativeWindows.SendAltKey('F');
        }
        else
        {
            // Check if Checkmark (PowerMic) or EoL (SpeechMike) button triggered this and is mapped to Sign Report
            // If so, Mosaic handles the actual signing - we only clean up impression state
            bool isHardcodedSignButton = (source == "Checkmark") || (source == "EoL");
            var signMappings = GetCurrentMicMappings();
            bool isButtonMappedToSign =
                signMappings.GetValueOrDefault(Actions.SignReport)?.MicButton == "Checkmark" ||
                signMappings.GetValueOrDefault(Actions.SignReport)?.MicButton == "EoL";

            if (isHardcodedSignButton && isButtonMappedToSign)
            {
                Logger.Trace($"Sign Report: {source} button - Mosaic handles signing, only cleaning up impression.");
            }
            else
            {
                // Standard behavior - send Alt+F
                NativeWindows.KeyUpModifiers();
                Thread.Sleep(50);

                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(100);
                NativeWindows.SendAltKey('F');
            }
        }

        // Close impression window on sign
        if (_config.ShowImpression)
        {
            _searchingForImpression = false;
            _impressionFromProcessReport = false; // Clear manual trigger flag
            InvokeUI(() => _mainForm.HideImpressionWindow());

            // Restore normal scrape rate
            RestartScrapeTimer(NormalScrapeIntervalMs);
        }

        // RecoMD: close report on sign
        if (_config.RecoMdEnabled)
        {
            _recoMdOpenedForAccession = null;
            _lastRecoMdSentText = null;
            SendRecoMdToBack();
            Task.Run(async () => await _recoMdService.CloseReportAsync());
        }
    }

    private void PerformDiscardStudy()
    {
        Logger.Trace("Discard Study action triggered");

        // Don't restore focus after discard - Mosaic should stay active.
        NativeWindows.ClearSavedFocus();

        // Close report popup if open
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible)
        {
            InvokeUI(() =>
            {
                try { _currentReportPopup?.Close(); } catch { }
                _currentReportPopup = null;
                _lastPopupReportText = null;
            });
        }
        // [RadAI] Close RadAI impression popup/overlay on discard
        if (_currentRadAiPopup != null && !_currentRadAiPopup.IsDisposed)
        {
            _pendingRadAiImpressionItems = null;
            InvokeUI(() =>
            {
                try { _currentRadAiPopup?.Close(); } catch { }
                _currentRadAiPopup = null;
            });
        }
        if (_currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed)
        {
            _pendingRadAiImpressionItems = null;
            InvokeUI(() =>
            {
                try { _currentRadAiOverlay?.Close(); } catch { }
                _currentRadAiOverlay = null;
            });
        }

        // Remember the accession we're discarding
        var accessionToDiscard = _lastNonEmptyAccession;

        // Mark that discard was explicitly requested via MosaicTools
        // This ensures CLOSED_UNSIGNED is sent even if the scrape loop handles it
        _discardDialogShownForCurrentAccession = true;

        // Perform the UI automation to discard
        bool success = _mosaicCommander.ClickDiscardStudy();

        if (success)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Study discarded", 2000));

            // Send CLOSED_UNSIGNED immediately after successful discard
            if (!string.IsNullOrEmpty(accessionToDiscard))
            {
                var hasCritical = HasCriticalNoteForAccession(accessionToDiscard);
                var msgType = hasCritical ? NativeWindows.MSG_STUDY_CLOSED_UNSIGNED_CRITICAL : NativeWindows.MSG_STUDY_CLOSED_UNSIGNED;
                Logger.Trace($"RVUCounter: Sending {(hasCritical ? "CLOSED_UNSIGNED_CRITICAL" : "CLOSED_UNSIGNED")} (discard action) for '{accessionToDiscard}'");
                NativeWindows.SendToRvuCounter(msgType, accessionToDiscard);
                _pipeService.SendStudyEvent(new StudyEventMessage("study_event", "unsigned", accessionToDiscard, hasCritical));

                // Clear state to prevent duplicate message from scrape loop
                _lastNonEmptyAccession = null;
                _discardDialogShownForCurrentAccession = false;
                _currentAccessionSigned = false;
            }

            // Hide clinical history window if configured to hide when no study
            if (_config.HideClinicalHistoryWhenNoStudy && _config.ShowClinicalHistory)
            {
                InvokeUI(() => _mainForm.ToggleClinicalHistory(false));
            }

            // Hide indicator window if configured to hide when no study
            if (_config.HideIndicatorWhenNoStudy && _config.IndicatorEnabled)
            {
                InvokeUI(() => _mainForm.ToggleIndicator(false));
            }
        }
        else
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Discard failed - try manually", 3000));
        }

        // Close impression window on discard (same as sign)
        if (_config.ShowImpression)
        {
            _searchingForImpression = false;
            _impressionFromProcessReport = false;
            InvokeUI(() => _mainForm.HideImpressionWindow());
        }

        // RecoMD: close report on discard
        if (_config.RecoMdEnabled)
        {
            _recoMdOpenedForAccession = null;
            _lastRecoMdSentText = null;
            SendRecoMdToBack();
            Task.Run(async () => await _recoMdService.CloseReportAsync());
        }
    }

    private void PerformCreateImpression()
    {
        Logger.Trace("Create Impression (via UI Automation)");

        var success = _mosaicCommander.ClickCreateImpression();
        if (success)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Create Impression", 1500));
        }
        else
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Create Impression button not found", 2500));
        }
    }

    private void InsertMacrosForStudy(string? studyDescription)
    {
        // Don't insert macros into addendums
        if (IsAddendumOpen())
        {
            Logger.Trace("Macros: Blocked - addendum is open");
            MarkMacrosCompleteForCurrentAccession();
            return;
        }

        // Session-wide check: don't re-insert macros for a study that was already processed
        var accession = _mosaicReader.LastAccession;
        if (!string.IsNullOrEmpty(accession) && _macrosInsertedForAccessions.ContainsKey(accession))
        {
            Logger.Trace($"Macros: Already inserted for accession {accession} this session, skipping");
            return;
        }

        // Find all matching enabled macros
        var matchingMacros = _config.Macros
            .Where(m => m.Enabled && m.MatchesStudy(studyDescription))
            .ToList();

        if (matchingMacros.Count == 0)
        {
            Logger.Trace($"Macros: No macros match study '{studyDescription}'");
            // Mark macros as complete even when none match (for Ignore Inpatient Drafted)
            MarkMacrosCompleteForCurrentAccession();
            return;
        }

        Logger.Trace($"Macros: {matchingMacros.Count} macro(s) match study '{studyDescription}'");

        // Build the text to paste
        var textBuilder = new System.Text.StringBuilder();

        // Add blank lines ONCE if enabled (for dictation space)
        // Use space character so Mosaic doesn't collapse empty lines
        if (_config.MacrosBlankLinesBefore)
        {
            for (int i = 0; i < 10; i++)
            {
                textBuilder.AppendLine(" ");
            }
        }

        // Append all matching macro texts
        foreach (var macro in matchingMacros)
        {
            if (!string.IsNullOrWhiteSpace(macro.Text))
            {
                textBuilder.Append(macro.Text);
                // Add a newline between macros if there are multiple
                if (matchingMacros.IndexOf(macro) < matchingMacros.Count - 1)
                {
                    textBuilder.AppendLine();
                }
            }
        }

        var textToPaste = textBuilder.ToString();
        if (string.IsNullOrWhiteSpace(textToPaste))
        {
            Logger.Trace("Macros: Matching macros have no text");
            return;
        }

        // Store the text BEFORE queuing the action (avoid race condition)
        _pendingMacroText = textToPaste;
        _pendingMacroCount = matchingMacros.Count;
        _pendingMacroInsertAccession = accession;

        // Queue the paste action (needs to run on STA thread)
        TriggerAction("__InsertMacros__", "Internal");
    }

    private volatile string? _pendingMacroText = null;
    private volatile int _pendingMacroCount = 0;
    private volatile string? _pendingMacroInsertAccession = null;

    private void PerformInsertMacros()
    {
        var text = _pendingMacroText;
        var count = _pendingMacroCount;
        var accessionToMark = _pendingMacroInsertAccession;
        _pendingMacroText = null;
        _pendingMacroCount = 0;
        _pendingMacroInsertAccession = null;

        if (string.IsNullOrEmpty(text)) return;

        if (IsAddendumOpen())
        {
            Logger.Trace("InsertMacros: Blocked - addendum is open");
            InvokeUI(() => _mainForm.ShowStatusToast("Cannot paste into addendum", 2500));
            return;
        }

        Logger.Trace($"InsertMacros: Pasting {count} macro(s)");

        // Use paste lock to prevent race conditions with clinical history auto-fix
        lock (PasteLock)
        {
            // Save focus
            var previousWindow = NativeWindows.GetForegroundWindow();

            try
            {
                // Activate Mosaic and paste
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(50);

                // Focus Transcript box to ensure paste goes to correct location
                _mosaicCommander.FocusTranscriptBox();
                Thread.Sleep(50);

                InsertTextToFocusedEditor(PrepareTextForPaste(text));
                Thread.Sleep(100); // Increased for reliability

                // Mark this accession as having had macros inserted (session-wide tracking)
                if (!string.IsNullOrEmpty(accessionToMark))
                {
                    _macrosInsertedForAccessions.TryAdd(accessionToMark, 0);
                    TrimTrackingSets();
                }

                InvokeUI(() => _mainForm.ShowStatusToast(
                    count == 1 ? "Macro inserted" : $"{count} macros inserted", 2000));

                // Mark macros complete for Ignore Inpatient Drafted feature
                MarkMacrosCompleteForCurrentAccession();
            }
            catch (Exception ex)
            {
                Logger.Trace($"InsertMacros error: {ex.Message}");
            }
            finally
            {
                LastPasteTime = DateTime.Now;

                // Restore focus
                if (previousWindow != IntPtr.Zero)
                {
                    Thread.Sleep(100);
                    NativeWindows.SetForegroundWindow(previousWindow);
                }
            }
        }
    }

    // Pick list popup reference
    private PickListPopupForm? _currentPickListPopup;
    private volatile string? _pendingPickListText;

    private void PerformShowPickLists()
    {
        Logger.Trace("Show Pick Lists");

        // Check if pick lists are enabled
        if (!_config.PickListsEnabled)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Pick lists are disabled", 2000));
            return;
        }

        // Check if there are any pick lists
        if (_config.PickLists.Count == 0)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No pick lists configured", 2000));
            return;
        }

        // Get current study description
        var studyDescription = _mosaicReader.LastDescription;

        // Filter pick lists by study criteria
        var matchingLists = _config.PickLists
            .Where(pl => pl.Enabled && pl.MatchesStudy(studyDescription))
            .ToList();

        if (matchingLists.Count == 0)
        {
            var studyInfo = string.IsNullOrEmpty(studyDescription) ? "no study" : $"'{studyDescription}'";
            InvokeUI(() => _mainForm.ShowStatusToast($"No pick lists match {studyInfo}", 2500));
            return;
        }

        Logger.Trace($"Pick Lists: {matchingLists.Count} list(s) match study '{studyDescription}'");

        // If keep-open is enabled and popup exists, just bring it to front
        if (_config.PickListKeepOpen && _currentPickListPopup != null && !_currentPickListPopup.IsDisposed)
        {
            InvokeUI(() =>
            {
                _currentPickListPopup?.Activate();
                _currentPickListPopup?.Focus();
            });
            return;
        }

        // Close existing popup if open (and keep-open is disabled)
        if (_currentPickListPopup != null && !_currentPickListPopup.IsDisposed)
        {
            InvokeUI(() =>
            {
                try { _currentPickListPopup?.Close(); } catch { }
                _currentPickListPopup = null;
            });
        }

        // Show popup on UI thread
        InvokeUI(() =>
        {
            _currentPickListPopup = new PickListPopupForm(_config, matchingLists, studyDescription, OnPickListItemSelected);
            _currentPickListPopup.FormClosed += (s, e) => _currentPickListPopup = null;
            _currentPickListPopup.Show();
        });
    }

    private void OnPickListItemSelected(string text)
    {
        Logger.Trace($"Pick list item selected: {text.Substring(0, Math.Min(50, text.Length))}...");

        // Store the text and queue the internal action for STA thread
        _pendingPickListText = text;
        TriggerAction("__InsertPickListText__", "Internal");
    }

    private void PerformInsertPickListText()
    {
        var text = _pendingPickListText;
        _pendingPickListText = null;

        if (string.IsNullOrEmpty(text)) return;

        if (IsAddendumOpen())
        {
            Logger.Trace("InsertPickListText: Blocked - addendum is open");
            InvokeUI(() => _mainForm.ShowStatusToast("Cannot paste into addendum", 2500));
            return;
        }

        Logger.Trace($"InsertPickListText: Pasting {text.Length} chars");

        // Use paste lock to prevent race conditions
        lock (PasteLock)
        {
            // Save focus
            var previousWindow = NativeWindows.GetForegroundWindow();

            try
            {
                // Activate Mosaic and paste
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(50);

                // Focus Transcript box to ensure paste goes to correct location
                _mosaicCommander.FocusTranscriptBox();
                Thread.Sleep(50);

                InsertTextToFocusedEditor(PrepareTextForPaste(text));
                Thread.Sleep(100);

                InvokeUI(() => _mainForm.ShowStatusToast("Pick list item inserted", 1500));
            }
            catch (Exception ex)
            {
                Logger.Trace($"InsertPickListText error: {ex.Message}");
                InvokeUI(() => _mainForm.ShowStatusToast($"Error: {ex.Message}", 2500));
            }
            finally
            {
                LastPasteTime = DateTime.Now;

                // Restore focus
                if (previousWindow != IntPtr.Zero)
                {
                    Thread.Sleep(100);
                    NativeWindows.SetForegroundWindow(previousWindow);
                }
            }
        }
    }

    private void PerformGetPrior()
    {
        Logger.Trace("Get Prior");

        if (IsAddendumOpen())
        {
            Logger.Trace("GetPrior: Blocked - addendum is open");
            InvokeUI(() => _mainForm.ShowStatusToast("Cannot paste into addendum", 2500));
            return;
        }

        InvokeUI(() => _mainForm.ShowStatusToast("Extracting Prior..."));

        {
            // Check if InteleViewer is active
            var foreground = NativeWindows.GetForegroundWindow();
            var activeTitle = NativeWindows.GetWindowTitle(foreground).ToLower();
            
            if (!activeTitle.Contains("inteleviewer"))
            {
                InvokeUI(() => _mainForm.ShowStatusToast("InteleViewer must be active!"));
                return;
            }
            
            // Save old clipboard
            var oldClip = ClipboardService.GetText() ?? "";
            ClipboardService.Clear();
            
            // Release modifier keys (Python Parity)
            NativeWindows.KeyUpModifiers();
            Thread.Sleep(300);
            
            // Send configured hotkey (e.g., "v" or "ctrl+shift+r")
            var hotkey = _config.IvReportHotkey;
            if (string.IsNullOrEmpty(hotkey)) hotkey = "v";
            
            Logger.Trace($"GetPrior: Sending hotkey '{hotkey}'");
            NativeWindows.SendHotkey(hotkey);
            
            // Wait for window to appear/load - reduced for speed
            Thread.Sleep(150);
            
            // Copy Attempt Loop
            string? rawText = null;
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                Logger.Trace($"GetPrior: Copy Attempt {attempt}...");
                ClipboardService.Clear();
                Thread.Sleep(50); // Snappy delay before copy
                
                NativeWindows.SendHotkey("ctrl+c");
                
                // Wait for clipboard (Snappy 50ms polling)
                for (int i = 0; i < 15; i++)
                {
                    rawText = ClipboardService.GetText();
                    if (!string.IsNullOrEmpty(rawText) && rawText.Length >= 5)
                        break;
                    Thread.Sleep(50);
                }

                if (!string.IsNullOrEmpty(rawText) && rawText.Length >= 5)
                    break;
                
                Thread.Sleep(200); // reduced wait between attempts
            }
            
            if (string.IsNullOrEmpty(rawText) || rawText.Length < 5)
            {
                Logger.Trace("GetPrior: Failed to retrieve text after 3 attempts.");
                InvokeUI(() => _mainForm.ShowStatusToast("No text retrieved"));
                return;
            }
            
            Logger.Trace($"Raw prior text: {rawText.Substring(0, Math.Min(100, rawText.Length))}...");
            
            // Process
            var formatted = _getPriorService.ProcessPriorText(rawText);
            if (string.IsNullOrEmpty(formatted))
            {
                InvokeUI(() => _mainForm.ShowStatusToast("Could not parse prior"));
                return;
            }
            
            // Paste into Mosaic (with leading and trailing newline for cleaner insertion)
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);

            // Focus Transcript box to ensure paste goes to correct location
            _mosaicCommander.FocusTranscriptBox();
            Thread.Sleep(100);

            InsertTextToFocusedEditor(PrepareTextForPaste(formatted + "\n"));

            Logger.Trace($"Get Prior complete: {formatted}");
            InvokeUI(() => _mainForm.ShowStatusToast("Prior inserted"));
        }
    }
    
    private void PerformCriticalFindings(bool debugMode = false)
    {
        if (debugMode)
        {
            // Debug mode: scrape but show dialog instead of pasting
            Logger.Trace("Critical Findings DEBUG MODE");
            InvokeUI(() => _mainForm.ShowStatusToast("Debug mode: Scraping Clario...", 2000));

            {
                var rawNote = _automationService.PerformClarioScrape(msg =>
                {
                    InvokeUI(() => _mainForm.ShowStatusToast(msg));
                });

                // Update clinical history window if visible
                InvokeUI(() => _mainForm.UpdateClinicalHistory(rawNote));

                var formatted = rawNote != null ? _noteFormatter.FormatNote(rawNote) : "No note found";

                _mainForm.ShowDebugResults(rawNote ?? "None", formatted);
                InvokeUI(() => _mainForm.ShowStatusToast(
                    "Debug complete. Review the raw data above to troubleshoot extraction.", 10000));
            }
            return;
        }

        // Normal mode: scrape and paste
        Logger.Trace("Critical Findings (Clario Scrape)");
        InvokeUI(() => _mainForm.ShowStatusToast("Scraping Clario..."));

        {
            // Scrape with repeating toast callback
            var rawNote = _automationService.PerformClarioScrape(msg =>
            {
                InvokeUI(() => _mainForm.ShowStatusToast(msg));
            });

            if (string.IsNullOrEmpty(rawNote))
            {
                InvokeUI(() => _mainForm.ShowStatusToast("No EXAM NOTE found"));
                return;
            }

            // Update clinical history window if visible
            InvokeUI(() => _mainForm.UpdateClinicalHistory(rawNote));

            // Format
            var formatted = _noteFormatter.FormatNote(rawNote);

            // Insert into Mosaic final report box (not transcript)
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(200);
            _mosaicCommander.FocusFinalReportBox();
            Thread.Sleep(100);

            InsertTextToFocusedEditor(formatted);

            Logger.Trace($"Critical Findings complete: {formatted}");

            // Remove from critical studies tracker (user has dealt with this study)
            UntrackCriticalStudy();

            InvokeUI(() => _mainForm.ShowStatusToast(
                "Critical findings inserted.\nHold Win key and trigger again to debug.", 20000));
        }
    }

    /// <summary>
    /// Track a critical study entry after successful critical note paste.
    /// </summary>
    private void TrackCriticalStudy()
    {
        if (!_config.TrackCriticalStudies)
            return;

        var accession = _mosaicReader.LastAccession;
        if (string.IsNullOrEmpty(accession))
        {
            Logger.Trace("TrackCriticalStudy: No accession to track");
            return;
        }

        // Avoid duplicate entries for the same accession
        lock (_criticalStudiesLock)
        {
            if (_criticalStudies.Any(s => s.Accession == accession))
            {
                Logger.Trace($"TrackCriticalStudy: Already tracking accession {accession}");
                return;
            }

            var entry = new CriticalStudyEntry
            {
                Accession = accession,
                PatientName = _mosaicReader.LastPatientName ?? "Unknown",
                SiteCode = _mosaicReader.LastSiteCode ?? "???",
                Description = _mosaicReader.LastDescription ?? "Unknown",
                Mrn = _mosaicReader.LastMrn ?? "",
                CriticalNoteTime = DateTime.Now
            };

            _criticalStudies.Add(entry);
            Logger.Trace($"TrackCriticalStudy: Added entry for {accession} @ {entry.SiteCode}");
        }

        // Notify UI
        InvokeUI(() => CriticalStudiesChanged?.Invoke());
    }

    /// <summary>
    /// Remove the current study from the critical studies tracker (user has dealt with it).
    /// </summary>
    private void UntrackCriticalStudy()
    {
        if (!_config.TrackCriticalStudies)
            return;

        var accession = _mosaicReader.LastAccession;
        if (string.IsNullOrEmpty(accession))
        {
            Logger.Trace("UntrackCriticalStudy: No accession to untrack");
            return;
        }

        lock (_criticalStudiesLock)
        {
            var entry = _criticalStudies.FirstOrDefault(s => s.Accession == accession);
            if (entry == null)
            {
                Logger.Trace($"UntrackCriticalStudy: Accession {accession} not in tracker");
                return;
            }

            _criticalStudies.Remove(entry);
            Logger.Trace($"UntrackCriticalStudy: Removed entry for {accession}");
        }

        // Notify UI
        InvokeUI(() => CriticalStudiesChanged?.Invoke());
    }

    public void RemoveCriticalStudy(CriticalStudyEntry entry)
    {
        if (entry == null) return;
        lock (_criticalStudiesLock)
        {
            _criticalStudies.Remove(entry);
        }
        Logger.Trace($"RemoveCriticalStudy: Manually removed entry for {entry.Accession}");
        InvokeUI(() => CriticalStudiesChanged?.Invoke());
    }

    // State for Report Popup Toggle (volatile: accessed from STA thread and UI thread)
    private volatile ReportPopupForm? _currentReportPopup;

    private volatile Form? _currentRadAiPopup;          // [RadAI]
    private volatile RadAiOverlayForm? _currentRadAiOverlay;  // [RadAI] overlay mode (ShowReportAfterProcess)
    private volatile string[]? _pendingRadAiImpressionItems;  // [RadAI]
    private volatile bool _radAiAutoInsertPending;             // [RadAI] waiting for scrape to stabilize after Process Report
    private string? _radAiPreProcessReport;                    // [RadAI] report text captured before Process Report
    private DateTime _radAiAutoInsertRequestTime;              // [RadAI] timeout for pending auto-insert

    private void PerformShowReport()
    {
        Logger.Trace("Show Report (scrape method)");

        // Toggle Logic: If open, try click cycle first, then close
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible)
        {
            InvokeUI(() =>
            {
                try
                {
                    if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && !_currentReportPopup.HandleClickCycle())
                    {
                        _currentReportPopup.Close();
                        _currentReportPopup = null;
                        _lastPopupReportText = null;
                    }
                }
                catch
                {
                    _currentReportPopup = null;
                    _lastPopupReportText = null;
                }
            });
            return;
        }

        try
        {
            // Use the scraped report text (much faster than Alt+C)
            string? reportText = _mosaicReader.LastFinalReport;

            if (string.IsNullOrEmpty(reportText))
            {
                // Distinguish between "no study open" and "study loading"
                var currentAccession = _mosaicReader.LastAccession;
                if (!string.IsNullOrEmpty(currentAccession))
                {
                    // We have an accession but no report yet - show popup with "loading" message
                    // The scrape timer will auto-update the popup when report becomes available
                    InvokeUI(() =>
                    {
                        _currentReportPopup = new ReportPopupForm(_config, "Report loading...", null,
                            changesEnabled: _config.ShowReportChanges,
                            correlationEnabled: _config.CorrelationEnabled,
                            baselineIsSectionOnly: false,
                            accession: currentAccession);
                        _lastPopupReportText = null; // Will be set when real report arrives

                        _currentReportPopup.ImpressionDeleteRequested += OnImpressionDeleteRequested;
                        _currentReportPopup.FormClosed += (s, e) =>
                        {
                            _currentReportPopup = null;
                            _lastPopupReportText = null;
                        };

                        _currentReportPopup.Show();
                    });
                }
                else
                {
                    InvokeUI(() => _mainForm.ShowStatusToast("No report available (scraping may be disabled)"));
                }
                return;
            }

            // Pass baseline for diff highlighting if: feature enabled AND (process report was pressed OR drafted study auto-processed)
            string? baselineForDiff = (_config.ShowReportChanges && (_processReportPressedForCurrentAccession || _draftedAutoProcessDetected))
                ? _baselineReport
                : null;

            Logger.Trace($"ShowReport: {reportText.Length} chars, ShowReportChanges={_config.ShowReportChanges}, ProcessPressed={_processReportPressedForCurrentAccession}, DraftedAutoProcess={_draftedAutoProcessDetected}, BaselineLen={_baselineReport?.Length ?? 0}");
            if (baselineForDiff != null)
            {
                Logger.Trace($"ShowReport: Passing baseline for diff ({baselineForDiff.Length} chars)");
            }

            InvokeUI(() =>
            {
                _currentReportPopup = new ReportPopupForm(_config, reportText, baselineForDiff,
                    changesEnabled: _config.ShowReportChanges,
                    correlationEnabled: _config.CorrelationEnabled,
                    baselineIsSectionOnly: baselineForDiff != null && _baselineIsFromTemplateDb,
                    accession: _mosaicReader.LastAccession);
                _lastPopupReportText = reportText;

                _currentReportPopup.ImpressionDeleteRequested += OnImpressionDeleteRequested;
                // Handle closure to clear references
                _currentReportPopup.FormClosed += (s, e) =>
                {
                    _currentReportPopup = null;
                    _lastPopupReportText = null;
                };

                _currentReportPopup.Show();

                // If Process Report was just pressed, show "Updating..." indicator immediately
                // (report is being processed, so current content may be stale)
                if (_processReportPressedForCurrentAccession)
                {
                    _staleSetTime = DateTime.UtcNow;
                    _currentReportPopup.SetStaleState(true);
                }
            });
        }
        catch (Exception ex)
        {
             Logger.Trace($"ShowReport error: {ex.Message}");
             InvokeUI(() => _mainForm.ShowStatusToast("Error showing report"));
        }
    }
    
    private void PerformCaptureSeries()
    {
        if (IsAddendumOpen())
        {
            Logger.Trace("CaptureSeries: Blocked - addendum is open");
            InvokeUI(() => _mainForm.ShowStatusToast("Cannot paste into addendum", 2500));
            return;
        }

        Logger.Trace("Capture Series/Image");
        InvokeUI(() => _mainForm.ShowStatusToast("Capturing..."));
        
        try
        {
            // Find yellow target box
            var targetRect = _ocrService.FindYellowTarget();
            Rectangle captureRect;
            
            if (targetRect.HasValue)
            {
                // User requested 450x450 crop (plenty for header info)
                captureRect = new Rectangle(targetRect.Value.X, targetRect.Value.Y, 450, 450);
                Logger.Trace($"Yellow box found at {targetRect.Value}, capturing 450x450");
            }
            else
            {
                // Fallback: capture around mouse cursor
                var cursorPos = Cursor.Position;
                captureRect = new Rectangle(cursorPos.X - 200, cursorPos.Y - 200, 400, 400);
                Logger.Trace($"Yellow box not found, using cursor: {cursorPos}");
            }
            
            // Run OCR synchronously to stay on STA thread for clipboard
            var ocrTask = _ocrService.CaptureAndRecognizeAsync(captureRect);
            ocrTask.Wait();
            var ocrText = ocrTask.Result;
            Logger.Trace($"OCR result: {ocrText}");
            
            var result = OcrService.ExtractSeriesImageNumbers(ocrText, _config.SeriesImageTemplate);
            
            if (string.IsNullOrEmpty(result))
            {
                InvokeUI(() => _mainForm.ShowStatusToast("Could not extract series/image"));
                return;
            }
            
            // Insert into Mosaic
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(200);
            InsertTextToFocusedEditor(result);
            
            InvokeUI(() => _mainForm.ShowStatusToast($"Inserted: {result}"));
        }
        catch (Exception ex)
        {
            Logger.Trace($"CaptureSeries error: {ex.Message}");
            InvokeUI(() => _mainForm.ShowStatusToast($"OCR Error: {ex.Message}"));
        }
    }

    private void PerformCycleWindowLevel()
    {
        // Check if we have any keys configured
        var keys = _config.WindowLevelKeys;
        if (keys == null || keys.Count == 0)
        {
            Logger.Trace("CycleWindowLevel: No keys configured");
            return;
        }

        // Check if InteleViewer is the active window (fast title check)
        var foreground = NativeWindows.GetForegroundWindow();
        var title = NativeWindows.GetWindowTitle(foreground);

        if (!title.Contains("InteleViewer", StringComparison.OrdinalIgnoreCase))
        {
            Logger.Trace($"CycleWindowLevel: Not InteleViewer (title='{title.Substring(0, Math.Min(60, title.Length))}')");
            return;
        }

        // Release modifiers without delay (F-keys work fine even if modifiers briefly held)
        NativeWindows.KeyUpModifiers();

        // Get the next key in the cycle
        if (_windowLevelCycleIndex >= keys.Count) _windowLevelCycleIndex = 0;
        var keyName = keys[_windowLevelCycleIndex];
        _windowLevelCycleIndex = (_windowLevelCycleIndex + 1) % keys.Count;

        // Convert key name to VK code and send
        var vk = KeyNameToVK(keyName);
        if (vk != 0)
        {
            NativeWindows.keybd_event(vk, 0, 0, UIntPtr.Zero);
            NativeWindows.keybd_event(vk, 0, NativeWindows.KEYEVENTF_KEYUP, UIntPtr.Zero);
            Logger.Trace($"CycleWindowLevel: Sent {keyName} (VK=0x{vk:X2})");
        }
        else
        {
            Logger.Trace($"CycleWindowLevel: Unknown key '{keyName}'");
        }
    }

    /// <summary>
    /// Convert a key name (e.g., "F4", "F5") to a virtual key code.
    /// </summary>
    private static byte KeyNameToVK(string keyName)
    {
        return keyName.ToUpperInvariant() switch
        {
            "F1" => 0x70, "F2" => 0x71, "F3" => 0x72, "F4" => 0x73,
            "F5" => 0x74, "F6" => 0x75, "F7" => 0x76, "F8" => 0x77,
            "F9" => 0x78, "F10" => 0x79, "F11" => 0x7A, "F12" => 0x7B,
            "1" => 0x31, "2" => 0x32, "3" => 0x33, "4" => 0x34,
            "5" => 0x35, "6" => 0x36, "7" => 0x37, "8" => 0x38, "9" => 0x39, "0" => 0x30,
            _ => 0
        };
    }

    private void PerformCreateCriticalNote()
    {
        Logger.Trace("Create Critical Note action triggered");

        var accession = _mosaicReader.LastAccession;
        if (string.IsNullOrEmpty(accession))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No study loaded", 2000));
            return;
        }

        if (HasCriticalNoteForAccession(accession))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Critical note already created", 2000));
            return;
        }

        bool success = _automationService.CreateCriticalCommunicationNote();
        if (success)
        {
            _criticalNoteCreatedForAccessions.TryAdd(accession, 0);

            // Track critical study for session-based tracker
            TrackCriticalStudy();

            InvokeUI(() =>
            {
                _mainForm.ShowStatusToast("Critical note created", 3000);
                _mainForm.SetNoteCreatedState(true);
            });
        }
        else
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Failed to create critical note - Clario may not be open", 3000));
        }
    }

    // [RadAI] — entire method; remove when RadAI integration is retired
    private void PerformRadAiImpression()
    {
        Logger.Trace("RadAI Impression action triggered");

        // If popup/overlay is already open, re-trigger acts as Insert
        bool hasOverlay = _currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed;
        bool hasPopup = _currentRadAiPopup != null && !_currentRadAiPopup.IsDisposed;
        if ((hasOverlay || hasPopup) && _pendingRadAiImpressionItems != null)
        {
            Logger.Trace("RadAI Impression: Display already open, triggering insert");
            PerformRadAiInsert();
            return;
        }

        if (_radAiService == null)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("RadAI not available (config not found)", 3000));
            return;
        }

        // Fresh scrape to ensure we have the latest report text (not stale cache)
        _mosaicReader.GetFinalReportFast();
        var reportText = _mosaicReader.LastFinalReport;
        if (string.IsNullOrEmpty(reportText))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No report text available", 2000));
            return;
        }

        InvokeUI(() => _mainForm.ShowStatusToast("Generating RadAI impression...", 10000));

        var result = Task.Run(async () => await _radAiService.GetImpressionAsync(
            reportText,
            _mosaicReader.LastDescription,
            _mosaicReader.LastPatientGender
        )).GetAwaiter().GetResult();

        if (result.Success && !string.IsNullOrEmpty(result.Impression))
        {
            _pendingRadAiImpressionItems = result.ImpressionItems;

            if (_config.ShowReportAfterProcess)
            {
                // Overlay mode: show below report popup, matching its visual style
                InvokeUI(() =>
                {
                    // Close existing overlay if open
                    if (_currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed)
                    {
                        try { _currentRadAiOverlay.Close(); } catch { }
                    }

                    var overlay = new RadAiOverlayForm(_config, result.ImpressionItems);

                    // Link to report popup if it's open (follow its position/size)
                    var reportPopup = _currentReportPopup;
                    if (reportPopup != null && !reportPopup.IsDisposed && reportPopup.Visible)
                    {
                        overlay.LinkToForm(reportPopup);
                    }
                    else
                    {
                        // Standalone: position at saved report popup location
                        overlay.Location = ScreenHelper.EnsureOnScreen(_config.ReportPopupX, _config.ReportPopupY + 200);
                    }

                    overlay.FormClosed += (s, e) => _currentRadAiOverlay = null;
                    _currentRadAiOverlay = overlay;
                    overlay.Show();

                    _mainForm.ShowStatusToast("RadAI impression ready — press again to insert", 3000);
                });
            }
            else
            {
                // Classic popup mode: standalone dialog with Insert/Copy/Close buttons
                InvokeUI(() =>
                {
                    if (_currentRadAiPopup != null && !_currentRadAiPopup.IsDisposed)
                    {
                        try { _currentRadAiPopup.Close(); } catch { }
                    }

                    var popup = RadAiService.ShowResultPopup(result.ImpressionItems, () =>
                    {
                        TriggerAction("__RadAiInsert__", "Internal");
                    }, _config);
                    _currentRadAiPopup = popup;
                    popup.FormClosed += (s, e) => _currentRadAiPopup = null;
                });
            }
        }
        else
        {
            var errorMsg = result.Success ? "RadAI returned empty impression" : $"RadAI error: {result.Error}";
            InvokeUI(() => _mainForm.ShowStatusToast(errorMsg, 4000));
        }
    }

    private void PerformRecoMd()
    {
        Logger.Trace("RecoMD action triggered");

        if (!_config.RecoMdEnabled)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("RecoMD is not enabled", 2000));
            return;
        }

        var alive = Task.Run(async () => await _recoMdService.IsAliveAsync()).GetAwaiter().GetResult();
        if (!alive)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("RecoMD is not running", 2000));
            return;
        }

        // Capture metadata BEFORE GetFinalReportFast (which resets patient fields)
        var accession = _mosaicReader.LastAccession;
        var description = _mosaicReader.LastDescription;
        var patientName = _mosaicReader.LastPatientName;
        var gender = _mosaicReader.LastPatientGender;
        var mrn = _mosaicReader.LastMrn;
        var age = _mosaicReader.LastPatientAge ?? 0;

        if (string.IsNullOrEmpty(accession))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No study open", 2000));
            return;
        }

        // Get current report text (this resets and re-scrapes patient fields)
        _mosaicReader.GetFinalReportFast();
        var reportText = _mosaicReader.LastFinalReport;
        if (string.IsNullOrEmpty(reportText))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No report text available", 2000));
            return;
        }

        // Clean report text: remove U+FFFC object replacement chars and collapse blank lines
        reportText = RecoMdService.CleanReportText(reportText);

        // Update tracking so continuous send in scrape timer knows report is opened
        _recoMdOpenedForAccession = accession;
        _lastRecoMdSentText = reportText;

        var success = Task.Run(async () =>
        {
            var opened = await _recoMdService.OpenReportAsync(accession,
                description, patientName, gender, mrn, age);
            if (!opened) return false;

            return await _recoMdService.SendReportTextAsync(accession, reportText);
        }).GetAwaiter().GetResult();

        if (success)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Sent to RecoMD", 2000));
            BringRecoMdToFront();
        }
        else
        {
            InvokeUI(() => _mainForm.ShowStatusToast("RecoMD: Failed to send", 3000));
        }
    }

    /// <summary>
    /// Find the RecoMD window, activate it, and pin it topmost so it stays visible.
    /// </summary>
    private static IntPtr _cachedRecoMdHwnd;

    private static IntPtr GetRecoMdWindow()
    {
        // Reuse cached handle if still valid
        if (_cachedRecoMdHwnd != IntPtr.Zero && NativeWindows.IsWindow(_cachedRecoMdHwnd) && NativeWindows.IsWindowVisible(_cachedRecoMdHwnd))
            return _cachedRecoMdHwnd;

        _cachedRecoMdHwnd = NativeWindows.FindWindowByProcessName("recoMD");
        return _cachedRecoMdHwnd;
    }

    private static void BringRecoMdToFront()
    {
        // Give RecoMD a moment to process the payload and show its UI
        Thread.Sleep(500);
        var hWnd = GetRecoMdWindow();
        if (hWnd != IntPtr.Zero)
        {
            // Activate first (handles minimized, thread attachment, alt-key trick)
            NativeWindows.ActivateWindow(hWnd, 1000);
            // Then pin topmost so it stays on top
            NativeWindows.ForceTopMost(hWnd);
            Logger.Trace($"RecoMD: Window (HWND={hWnd}) activated and set topmost");
        }
        else
        {
            Logger.Trace("RecoMD: Window not found by process name");
        }
    }

    /// <summary>
    /// Remove topmost from the RecoMD window so it goes back to normal z-order.
    /// </summary>
    private static void SendRecoMdToBack()
    {
        var hWnd = GetRecoMdWindow();
        if (hWnd != IntPtr.Zero)
        {
            NativeWindows.SetWindowPos(hWnd, NativeWindows.HWND_NOTOPMOST,
                0, 0, 0, 0, NativeWindows.SWP_NOMOVE | NativeWindows.SWP_NOSIZE);
            Logger.Trace("RecoMD: Window topmost removed");
        }
    }

    /// <summary>
    /// Click the green "ALL" accept button in RecoMD to copy recommendations,
    /// then paste them into the IMPRESSION section in Mosaic.
    /// </summary>
    private void PerformPasteRecoMd()
    {
        // Skip if another paste is already in progress (user pressed button multiple times)
        if (_recoMdPasteInProgress)
        {
            Logger.Trace("Paste RecoMD skipped — already in progress");
            return;
        }
        _recoMdPasteInProgress = true;
        try
        {
        PerformPasteRecoMdCore();
        }
        finally
        {
            _recoMdPasteInProgress = false;
        }
    }

    private void PerformPasteRecoMdCore()
    {
        Logger.Trace("Paste RecoMD action triggered");

        if (!_config.RecoMdEnabled)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("RecoMD is not enabled", 2000));
            return;
        }

        var hWnd = GetRecoMdWindow();
        if (hWnd == IntPtr.Zero)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("RecoMD window not found", 2000));
            return;
        }

        // Step 1: Click RecoMD ALL button — puts recommendations on clipboard.
        NativeWindows.ActivateWindow(hWnd, 500);
        NativeWindows.ForceTopMost(hWnd);
        Thread.Sleep(300);

        if (!NativeWindows.GetWindowRect(hWnd, out var winRect))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("RecoMD window not found", 2000));
            return;
        }

        var greenCenter = FindGreenAllButton(winRect);
        if (greenCenter == null)
        {
            Logger.Trace("RecoMD: Could not find green ALL button");
            InvokeUI(() => _mainForm.ShowStatusToast("RecoMD: No recommendations available", 2000));
            return;
        }

        Logger.Trace($"RecoMD: Clicking ALL button at ({greenCenter.Value.X},{greenCenter.Value.Y})");
        NativeWindows.ClickAtScreenPos(greenCenter.Value.X, greenCenter.Value.Y, restoreCursor: false);
        Thread.Sleep(500);
        SendRecoMdToBack();

        // Step 2: Read RecoMD clipboard via Win32 API (bypasses .NET's broken Chromium support).
        var recoText = NativeWindows.GetClipboardTextWin32();
        if (string.IsNullOrWhiteSpace(recoText))
        {
            Logger.Trace("RecoMD: Win32 clipboard read returned empty");
            InvokeUI(() => _mainForm.ShowStatusToast("RecoMD: Could not read recommendations from clipboard", 3000));
            return;
        }
        Logger.Trace($"RecoMD: Got {recoText.Length} chars from clipboard via Win32");

        // Step 3: Get existing impression from the already-scraped report.
        var reportText = _mosaicReader.LastFinalReport;
        if (string.IsNullOrEmpty(reportText))
        {
            Logger.Trace("RecoMD: No scraped report available, attempting fresh scrape");
            reportText = _mosaicReader.GetFinalReportFast();
        }

        string existingImpression = "";
        if (!string.IsNullOrEmpty(reportText))
        {
            var (_, impression) = Services.CorrelationService.ExtractSections(reportText);
            existingImpression = impression;
        }

        // Clean existing impression: remove U+FFFC chars, blank lines, and strip
        // auto-numbering (e.g., "1. text" → "text") since Mosaic re-adds numbering on paste.
        if (!string.IsNullOrEmpty(existingImpression))
        {
            var cleanLines = existingImpression.Split('\n')
                .Select(l => l.Replace("\uFFFC", "").Trim())
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .Select(l => System.Text.RegularExpressions.Regex.Replace(l, @"^\d+[.)]\s*", ""))
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .ToList();
            existingImpression = string.Join("\r\n", cleanLines);
        }

        // Step 4: Combine existing impression + RecoMD recommendations.
        string combined;
        if (string.IsNullOrWhiteSpace(existingImpression))
        {
            combined = recoText.Trim();
        }
        else
        {
            combined = existingImpression.TrimEnd() + "\r\n" + recoText.Trim();
        }

        Logger.Trace($"RecoMD: Combined impression ({existingImpression.Length} existing + {recoText.Trim().Length} reco = {combined.Length} chars)");

        // Step 5: Replace impression content — exact same pattern as RadAI insert.
        lock (PasteLock)
        {
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);

            _mosaicCommander.FocusFinalReportBox();
            Thread.Sleep(100);
            NativeWindows.SendHotkey("ctrl+end");
            Thread.Sleep(100);

            bool selected = _mosaicCommander.SelectImpressionContent();
            if (!selected)
            {
                InvokeUI(() => _mainForm.ShowStatusToast("Could not find IMPRESSION section", 3000));
                return;
            }
            Thread.Sleep(50);

            InsertTextToFocusedEditor(combined);
            Thread.Sleep(100);
        }

        Logger.Trace("RecoMD: Pasted recommendations into Mosaic impression");
        InvokeUI(() => _mainForm.ShowStatusToast("RecoMD recommendations pasted", 2000));
    }

    /// <summary>
    /// Scan the top portion of the RecoMD window for the green "ALL" thumbs-up button.
    /// Uses BitBlt to capture the region in one shot (fast), then scans pixels in memory.
    /// Returns screen coordinates of the button center, or null if not found.
    /// </summary>
    private static System.Drawing.Point? FindGreenAllButton(NativeWindows.RECT winRect)
    {
        int scanHeight = Math.Min(50, winRect.Height);
        int scanWidth = Math.Min(400, winRect.Width);

        var pixels = NativeWindows.CaptureScreenRegion(winRect.Left, winRect.Top, scanWidth, scanHeight, out int stride);
        if (pixels == null)
        {
            Logger.Trace("RecoMD pixel scan: BitBlt capture failed");
            return null;
        }

        int greenMinX = int.MaxValue, greenMaxX = 0;
        int greenMinY = int.MaxValue, greenMaxY = 0;
        int greenCount = 0;

        // Scan with 2px step — pixels are BGRA (4 bytes per pixel, top-down)
        for (int dy = 0; dy < scanHeight; dy += 2)
        {
            int rowOffset = dy * stride;
            for (int dx = 0; dx < scanWidth; dx += 2)
            {
                int idx = rowOffset + dx * 4;
                int b = pixels[idx];
                int g = pixels[idx + 1];
                int r = pixels[idx + 2];

                // Bright green: G is dominant, R and B are low
                if (g > 140 && r < 130 && b < 130 && g > r && g > b)
                {
                    greenCount++;
                    int sx = winRect.Left + dx;
                    int sy = winRect.Top + dy;
                    if (sx < greenMinX) greenMinX = sx;
                    if (sx > greenMaxX) greenMaxX = sx;
                    if (sy < greenMinY) greenMinY = sy;
                    if (sy > greenMaxY) greenMaxY = sy;
                }
            }
        }

        Logger.Trace($"RecoMD pixel scan: {greenCount} green pixels found in top {scanHeight}px");

        if (greenCount < 5) return null;

        int cx = (greenMinX + greenMaxX) / 2;
        int cy = (greenMinY + greenMaxY) / 2;
        return new System.Drawing.Point(cx, cy);
    }

    // [RadAI] — entire method; remove when RadAI integration is retired
    private void PerformRadAiInsert()
    {
        Logger.Trace("RadAI Insert action triggered");

        var items = _pendingRadAiImpressionItems;

        if (items == null || items.Length == 0)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No RadAI impression to insert", 2000));
            return;
        }

        // Join items without numbering — Mosaic auto-numbers impression lines
        var newImpression = string.Join("\r\n", items);

        lock (PasteLock)
        {
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);

            // Focus report editor and scroll to bottom
            _mosaicCommander.FocusFinalReportBox();
            Thread.Sleep(100);
            NativeWindows.SendHotkey("ctrl+end");
            Thread.Sleep(100);

            // Use UIA to find IMPRESSION content elements and select them
            // Retry several times — after Process Report, Mosaic may still be rebuilding the editor
            bool selected = false;
            for (int attempt = 0; attempt < 5; attempt++)
            {
                selected = _mosaicCommander.SelectImpressionContent();
                if (selected) break;
                Logger.Trace($"RadAI Insert: IMPRESSION not found, retry {attempt + 1}/5...");
                Thread.Sleep(1500);
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(100);
                _mosaicCommander.FocusFinalReportBox();
                Thread.Sleep(100);
                NativeWindows.SendHotkey("ctrl+end");
                Thread.Sleep(100);
            }
            if (!selected)
            {
                InvokeUI(() => _mainForm.ShowStatusToast("Could not find IMPRESSION section in editor", 3000));
                return;
            }

            Thread.Sleep(50);

            // Replace selection or insert at cursor
            InsertTextToFocusedEditor(newImpression);
            Thread.Sleep(100);
        }

        // Overlay mode: verify the paste succeeded, then auto-close
        if (_currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed)
        {
            Thread.Sleep(500); // Give Mosaic time to process the paste
            _mosaicReader.GetFinalReportFast();
            var updatedReport = _mosaicReader.LastFinalReport ?? "";

            // Verify: check if the first impression item appears in the report
            bool verified = false;
            if (items.Length > 0 && !string.IsNullOrEmpty(updatedReport))
            {
                var checkText = items[0].Trim();
                if (checkText.Length > 30) checkText = checkText.Substring(0, 30);
                verified = updatedReport.Contains(checkText, StringComparison.OrdinalIgnoreCase);
            }

            if (verified)
            {
                Logger.Trace("RadAI Insert: Verified — impression found in report");
                _pendingRadAiImpressionItems = null;
                InvokeUI(() =>
                {
                    if (_currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed)
                    {
                        try { _currentRadAiOverlay.Close(); } catch { }
                        _currentRadAiOverlay = null;
                    }
                    _mainForm.ShowStatusToast("RadAI impression inserted", 3000);
                });
            }
            else
            {
                Logger.Trace("RadAI Insert: Verification failed — impression not found in report");
                InvokeUI(() => _mainForm.ShowStatusToast("RadAI insert may have failed — try again", 4000));
            }
        }
        else
        {
            // Classic popup mode: no verification, just toast
            Logger.Trace("RadAI Insert: Impression replaced in report");
            InvokeUI(() => _mainForm.ShowStatusToast("RadAI impression inserted", 3000));
        }
    }

    private void OnImpressionDeleteRequested(string newText)
    {
        _pendingImpressionReplaceText = newText;
        _impressionDeletePending = true;
        TriggerAction("__ReplaceImpression__", "ImpressionDelete");
    }

    private void PerformReplaceImpression()
    {
        Logger.Trace("PerformReplaceImpression triggered");

        var newText = _pendingImpressionReplaceText;
        if (newText == null)
        {
            _impressionDeletePending = false;
            InvokeUI(() => _currentReportPopup?.ClearDeletePending());
            return;
        }

        lock (PasteLock)
        {
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);

            // Ctrl+End to scroll to bottom of report (works better than Page Down for short reports)
            NativeWindows.SendHotkey("ctrl+end");
            Thread.Sleep(100);

            // Select impression content
            bool selected = _mosaicCommander.SelectImpressionContent();
            if (!selected)
            {
                Logger.Trace("PerformReplaceImpression: Could not find IMPRESSION section");
                InvokeUI(() => _mainForm.ShowStatusToast("Could not find IMPRESSION section", 3000));
                _impressionDeletePending = false;
                _pendingImpressionReplaceText = null;
                InvokeUI(() => _currentReportPopup?.ClearDeletePending());
                return;
            }

            Thread.Sleep(50);

            if (string.IsNullOrEmpty(newText))
            {
                // All points deleted — send Delete key to clear impression
                const byte VK_DELETE = 0x2E;
                NativeWindows.keybd_event(VK_DELETE, 0, 0, UIntPtr.Zero);
                Thread.Sleep(10);
                NativeWindows.keybd_event(VK_DELETE, 0, NativeWindows.KEYEVENTF_KEYUP, UIntPtr.Zero);
            }
            else
            {
                InsertTextToFocusedEditor(newText);
            }
            Thread.Sleep(100);
        }

        // Verify: re-scrape and update report popup display
        Thread.Sleep(500);
        _mosaicReader.GetFinalReportFast();
        var updatedReport = _mosaicReader.LastFinalReport ?? "";

        _pendingImpressionReplaceText = null;
        _impressionDeletePending = false;

        InvokeUI(() =>
        {
            _currentReportPopup?.ClearDeletePending();
            // Update the report popup with the re-scraped content
            if (!string.IsNullOrEmpty(updatedReport) && _currentReportPopup != null && !_currentReportPopup.IsDisposed)
            {
                _currentReportPopup.UpdateReport(updatedReport);
                _lastPopupReportText = updatedReport;
            }
            // Also update impression window if visible
            var updatedImpression = ImpressionForm.ExtractImpression(updatedReport);
            if (!string.IsNullOrEmpty(updatedImpression))
                _mainForm.UpdateImpression(updatedImpression);
        });

        Logger.Trace("PerformReplaceImpression completed");
    }

    #endregion

    #region Dictation Sync

    private System.Threading.Timer? _dictationSyncTimer;

    private void StartDictationSync()
    {
        _dictationSyncTimer = new System.Threading.Timer(_ =>
        {
            if (_isUserActive) return; // Don't check during user actions

            try
            {
                bool? state = NativeWindows.IsMicrophoneActiveFromRegistry();
                if (!state.HasValue) return;

                // Only sync ON immediately; sync OFF only if main sync timer agrees
                // (avoids contradicting the debounced sync timer's OFF logic)
                if (state.Value && !_dictationActive)
                {
                    _dictationActive = true;
                    InvokeUI(() => _mainForm.UpdateIndicatorState(true));
                }
            }
            catch { }
        }, null, 1000, 2000);
    }
    
    #endregion
    
    #region Mosaic Scrape Timer
    
    private void StartMosaicScrapeTimer()
    {
        if (_scrapeTimer != null) return; // Already running

        Logger.Trace($"Starting Mosaic scrape timer ({_config.ScrapeIntervalSeconds}s interval)...");
        _scrapeTimer = new System.Threading.Timer(_ =>
        {
            if (_isUserActive) return; // Don't scrape during user actions
            if (Interlocked.CompareExchange(ref _scrapeRunning, 1, 0) != 0)
            {
                // Watchdog: if a scrape has been running for >30s, it's probably hung in a COM call.
                // Force-release the guard so future scrapes aren't permanently blocked.
                var startTicks = Interlocked.Read(ref _scrapeStartedTicks);
                if (startTicks > 0)
                {
                    var elapsed = (DateTime.UtcNow.Ticks - startTicks) / (double)TimeSpan.TicksPerSecond;
                    if (elapsed > 30)
                    {
                        Logger.Trace($"Scrape watchdog: force-releasing hung scrape after {elapsed:F0}s");
                        Interlocked.Exchange(ref _scrapeRunning, 0);
                        Interlocked.Exchange(ref _scrapeStartedTicks, 0);
                        // Don't proceed on this tick — let the next timer fire start fresh
                    }
                }
                return;
            }
            Interlocked.Exchange(ref _scrapeStartedTicks, DateTime.UtcNow.Ticks);

            try
            {
                // Only check drafted status if we need it for features
                // Also need it for macros since Description extraction happens during drafted check
                bool needDraftedCheck = _config.ShowDraftedIndicator ||
                                        _config.ShowImpression ||
                                        (_config.MacrosEnabled && _config.Macros.Count > 0);

                // Scrape Mosaic for report data
                var reportText = _mosaicReader.GetFinalReportFast(needDraftedCheck);

                // Bail out if user action started during the scrape
                if (_isUserActive) return;

                // Idle backoff: slow down scraping when no report content is found
                // (either no study open, or study visible but ProseMirror not accessible)
                // Two tiers: 10s after 3 idle scrapes, 30s after 10 idle scrapes.
                // Heavy FlaUI/UIA calls every few seconds cause system-wide input lag
                // (mouse jumpiness, keyboard hook death) even when finding nothing.
                if (string.IsNullOrEmpty(reportText))
                {
                    _consecutiveIdleScrapes++;
                    if (!_searchingForImpression && !_needsBaselineCapture)
                    {
                        if (_consecutiveIdleScrapes >= 10)
                        {
                            RestartScrapeTimer(DeepIdleScrapeIntervalMs);
                            if (_consecutiveIdleScrapes == 10)
                                Logger.Trace("Scrape deep idle: slowing to 30s (extended no report content)");
                        }
                        else if (_consecutiveIdleScrapes >= 3)
                        {
                            RestartScrapeTimer(IdleScrapeIntervalMs);
                            if (_consecutiveIdleScrapes == 3)
                                Logger.Trace("Scrape idle backoff: slowing to 10s (no report content)");
                        }
                    }
                }
                else if (_consecutiveIdleScrapes > 0)
                {
                    _consecutiveIdleScrapes = 0;
                    // Restore normal rate (unless fast/study-load mode is active)
                    if (!_searchingForImpression && !_needsBaselineCapture)
                        RestartScrapeTimer(NormalScrapeIntervalMs);
                }

                // Scrape heartbeat: log every ~120 ticks (~4 min at 2s) to confirm timer is alive
                _scrapeHeartbeatCount++;
                if (_scrapeHeartbeatCount >= 120)
                {
                    _scrapeHeartbeatCount = 0;
                    var acc = _mosaicReader.LastAccession ?? "(none)";
                    Logger.Trace($"Scrape heartbeat: acc={acc}, idle={_consecutiveIdleScrapes}, clinHist={_clinicalHistoryVisible}");
                }

                // Periodic GC: safety net for any COM wrappers that escape deterministic release
                // (e.g., _cachedSlimHubWindow, elements accessed via indexed properties).
                // Most COM objects are now released immediately via ReleaseElement/ReleaseElements,
                // so this runs much less frequently as a backstop.
                _scrapesSinceLastGc++;
                if (_scrapesSinceLastGc >= 120)
                {
                    _scrapesSinceLastGc = 0;
                    GC.Collect(2, GCCollectionMode.Forced);
                    GC.WaitForPendingFinalizers();
                }

                // [RadAI] Auto-insert: trigger when scrape succeeds after Process Report.
                // After Alt+P, Mosaic rebuilds the editor (15-25s). During rebuild, GetFinalReportFast
                // returns stale cached text (U+FFFC filter). Once the editor stabilizes, the scrape
                // returns fresh text that differs from the pre-process snapshot — that's our signal.
                if (_radAiAutoInsertPending && !string.IsNullOrEmpty(reportText))
                {
                    if ((DateTime.Now - _radAiAutoInsertRequestTime).TotalSeconds > 60)
                    {
                        Logger.Trace("RadAI: Auto-insert timed out (60s), cancelling");
                        _radAiAutoInsertPending = false;
                    }
                    else
                    {
                        var preProcess = _radAiPreProcessReport ?? "";
                        if (reportText != preProcess
                            && reportText.Contains("IMPRESSION", StringComparison.OrdinalIgnoreCase))
                        {
                            Logger.Trace($"RadAI: Report stabilized after Process Report ({reportText.Length} chars), triggering auto-insert");
                            _radAiAutoInsertPending = false;
                            _actionQueue.Enqueue(new ActionRequest { Action = Actions.RadAiImpression, Source = "AutoProcess" });
                            _actionQueue.Enqueue(new ActionRequest { Action = "__RadAiInsert__", Source = "AutoProcess" });
                        }
                    }
                }

                // Check for new study (non-empty accession different from last non-empty)
                var currentAccession = _mosaicReader.LastAccession;

                // Clear stale Clario data when accession changes, before pipe send
                if (!string.IsNullOrEmpty(currentAccession) && currentAccession != _lastNonEmptyAccession)
                {
                    _automationService.ClearStrokeState();
                }

                // Send study data over pipe (only sends when changed, record equality)
                _pipeService.SendStudyData(new StudyDataMessage(
                    Type: "study_data",
                    Accession: currentAccession,
                    Description: _mosaicReader.LastDescription,
                    TemplateName: _mosaicReader.LastTemplateName,
                    PatientName: _mosaicReader.LastPatientName,
                    PatientGender: _mosaicReader.LastPatientGender,
                    Mrn: _mosaicReader.LastMrn,
                    SiteCode: _mosaicReader.LastSiteCode,
                    ClarioPriority: _automationService.LastClarioPriority,
                    ClarioClass: _automationService.LastClarioClass,
                    Drafted: _mosaicReader.LastDraftedState,
                    HasCritical: !string.IsNullOrEmpty(currentAccession) && HasCriticalNoteForAccession(currentAccession),
                    Timestamp: DateTime.UtcNow.ToString("o")
                ));

                // Check for discard dialog (RVUCounter integration)
                // Must check BEFORE accession change detection - dialog disappears when user clicks YES
                // Only check when a study is open (avoids unnecessary FlaUI tree search)
                if (!string.IsNullOrEmpty(currentAccession) && _mosaicReader.IsDiscardDialogVisible())
                {
                    _discardDialogShownForCurrentAccession = true;
                    Logger.Trace("RVUCounter: Discard dialog detected for current accession");
                }

                // Detect accession change with flap debounce
                var (accessionChanged, studyClosed) = DetectAccessionChange(currentAccession);

                if (accessionChanged)
                    OnStudyChanged(currentAccession, reportText, studyClosed);

                UpdateRecoMd(currentAccession, reportText);

                RecordTemplateIfNeeded(reportText);

                CaptureBaselineReport(reportText);

                UpdateReportPopup(reportText);

                InsertPendingMacros(currentAccession, reportText);

                UpdateClinicalHistoryAndAlerts(currentAccession, reportText);

                UpdateImpressionDisplay(reportText);

                // Re-assert topmost on all tool windows periodically.
                // Other apps (Chrome, InteleViewer) can steal topmost status;
                // without this, our windows stay behind until an action is triggered.
                BatchUI(() => _mainForm.EnsureWindowsOnTop());
            }
            catch (Exception ex)
            {
                Logger.Trace($"Mosaic scrape error: {ex.Message}");
            }
            finally
            {
                // Flush all batched UI updates in a single BeginInvoke call.
                // This reduces WM_USER message spam that can cause Windows to
                // silently remove the low-level keyboard hook.
                FlushUI();
                Interlocked.Exchange(ref _scrapeStartedTicks, 0);
                Interlocked.Exchange(ref _scrapeRunning, 0);
            }
        }, null, 0, NormalScrapeIntervalMs);
    }
    
    private void StopMosaicScrapeTimer()
    {
        if (_scrapeTimer != null)
        {
            Logger.Trace("Stopping Mosaic scrape timer...");
            _scrapeTimer.Dispose();
            _scrapeTimer = null;
        }
    }
    
    public void ToggleMosaicScraper(bool enabled)
    {
        if (enabled)
        {
            // Stop and restart to pick up any interval changes from settings
            StopMosaicScrapeTimer();
            StartMosaicScrapeTimer();
        }
        else
        {
            StopMosaicScrapeTimer();
        }
    }

    /// <summary>
    /// Flap-debounce logic for accession transitions.
    /// Mosaic briefly sets accession to empty during study transitions — this detects
    /// real changes vs transient flaps by waiting for the empty state to persist.
    /// </summary>
    private (bool changed, bool closed) DetectAccessionChange(string? currentAccession)
    {
        bool accessionChanged = false;
        bool studyClosed = false;

        if (!string.IsNullOrEmpty(currentAccession))
        {
            // Non-empty accession — cancel any pending close since accession came back
            if (_pendingCloseAccession != null)
            {
                if (currentAccession == _pendingCloseAccession)
                {
                    // Same accession returned after brief empty — this was a flap, ignore it
                    Logger.Trace($"Accession flap cancelled: '{currentAccession}' returned after {_pendingCloseTickCount} tick(s)");
                    _pendingCloseAccession = null;
                    _pendingCloseTickCount = 0;
                    // Don't process as a study change — nothing actually changed
                }
                else
                {
                    // Different accession appeared — process the pending close first, then the new study
                    // The pending close's RVU event and state reset will happen as part of the new study change
                    _pendingCloseAccession = null;
                    _pendingCloseTickCount = 0;
                    accessionChanged = true;
                }
            }
            else if (currentAccession != _lastNonEmptyAccession)
            {
                // New non-empty accession (direct transition, no empty gap)
                accessionChanged = true;
            }
        }
        else if (!string.IsNullOrEmpty(_lastNonEmptyAccession))
        {
            // Accession just went empty — start debounce, don't process yet
            if (_pendingCloseAccession == null)
            {
                _pendingCloseAccession = _lastNonEmptyAccession;
                _pendingCloseTickCount = 1;
                Logger.Trace($"Accession went empty, deferring close for '{_pendingCloseAccession}' (tick 1)");
            }
            else
            {
                _pendingCloseTickCount++;
                if (_pendingCloseTickCount >= 3)
                {
                    // Empty for 3+ ticks — this is a real close, not a flap
                    Logger.Trace($"Accession confirmed closed after {_pendingCloseTickCount} ticks: '{_pendingCloseAccession}'");
                    accessionChanged = true;
                    studyClosed = true;
                    _pendingCloseAccession = null;
                    _pendingCloseTickCount = 0;
                }
            }
        }
        else if (_pendingCloseAccession != null)
        {
            // Still empty and we have a pending close — increment tick count
            _pendingCloseTickCount++;
            if (_pendingCloseTickCount >= 3)
            {
                Logger.Trace($"Accession confirmed closed after {_pendingCloseTickCount} ticks: '{_pendingCloseAccession}'");
                accessionChanged = true;
                studyClosed = true;
                _pendingCloseAccession = null;
                _pendingCloseTickCount = 0;
            }
        }

        return (accessionChanged, studyClosed);
    }

    /// <summary>
    /// Full state reset on study change: RVU counter notification, RecoMD close, UI reset, new study setup.
    /// Called when accession changes (new study opened or study closed).
    /// Future API: replaced by study.opened / study.closed event handlers.
    /// </summary>
    private void OnStudyChanged(string? currentAccession, string? reportText, bool studyClosed)
    {
        Logger.Trace($"Study change detected: '{_lastNonEmptyAccession}' -> '{currentAccession ?? "(empty)"}'");

        // RVUCounter integration: Notify about previous study
        // Logic:
        //   1. If _currentAccessionSigned → SIGNED (MosaicTools triggered sign)
        //   2. Else if discard dialog was shown for this accession → CLOSED_UNSIGNED
        //   3. Else → SIGNED (no dialog = manual sign via Alt+F or button click)
        //   Each type has a _CRITICAL variant if a critical note was created
        if (!string.IsNullOrEmpty(_lastNonEmptyAccession))
        {
            var hasCritical = HasCriticalNoteForAccession(_lastNonEmptyAccession);
            var criticalSuffix = hasCritical ? "_CRITICAL" : "";

            if (_currentAccessionSigned)
            {
                // Explicitly signed via MosaicTools
                var msgType = hasCritical ? NativeWindows.MSG_STUDY_SIGNED_CRITICAL : NativeWindows.MSG_STUDY_SIGNED;
                Logger.Trace($"RVUCounter: Sending SIGNED{criticalSuffix} (MosaicTools) for '{_lastNonEmptyAccession}'");
                NativeWindows.SendToRvuCounter(msgType, _lastNonEmptyAccession);
                _pipeService.SendStudyEvent(new StudyEventMessage("study_event", "signed", _lastNonEmptyAccession, hasCritical));
            }
            else if (_discardDialogShownForCurrentAccession)
            {
                // Discard dialog was shown for this accession → study was discarded
                var msgType = hasCritical ? NativeWindows.MSG_STUDY_CLOSED_UNSIGNED_CRITICAL : NativeWindows.MSG_STUDY_CLOSED_UNSIGNED;
                Logger.Trace($"RVUCounter: Sending CLOSED_UNSIGNED{criticalSuffix} (dialog was shown) for '{_lastNonEmptyAccession}'");
                NativeWindows.SendToRvuCounter(msgType, _lastNonEmptyAccession);
                _pipeService.SendStudyEvent(new StudyEventMessage("study_event", "unsigned", _lastNonEmptyAccession, hasCritical));
            }
            else
            {
                // No dialog → user signed manually (Alt+F or clicked Sign button)
                var msgType = hasCritical ? NativeWindows.MSG_STUDY_SIGNED_CRITICAL : NativeWindows.MSG_STUDY_SIGNED;
                Logger.Trace($"RVUCounter: Sending SIGNED{criticalSuffix} (manual) for '{_lastNonEmptyAccession}'");
                NativeWindows.SendToRvuCounter(msgType, _lastNonEmptyAccession);
                _pipeService.SendStudyEvent(new StudyEventMessage("study_event", "signed", _lastNonEmptyAccession, hasCritical));
            }
        }

        // RecoMD: close previous study on accession change
        if (_config.RecoMdEnabled)
        {
            _recoMdOpenedForAccession = null;
            _lastRecoMdSentText = null;
            SendRecoMdToBack();
            Task.Run(async () => await _recoMdService.CloseReportAsync());
        }

        // Reset state for new study
        _currentAccessionSigned = false;
        _discardDialogShownForCurrentAccession = false;
        _processReportPressedForCurrentAccession = false;
        _draftedAutoProcessDetected = false;
        _autoShowReportDoneForAccession = false;
        // Note: _criticalNoteCreatedForAccessions is session-scoped, NOT reset on study change
        // This prevents duplicate notes caused by transient accession-null scrape glitches
        _baselineReport = null;
        _baselineIsFromTemplateDb = false;
        _baselineCaptureAttempts = 0;
        _templateRecordedForStudy = false;
        _lastPopupReportText = null;
        _mosaicReader.ClearLastReport();
        // Close stale report popup from prior study
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed)
        {
            var stalePopup = _currentReportPopup;
            _currentReportPopup = null;
            BatchUI(() => { try { stalePopup.Close(); } catch { } });
        }
        // [RadAI] Cancel pending auto-insert and close stale popup/overlay from prior study
        _radAiAutoInsertPending = false;
        if (_currentRadAiPopup != null && !_currentRadAiPopup.IsDisposed)
        {
            var staleRadAi = _currentRadAiPopup;
            _currentRadAiPopup = null;
            _pendingRadAiImpressionItems = null;
            BatchUI(() => { try { staleRadAi.Close(); } catch { } });
        }
        if (_currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed)
        {
            var staleOverlay = _currentRadAiOverlay;
            _currentRadAiOverlay = null;
            _pendingRadAiImpressionItems = null;
            BatchUI(() => { try { staleOverlay.Close(); } catch { } });
        }
        // Reset alert state tracking
        _templateMismatchActive = false;
        _genderMismatchActive = false;
        _strokeDetectedActive = false;
        _pendingClarioPriorityRetry = false;
        _lastAidocFindings = null;
        _lastAidocRelevant = false;
        _aidocConfirmedNegative.Clear();

        // Reset Ignore Inpatient Drafted state
        _macrosCompleteForCurrentAccession = false;
        _autoFixCompleteForCurrentAccession = false;
        _ctrlASentForCurrentAccession = false;

        // Update tracking - only update to new non-empty accession
        if (!studyClosed)
        {
            _lastNonEmptyAccession = currentAccession;

            _needsBaselineCapture = _config.ShowReportChanges || _config.CorrelationEnabled;
            // Speed up scraping to catch template flash before auto-processing
            if (_needsBaselineCapture && !_searchingForImpression)
                RestartScrapeTimer(_studyLoadScrapeIntervalMs);

            // Baseline capture deferred to scrape timer to catch template flash
            // (removing immediate capture fixes issue where dictated content like "appendix is normal"
            // gets captured as baseline if already present in template/macros)
            Logger.Trace($"New study - ShowReportChanges={_config.ShowReportChanges}, CorrelationEnabled={_config.CorrelationEnabled}, reportText null={string.IsNullOrEmpty(reportText)}");

            // Show toast (disabled - too noisy)
            // Logger.Trace($"Showing New Study toast for {currentAccession}");
            // InvokeUI(() => _mainForm.ShowStatusToast($"New Study: {currentAccession}", 3000));

            // Re-show clinical history window if it was hidden due to no study
            // (only in always-show mode; alerts-only mode will show when alert triggers)
            if (_config.HideClinicalHistoryWhenNoStudy && _config.ShowClinicalHistory && _config.AlwaysShowClinicalHistory)
            {
                BatchUI(() => _mainForm.ToggleClinicalHistory(true));
                _clinicalHistoryVisible = true;
            }

            // Re-show indicator window if it was hidden due to no study
            if (_config.HideIndicatorWhenNoStudy && _config.IndicatorEnabled)
            {
                BatchUI(() => _mainForm.ToggleIndicator(true));
            }

            // Queue macros for insertion - they'll be inserted when clinical history is visible
            // This handles the case where Mosaic auto-processes the report on open
            if (_config.MacrosEnabled && _config.Macros.Count > 0)
            {
                var studyDescription = _mosaicReader.LastDescription;
                Logger.Trace($"Macros: Queuing for study '{studyDescription}' ({_config.Macros.Count} macros configured)");
                _pendingMacroAccession = currentAccession;
                _pendingMacroDescription = studyDescription;
            }

            // Reset clinical history state on study change
            BatchUI(() => _mainForm.OnClinicalHistoryStudyChanged());
            // Hide impression window on new study
            BatchUI(() => _mainForm.HideImpressionWindow());
            _searchingForImpression = false;
            _impressionFromProcessReport = false;

            // Always extract Clario Priority/Class for pipe broadcasts
            // Stroke detection UI logic only runs if that feature is enabled
            ExtractClarioPriorityAndClass(currentAccession);

            // If Clario hasn't caught up yet (accession mismatch / null priority),
            // flag for retry on subsequent scrape ticks
            _pendingClarioPriorityRetry = string.IsNullOrEmpty(_automationService.LastClarioPriority);

            if (_config.StrokeDetectionEnabled)
            {
                PerformStrokeDetection(currentAccession, reportText);
            }
            else
            {
                InvokeUI(() => _mainForm.SetStrokeState(false));
            }
        }
        else
        {
            // Study closed without new one opening - clear the tracked accession
            _lastNonEmptyAccession = null;
            _needsBaselineCapture = false;
            Logger.Trace("Study closed, no new study opened");

            // Revert to idle scrape interval — no study open, no reason to scrape fast
            if (!_searchingForImpression)
                RestartScrapeTimer(IdleScrapeIntervalMs);

            // Hide clinical history window if configured to hide when no study
            // (or always hide in alerts-only mode when no alerts)
            if (_config.ShowClinicalHistory && (_config.HideClinicalHistoryWhenNoStudy || !_config.AlwaysShowClinicalHistory))
            {
                BatchUI(() => _mainForm.ToggleClinicalHistory(false));
                _clinicalHistoryVisible = false;
            }

            // Hide indicator window if configured to hide when no study
            if (_config.HideIndicatorWhenNoStudy && _config.IndicatorEnabled)
            {
                BatchUI(() => _mainForm.ToggleIndicator(false));
            }
        }
    }

    /// <summary>
    /// Deferred baseline capture for diff highlighting.
    /// Captures the clean template text before user dictation, with DB fallback.
    /// Future API: replaced by report.changed event handler.
    /// </summary>
    private void CaptureBaselineReport(string? reportText)
    {
        if (!_needsBaselineCapture || _processReportPressedForCurrentAccession)
            return;

        if (_mosaicReader.LastDraftedState)
        {
            _baselineCaptureAttempts++;

            if (!string.IsNullOrEmpty(reportText) && _baselineCaptureAttempts <= 1)
            {
                // First tick: try real-time capture (template flash)
                // Only capture if report seems like a clean template (no long dictated content)
                // Templates are typically < 2500 chars; anything significantly longer likely has dictation
                bool seemsLikeCleanTemplate = reportText.Length < 2500;

                if (seemsLikeCleanTemplate)
                {
                    _needsBaselineCapture = false;
                    _baselineReport = reportText;
                    Logger.Trace($"Captured baseline from scrape (DRAFTED, immediate): {reportText.Length} chars, drafted={_mosaicReader.LastDraftedState}");


                    // Revert to normal scrape interval now that baseline is captured
                    if (!_searchingForImpression)
                        RestartScrapeTimer(NormalScrapeIntervalMs);
                }
                else
                {
                    Logger.Trace($"Skipping baseline capture - report too long ({reportText.Length} chars), likely has dictation. Will try DB fallback.");
                    // Skip ahead to DB fallback
                    _baselineCaptureAttempts = 4;
                }
            }
            else if (_baselineCaptureAttempts >= 4)
            {
                // ~2 seconds elapsed, give up on real-time capture, try DB fallback
                _needsBaselineCapture = false;
                if (!_searchingForImpression)
                    RestartScrapeTimer(NormalScrapeIntervalMs);

                if (_config.TemplateDatabaseEnabled)
                {
                    var desc = _mosaicReader.LastDescription;
                    if (!string.IsNullOrEmpty(desc))
                    {
                        var fallback = _templateDatabase.GetFallbackTemplate(desc);
                        if (fallback != null)
                        {
                            _baselineReport = fallback;
                            _baselineIsFromTemplateDb = true;
                            Logger.Trace($"Using template DB fallback for '{desc}' ({fallback.Length} chars)");
                        }
                        else
                        {
                            Logger.Trace($"TemplateDB: No confident fallback for '{desc}'");
                        }
                    }
                }
            }
        }
        else
        {
            _baselineCaptureAttempts++;

            if (!string.IsNullOrEmpty(reportText))
            {
                // Non-drafted: wait for impression to appear - report is generated top-to-bottom after Process Report
                var impression = ImpressionForm.ExtractImpression(reportText);
                if (!string.IsNullOrEmpty(impression))
                {
                    _needsBaselineCapture = false;
                    _baselineReport = reportText;
                    Logger.Trace($"Captured baseline from scrape ({reportText.Length} chars)");

                    // Revert to normal scrape interval now that baseline is captured
                    if (!_searchingForImpression)
                        RestartScrapeTimer(NormalScrapeIntervalMs);
                }
            }

            // Timeout: give up after ~4 seconds (8 attempts at 500ms) to prevent stuck fast timer
            if (_needsBaselineCapture && _baselineCaptureAttempts >= 8)
            {
                _needsBaselineCapture = false;
                Logger.Trace($"Baseline capture timed out (non-drafted, {_baselineCaptureAttempts} attempts). Restoring normal scrape interval.");
                if (!_searchingForImpression)
                    RestartScrapeTimer(NormalScrapeIntervalMs);
            }
        }
    }

    /// <summary>
    /// Save clean template to DB on first scrape tick before Process Report.
    /// Future API: replaced by study.opened event handler.
    /// </summary>
    private void RecordTemplateIfNeeded(string? reportText)
    {
        if (!_templateRecordedForStudy && !_processReportPressedForCurrentAccession
            && _config.TemplateDatabaseEnabled && !string.IsNullOrEmpty(reportText)
            && !string.IsNullOrEmpty(_mosaicReader.LastDescription))
        {
            _templateRecordedForStudy = true;
            _templateDatabase.RecordTemplate(_mosaicReader.LastDescription, reportText,
                isDrafted: _mosaicReader.LastDraftedState);
        }
    }

    /// <summary>
    /// Check and insert queued macros when clinical history becomes visible.
    /// Future API: replaced by clinical_history.available event handler.
    /// </summary>
    private void InsertPendingMacros(string? currentAccession, string? reportText)
    {
        if (string.IsNullOrEmpty(_pendingMacroAccession) ||
            _pendingMacroAccession != currentAccession ||
            string.IsNullOrWhiteSpace(reportText) ||
            !reportText.Contains("CLINICAL HISTORY", StringComparison.OrdinalIgnoreCase))
            return;

        // Don't insert macros into addendums
        if (reportText.TrimStart().StartsWith("Addendum", StringComparison.OrdinalIgnoreCase))
        {
            Logger.Trace($"Macros: Blocked pending insert - addendum detected for {_pendingMacroAccession}");
            MarkMacrosCompleteForCurrentAccession();
            _pendingMacroAccession = null;
            _pendingMacroDescription = null;
        }
        else
        {
            Logger.Trace($"Macros: Clinical history now visible, inserting for {_pendingMacroAccession}");
            InsertMacrosForStudy(_pendingMacroDescription);
            _pendingMacroAccession = null;
            _pendingMacroDescription = null;
        }
    }

    /// <summary>
    /// Continuous RecoMD sync — send report text on every scrape tick.
    /// Future API: replaced by report.changed event handler.
    /// </summary>
    private void UpdateRecoMd(string? currentAccession, string? reportText)
    {
        if (!_config.RecoMdEnabled || string.IsNullOrEmpty(currentAccession) || string.IsNullOrEmpty(reportText))
            return;

        var recoText = RecoMdService.CleanReportText(reportText);
        var recoAcc = currentAccession;
        var needsOpen = _recoMdOpenedForAccession != recoAcc;
        var textChanged = recoText != _lastRecoMdSentText;
        _lastRecoMdSentText = recoText;

        if (needsOpen)
        {
            // First send for this accession — open report then send text
            _recoMdOpenedForAccession = recoAcc;
            _recoMdSendTickCounter = 0;
            var recoDesc = _mosaicReader.LastDescription;
            var recoName = _mosaicReader.LastPatientName;
            var recoGender = _mosaicReader.LastPatientGender;
            var recoMrn = _mosaicReader.LastMrn;
            var recoAge = _mosaicReader.LastPatientAge ?? 0;
            Logger.Trace($"RecoMD: Opening + sending for {recoAcc}");
            Task.Run(async () =>
            {
                await _recoMdService.OpenReportAsync(recoAcc, recoDesc, recoName, recoGender, recoMrn, recoAge);
                await _recoMdService.SendReportTextAsync(recoAcc, recoText);
            });
        }
        else
        {
            // Send immediately on text change, otherwise throttle to every 3rd tick
            _recoMdSendTickCounter++;
            if (textChanged)
            {
                _recoMdSendTickCounter = 0;
                Logger.Trace("RecoMD: Text changed, sending update");
                Task.Run(async () => await _recoMdService.SendReportTextAsync(recoAcc, recoText));
            }
            else if (_recoMdSendTickCounter >= 3)
            {
                _recoMdSendTickCounter = 0;
                Task.Run(async () => await _recoMdService.SendReportTextAsync(recoAcc, recoText));
            }
        }
    }

    /// <summary>
    /// Auto-updating report popup with diff highlighting.
    /// Also detects auto-processing on drafted studies.
    /// Future API: replaced by report.changed event handler.
    /// </summary>
    private void UpdateReportPopup(string? reportText)
    {
        // Detect auto-processing on drafted studies: if baseline was captured and report changed, enable diff
        if (!_draftedAutoProcessDetected && !_processReportPressedForCurrentAccession
            && _config.ShowReportChanges && _baselineReport != null
            && _mosaicReader.LastDraftedState
            && !string.IsNullOrEmpty(reportText) && reportText != _baselineReport)
        {
            _draftedAutoProcessDetected = true;
            Logger.Trace($"Drafted study auto-process detected: report changed from baseline ({_baselineReport.Length} → {reportText.Length} chars)");


        }

        // Auto-update popup when report text changes (continuous updates with diff highlighting)
        var popup = _currentReportPopup;
        if (popup != null && !popup.IsDisposed && popup.Visible)
        {
            if (!string.IsNullOrEmpty(reportText) && reportText != _lastPopupReportText)
            {
                Logger.Trace($"Auto-updating popup: report changed ({reportText.Length} chars vs {_lastPopupReportText?.Length ?? 0} chars), baseline={_baselineReport?.Length ?? 0} chars");
                _lastPopupReportText = reportText;
                BatchUI(() => { if (!popup.IsDisposed) popup.UpdateReport(reportText, _baselineReport, _baselineIsFromTemplateDb); });
            }
            else if (string.IsNullOrEmpty(reportText) && !string.IsNullOrEmpty(_lastPopupReportText))
            {
                // Report text is gone (being updated in Mosaic) but popup is visible with cached content
                Logger.Trace("Popup showing stale content - report being updated");
                _staleSetTime = DateTime.UtcNow;
                BatchUI(() => { if (!popup.IsDisposed) popup.SetStaleState(true); });
            }
            else if (!string.IsNullOrEmpty(reportText) &&
                (!_processReportPressedForCurrentAccession || (DateTime.UtcNow - _staleSetTime).TotalSeconds > 10))
            {
                // Report text is available and matches last text - clear stale indicator if showing.
                // Don't clear if Process Report was just pressed (still waiting for report to regenerate),
                // but force-clear after 10s to prevent stuck stale indicator.
                BatchUI(() => { if (!popup.IsDisposed) popup.SetStaleState(false); });
            }
        }
    }

    /// <summary>
    /// Update clinical history display, template/gender/stroke/Aidoc alerts.
    /// Future API: replaced by study.opened + report.changed event handlers.
    /// </summary>
    private void UpdateClinicalHistoryAndAlerts(string? currentAccession, string? reportText)
    {
        if (!_config.ShowClinicalHistory)
            return;

        // First, evaluate all alert conditions
        bool newTemplateMismatch = false;
        bool newGenderMismatch = false;
        string? templateDescription = null;
        string? templateName = null;
        string? patientGender = null;
        List<string>? genderMismatches = null;

        // Check template matching (red border when mismatch) if enabled
        if (_config.ShowTemplateMismatch)
        {
            templateDescription = _mosaicReader.LastDescription;
            templateName = _mosaicReader.LastTemplateName;
            bool bodyPartsMatch = AutomationService.DoBodyPartsMatch(templateDescription, templateName);
            newTemplateMismatch = !bodyPartsMatch;
        }

        // Check for gender mismatch
        if (_config.GenderCheckEnabled && !string.IsNullOrWhiteSpace(reportText))
        {
            patientGender = _mosaicReader.LastPatientGender;
            genderMismatches = ClinicalHistoryForm.CheckGenderMismatch(reportText, patientGender);
            newGenderMismatch = genderMismatches.Count > 0;
        }

        // Retry Clario priority extraction if it failed on accession change
        // (Clario may not have updated to the new study yet)
        if (_pendingClarioPriorityRetry && !string.IsNullOrEmpty(currentAccession))
        {
            ExtractClarioPriorityAndClass(currentAccession);
            if (!string.IsNullOrEmpty(_automationService.LastClarioPriority))
            {
                _pendingClarioPriorityRetry = false;
                Logger.Trace($"Clario priority retry succeeded: '{_automationService.LastClarioPriority}'");
                if (_config.StrokeDetectionEnabled)
                {
                    PerformStrokeDetection(currentAccession, reportText);
                }
            }
        }

        // Aidoc scraping - check for AI-detected findings
        bool aidocAlertActive = false;
        string? aidocFindingText = null;
        List<string>? relevantFindings = null;
        bool prevAidocRelevant = _lastAidocRelevant;
        if (_config.AidocScrapeEnabled && !string.IsNullOrEmpty(currentAccession))
        {
            try
            {
                var aidocResult = _aidocService.ScrapeShortcutWidget();
                if (aidocResult != null && aidocResult.Findings.Count > 0)
                {
                    var studyDescription = _mosaicReader.LastDescription;

                    // "Once negative, stay negative" — latch findings that sample as negative.
                    // The Aidoc widget has a red/orange flashing animation when ANY positive
                    // finding exists, which can cause false positives on individual icon pixels.
                    // Real positive icons consistently show orange; false positives flicker.
                    foreach (var f in aidocResult.Findings)
                    {
                        if (!f.IsPositive)
                            _aidocConfirmedNegative.Add(f.FindingType);
                    }

                    // Only show findings that are relevant, sampled positive, AND not latched negative
                    relevantFindings = aidocResult.Findings
                        .Where(f => f.IsPositive && !_aidocConfirmedNegative.Contains(f.FindingType)
                            && AidocService.IsRelevantFinding(f.FindingType, studyDescription))
                        .Select(f => f.FindingType)
                        .ToList();

                    if (relevantFindings.Count > 0)
                    {
                        aidocAlertActive = true;
                        aidocFindingText = string.Join(", ", relevantFindings);

                        // Toast on first detection or change
                        if (!_lastAidocRelevant || _lastAidocFindings != aidocFindingText)
                        {
                            Logger.Trace($"Aidoc: Relevant findings '{aidocFindingText}' for study '{studyDescription}'");
                            BatchUI(() => _mainForm.ShowStatusToast($"Aidoc: {aidocFindingText} detected", 5000));
                        }
                    }

                    _lastAidocFindings = aidocAlertActive ? aidocFindingText : null;
                    _lastAidocRelevant = aidocAlertActive;
                }
                else
                {
                    _lastAidocFindings = null;
                    _lastAidocRelevant = false;
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Aidoc scrape error: {ex.Message}");
            }
        }

        // Verify Aidoc findings against report text
        // Only verify against text with U+FFFC (real report editor), not the transcript
        List<FindingVerification>? aidocVerifications = null;
        if (aidocAlertActive && relevantFindings != null && !string.IsNullOrEmpty(reportText)
            && reportText.Contains('\uFFFC'))
        {
            aidocVerifications = AidocFindingVerifier.VerifyFindings(relevantFindings, reportText);
        }

        // Determine if any alerts are active
        bool anyAlertActive = newTemplateMismatch || newGenderMismatch || _strokeDetectedActive || aidocAlertActive;

        // Handle visibility based on always-show vs alerts-only mode
        if (_config.AlwaysShowClinicalHistory)
        {
            // ALWAYS-SHOW MODE: Current behavior - window always visible, show clinical history + border colors
            // Uses BatchUI to consolidate into single BeginInvoke (reduces WM_USER spam that kills keyboard hook)

            // Only update if we have content - don't clear during brief processing gaps
            if (!string.IsNullOrWhiteSpace(reportText))
            {
                BatchUI(() => _mainForm.UpdateClinicalHistory(reportText, currentAccession));
                BatchUI(() => _mainForm.UpdateClinicalHistoryTextColor(reportText));
            }

            // Update template mismatch state
            BatchUI(() => _mainForm.UpdateClinicalHistoryTemplateMismatch(newTemplateMismatch, templateDescription, templateName));

            // Update drafted state (green border when drafted) if enabled
            if (_config.ShowDraftedIndicator)
            {
                bool isDrafted = _mosaicReader.LastDraftedState;
                BatchUI(() => _mainForm.UpdateClinicalHistoryDraftedState(isDrafted));
            }

            // Update gender check
            if (_config.GenderCheckEnabled)
            {
                BatchUI(() => _mainForm.UpdateGenderCheck(reportText, patientGender));
            }
            else
            {
                BatchUI(() => _mainForm.UpdateGenderCheck(null, null));
            }

            // Show Aidoc finding appended to clinical history in orange (not replacing it)
            if (aidocAlertActive && aidocFindingText != null)
            {
                var captured = aidocVerifications;
                BatchUI(() => _mainForm.SetAidocAppend(captured));
            }
            else if (!aidocAlertActive && prevAidocRelevant)
            {
                // Aidoc finding cleared - remove orange append
                BatchUI(() => _mainForm.SetAidocAppend(null));
            }
        }
        else
        {
            // ALERTS-ONLY MODE: Window only appears when an alert triggers

            // Always update clinical history text even in alerts-only mode
            // (needed for auto-fix recheck to detect Mosaic self-corrections)
            if (!string.IsNullOrWhiteSpace(reportText))
            {
                BatchUI(() => _mainForm.UpdateClinicalHistory(reportText, currentAccession));
            }

            // Determine highest priority alert to show
            AlertType? alertToShow = null;
            string alertDetails = "";

            if (newGenderMismatch && genderMismatches != null)
            {
                alertToShow = AlertType.GenderMismatch;
                alertDetails = string.Join(", ", genderMismatches);
            }
            else if (newTemplateMismatch)
            {
                alertToShow = AlertType.TemplateMismatch;
                if (templateDescription != null && templateName != null)
                    alertDetails = $"Study: {templateDescription}\nTemplate: {templateName}";
            }
            else if (_strokeDetectedActive)
            {
                alertToShow = AlertType.StrokeDetected;
                alertDetails = "Study flagged as stroke protocol";
            }
            else if (aidocAlertActive && aidocFindingText != null)
            {
                alertToShow = AlertType.AidocFinding;
                alertDetails = aidocFindingText;
            }

            if (anyAlertActive)
            {
                // Show notification box with alert
                if (!_clinicalHistoryVisible)
                {
                    BatchUI(() => _mainForm.ToggleClinicalHistory(true));
                    _clinicalHistoryVisible = true;
                }

                // Display alert content
                if (alertToShow == AlertType.GenderMismatch)
                {
                    // Gender mismatch uses the blinking display
                    BatchUI(() => _mainForm.UpdateGenderCheck(reportText, patientGender));
                }
                else if (alertToShow.HasValue)
                {
                    // Clear gender warning if not active
                    BatchUI(() => _mainForm.UpdateGenderCheck(null, null));
                    // Show the alert
                    BatchUI(() => _mainForm.ShowAlertOnly(alertToShow.Value, alertDetails));
                }

                // Also update template mismatch border (for non-gender alerts)
                if (alertToShow != AlertType.GenderMismatch)
                {
                    BatchUI(() => _mainForm.UpdateClinicalHistoryTemplateMismatch(newTemplateMismatch, templateDescription, templateName));
                }
            }
            else if (_clinicalHistoryVisible)
            {
                // No alerts active - hide the notification box
                BatchUI(() => _mainForm.UpdateGenderCheck(null, null));
                BatchUI(() => _mainForm.ClearAlert());
                BatchUI(() => _mainForm.ToggleClinicalHistory(false));
                _clinicalHistoryVisible = false;
            }
        }

        // Update tracking for next iteration
        _templateMismatchActive = newTemplateMismatch;
        _genderMismatchActive = newGenderMismatch;
    }

    /// <summary>
    /// Impression search/show/hide after Process Report.
    /// Future API: replaced by report.changed event handler.
    /// </summary>
    private void UpdateImpressionDisplay(string? reportText)
    {
        if (!_config.ShowImpression)
            return;

        // Skip impression updates while user is deleting points
        if (_impressionDeletePending)
            return;

        var impression = ImpressionForm.ExtractImpression(reportText);
        bool isDrafted = _mosaicReader.LastDraftedState;

        if (_searchingForImpression)
        {
            // Fast search mode after Process Report - looking for impression
            // Wait 2 seconds before showing to let RadPair finish initial processing
            if (!string.IsNullOrEmpty(impression))
            {
                var elapsed = (DateTime.Now - _impressionSearchStartTime).TotalSeconds;
                if (elapsed >= 2.0)
                {
                    OnImpressionFound(impression);
                }
            }
        }
        else if (_impressionFromProcessReport)
        {
            // Process Report triggered - keep updating impression until Sign Report
            // Don't auto-hide, just update if we have new content
            if (!string.IsNullOrEmpty(impression))
            {
                BatchUI(() => _mainForm.UpdateImpression(impression));
            }
        }
        else if (isDrafted && !string.IsNullOrEmpty(impression))
        {
            // Auto-show impression when study is drafted (passive mode)
            // Only show window if not already visible to avoid flashing
            BatchUI(() =>
            {
                _mainForm.ShowImpressionWindowIfNotVisible();
                _mainForm.UpdateImpression(impression);
            });
        }
        else if (!isDrafted)
        {
            // Hide impression window when study is not drafted (only for auto-shown)
            // Don't hide if it was manually triggered by Process Report
            BatchUI(() => _mainForm.HideImpressionWindow());
        }
    }

    #endregion

    #region Stroke Detection

    // Default keywords for stroke detection in clinical history (used if config list is empty)
    private static readonly string[] DefaultStrokeKeywords = new[]
    {
        "stroke", "CVA", "TIA", "hemiparesis", "hemiplegia", "aphasia",
        "dysarthria", "facial droop", "weakness", "numbness", "code stroke",
        "NIH stroke scale", "NIHSS"
    };

    /// <summary>
    /// Extract Clario Priority/Class for pipe broadcasts. Called on every study change.
    /// </summary>
    private void ExtractClarioPriorityAndClass(string? accession)
    {
        try
        {
            var priorityData = _automationService.ExtractClarioPriorityAndClass(accession);
            if (priorityData != null)
            {
                Logger.Trace($"Extracted Clario Priority='{priorityData.Priority}', Class='{priorityData.Class}'");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"ExtractClarioPriorityAndClass error: {ex.Message}");
        }
    }

    /// <summary>
    /// [CustomSTT] Convert spoken punctuation words to symbols, except "colon" (medical term).
    /// Replaces Deepgram's dictation mode so we have full control over which words convert.
    /// </summary>
    private static string ApplySpokenPunctuation(string text)
    {
        // Order matters: multi-word phrases first, then single words.
        // Case-insensitive replacements with word boundaries.
        // "colon" intentionally omitted — it's a body part, not punctuation.
        var replacements = new (string pattern, string replacement)[]
        {
            (@"\bnew\s+paragraph\b",  "\n\n"),
            (@"\bnew\s+line\b",       "\n"),
            (@"\bexclamation\s+mark\b", "!"),
            (@"\bquestion\s+mark\b",  "?"),
            (@"\bperiod\b",           "."),
            (@"\bcomma\b",            ","),
            (@"\bsemicolon\b",        ";"),
            (@"\bhyphen\b",           "-"),
        };

        var result = text;
        foreach (var (pattern, replacement) in replacements)
        {
            result = System.Text.RegularExpressions.Regex.Replace(
                result, pattern, replacement, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        // Expand contractions that Deepgram favors over spoken full forms
        var contractions = new (string pattern, string replacement)[]
        {
            (@"\bthere's\b",   "there is"),
            (@"\bit's\b",      "it is"),
            (@"\bthat's\b",    "that is"),
            (@"\bdoesn't\b",   "does not"),
            (@"\bdon't\b",     "do not"),
            (@"\bdidn't\b",    "did not"),
            (@"\bwasn't\b",    "was not"),
            (@"\bisn't\b",     "is not"),
            (@"\baren't\b",    "are not"),
            (@"\bweren't\b",   "were not"),
            (@"\bwouldn't\b",  "would not"),
            (@"\bcouldn't\b",  "could not"),
            (@"\bshouldn't\b", "should not"),
            (@"\bhasn't\b",    "has not"),
            (@"\bhaven't\b",   "have not"),
            (@"\bhadn't\b",    "had not"),
            (@"\bcan't\b",     "cannot"),
            (@"\bwon't\b",     "will not"),
        };
        foreach (var (pattern, replacement) in contractions)
        {
            result = System.Text.RegularExpressions.Regex.Replace(
                result, pattern, replacement, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        // Clean up spaces before punctuation marks (e.g., "word . next" → "word. next")
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+([.,;!?])", "$1");

        return result;
    }

    // ── Radiology transcript cleanup ──────────────────────────────────────
    // Converts spoken radiology shorthand into standard written forms.
    // Runs on every STT fragment before paste — shared across all providers.

    private static readonly Dictionary<string, string> NumberWords = new(StringComparer.OrdinalIgnoreCase)
    {
        ["zero"] = "0", ["one"] = "1", ["two"] = "2", ["three"] = "3",
        ["four"] = "4", ["five"] = "5", ["six"] = "6", ["seven"] = "7",
        ["eight"] = "8", ["nine"] = "9", ["ten"] = "10", ["eleven"] = "11",
        ["twelve"] = "12", ["thirteen"] = "13", ["fourteen"] = "14",
        ["fifteen"] = "15", ["sixteen"] = "16", ["seventeen"] = "17",
        ["eighteen"] = "18", ["nineteen"] = "19", ["twenty"] = "20",
        ["thirty"] = "30", ["forty"] = "40", ["fifty"] = "50",
        ["sixty"] = "60", ["seventy"] = "70", ["eighty"] = "80", ["ninety"] = "90",
    };

    // Matches a number word or bare digit(s)
    private const string NumPat = @"(?<n>\d+|zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)";

    private static string NumReplace(System.Text.RegularExpressions.Match m, string groupName = "n")
    {
        var val = m.Groups[groupName].Value;
        return NumberWords.TryGetValue(val, out var d) ? d : val;
    }

    private static readonly Dictionary<string, int> OrdinalWords = new(StringComparer.OrdinalIgnoreCase)
    {
        ["first"] = 1, ["second"] = 2, ["third"] = 3, ["fourth"] = 4, ["fifth"] = 5,
        ["sixth"] = 6, ["seventh"] = 7, ["eighth"] = 8, ["ninth"] = 9, ["tenth"] = 10,
        ["eleventh"] = 11, ["twelfth"] = 12, ["thirteenth"] = 13, ["fourteenth"] = 14,
        ["fifteenth"] = 15, ["sixteenth"] = 16, ["seventeenth"] = 17, ["eighteenth"] = 18,
        ["nineteenth"] = 19, ["twentieth"] = 20, ["thirtieth"] = 30,
    };

    /// <summary>
    /// Parse spoken number (0-99) from words or digits. Returns -1 if not parseable.
    /// Handles: "five", "fifteen", "twenty three", "twenty first", "oh three", "5th", "23".
    /// </summary>
    private static int TryParseSpokenInt(string text)
    {
        text = text.Trim();
        if (string.IsNullOrEmpty(text)) return -1;

        // Bare digits (strip ordinal suffix: "5th" → "5")
        var stripped = System.Text.RegularExpressions.Regex.Replace(text, @"(?:st|nd|rd|th)$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (int.TryParse(stripped, out var n)) return n;

        // Simple cardinal
        if (NumberWords.TryGetValue(text, out var nw) && int.TryParse(nw, out var nwi)) return nwi;

        // Simple ordinal
        if (OrdinalWords.TryGetValue(text, out var ow)) return ow;

        // Compound (two words)
        var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 2)
        {
            // "oh three" → 3
            if (string.Equals(parts[0], "oh", StringComparison.OrdinalIgnoreCase))
            {
                var r = TryParseSpokenInt(parts[1]);
                if (r >= 0 && r <= 9) return r;
            }

            // "twenty three" or "twenty first"
            int tens = parts[0].ToLower() switch
            {
                "twenty" => 20, "thirty" => 30, "forty" => 40, "fifty" => 50,
                "sixty" => 60, "seventy" => 70, "eighty" => 80, "ninety" => 90, _ => -1
            };
            if (tens > 0)
            {
                if (OrdinalWords.TryGetValue(parts[1], out var o) && o >= 1 && o <= 9) return tens + o;
                if (NumberWords.TryGetValue(parts[1], out var c) && int.TryParse(c, out var ci) && ci >= 1 && ci <= 9) return tens + ci;
            }
        }

        return -1;
    }

    /// <summary>
    /// Parse spoken year (1900-2099) from words or digits. Returns -1 if not parseable.
    /// Handles: "twenty twenty five" (2025), "nineteen ninety nine" (1999),
    /// "two thousand and five" (2005), "2025", and bare two-digit numbers
    /// like "twenty five" (→ 2025, assumes 20xx prefix).
    /// </summary>
    private static int TryParseSpokenYear(string text)
    {
        text = text.Trim();
        if (string.IsNullOrEmpty(text)) return -1;
        if (int.TryParse(text, out var n) && n >= 1900 && n <= 2099) return n;

        var words = text.ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length < 1) return -1;

        // "two thousand [and] [XX]"
        if (words.Length >= 2 && words[0] == "two" && words[1] == "thousand")
        {
            if (words.Length == 2) return 2000;
            int start = 2;
            if (start < words.Length && words[start] == "and") start++;
            if (start >= words.Length) return 2000;
            var offset = TryParseSpokenInt(string.Join(" ", words[start..]));
            if (offset >= 1 && offset <= 99) return 2000 + offset;
            return -1;
        }

        // "twenty twenty [one-nine]" → 2020-2029
        if (words.Length >= 2 && words[0] == "twenty" && words[1] == "twenty")
        {
            if (words.Length == 2) return 2020;
            if (words.Length == 3)
            {
                var ones = TryParseSpokenInt(words[2]);
                if (ones >= 1 && ones <= 9) return 2020 + ones;
            }
            return -1;
        }

        // "twenty ten-nineteen" → 2010-2019
        if (words.Length == 2 && words[0] == "twenty")
        {
            var second = TryParseSpokenInt(words[1]);
            if (second >= 10 && second <= 19) return 2000 + second;
        }

        // "nineteen XX" → 19XX
        if (words.Length >= 2 && words[0] == "nineteen")
        {
            var offset = TryParseSpokenInt(string.Join(" ", words[1..]));
            if (offset >= 0 && offset <= 99) return 1900 + offset;
            return -1;
        }

        // Bare two-digit number assumed 20xx: "twenty five" → 2025
        // (used when year is the 3rd number in a numeric date — context implies year)
        var bare = TryParseSpokenInt(text);
        if (bare >= 0 && bare <= 99) return 2000 + bare;

        return -1;
    }

    private static string ApplyRadiologyCleanup(string text)
    {
        const System.Text.RegularExpressions.RegexOptions IC = System.Text.RegularExpressions.RegexOptions.IgnoreCase;

        var result = text;

        // ── Spine levels ──
        // "C one two" → "C1-2", "L five S one" → "L5-S1", "T twelve L one" → "T12-L1"
        // Pattern: [CTLS] + number + optional(dash/through + [CTLS]? + number)
        result = System.Text.RegularExpressions.Regex.Replace(result,
            @"\b(?<seg>[CTLS])\s*" + NumPat.Replace("<n>", "<n1>") +
            @"(?:\s*(?:-|through)?\s*(?<seg2>[CTLS])?\s*" + NumPat.Replace("<n>", "<n2>") + @")?\b",
            m =>
            {
                var seg = m.Groups["seg"].Value.ToUpper();
                var n1 = m.Groups["n1"].Value;
                if (NumberWords.TryGetValue(n1, out var d1)) n1 = d1;

                if (!m.Groups["n2"].Success)
                    return seg + n1;

                var n2 = m.Groups["n2"].Value;
                if (NumberWords.TryGetValue(n2, out var d2)) n2 = d2;

                var seg2 = m.Groups["seg2"].Success ? m.Groups["seg2"].Value.ToUpper() : "";
                return seg + n1 + "-" + seg2 + n2;
            }, IC);

        // Fix concatenated digit spine levels: C34 → C3-4, T12 stays T12
        // Only splits when the 2-digit value exceeds the segment's max vertebra count
        result = System.Text.RegularExpressions.Regex.Replace(result,
            @"\b([CTLS])(\d)(\d)\b",
            m =>
            {
                var seg = m.Groups[1].Value.ToUpper();
                var d1 = m.Groups[2].Value;
                var d2 = m.Groups[3].Value;
                var combined = int.Parse(d1 + d2);
                var max = seg[0] switch { 'C' => 7, 'T' => 12, 'L' => 5, 'S' => 5, _ => 0 };
                if (combined <= max) return seg + d1 + d2; // valid single vertebra (e.g., T12)
                return seg + d1 + "-" + d2;
            }, IC);

        // ── "point" between numbers → decimal ──
        // "three point two" → "3.2", "0 point five" → "0.5"
        result = System.Text.RegularExpressions.Regex.Replace(result,
            NumPat.Replace("<n>", "<n1>") + @"\s+point\s+" + NumPat.Replace("<n>", "<n2>"),
            m =>
            {
                var n1 = m.Groups["n1"].Value;
                if (NumberWords.TryGetValue(n1, out var d1)) n1 = d1;
                var n2 = m.Groups["n2"].Value;
                if (NumberWords.TryGetValue(n2, out var d2)) n2 = d2;
                return n1 + "." + n2;
            }, IC);

        // ── Unit words → abbreviations (only after a number) ──
        // "12 centimeters" → "12 cm", "three millimeters" → "3 mm"
        var unitMap = new (string pattern, string abbrev)[]
        {
            ("centimeters?", "cm"), ("millimeters?", "mm"), ("meters?", "m"),
        };
        foreach (var (unitPat, abbrev) in unitMap)
        {
            result = System.Text.RegularExpressions.Regex.Replace(result,
                NumPat + @"\s+" + unitPat + @"\b",
                m => NumReplace(m) + " " + abbrev, IC);
        }

        // ── Dimensions: "by" between numbers → "x" ──
        // "3.2 by 2.1 by 1.0" → "3.2 x 2.1 x 1.0"
        // Only when both sides are digits/decimals (post-conversion)
        result = System.Text.RegularExpressions.Regex.Replace(result,
            @"(\d+(?:\.\d+)?)\s+by\s+(\d+(?:\.\d+)?)", "$1 x $2", IC);

        // Handle cross-fragment dimensions:
        // Trailing "by" after a number at end of fragment: "3.2 by" → "3.2 x"
        result = System.Text.RegularExpressions.Regex.Replace(result,
            @"(\d+(?:\.\d+)?)\s+by\s*$", "$1 x", IC);
        // Leading "by" before a number at start of fragment: "by 2.0" → "x 2.0"
        result = System.Text.RegularExpressions.Regex.Replace(result,
            @"^\s*by\s+(\d+(?:\.\d+)?)", "x $1", IC);

        // ── Dates ──
        // Regex building blocks for day/year matching
        var onesW = @"one|two|three|four|five|six|seven|eight|nine";
        var teensW = @"ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen";
        var tensW = @"twenty|thirty";
        var ordOnesW = @"first|second|third|fourth|fifth|sixth|seventh|eighth|ninth";
        var ordTeensW = @"tenth|eleventh|twelfth|thirteenth|fourteenth|fifteenth|sixteenth|seventeenth|eighteenth|nineteenth";

        // Day: ordinal or cardinal for 1-31 (compound forms listed first for greedy match)
        var dayPat =
            $@"(?:(?:{tensW})\s+(?:{ordOnesW})" +   // "twenty first"
            $@"|(?:{ordOnesW}|{ordTeensW}|twentieth|thirtieth)" + // "fifth", "twentieth"
            $@"|(?:{tensW})\s+(?:{onesW})" +         // "twenty three"
            $@"|(?:{onesW}|{teensW}|twenty|thirty)" + // "five", "fifteen"
            @"|\d{1,2}(?:st|nd|rd|th)?)";            // "5th", "21"

        // Year: spoken year patterns (most specific first)
        var year2020 = $@"twenty\s+twenty(?:\s+(?:{onesW}))?";
        var year2010 = $@"twenty\s+(?:{teensW})";
        var year19xx = $@"nineteen\s+(?:(?:ninety|eighty|seventy|sixty|fifty|forty|thirty|twenty|ten)(?:\s+(?:{onesW}))?|(?:{teensW})|(?:{onesW}))";
        var year2000 = $@"two\s+thousand(?:\s+(?:and\s+)?(?:twenty(?:\s+(?:{onesW}))?|(?:{teensW})|(?:{onesW})))?";
        var yearDigit = @"20[0-9]\d|19\d\d";
        var yearPat = $@"(?:{year2020}|{year2010}|{year19xx}|{year2000}|{yearDigit})";

        // Also match bare two-digit spoken number as year (assumes 20xx): "twenty five" → 2025
        var yearOrBare = $@"(?:{yearPat}|(?:{tensW})\s+(?:{onesW})|(?:{teensW})|(?:{onesW})|\d{{1,2}})";

        var monthNamePat = @"(?:january|february|march|april|may|june|july|august|september|october|november|december)";
        var numMonthPat = $@"(?:oh\s+(?:{onesW})|{onesW}|ten|eleven|twelve|\d{{1,2}})";

        // 1. Month-name + day + year: "january fifth twenty twenty five" → "January 5, 2025"
        result = System.Text.RegularExpressions.Regex.Replace(result,
            $@"\b({monthNamePat})\s+({dayPat})\s*,?\s+({yearPat})\b",
            m =>
            {
                var month = char.ToUpper(m.Groups[1].Value[0]) + m.Groups[1].Value[1..].ToLower();
                var day = TryParseSpokenInt(m.Groups[2].Value);
                var year = TryParseSpokenYear(m.Groups[3].Value);
                if (day >= 1 && day <= 31 && year >= 1900)
                    return $"{month} {day}, {year}";
                return m.Value;
            }, IC);

        // 2. Month-name + year (no day): "january twenty twenty five" → "January 2025"
        result = System.Text.RegularExpressions.Regex.Replace(result,
            $@"\b({monthNamePat})\s+({yearPat})\b",
            m =>
            {
                var month = char.ToUpper(m.Groups[1].Value[0]) + m.Groups[1].Value[1..].ToLower();
                var year = TryParseSpokenYear(m.Groups[2].Value);
                if (year >= 1900)
                    return $"{month} {year}";
                return m.Value;
            }, IC);

        // 3. Month-name + day (no year): "january fifth" → "January 5"
        result = System.Text.RegularExpressions.Regex.Replace(result,
            $@"\b({monthNamePat})\s+({dayPat})\b",
            m =>
            {
                var month = char.ToUpper(m.Groups[1].Value[0]) + m.Groups[1].Value[1..].ToLower();
                var day = TryParseSpokenInt(m.Groups[2].Value);
                if (day >= 1 && day <= 31)
                    return $"{month} {day}";
                return m.Value;
            }, IC);

        // 4. Numeric date with 4-number year anchor: "one fifteen twenty twenty five" → "1/15/2025"
        result = System.Text.RegularExpressions.Regex.Replace(result,
            $@"\b({numMonthPat})\s+({dayPat})\s+({yearPat})\b",
            m =>
            {
                var month = TryParseSpokenInt(m.Groups[1].Value);
                var day = TryParseSpokenInt(m.Groups[2].Value);
                var year = TryParseSpokenYear(m.Groups[3].Value);
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 1900)
                    return $"{month}/{day}/{year}";
                return m.Value;
            }, IC);

        // 5. Numeric date with 3-number year (assumes 20xx): "one fifteen twenty five" → "1/15/2025"
        //    Only matches when the third number could be a year (>= the day value or >= 20)
        result = System.Text.RegularExpressions.Regex.Replace(result,
            $@"\b({numMonthPat})\s+({dayPat})\s+({yearOrBare})\b",
            m =>
            {
                var month = TryParseSpokenInt(m.Groups[1].Value);
                var day = TryParseSpokenInt(m.Groups[2].Value);
                var year = TryParseSpokenYear(m.Groups[3].Value);
                // Only convert if this looks like a valid date
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 2000 && year <= 2099)
                    return $"{month}/{day}/{year}";
                return m.Value;
            }, IC);

        return result;
    }

    private void PerformStrokeDetection(string? accession, string? reportText)
    {
        // Use already-extracted values from ExtractClarioPriorityAndClass()
        bool isStroke = _automationService.IsStrokeStudy;

        // Stroke priority only applies to CT and MRI, not XR/CR/US etc.
        if (isStroke)
        {
            var desc = _mosaicReader.LastDescription;
            if (!string.IsNullOrEmpty(desc))
            {
                bool isCrossSection = desc.StartsWith("CT ", StringComparison.OrdinalIgnoreCase) ||
                                      desc.StartsWith("MR ", StringComparison.OrdinalIgnoreCase) ||
                                      desc.StartsWith("MRI ", StringComparison.OrdinalIgnoreCase) ||
                                      desc.Contains(" CT ", StringComparison.OrdinalIgnoreCase) ||
                                      desc.Contains(" MR ", StringComparison.OrdinalIgnoreCase) ||
                                      desc.Contains(" MRI ", StringComparison.OrdinalIgnoreCase);
                if (!isCrossSection)
                {
                    Logger.Trace($"PerformStrokeDetection: Suppressing stroke for non-CT/MR study: '{desc}'");
                    isStroke = false;
                }
            }
        }

        // Also check clinical history for stroke keywords if enabled
        if (!isStroke && _config.StrokeDetectionUseClinicalHistory && !string.IsNullOrEmpty(reportText))
        {
            // Extract clinical history from report text
            var (_, clinicalHistory) = ClinicalHistoryForm.ExtractClinicalHistoryWithFixInfo(reportText);
            if (!string.IsNullOrEmpty(clinicalHistory))
            {
                isStroke = ContainsStrokeKeywords(clinicalHistory);
                if (isStroke)
                {
                    Logger.Trace("Stroke detection: Found stroke keywords in clinical history");
                }
            }
        }

        // Update tracking and the clinical history form
        _strokeDetectedActive = isStroke;

        if (isStroke)
        {
            Logger.Trace($"Stroke study detected for accession {accession}");

            // Auto-show clinical history if stroke detected (even if setting is off)
            InvokeUI(() =>
            {
                _mainForm.ToggleClinicalHistory(true);
                _clinicalHistoryVisible = true;
                _mainForm.SetStrokeState(true);
            });

            InvokeUI(() => _mainForm.ShowStatusToast("Stroke Protocol Detected", 4000));
        }
        else
        {
            InvokeUI(() => _mainForm.SetStrokeState(false));
        }
    }

    private bool ContainsStrokeKeywords(string text)
    {
        var lowerText = text.ToLowerInvariant();
        // Use config keywords if available, otherwise use defaults
        var keywords = _config.StrokeClinicalHistoryKeywords.Count > 0
            ? _config.StrokeClinicalHistoryKeywords
            : DefaultStrokeKeywords.ToList();
        foreach (var keyword in keywords)
        {
            if (lowerText.Contains(keyword.ToLowerInvariant()))
            {
                return true;
            }
        }
        return false;
    }

    #endregion

    #region Critical Communication Note

    /// <summary>
    /// Get the current accession number.
    /// </summary>
    public string? GetCurrentAccession()
    {
        return _mosaicReader.LastAccession;
    }

    /// <summary>
    /// Create critical communication note for stroke case.
    /// Returns true if created, false if already exists or failed.
    /// </summary>
    public bool CreateStrokeCriticalNote(string? accession)
    {
        if (string.IsNullOrEmpty(accession))
        {
            Logger.Trace("CreateStrokeCriticalNote: No accession provided");
            return false;
        }

        if (_criticalNoteCreatedForAccessions.ContainsKey(accession))
        {
            Logger.Trace($"CreateStrokeCriticalNote: Note already created for {accession}");
            return false; // Already created
        }

        var success = _automationService.CreateCriticalCommunicationNote();
        if (success)
        {
            _criticalNoteCreatedForAccessions.TryAdd(accession, 0);

            // Track critical study for session-based tracker
            TrackCriticalStudy();

            InvokeUI(() =>
            {
                _mainForm.ShowStatusToast("Critical note created", 3000);
                _mainForm.SetNoteCreatedState(true);
            });
        }
        return success;
    }

    /// <summary>
    /// Check if a critical note has already been created for the given accession.
    /// </summary>
    public bool HasCriticalNoteForAccession(string? accession)
    {
        return accession != null && _criticalNoteCreatedForAccessions.ContainsKey(accession);
    }

    #endregion

    #region Session-Wide Tracking

    /// <summary>
    /// Check if clinical history has already been auto-fixed for the given accession this session.
    /// </summary>
    public bool HasClinicalHistoryFixedForAccession(string? accession)
    {
        return accession != null && _clinicalHistoryFixedForAccessions.ContainsKey(accession);
    }

    /// <summary>
    /// Mark that clinical history has been auto-fixed for the given accession.
    /// </summary>
    public void MarkClinicalHistoryFixedForAccession(string? accession)
    {
        if (!string.IsNullOrEmpty(accession))
        {
            _clinicalHistoryFixedForAccessions.TryAdd(accession, 0);
            TrimTrackingSets();
        }
    }

    /// <summary>
    /// Prevent unbounded growth of tracking sets by clearing when they get too large.
    /// </summary>
    private void TrimTrackingSets()
    {
        const int maxSize = 100;
        const int keepSize = 50;
        if (_macrosInsertedForAccessions.Count > maxSize)
        {
            Logger.Trace($"Trimming macro tracking set (was {_macrosInsertedForAccessions.Count})");
            // Keep the most recent entries by removing oldest (arbitrary order for ConcurrentDictionary)
            var toRemove = _macrosInsertedForAccessions.Keys.Take(_macrosInsertedForAccessions.Count - keepSize).ToList();
            foreach (var key in toRemove)
                _macrosInsertedForAccessions.TryRemove(key, out _);
        }
        if (_clinicalHistoryFixedForAccessions.Count > maxSize)
        {
            Logger.Trace($"Trimming clinical history tracking set (was {_clinicalHistoryFixedForAccessions.Count})");
            var toRemove = _clinicalHistoryFixedForAccessions.Keys.Take(_clinicalHistoryFixedForAccessions.Count - keepSize).ToList();
            foreach (var key in toRemove)
                _clinicalHistoryFixedForAccessions.TryRemove(key, out _);
        }
    }

    #endregion

    #region Ignore Inpatient Drafted

    /// <summary>
    /// Check if both auto-insertions (macros and clinical history auto-fix) have completed
    /// for an inpatient XR study, and if so, send Ctrl+A to select all text.
    /// </summary>
    private void TryTriggerIgnoreInpatientDrafted()
    {
        if (_ctrlASentForCurrentAccession) return;

        // Wait for whichever features are enabled to complete
        // If macros enabled, must be done. If auto-fix enabled, must be done.
        if (_config.MacrosEnabled && !_macrosCompleteForCurrentAccession) return;
        if (_config.AutoFixClinicalHistory && !_autoFixCompleteForCurrentAccession) return;

        // Check if study matches criteria
        if (!_config.ShouldIgnoreInpatientDrafted(
                _automationService.LastClarioClass,
                _mosaicReader.LastDescription))
            return;

        // Only trigger for drafted studies
        if (!_mosaicReader.LastDraftedState) return;

        _ctrlASentForCurrentAccession = true;
        Logger.Trace("Ignore Inpatient Drafted: Sending Ctrl+A");

        // Small delay to ensure paste operations settle
        Thread.Sleep(100);

        // Use paste lock to prevent race conditions
        lock (PasteLock)
        {
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(50);
            NativeWindows.SendCtrlA();
        }
    }

    /// <summary>
    /// Mark that macros are complete for current accession (for Ignore Inpatient Drafted feature).
    /// Call this after macros are inserted OR when no macros match the study.
    /// </summary>
    public void MarkMacrosCompleteForCurrentAccession()
    {
        _macrosCompleteForCurrentAccession = true;
        TryTriggerIgnoreInpatientDrafted();
    }

    /// <summary>
    /// Mark that clinical history auto-fix is complete for current accession (for Ignore Inpatient Drafted feature).
    /// Call this after auto-fix pastes OR when auto-fix is not needed.
    /// </summary>
    public void MarkAutoFixCompleteForCurrentAccession()
    {
        _autoFixCompleteForCurrentAccession = true;
        TryTriggerIgnoreInpatientDrafted();
    }

    #endregion


    public void Dispose()
    {
        _stopThread = true;
        _actionEvent.Set();
        _actionThread.Join(500);
        _syncTimer?.Dispose();
        _dictationSyncTimer?.Dispose();
        _scrapeTimer?.Dispose();
        _hidService?.Dispose();
        _keyboardService?.Dispose();
        _automationService.Dispose();
        _pipeService?.Dispose();
        _sttService?.Dispose();  // [CustomSTT]
        _actionEvent.Dispose();
    }
}
