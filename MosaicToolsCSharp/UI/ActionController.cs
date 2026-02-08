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
    private NoteFormatter _noteFormatter;
    private GetPriorService _getPriorService;
    private readonly OcrService _ocrService;
    private readonly PipeService _pipeService;
    private readonly TemplateDatabase _templateDatabase;

    public PipeService PipeService => _pipeService;

    // Action Queue (Must be STA for Clipboard and SendKeys)
    private readonly ConcurrentQueue<ActionRequest> _actionQueue = new();
    private readonly AutoResetEvent _actionEvent = new(false);
    private readonly Thread _actionThread;
    
    // State
    private bool _dictationActive = false;
    private volatile bool _isUserActive = false;
    private int _scrapeRunning = 0; // Reentrancy guard for scrape timer
    private volatile bool _stopThread = false;

    // Impression search state
    private bool _searchingForImpression = false;
    private bool _impressionFromProcessReport = false; // True if opened by Process Report (stays until Sign)
    private DateTime _impressionSearchStartTime; // When we started searching - used for initial delay
    private int NormalScrapeIntervalMs => _config.ScrapeIntervalSeconds * 1000;
    private int _fastScrapeIntervalMs = 1000;
    private int _studyLoadScrapeIntervalMs = 500;

    // PTT (Push-to-talk) state
    private bool _pttBusy = false;
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
    private string? _lastPopupReportText; // Track what's currently displayed in popup for change detection

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
    
    public ActionController(Configuration config, MainForm mainForm)
    {
        _config = config;
        _mainForm = mainForm;
        
        _hidService = new HidService();
        _keyboardService = new KeyboardService();
        _automationService = new AutomationService();
        _noteFormatter = new NoteFormatter(config.DoctorName, config.CriticalFindingsTemplate, config.TargetTimezone);
        _getPriorService = new GetPriorService(_config.ComparisonTemplate);
        _ocrService = new OcrService();
        _pipeService = new PipeService();
        _pipeService.Start();
        _templateDatabase = new TemplateDatabase();

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
                    _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast($"Error: {ex.Message}"));
                }
                finally
                {
                    _isUserActive = false;
                    
                    // Restore focus after action completes
                    NativeWindows.RestorePreviousFocus(50);
                    
                    _mainForm.BeginInvoke(() => _mainForm.EnsureWindowsOnTop());
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
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast(msg));

        _hidService.SetPreferredDevice(_config.PreferredMicrophone);
        _hidService.Start();

        // Register hotkeys (skip in headless mode)
        if (!App.IsHeadless)
        {
            RegisterHotkeys();
            _keyboardService.Start();
        }

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

        // Restart scraper to pick up any interval changes
        ToggleMosaicScraper(true);
    }

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
        // If PTT is on, the Record Button is handled by OnRecordButtonStateChanged
        // and should not trigger its mapped action (usually System Beep or Toggle)
        // "Record Button" = PowerMic, "Record" = SpeechMike
        if (_config.DeadManSwitch && (button == "Record Button" || button == "Record"))
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
        // 1. Dead Man's Switch (Push-to-Talk) Active Logic
        if (_config.DeadManSwitch)
        {
            if (isDown)
            {
                if (!_pttBusy && !_dictationActive)
                {
                    _pttBusy = true;
                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        try { PerformToggleRecord(true, sendKey: true); }
                        finally { _pttBusy = false; }
                    });
                }
            }
            else
            {
                if (!_pttBusy && _dictationActive)
                {
                    _pttBusy = true;
                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        try { 
                            Thread.Sleep(50); 
                            PerformToggleRecord(false, sendKey: true); 
                        }
                        finally { _pttBusy = false; }
                    });
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
            case "__InsertMacros__":
                PerformInsertMacros();
                break;
            case "__InsertPickListText__":
                PerformInsertPickListText();
                break;
        }
    }
    
    #region Action Implementations
    
    private int _syncCheckToken = 0;
    private string PrepareTextForPaste(string text)
    {
        if (_config.SeparatePastedItems && !text.StartsWith("\n"))
            return "\n" + text;
        return text;
    }

    public bool IsAddendumOpen()
    {
        // Primary: check the flag set during ProseMirror scanning (works even when
        // LastFinalReport is stale due to the U+FFFC fallback logic)
        if (_automationService.IsAddendumDetected)
            return true;

        // Fallback: check LastFinalReport text directly
        var report = _automationService.LastFinalReport;
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
        _mainForm.BeginInvoke(() => _mainForm.UpdateIndicatorState(_dictationActive));
        Logger.Trace($"System Beep: State toggled to {(_dictationActive ? "ON" : "OFF")}");

        // 4. Reality Check (Python Parity)
        // We only correct if reality says TRUE but we are FALSE.
        // We wait for the system to settle before checking.
        int currentToken = ++_syncCheckToken;
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
                _mainForm.BeginInvoke(() => _mainForm.UpdateIndicatorState(_dictationActive));
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
                _mainForm.BeginInvoke(() => _mainForm.UpdateIndicatorState(true));
            }
            else
            {
                // 2. Sticky OFF: Require 3 consecutive inactive checks (~750ms) to turn off
                _consecutiveInactiveCount++;
                if (_consecutiveInactiveCount >= 3)
                {
                    _mainForm.BeginInvoke(() => _mainForm.UpdateIndicatorState(false));
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
    
    private void PerformProcessReport(string source = "Manual")
    {
        Logger.Trace($"Process Report (Source: {source})");

        // Mark that Process Report was pressed for this accession (for diff highlighting)
        _processReportPressedForCurrentAccession = true;

        // Safety: Release all modifiers before starting automated sequence
        NativeWindows.KeyUpModifiers();
        Thread.Sleep(50);

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
        
        // Scroll down if enabled (3 rapid Page Down presses)
        // Scroll down if enabled (Smart Scroll)
        if (_config.ScrollToBottomOnProcess)
        {
            int pageDowns = 0;
            
            {
                string? report = _automationService.LastFinalReport;
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
            bool popupAlreadyOpen = _currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible;
            if (popupAlreadyOpen)
            {
                Logger.Trace("Process Report: Skipping auto-show (popup already open), marking as stale");
                var popup = _currentReportPopup;
                _mainForm.BeginInvoke(() => { if (popup != null && !popup.IsDisposed) popup.SetStaleState(true); });
            }
            else
            {
                Logger.Trace("Process Report: Auto-showing report overlay (first time for accession)");
                PerformShowReport();
            }
        }
        else if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible)
        {
            // Popup is open but auto-show is disabled or already done - still mark as stale during processing
            Logger.Trace("Process Report: Marking popup as stale during processing");
            var popup = _currentReportPopup;
            _mainForm.BeginInvoke(() => { if (popup != null && !popup.IsDisposed) popup.SetStaleState(true); });
        }

        // Auto-create critical note for stroke cases if enabled
        if (_config.StrokeAutoCreateNote && _strokeDetectedActive)
        {
            var accession = _automationService.LastAccession;
            if (!HasCriticalNoteForAccession(accession))
            {
                Logger.Trace($"Process Report: Auto-creating critical note for stroke case {accession}");
                CreateStrokeCriticalNote(accession);
            }
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
        _mainForm.BeginInvoke(() => _mainForm.ShowImpressionWindow());

        // Switch to fast scrape rate (1 second)
        RestartScrapeTimer(_fastScrapeIntervalMs);
    }

    private void OnImpressionFound(string impression)
    {
        Logger.Trace("Impression found! Switching to slow scrape mode.");
        _searchingForImpression = false;

        // Update the impression window with content
        _mainForm.BeginInvoke(() => _mainForm.UpdateImpression(impression));

        // Revert to configured scrape rate
        RestartScrapeTimer(NormalScrapeIntervalMs);
    }

    private void RestartScrapeTimer(int intervalMs)
    {
        _scrapeTimer?.Change(intervalMs, intervalMs);
        Logger.Trace($"Scrape timer interval changed to {intervalMs}ms");
    }

    private void PerformSignReport(string source = "Manual")
    {
        Logger.Trace($"Sign Report (Source: {source})");

        // Close report popup if open
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible)
        {
            _mainForm.BeginInvoke(() =>
            {
                try { _currentReportPopup?.Close(); } catch { }
                _currentReportPopup = null;
                _lastPopupReportText = null;
            });
        }

        // Mark current accession as signed for RVUCounter integration
        // This flag will be used when accession changes to send the appropriate notification
        _currentAccessionSigned = true;
        Logger.Trace($"RVUCounter: Marked accession '{_lastNonEmptyAccession}' as signed");

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

        // Close impression window on sign
        if (_config.ShowImpression)
        {
            _searchingForImpression = false;
            _impressionFromProcessReport = false; // Clear manual trigger flag
            _mainForm.BeginInvoke(() => _mainForm.HideImpressionWindow());

            // Restore normal scrape rate
            RestartScrapeTimer(NormalScrapeIntervalMs);
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
            _mainForm.BeginInvoke(() =>
            {
                try { _currentReportPopup?.Close(); } catch { }
                _currentReportPopup = null;
                _lastPopupReportText = null;
            });
        }

        // Remember the accession we're discarding
        var accessionToDiscard = _lastNonEmptyAccession;

        // Mark that discard was explicitly requested via MosaicTools
        // This ensures CLOSED_UNSIGNED is sent even if the scrape loop handles it
        _discardDialogShownForCurrentAccession = true;

        // Perform the UI automation to discard
        bool success = _automationService.ClickDiscardStudy();

        if (success)
        {
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Study discarded", 2000));

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
                _mainForm.BeginInvoke(() => _mainForm.ToggleClinicalHistory(false));
            }

            // Hide indicator window if configured to hide when no study
            if (_config.HideIndicatorWhenNoStudy && _config.IndicatorEnabled)
            {
                _mainForm.BeginInvoke(() => _mainForm.ToggleIndicator(false));
            }
        }
        else
        {
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Discard failed - try manually", 3000));
        }

        // Close impression window on discard (same as sign)
        if (_config.ShowImpression)
        {
            _searchingForImpression = false;
            _impressionFromProcessReport = false;
            _mainForm.BeginInvoke(() => _mainForm.HideImpressionWindow());
        }

    }

    private void PerformCreateImpression()
    {
        Logger.Trace("Create Impression (via UI Automation)");

        var success = _automationService.ClickCreateImpression();
        if (success)
        {
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Create Impression", 1500));
        }
        else
        {
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Create Impression button not found", 2500));
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
        var accession = _automationService.LastAccession;
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

    private string? _pendingMacroText = null;
    private int _pendingMacroCount = 0;
    private string? _pendingMacroInsertAccession = null;

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
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Cannot paste into addendum", 2500));
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
                // Copy macro text to clipboard
                ClipboardService.SetText(PrepareTextForPaste(text));
                Thread.Sleep(50);

                // Activate Mosaic and paste
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(50);

                // Focus Transcript box to ensure paste goes to correct location
                _automationService.FocusTranscriptBox();
                Thread.Sleep(50);

                // Paste
                NativeWindows.SendHotkey("ctrl+v");
                Thread.Sleep(100); // Increased for reliability

                // Mark this accession as having had macros inserted (session-wide tracking)
                if (!string.IsNullOrEmpty(accessionToMark))
                {
                    _macrosInsertedForAccessions.TryAdd(accessionToMark, 0);
                    TrimTrackingSets();
                }

                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast(
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
    private string? _pendingPickListText;

    private void PerformShowPickLists()
    {
        Logger.Trace("Show Pick Lists");

        // Check if pick lists are enabled
        if (!_config.PickListsEnabled)
        {
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Pick lists are disabled", 2000));
            return;
        }

        // Check if there are any pick lists
        if (_config.PickLists.Count == 0)
        {
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("No pick lists configured", 2000));
            return;
        }

        // Get current study description
        var studyDescription = _automationService.LastDescription;

        // Filter pick lists by study criteria
        var matchingLists = _config.PickLists
            .Where(pl => pl.Enabled && pl.MatchesStudy(studyDescription))
            .ToList();

        if (matchingLists.Count == 0)
        {
            var studyInfo = string.IsNullOrEmpty(studyDescription) ? "no study" : $"'{studyDescription}'";
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast($"No pick lists match {studyInfo}", 2500));
            return;
        }

        Logger.Trace($"Pick Lists: {matchingLists.Count} list(s) match study '{studyDescription}'");

        // If keep-open is enabled and popup exists, just bring it to front
        if (_config.PickListKeepOpen && _currentPickListPopup != null && !_currentPickListPopup.IsDisposed)
        {
            _mainForm.BeginInvoke(() =>
            {
                _currentPickListPopup?.Activate();
                _currentPickListPopup?.Focus();
            });
            return;
        }

        // Close existing popup if open (and keep-open is disabled)
        if (_currentPickListPopup != null && !_currentPickListPopup.IsDisposed)
        {
            _mainForm.BeginInvoke(() =>
            {
                try { _currentPickListPopup?.Close(); } catch { }
                _currentPickListPopup = null;
            });
        }

        // Show popup on UI thread
        _mainForm.BeginInvoke(() =>
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
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Cannot paste into addendum", 2500));
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
                // Copy text to clipboard
                ClipboardService.SetText(PrepareTextForPaste(text));
                Thread.Sleep(50);

                // Activate Mosaic and paste
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(50);

                // Focus Transcript box to ensure paste goes to correct location
                _automationService.FocusTranscriptBox();
                Thread.Sleep(50);

                // Paste
                NativeWindows.SendHotkey("ctrl+v");
                Thread.Sleep(100);

                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Pick list item inserted", 1500));
            }
            catch (Exception ex)
            {
                Logger.Trace($"InsertPickListText error: {ex.Message}");
                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast($"Error: {ex.Message}", 2500));
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
        _isUserActive = true;
        Logger.Trace("Get Prior");

        if (IsAddendumOpen())
        {
            Logger.Trace("GetPrior: Blocked - addendum is open");
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Cannot paste into addendum", 2500));
            _isUserActive = false;
            return;
        }

        _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Extracting Prior..."));

        try
        {
            // Check if InteleViewer is active
            var foreground = NativeWindows.GetForegroundWindow();
            var activeTitle = NativeWindows.GetWindowTitle(foreground).ToLower();
            
            if (!activeTitle.Contains("inteleviewer"))
            {
                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("InteleViewer must be active!"));
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
                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("No text retrieved"));
                return;
            }
            
            Logger.Trace($"Raw prior text: {rawText.Substring(0, Math.Min(100, rawText.Length))}...");
            
            // Process
            var formatted = _getPriorService.ProcessPriorText(rawText);
            if (string.IsNullOrEmpty(formatted))
            {
                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Could not parse prior"));
                return;
            }
            
            // Paste into Mosaic (with leading and trailing newline for cleaner insertion)
            ClipboardService.SetText(PrepareTextForPaste(formatted + "\n"));
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);

            // Focus Transcript box to ensure paste goes to correct location
            _automationService.FocusTranscriptBox();
            Thread.Sleep(100);

            NativeWindows.SendHotkey("ctrl+v");

            Logger.Trace($"Get Prior complete: {formatted}");
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Prior inserted"));
        }
        finally
        {
            _isUserActive = false;
        }
    }
    
    private void PerformCriticalFindings(bool debugMode = false)
    {
        _isUserActive = true;

        if (debugMode)
        {
            // Debug mode: scrape but show dialog instead of pasting
            Logger.Trace("Critical Findings DEBUG MODE");
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Debug mode: Scraping Clario...", 2000));

            try
            {
                var rawNote = _automationService.PerformClarioScrape(msg =>
                {
                    _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast(msg));
                });

                // Update clinical history window if visible
                _mainForm.BeginInvoke(() => _mainForm.UpdateClinicalHistory(rawNote));

                var formatted = rawNote != null ? _noteFormatter.FormatNote(rawNote) : "No note found";

                _mainForm.ShowDebugResults(rawNote ?? "None", formatted);
                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast(
                    "Debug complete. Review the raw data above to troubleshoot extraction.", 10000));
            }
            finally
            {
                _isUserActive = false;
            }
            return;
        }

        // Normal mode: scrape and paste
        Logger.Trace("Critical Findings (Clario Scrape)");
        _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Scraping Clario..."));

        try
        {
            // Scrape with repeating toast callback
            var rawNote = _automationService.PerformClarioScrape(msg =>
            {
                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast(msg));
            });

            if (string.IsNullOrEmpty(rawNote))
            {
                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("No EXAM NOTE found"));
                return;
            }

            // Update clinical history window if visible
            _mainForm.BeginInvoke(() => _mainForm.UpdateClinicalHistory(rawNote));

            // Format
            var formatted = _noteFormatter.FormatNote(rawNote);

            // Paste into Mosaic
            ClipboardService.SetText(formatted);
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(200);

            NativeWindows.SendHotkey("ctrl+v");

            Logger.Trace($"Critical Findings complete: {formatted}");

            // Remove from critical studies tracker (user has dealt with this study)
            UntrackCriticalStudy();

            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast(
                "Critical findings inserted.\nHold Win key and trigger again to debug.", 20000));
        }
        finally
        {
            _isUserActive = false;
        }
    }

    /// <summary>
    /// Track a critical study entry after successful critical note paste.
    /// </summary>
    private void TrackCriticalStudy()
    {
        if (!_config.TrackCriticalStudies)
            return;

        var accession = _automationService.LastAccession;
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
                PatientName = _automationService.LastPatientName ?? "Unknown",
                SiteCode = _automationService.LastSiteCode ?? "???",
                Description = _automationService.LastDescription ?? "Unknown",
                Mrn = _automationService.LastMrn ?? "",
                CriticalNoteTime = DateTime.Now
            };

            _criticalStudies.Add(entry);
            Logger.Trace($"TrackCriticalStudy: Added entry for {accession} ({entry.PatientName} @ {entry.SiteCode}, MRN={entry.Mrn})");
        }

        // Notify UI
        _mainForm.BeginInvoke(() => CriticalStudiesChanged?.Invoke());
    }

    /// <summary>
    /// Remove the current study from the critical studies tracker (user has dealt with it).
    /// </summary>
    private void UntrackCriticalStudy()
    {
        if (!_config.TrackCriticalStudies)
            return;

        var accession = _automationService.LastAccession;
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
        _mainForm.BeginInvoke(() => CriticalStudiesChanged?.Invoke());
    }

    public void RemoveCriticalStudy(CriticalStudyEntry entry)
    {
        if (entry == null) return;
        lock (_criticalStudiesLock)
        {
            _criticalStudies.Remove(entry);
        }
        Logger.Trace($"RemoveCriticalStudy: Manually removed entry for {entry.Accession}");
        _mainForm.BeginInvoke(() => CriticalStudiesChanged?.Invoke());
    }

    // State for Report Popup Toggle
    private ReportPopupForm? _currentReportPopup;

    private void PerformShowReport()
    {
        Logger.Trace("Show Report (scrape method)");

        // Toggle Logic: If open, try click cycle first, then close
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible)
        {
            _mainForm.BeginInvoke(() =>
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
            string? reportText = _automationService.LastFinalReport;

            if (string.IsNullOrEmpty(reportText))
            {
                // Distinguish between "no study open" and "study loading"
                var currentAccession = _automationService.LastAccession;
                if (!string.IsNullOrEmpty(currentAccession))
                {
                    // We have an accession but no report yet - show popup with "loading" message
                    // The scrape timer will auto-update the popup when report becomes available
                    _mainForm.BeginInvoke(() =>
                    {
                        _currentReportPopup = new ReportPopupForm(_config, "Report loading...", null,
                            changesEnabled: _config.ShowReportChanges,
                            correlationEnabled: _config.CorrelationEnabled,
                            baselineIsSectionOnly: false,
                            accession: currentAccession);
                        _lastPopupReportText = null; // Will be set when real report arrives

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
                    _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("No report available (scraping may be disabled)"));
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

            _mainForm.BeginInvoke(() =>
            {
                _currentReportPopup = new ReportPopupForm(_config, reportText, baselineForDiff,
                    changesEnabled: _config.ShowReportChanges,
                    correlationEnabled: _config.CorrelationEnabled,
                    baselineIsSectionOnly: baselineForDiff != null && _baselineIsFromTemplateDb,
                    accession: _automationService.LastAccession);
                _lastPopupReportText = reportText;

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
                    _currentReportPopup.SetStaleState(true);
                }
            });
        }
        catch (Exception ex)
        {
             Logger.Trace($"ShowReport error: {ex.Message}");
             _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Error showing report"));
        }
    }
    
    private void PerformCaptureSeries()
    {
        _isUserActive = true;

        if (IsAddendumOpen())
        {
            Logger.Trace("CaptureSeries: Blocked - addendum is open");
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Cannot paste into addendum", 2500));
            _isUserActive = false;
            return;
        }

        Logger.Trace("Capture Series/Image");
        _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Capturing..."));
        
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
                _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Could not extract series/image"));
                return;
            }
            
            // Clear clipboard and set new text
            ClipboardService.Clear();
            Thread.Sleep(50);
            
            Logger.Trace($"Setting clipboard: {result}");
            var clipSuccess = ClipboardService.SetText(result);
            Logger.Trace($"Clipboard set success: {clipSuccess}");
            
            // Paste into Mosaic
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(200);
            
            NativeWindows.SendHotkey("ctrl+v");
            
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast($"Inserted: {result}"));
        }
        catch (Exception ex)
        {
            Logger.Trace($"CaptureSeries error: {ex.Message}");
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast($"OCR Error: {ex.Message}"));
        }
        finally
        {
            _isUserActive = false;
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

        var accession = _automationService.LastAccession;
        if (string.IsNullOrEmpty(accession))
        {
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("No study loaded", 2000));
            return;
        }

        if (HasCriticalNoteForAccession(accession))
        {
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Critical note already created", 2000));
            return;
        }

        bool success = _automationService.CreateCriticalCommunicationNote();
        if (success)
        {
            _criticalNoteCreatedForAccessions.TryAdd(accession, 0);

            // Track critical study for session-based tracker
            TrackCriticalStudy();

            _mainForm.BeginInvoke(() =>
            {
                _mainForm.ShowStatusToast("Critical note created", 3000);
                _mainForm.SetNoteCreatedState(true);
            });
        }
        else
        {
            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Failed to create critical note - Clario may not be open", 3000));
        }
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
                if (state.HasValue && state.Value != _dictationActive)
                {
                    _dictationActive = state.Value;
                    _mainForm.BeginInvoke(() => _mainForm.UpdateIndicatorState(_dictationActive));
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
            if (Interlocked.CompareExchange(ref _scrapeRunning, 1, 0) != 0) return; // Prevent overlapping scrapes

            try
            {
                // Only check drafted status if we need it for features
                // Also need it for macros since Description extraction happens during drafted check
                bool needDraftedCheck = _config.ShowDraftedIndicator ||
                                        _config.ShowImpression ||
                                        (_config.MacrosEnabled && _config.Macros.Count > 0);

                // Scrape Mosaic for report data
                var reportText = _automationService.GetFinalReportFast(needDraftedCheck);

                // Bail out if user action started during the scrape
                if (_isUserActive) return;

                // Check for new study (non-empty accession different from last non-empty)
                var currentAccession = _automationService.LastAccession;

                // Clear stale Clario data when accession changes, before pipe send
                if (!string.IsNullOrEmpty(currentAccession) && currentAccession != _lastNonEmptyAccession)
                {
                    _automationService.ClearStrokeState();
                }

                // Send study data over pipe (only sends when changed, record equality)
                _pipeService.SendStudyData(new StudyDataMessage(
                    Type: "study_data",
                    Accession: currentAccession,
                    Description: _automationService.LastDescription,
                    TemplateName: _automationService.LastTemplateName,
                    PatientName: _automationService.LastPatientName,
                    PatientGender: _automationService.LastPatientGender,
                    Mrn: _automationService.LastMrn,
                    SiteCode: _automationService.LastSiteCode,
                    ClarioPriority: _automationService.LastClarioPriority,
                    ClarioClass: _automationService.LastClarioClass,
                    Drafted: _automationService.LastDraftedState,
                    HasCritical: !string.IsNullOrEmpty(currentAccession) && HasCriticalNoteForAccession(currentAccession),
                    Timestamp: DateTime.UtcNow.ToString("o")
                ));

                // Check for discard dialog (RVUCounter integration)
                // Must check BEFORE accession change detection - dialog disappears when user clicks YES
                if (_automationService.IsDiscardDialogVisible())
                {
                    _discardDialogShownForCurrentAccession = true;
                    Logger.Trace("RVUCounter: Discard dialog detected for current accession");
                }

                // Detect study change with flap debounce
                // Mosaic briefly sets accession to empty during study transitions.
                // Without debounce, each flap triggers a full study close + reopen cycle,
                // causing duplicate macro inserts and false RVU counter events.
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

                if (accessionChanged)
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
                    _lastPopupReportText = null;
                    _automationService.ClearLastReport();
                    // Close stale report popup from prior study
                    if (_currentReportPopup != null && !_currentReportPopup.IsDisposed)
                    {
                        var stalePopup = _currentReportPopup;
                        _currentReportPopup = null;
                        _mainForm.BeginInvoke(() => { try { stalePopup.Close(); } catch { } });
                    }
                    _needsBaselineCapture = _config.ShowReportChanges || _config.CorrelationEnabled; // Baseline needed for Changes mode diff AND Rainbow mode dictated-sentence detection
                    // Speed up scraping to catch template flash before auto-processing
                    if (_needsBaselineCapture && !_searchingForImpression)
                        RestartScrapeTimer(_studyLoadScrapeIntervalMs);

                    // Reset alert state tracking
                    _templateMismatchActive = false;
                    _genderMismatchActive = false;
                    _strokeDetectedActive = false;
                    _pendingClarioPriorityRetry = false;

                    // Reset Ignore Inpatient Drafted state
                    _macrosCompleteForCurrentAccession = false;
                    _autoFixCompleteForCurrentAccession = false;
                    _ctrlASentForCurrentAccession = false;

                    // Update tracking - only update to new non-empty accession
                    if (!studyClosed)
                    {
                        _lastNonEmptyAccession = currentAccession;

                        // Baseline capture deferred to scrape timer to catch template flash
                        // (removing immediate capture fixes issue where dictated content like "appendix is normal"
                        // gets captured as baseline if already present in template/macros)
                        Logger.Trace($"New study - ShowReportChanges={_config.ShowReportChanges}, CorrelationEnabled={_config.CorrelationEnabled}, reportText null={string.IsNullOrEmpty(reportText)}");

                        // Show toast (disabled - too noisy)
                        // Logger.Trace($"Showing New Study toast for {currentAccession}");
                        // _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast($"New Study: {currentAccession}", 3000));

                        // Re-show clinical history window if it was hidden due to no study
                        // (only in always-show mode; alerts-only mode will show when alert triggers)
                        if (_config.HideClinicalHistoryWhenNoStudy && _config.ShowClinicalHistory && _config.AlwaysShowClinicalHistory)
                        {
                            _mainForm.BeginInvoke(() => _mainForm.ToggleClinicalHistory(true));
                            _clinicalHistoryVisible = true;
                        }

                        // Re-show indicator window if it was hidden due to no study
                        if (_config.HideIndicatorWhenNoStudy && _config.IndicatorEnabled)
                        {
                            _mainForm.BeginInvoke(() => _mainForm.ToggleIndicator(true));
                        }

                        // Queue macros for insertion - they'll be inserted when clinical history is visible
                        // This handles the case where Mosaic auto-processes the report on open
                        if (_config.MacrosEnabled && _config.Macros.Count > 0)
                        {
                            var studyDescription = _automationService.LastDescription;
                            Logger.Trace($"Macros: Queuing for study '{studyDescription}' ({_config.Macros.Count} macros configured)");
                            _pendingMacroAccession = currentAccession;
                            _pendingMacroDescription = studyDescription;
                        }

                        // Reset clinical history state on study change
                        _mainForm.BeginInvoke(() => _mainForm.OnClinicalHistoryStudyChanged());
                        // Hide impression window on new study
                        _mainForm.BeginInvoke(() => _mainForm.HideImpressionWindow());
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
                    }
                    else
                    {
                        // Study closed without new one opening - clear the tracked accession
                        _lastNonEmptyAccession = null;
                        Logger.Trace("Study closed, no new study opened");

                        // Hide clinical history window if configured to hide when no study
                        // (or always hide in alerts-only mode when no alerts)
                        if (_config.ShowClinicalHistory && (_config.HideClinicalHistoryWhenNoStudy || !_config.AlwaysShowClinicalHistory))
                        {
                            _mainForm.BeginInvoke(() => _mainForm.ToggleClinicalHistory(false));
                            _clinicalHistoryVisible = false;
                        }

                        // Hide indicator window if configured to hide when no study
                        if (_config.HideIndicatorWhenNoStudy && _config.IndicatorEnabled)
                        {
                            _mainForm.BeginInvoke(() => _mainForm.ToggleIndicator(false));
                        }
                    }
                }

                // Capture baseline if needed (deferred from study change detection)
                if (_needsBaselineCapture && !_processReportPressedForCurrentAccession)
                {
                    if (_automationService.LastDraftedState)
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
                                Logger.Trace($"Captured baseline from scrape (DRAFTED, immediate): {reportText.Length} chars, drafted={_automationService.LastDraftedState}");
                                Logger.Trace($"Baseline content: {reportText.Replace("\r", "").Replace("\n", " | ")}");

                                // Record to template database (drafted = reduced weight)
                                if (_config.TemplateDatabaseEnabled && !string.IsNullOrEmpty(_automationService.LastDescription))
                                {
                                    _templateDatabase.RecordTemplate(_automationService.LastDescription, reportText, isDrafted: true);
                                }

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
                                var desc = _automationService.LastDescription;
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
                    else if (!string.IsNullOrEmpty(reportText))
                    {
                        // Non-drafted: wait for impression to appear - report is generated top-to-bottom after Process Report
                        var impression = ImpressionForm.ExtractImpression(reportText);
                        if (!string.IsNullOrEmpty(impression))
                        {
                            _needsBaselineCapture = false;
                            _baselineReport = reportText;
                            Logger.Trace($"Captured baseline from scrape ({reportText.Length} chars)");

                            // Record to template database for future fallback
                            if (_config.TemplateDatabaseEnabled && !string.IsNullOrEmpty(_automationService.LastDescription))
                            {
                                _templateDatabase.RecordTemplate(_automationService.LastDescription, reportText);
                            }

                            // Revert to normal scrape interval now that baseline is captured
                            if (!_searchingForImpression)
                                RestartScrapeTimer(NormalScrapeIntervalMs);
                        }
                    }
                }

                // Detect auto-processing on drafted studies: if baseline was captured and report changed, enable diff
                if (!_draftedAutoProcessDetected && !_processReportPressedForCurrentAccession
                    && _config.ShowReportChanges && _baselineReport != null
                    && _automationService.LastDraftedState
                    && !string.IsNullOrEmpty(reportText) && reportText != _baselineReport)
                {
                    _draftedAutoProcessDetected = true;
                    Logger.Trace($"Drafted study auto-process detected: report changed from baseline ({_baselineReport.Length} → {reportText.Length} chars)");
                    Logger.Trace($"Post-process content: {reportText.Replace("\r", "").Replace("\n", " | ")}");

                }

                // Auto-update popup when report text changes (continuous updates with diff highlighting)
                var popup = _currentReportPopup;
                if (popup != null && !popup.IsDisposed && popup.Visible)
                {
                    if (!string.IsNullOrEmpty(reportText) && reportText != _lastPopupReportText)
                    {
                        Logger.Trace($"Auto-updating popup: report changed ({reportText.Length} chars vs {_lastPopupReportText?.Length ?? 0} chars), baseline={_baselineReport?.Length ?? 0} chars");
                        _lastPopupReportText = reportText;
                        _mainForm.BeginInvoke(() => { if (!popup.IsDisposed) popup.UpdateReport(reportText, _baselineReport, _baselineIsFromTemplateDb); });
                    }
                    else if (string.IsNullOrEmpty(reportText) && !string.IsNullOrEmpty(_lastPopupReportText))
                    {
                        // Report text is gone (being updated in Mosaic) but popup is visible with cached content
                        Logger.Trace("Popup showing stale content - report being updated");
                        _mainForm.BeginInvoke(() => { if (!popup.IsDisposed) popup.SetStaleState(true); });
                    }
                    else if (!string.IsNullOrEmpty(reportText))
                    {
                        // Report text is available and matches last text - clear stale indicator if showing
                        _mainForm.BeginInvoke(() => { if (!popup.IsDisposed) popup.SetStaleState(false); });
                    }
                }

                // Check for pending macros - insert when clinical history becomes visible
                // This handles the case where Mosaic auto-processes the report on study open
                if (!string.IsNullOrEmpty(_pendingMacroAccession) &&
                    _pendingMacroAccession == currentAccession &&
                    !string.IsNullOrWhiteSpace(reportText) &&
                    reportText.Contains("CLINICAL HISTORY", StringComparison.OrdinalIgnoreCase))
                {
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

                // Update clinical history / notification box from Mosaic report
                if (_config.ShowClinicalHistory)
                {
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
                        templateDescription = _automationService.LastDescription;
                        templateName = _automationService.LastTemplateName;
                        bool bodyPartsMatch = AutomationService.DoBodyPartsMatch(templateDescription, templateName);
                        newTemplateMismatch = !bodyPartsMatch;
                    }

                    // Check for gender mismatch
                    if (_config.GenderCheckEnabled && !string.IsNullOrWhiteSpace(reportText))
                    {
                        patientGender = _automationService.LastPatientGender;
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

                    // Determine if any alerts are active
                    bool anyAlertActive = newTemplateMismatch || newGenderMismatch || _strokeDetectedActive;

                    // Handle visibility based on always-show vs alerts-only mode
                    if (_config.AlwaysShowClinicalHistory)
                    {
                        // ALWAYS-SHOW MODE: Current behavior - window always visible, show clinical history + border colors

                        // Only update if we have content - don't clear during brief processing gaps
                        if (!string.IsNullOrWhiteSpace(reportText))
                        {
                            _mainForm.BeginInvoke(() => _mainForm.UpdateClinicalHistory(reportText, currentAccession));
                            _mainForm.BeginInvoke(() => _mainForm.UpdateClinicalHistoryTextColor(reportText));
                        }

                        // Update template mismatch state
                        _mainForm.BeginInvoke(() => _mainForm.UpdateClinicalHistoryTemplateMismatch(newTemplateMismatch, templateDescription, templateName));

                        // Update drafted state (green border when drafted) if enabled
                        if (_config.ShowDraftedIndicator)
                        {
                            bool isDrafted = _automationService.LastDraftedState;
                            _mainForm.BeginInvoke(() => _mainForm.UpdateClinicalHistoryDraftedState(isDrafted));
                        }

                        // Update gender check
                        if (_config.GenderCheckEnabled)
                        {
                            _mainForm.BeginInvoke(() => _mainForm.UpdateGenderCheck(reportText, patientGender));
                        }
                        else
                        {
                            _mainForm.BeginInvoke(() => _mainForm.UpdateGenderCheck(null, null));
                        }
                    }
                    else
                    {
                        // ALERTS-ONLY MODE: Window only appears when an alert triggers

                        // Always update clinical history text even in alerts-only mode
                        // (needed for auto-fix recheck to detect Mosaic self-corrections)
                        if (!string.IsNullOrWhiteSpace(reportText))
                        {
                            _mainForm.BeginInvoke(() => _mainForm.UpdateClinicalHistory(reportText, currentAccession));
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

                        if (anyAlertActive)
                        {
                            // Show notification box with alert
                            if (!_clinicalHistoryVisible)
                            {
                                _mainForm.BeginInvoke(() => _mainForm.ToggleClinicalHistory(true));
                                _clinicalHistoryVisible = true;
                            }

                            // Display alert content
                            if (alertToShow == AlertType.GenderMismatch)
                            {
                                // Gender mismatch uses the blinking display
                                _mainForm.BeginInvoke(() => _mainForm.UpdateGenderCheck(reportText, patientGender));
                            }
                            else if (alertToShow.HasValue)
                            {
                                // Clear gender warning if not active
                                _mainForm.BeginInvoke(() => _mainForm.UpdateGenderCheck(null, null));
                                // Show the alert
                                _mainForm.BeginInvoke(() => _mainForm.ShowAlertOnly(alertToShow.Value, alertDetails));
                            }

                            // Also update template mismatch border (for non-gender alerts)
                            if (alertToShow != AlertType.GenderMismatch)
                            {
                                _mainForm.BeginInvoke(() => _mainForm.UpdateClinicalHistoryTemplateMismatch(newTemplateMismatch, templateDescription, templateName));
                            }
                        }
                        else if (_clinicalHistoryVisible)
                        {
                            // No alerts active - hide the notification box
                            _mainForm.BeginInvoke(() => _mainForm.UpdateGenderCheck(null, null));
                            _mainForm.BeginInvoke(() => _mainForm.ClearAlert());
                            _mainForm.BeginInvoke(() => _mainForm.ToggleClinicalHistory(false));
                            _clinicalHistoryVisible = false;
                        }
                    }

                    // Update tracking for next iteration
                    _templateMismatchActive = newTemplateMismatch;
                    _genderMismatchActive = newGenderMismatch;
                }

                // Handle impression display
                if (_config.ShowImpression)
                {
                    var impression = ImpressionForm.ExtractImpression(reportText);
                    bool isDrafted = _automationService.LastDraftedState;

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
                            _mainForm.BeginInvoke(() => _mainForm.UpdateImpression(impression));
                        }
                    }
                    else if (isDrafted && !string.IsNullOrEmpty(impression))
                    {
                        // Auto-show impression when study is drafted (passive mode)
                        // Only show window if not already visible to avoid flashing
                        _mainForm.BeginInvoke(() =>
                        {
                            _mainForm.ShowImpressionWindowIfNotVisible();
                            _mainForm.UpdateImpression(impression);
                        });
                    }
                    else if (!isDrafted)
                    {
                        // Hide impression window when study is not drafted (only for auto-shown)
                        // Don't hide if it was manually triggered by Process Report
                        _mainForm.BeginInvoke(() => _mainForm.HideImpressionWindow());
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Mosaic scrape error: {ex.Message}");
            }
            finally
            {
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

    private void PerformStrokeDetection(string? accession, string? reportText)
    {
        // Use already-extracted values from ExtractClarioPriorityAndClass()
        bool isStroke = _automationService.IsStrokeStudy;

        // Stroke priority only applies to CT and MRI, not XR/CR/US etc.
        if (isStroke)
        {
            var desc = _automationService.LastDescription;
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
            _mainForm.BeginInvoke(() =>
            {
                _mainForm.ToggleClinicalHistory(true);
                _clinicalHistoryVisible = true;
                _mainForm.SetStrokeState(true);
            });

            _mainForm.BeginInvoke(() => _mainForm.ShowStatusToast("Stroke Protocol Detected", 4000));
        }
        else
        {
            _mainForm.BeginInvoke(() => _mainForm.SetStrokeState(false));
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
        return _automationService.LastAccession;
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

            _mainForm.BeginInvoke(() =>
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
        if (_macrosInsertedForAccessions.Count > maxSize)
        {
            Logger.Trace($"Trimming macro tracking set (was {_macrosInsertedForAccessions.Count})");
            _macrosInsertedForAccessions.Clear();
        }
        if (_clinicalHistoryFixedForAccessions.Count > maxSize)
        {
            Logger.Trace($"Trimming clinical history tracking set (was {_clinicalHistoryFixedForAccessions.Count})");
            _clinicalHistoryFixedForAccessions.Clear();
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
                _automationService.LastDescription))
            return;

        // Only trigger for drafted studies
        if (!_automationService.LastDraftedState) return;

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
    }
}
