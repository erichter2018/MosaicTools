using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
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
    private readonly GetPriorService _getPriorService;
    private readonly OcrService _ocrService;
    
    // Action Queue (Must be STA for Clipboard and SendKeys)
    private readonly ConcurrentQueue<ActionRequest> _actionQueue = new();
    private readonly AutoResetEvent _actionEvent = new(false);
    private readonly Thread _actionThread;
    
    // State
    private bool _dictationActive = false;
    private bool _isUserActive = false;
    private volatile bool _stopThread = false;

    // Impression search state
    private bool _searchingForImpression = false;
    private bool _impressionFromProcessReport = false; // True if opened by Process Report (stays until Sign)
    private DateTime _impressionSearchStartTime; // When we started searching - used for initial delay
    private int NormalScrapeIntervalMs => _config.ScrapeIntervalSeconds * 1000;
    private int _fastScrapeIntervalMs = 1000;
    private int _postImpressionScrapeIntervalMs = 3000;

    // PTT (Push-to-talk) state
    private bool _pttBusy = false;
    private DateTime _lastSyncTime = DateTime.Now;
    private readonly System.Threading.Timer? _syncTimer;
    private System.Threading.Timer? _scrapeTimer;
    private DateTime _lastManualToggleTime = DateTime.MinValue;
    private int _consecutiveInactiveCount = 0; // For indicator debounce

    // Accession tracking - only track non-empty accessions
    private string? _lastNonEmptyAccession;

    /// <summary>
    /// Returns true if a study is currently open (has a non-empty accession).
    /// </summary>
    public bool IsStudyOpen => !string.IsNullOrEmpty(_lastNonEmptyAccession);

    // RVUCounter integration - track whether current accession was signed
    private bool _currentAccessionSigned = false;

    // Report changes tracking - baseline report for diff highlighting
    private string? _baselineReport;
    private bool _processReportPressedForCurrentAccession = false;
    private bool _needsBaselineCapture = false; // Set true on new study, cleared when baseline captured
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
    private bool _clinicalHistoryVisible = false;

    // Critical note tracking - which accession already has a critical note created
    private string? _criticalNoteCreatedForAccession = null;

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
        _noteFormatter = new NoteFormatter(config.DoctorName, config.CriticalFindingsTemplate);
        _getPriorService = new GetPriorService();
        _ocrService = new OcrService();

        _actionThread = new Thread(ActionLoop) { IsBackground = true };
        _actionThread.SetApartmentState(ApartmentState.STA);
        _actionThread.Start();

        // 250ms heartbeat for registry sync (high frequency for "Instant ON")
        _syncTimer = new System.Threading.Timer(OnSyncTimerCallback, null, 250, 250);
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
                if (_config.RestoreFocusAfterAction)
                {
                    NativeWindows.SavePreviousFocus();
                }
                
                try
                {
                    ExecuteAction(req);
                }
                catch (Exception ex)
                {
                    Logger.Trace($"Action loop error ({req.Action}): {ex.Message}");
                    _mainForm.Invoke(() => _mainForm.ShowStatusToast($"Error: {ex.Message}"));
                }
                finally
                {
                    _isUserActive = false;
                    
                    // Restore focus after action completes
                    if (_config.RestoreFocusAfterAction)
                    {
                        NativeWindows.RestorePreviousFocus(50);
                    }
                    
                    _mainForm.Invoke(() => _mainForm.EnsureWindowsOnTop());
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
            _mainForm.Invoke(() => _mainForm.ShowStatusToast(msg));
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
        
        // Start Mosaic scrape timer if enabled
        if (_config.ScrapeMosaicEnabled)
        {
            StartMosaicScrapeTimer();
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
        _noteFormatter = new NoteFormatter(_config.DoctorName, _config.CriticalFindingsTemplate);
        RegisterHotkeys();
        
        // Toggle scraper based on setting
        ToggleMosaicScraper(_config.ScrapeMosaicEnabled);
    }
    
    private void RegisterHotkeys()
    {
        _keyboardService.ClearHotkeys();
        
        foreach (var (action, mapping) in _config.ActionMappings)
        {
            if (!string.IsNullOrEmpty(mapping.Hotkey))
            {
                var actionName = action; // Capture for closure
                _keyboardService.RegisterHotkey(mapping.Hotkey, () => TriggerAction(actionName, "Hotkey"));
            }
        }
    }
    
    private void OnMicButtonPressed(string button)
    {
        // If PTT is on, the Record Button is handled by OnRecordButtonStateChanged 
        // and should not trigger its mapped action (usually System Beep or Toggle)
        if (_config.DeadManSwitch && button == "Record Button")
        {
            return;
        }

        // Find action for this button
        foreach (var (action, mapping) in _config.ActionMappings)
        {
            if (mapping.MicButton == button)
            {
                TriggerAction(action, button);
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
        if (req.Source == "Hotkey")
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
                // Skip Forward is hardcoded in Mosaic to trigger Create Impression
                // Only invoke via UI automation if triggered by other sources
                if (req.Source != "Skip Forward")
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
    private void PerformBeep()
    {
        // 1. Determine STARTING state (The state we are moving TO)
        bool startingActive = !_dictationActive;
        
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
        _mainForm.Invoke(() => _mainForm.UpdateIndicatorState(_dictationActive));
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
                _mainForm.Invoke(() => _mainForm.UpdateIndicatorState(_dictationActive));
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
                _mainForm.Invoke(() => _mainForm.UpdateIndicatorState(true));
            }
            else
            {
                // 2. Sticky OFF: Require 3 consecutive inactive checks (~750ms) to turn off
                _consecutiveInactiveCount++;
                if (_consecutiveInactiveCount >= 3)
                {
                    _mainForm.Invoke(() => _mainForm.UpdateIndicatorState(false));
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
        
        // 2. Conditional Alt+P logic for PowerMic 2 hardcoded buttons
        bool isSkipBackTrigger = (source == "Skip Back");
        bool isSkipBackMappedToProcess = _config.ActionMappings.GetValueOrDefault(Actions.ProcessReport)?.MicButton == "Skip Back";

        if (isSkipBackTrigger && isSkipBackMappedToProcess)
        {
            // If the hardware button is pressed, and it's mapped to Process Report
            if (dictationWasActive)
            {
                // If dictation was ON, the hardware button might fail to process the report.
                // We send it manually AFTER stopping dictation.
                Logger.Trace("Process Report: Dictation was ON + Skip Back. Sending Alt+P manually.");
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(100);
                NativeWindows.SendAltKey('P');
            }
            else
            {
                // Hardware handles it when dictation is OFF.
                Logger.Trace("Process Report: Dictation was OFF + Skip Back. Skipping redundant Alt+P.");
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
            
            // Only smart scroll if we have data from the scraper
            if (_config.ScrapeMosaicEnabled)
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
            else
            {
                 Logger.Trace("Smart Scroll: Scrape disabled. Skipping scroll.");
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
        _mainForm.Invoke(() => _mainForm.ShowImpressionWindow());

        // Switch to fast scrape rate (1 second)
        RestartScrapeTimer(_fastScrapeIntervalMs);
    }

    private void OnImpressionFound(string impression)
    {
        Logger.Trace("Impression found! Switching to slow scrape mode.");
        _searchingForImpression = false;

        // Update the impression window with content
        _mainForm.Invoke(() => _mainForm.UpdateImpression(impression));

        // Switch to slow scrape rate (3 seconds)
        RestartScrapeTimer(_postImpressionScrapeIntervalMs);
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
            _mainForm.Invoke(() =>
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

        // Check if Checkmark button triggered this and is mapped to Sign Report
        // If so, Mosaic handles the actual signing - we only clean up impression state
        bool isCheckmarkTrigger = (source == "Checkmark");
        bool isCheckmarkMappedToSign = _config.ActionMappings.GetValueOrDefault(Actions.SignReport)?.MicButton == "Checkmark";

        if (isCheckmarkTrigger && isCheckmarkMappedToSign)
        {
            Logger.Trace("Sign Report: Checkmark button - Mosaic handles signing, only cleaning up impression.");
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
            _mainForm.Invoke(() => _mainForm.HideImpressionWindow());

            // Restore normal scrape rate
            RestartScrapeTimer(NormalScrapeIntervalMs);
        }
    }

    private void PerformDiscardStudy()
    {
        Logger.Trace("Discard Study action triggered");

        // Close report popup if open
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible)
        {
            _mainForm.Invoke(() =>
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
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Study discarded", 2000));

            // Send CLOSED_UNSIGNED immediately after successful discard
            if (!string.IsNullOrEmpty(accessionToDiscard))
            {
                Logger.Trace($"RVUCounter: Sending CLOSED_UNSIGNED (discard action) for '{accessionToDiscard}'");
                NativeWindows.SendToRvuCounter(NativeWindows.MSG_STUDY_CLOSED_UNSIGNED, accessionToDiscard);

                // Clear state to prevent duplicate message from scrape loop
                _lastNonEmptyAccession = null;
                _discardDialogShownForCurrentAccession = false;
                _currentAccessionSigned = false;
            }

            // Hide clinical history window if configured to hide when no study
            if (_config.HideClinicalHistoryWhenNoStudy && _config.ShowClinicalHistory)
            {
                _mainForm.Invoke(() => _mainForm.ToggleClinicalHistory(false));
            }

            // Hide indicator window if configured to hide when no study
            if (_config.HideIndicatorWhenNoStudy && _config.IndicatorEnabled)
            {
                _mainForm.Invoke(() => _mainForm.ToggleIndicator(false));
            }
        }
        else
        {
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Discard failed - try manually", 3000));
        }

        // Close impression window on discard (same as sign)
        if (_config.ShowImpression)
        {
            _searchingForImpression = false;
            _impressionFromProcessReport = false;
            _mainForm.Invoke(() => _mainForm.HideImpressionWindow());
            RestartScrapeTimer(NormalScrapeIntervalMs);
        }
    }

    private void PerformCreateImpression()
    {
        Logger.Trace("Create Impression (via UI Automation)");

        var success = _automationService.ClickCreateImpression();
        if (success)
        {
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Create Impression", 1500));
        }
        else
        {
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Create Impression button not found", 2500));
        }
    }

    private void InsertMacrosForStudy(string? studyDescription)
    {
        // Find all matching enabled macros
        var matchingMacros = _config.Macros
            .Where(m => m.Enabled && m.MatchesStudy(studyDescription))
            .ToList();

        if (matchingMacros.Count == 0)
        {
            Logger.Trace($"Macros: No macros match study '{studyDescription}'");
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

        // Queue the paste action (needs to run on STA thread)
        TriggerAction("__InsertMacros__", "Internal");
    }

    private string? _pendingMacroText = null;
    private int _pendingMacroCount = 0;

    private void PerformInsertMacros()
    {
        var text = _pendingMacroText;
        var count = _pendingMacroCount;
        _pendingMacroText = null;
        _pendingMacroCount = 0;

        if (string.IsNullOrEmpty(text)) return;

        Logger.Trace($"InsertMacros: Pasting {count} macro(s)");

        // Use paste lock to prevent race conditions with clinical history auto-fix
        lock (PasteLock)
        {
            // Save focus
            var previousWindow = IntPtr.Zero;
            if (_config.RestoreFocusAfterAction)
            {
                previousWindow = NativeWindows.GetForegroundWindow();
            }

            try
            {
                // Copy macro text to clipboard
                ClipboardService.SetText(text);
                Thread.Sleep(50);

                // Activate Mosaic and paste
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(100);

                // Paste
                NativeWindows.SendHotkey("ctrl+v");
                Thread.Sleep(100); // Increased for reliability

                _mainForm.Invoke(() => _mainForm.ShowStatusToast(
                    count == 1 ? "Macro inserted" : $"{count} macros inserted", 2000));
            }
            catch (Exception ex)
            {
                Logger.Trace($"InsertMacros error: {ex.Message}");
            }
            finally
            {
                LastPasteTime = DateTime.Now;

                // Restore focus
                if (_config.RestoreFocusAfterAction && previousWindow != IntPtr.Zero)
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
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Pick lists are disabled", 2000));
            return;
        }

        // Check if there are any pick lists
        if (_config.PickLists.Count == 0)
        {
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("No pick lists configured", 2000));
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
            _mainForm.Invoke(() => _mainForm.ShowStatusToast($"No pick lists match {studyInfo}", 2500));
            return;
        }

        Logger.Trace($"Pick Lists: {matchingLists.Count} list(s) match study '{studyDescription}'");

        // Close existing popup if open
        if (_currentPickListPopup != null && !_currentPickListPopup.IsDisposed)
        {
            _mainForm.Invoke(() =>
            {
                try { _currentPickListPopup?.Close(); } catch { }
                _currentPickListPopup = null;
            });
        }

        // Show popup on UI thread
        _mainForm.Invoke(() =>
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

        Logger.Trace($"InsertPickListText: Pasting {text.Length} chars");

        // Use paste lock to prevent race conditions
        lock (PasteLock)
        {
            // Save focus
            var previousWindow = IntPtr.Zero;
            if (_config.RestoreFocusAfterAction)
            {
                previousWindow = NativeWindows.GetForegroundWindow();
            }

            try
            {
                // Copy text to clipboard
                ClipboardService.SetText(text);
                Thread.Sleep(50);

                // Activate Mosaic and paste
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(100);

                // Paste
                NativeWindows.SendHotkey("ctrl+v");
                Thread.Sleep(100);

                _mainForm.Invoke(() => _mainForm.ShowStatusToast("Pick list item inserted", 1500));
            }
            catch (Exception ex)
            {
                Logger.Trace($"InsertPickListText error: {ex.Message}");
                _mainForm.Invoke(() => _mainForm.ShowStatusToast($"Error: {ex.Message}", 2500));
            }
            finally
            {
                LastPasteTime = DateTime.Now;

                // Restore focus
                if (_config.RestoreFocusAfterAction && previousWindow != IntPtr.Zero)
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
        _mainForm.Invoke(() => _mainForm.ShowStatusToast("Extracting Prior..."));
        
        try
        {
            // Check if InteleViewer is active
            var foreground = NativeWindows.GetForegroundWindow();
            var activeTitle = NativeWindows.GetWindowTitle(foreground).ToLower();
            
            if (!activeTitle.Contains("inteleviewer"))
            {
                _mainForm.Invoke(() => _mainForm.ShowStatusToast("InteleViewer must be active!"));
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
                _mainForm.Invoke(() => _mainForm.ShowStatusToast("No text retrieved"));
                return;
            }
            
            Logger.Trace($"Raw prior text: {rawText.Substring(0, Math.Min(100, rawText.Length))}...");
            
            // Process
            var formatted = _getPriorService.ProcessPriorText(rawText);
            if (string.IsNullOrEmpty(formatted))
            {
                _mainForm.Invoke(() => _mainForm.ShowStatusToast("Could not parse prior"));
                return;
            }
            
            // Paste into Mosaic (with leading newline for cleaner insertion)
            ClipboardService.SetText("\n" + formatted);
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(200);

            NativeWindows.SendHotkey("ctrl+v");

            Logger.Trace($"Get Prior complete: {formatted}");
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Prior inserted"));
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
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Debug mode: Scraping Clario...", 2000));

            try
            {
                var rawNote = _automationService.PerformClarioScrape(msg =>
                {
                    _mainForm.Invoke(() => _mainForm.ShowStatusToast(msg));
                });

                // Update clinical history window if visible
                _mainForm.Invoke(() => _mainForm.UpdateClinicalHistory(rawNote));

                var formatted = rawNote != null ? _noteFormatter.FormatNote(rawNote) : "No note found";

                _mainForm.ShowDebugResults(rawNote ?? "None", formatted);
                _mainForm.Invoke(() => _mainForm.ShowStatusToast(
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
        _mainForm.Invoke(() => _mainForm.ShowStatusToast("Scraping Clario..."));

        try
        {
            // Scrape with repeating toast callback
            var rawNote = _automationService.PerformClarioScrape(msg =>
            {
                _mainForm.Invoke(() => _mainForm.ShowStatusToast(msg));
            });

            if (string.IsNullOrEmpty(rawNote))
            {
                _mainForm.Invoke(() => _mainForm.ShowStatusToast("No EXAM NOTE found"));
                return;
            }

            // Update clinical history window if visible
            _mainForm.Invoke(() => _mainForm.UpdateClinicalHistory(rawNote));

            // Format
            var formatted = _noteFormatter.FormatNote(rawNote);

            // Paste into Mosaic
            ClipboardService.SetText(formatted);
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(200);

            NativeWindows.SendHotkey("ctrl+v");

            Logger.Trace($"Critical Findings complete: {formatted}");
            _mainForm.Invoke(() => _mainForm.ShowStatusToast(
                "Critical findings inserted.\nHold Win key and trigger again to debug.", 20000));
        }
        finally
        {
            _isUserActive = false;
        }
    }
    
    // State for Report Popup Toggle
    private ReportPopupForm? _currentReportPopup;

    private void PerformShowReport()
    {
        Logger.Trace("Show Report (scrape method)");

        // Toggle Logic: If open, close it and return
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible)
        {
            _mainForm.Invoke(() =>
            {
               try { _currentReportPopup.Close(); } catch {}
               _currentReportPopup = null;
               _lastPopupReportText = null;
            });
            return;
        }

        try
        {
            // Use the scraped report text (much faster than Alt+C)
            string? reportText = _automationService.LastFinalReport;

            if (string.IsNullOrEmpty(reportText))
            {
                 _mainForm.Invoke(() => _mainForm.ShowStatusToast("No report available (scraping may be disabled)"));
                 return;
            }

            // Pass baseline for diff highlighting if: feature enabled AND process report was pressed
            string? baselineForDiff = (_config.ShowReportChanges && _processReportPressedForCurrentAccession)
                ? _baselineReport
                : null;

            Logger.Trace($"ShowReport: {reportText.Length} chars, ShowReportChanges={_config.ShowReportChanges}, ProcessPressed={_processReportPressedForCurrentAccession}, BaselineLen={_baselineReport?.Length ?? 0}");
            if (baselineForDiff != null)
            {
                Logger.Trace($"ShowReport: Passing baseline for diff ({baselineForDiff.Length} chars)");
            }

            _mainForm.Invoke(() =>
            {
                _currentReportPopup = new ReportPopupForm(_config, reportText, baselineForDiff);
                _lastPopupReportText = reportText;

                // Handle closure to clear references
                _currentReportPopup.FormClosed += (s, e) =>
                {
                    _currentReportPopup = null;
                    _lastPopupReportText = null;
                };

                _currentReportPopup.Show();
            });
        }
        catch (Exception ex)
        {
             Logger.Trace($"ShowReport error: {ex.Message}");
             _mainForm.Invoke(() => _mainForm.ShowStatusToast("Error showing report"));
        }
    }
    
    private void PerformCaptureSeries()
    {
        _isUserActive = true;
        Logger.Trace("Capture Series/Image");
        _mainForm.Invoke(() => _mainForm.ShowStatusToast("Capturing..."));
        
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
                _mainForm.Invoke(() => _mainForm.ShowStatusToast("Could not extract series/image"));
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
            
            _mainForm.Invoke(() => _mainForm.ShowStatusToast($"Inserted: {result}"));
        }
        catch (Exception ex)
        {
            Logger.Trace($"CaptureSeries error: {ex.Message}");
            _mainForm.Invoke(() => _mainForm.ShowStatusToast($"OCR Error: {ex.Message}"));
        }
        finally
        {
            _isUserActive = false;
        }
    }
    
    #endregion
    
    #region Dictation Sync
    
    private void StartDictationSync()
    {
        var timer = new System.Threading.Timer(_ =>
        {
            if (_isUserActive) return; // Don't check during user actions
            
            try
            {
                bool? state = NativeWindows.IsMicrophoneActiveFromRegistry();
                if (state.HasValue && state.Value != _dictationActive)
                {
                    _dictationActive = state.Value;
                    _mainForm.Invoke(() => _mainForm.UpdateIndicatorState(_dictationActive));
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

            try
            {
                // Only check drafted status if we need it for features
                // Also need it for macros since Description extraction happens during drafted check
                bool needDraftedCheck = _config.ShowDraftedIndicator ||
                                        _config.ShowImpression ||
                                        (_config.MacrosEnabled && _config.Macros.Count > 0);

                // Scrape Mosaic for report data
                var reportText = _automationService.GetFinalReportFast(needDraftedCheck);

                // Check for new study (non-empty accession different from last non-empty)
                var currentAccession = _automationService.LastAccession;

                // Check for discard dialog (RVUCounter integration)
                // Must check BEFORE accession change detection - dialog disappears when user clicks YES
                if (_automationService.IsDiscardDialogVisible())
                {
                    _discardDialogShownForCurrentAccession = true;
                    Logger.Trace("RVUCounter: Discard dialog detected for current accession");
                }

                // Detect study change: either new accession OR accession went empty (study closed)
                bool accessionChanged = false;
                bool studyClosed = false;

                if (!string.IsNullOrEmpty(currentAccession))
                {
                    // New non-empty accession
                    if (currentAccession != _lastNonEmptyAccession)
                    {
                        accessionChanged = true;
                    }
                }
                else if (!string.IsNullOrEmpty(_lastNonEmptyAccession))
                {
                    // Accession went from non-empty to empty (study closed without opening new one)
                    accessionChanged = true;
                    studyClosed = true;
                }

                if (accessionChanged)
                {
                    Logger.Trace($"Study change detected: '{_lastNonEmptyAccession}' -> '{currentAccession ?? "(empty)"}'");

                    // RVUCounter integration: Notify about previous study
                    // Logic:
                    //   1. If _currentAccessionSigned → SIGNED (MosaicTools triggered sign)
                    //   2. Else if discard dialog was shown for this accession → CLOSED_UNSIGNED
                    //   3. Else → SIGNED (no dialog = manual sign via Alt+F or button click)
                    if (!string.IsNullOrEmpty(_lastNonEmptyAccession))
                    {
                        if (_currentAccessionSigned)
                        {
                            // Explicitly signed via MosaicTools
                            Logger.Trace($"RVUCounter: Sending SIGNED (MosaicTools) for '{_lastNonEmptyAccession}'");
                            NativeWindows.SendToRvuCounter(NativeWindows.MSG_STUDY_SIGNED, _lastNonEmptyAccession);
                        }
                        else if (_discardDialogShownForCurrentAccession)
                        {
                            // Discard dialog was shown for this accession → study was discarded
                            Logger.Trace($"RVUCounter: Sending CLOSED_UNSIGNED (dialog was shown) for '{_lastNonEmptyAccession}'");
                            NativeWindows.SendToRvuCounter(NativeWindows.MSG_STUDY_CLOSED_UNSIGNED, _lastNonEmptyAccession);
                        }
                        else
                        {
                            // No dialog → user signed manually (Alt+F or clicked Sign button)
                            Logger.Trace($"RVUCounter: Sending SIGNED (manual) for '{_lastNonEmptyAccession}'");
                            NativeWindows.SendToRvuCounter(NativeWindows.MSG_STUDY_SIGNED, _lastNonEmptyAccession);
                        }
                    }

                    // Reset state for new study
                    _currentAccessionSigned = false;
                    _discardDialogShownForCurrentAccession = false;
                    _processReportPressedForCurrentAccession = false;
                    _criticalNoteCreatedForAccession = null;
                    _baselineReport = null;
                    _lastPopupReportText = null;
                    _needsBaselineCapture = _config.ShowReportChanges; // Will capture on next scrape with report text

                    // Reset alert state tracking
                    _templateMismatchActive = false;
                    _genderMismatchActive = false;
                    _strokeDetectedActive = false;

                    // Update tracking - only update to new non-empty accession
                    if (!studyClosed)
                    {
                        _lastNonEmptyAccession = currentAccession;

                        // Capture baseline report for diff highlighting (if enabled)
                        Logger.Trace($"New study - ShowReportChanges={_config.ShowReportChanges}, reportText null={string.IsNullOrEmpty(reportText)}");
                        if (_config.ShowReportChanges && !string.IsNullOrEmpty(reportText))
                        {
                            _baselineReport = reportText;
                            Logger.Trace($"Captured baseline report ({reportText.Length} chars) for changes tracking");
                        }

                        // Show toast
                        Logger.Trace($"Showing New Study toast for {currentAccession}");
                        _mainForm.Invoke(() => _mainForm.ShowStatusToast($"New Study: {currentAccession}", 3000));

                        // Re-show clinical history window if it was hidden due to no study
                        // (only in always-show mode; alerts-only mode will show when alert triggers)
                        if (_config.HideClinicalHistoryWhenNoStudy && _config.ShowClinicalHistory && _config.AlwaysShowClinicalHistory)
                        {
                            _mainForm.Invoke(() => _mainForm.ToggleClinicalHistory(true));
                            _clinicalHistoryVisible = true;
                        }

                        // Re-show indicator window if it was hidden due to no study
                        if (_config.HideIndicatorWhenNoStudy && _config.IndicatorEnabled)
                        {
                            _mainForm.Invoke(() => _mainForm.ToggleIndicator(true));
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
                        _mainForm.Invoke(() => _mainForm.OnClinicalHistoryStudyChanged());
                        // Hide impression window on new study
                        _mainForm.Invoke(() => _mainForm.HideImpressionWindow());
                        _searchingForImpression = false;
                        _impressionFromProcessReport = false;

                        // Stroke detection - check Clario Priority/Class for stroke protocol
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
                            _mainForm.Invoke(() => _mainForm.ToggleClinicalHistory(false));
                            _clinicalHistoryVisible = false;
                        }

                        // Hide indicator window if configured to hide when no study
                        if (_config.HideIndicatorWhenNoStudy && _config.IndicatorEnabled)
                        {
                            _mainForm.Invoke(() => _mainForm.ToggleIndicator(false));
                        }
                    }
                }

                // Capture baseline if needed (deferred from study change detection)
                // Wait for impression to appear before capturing - that's when report is fully generated
                if (_needsBaselineCapture && !string.IsNullOrEmpty(reportText) && !_processReportPressedForCurrentAccession)
                {
                    var impression = ImpressionForm.ExtractImpression(reportText);
                    if (!string.IsNullOrEmpty(impression))
                    {
                        _needsBaselineCapture = false;
                        _baselineReport = reportText;
                        Logger.Trace($"Captured baseline from scrape ({reportText.Length} chars)");
                    }
                }

                // Auto-update popup when report text changes (continuous updates with diff highlighting)
                if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible &&
                    !string.IsNullOrEmpty(reportText) && reportText != _lastPopupReportText)
                {
                    Logger.Trace($"Auto-updating popup: report changed ({reportText.Length} chars vs {_lastPopupReportText?.Length ?? 0} chars), baseline={_baselineReport?.Length ?? 0} chars");
                    _lastPopupReportText = reportText;
                    _mainForm.Invoke(() => _currentReportPopup?.UpdateReport(reportText, _baselineReport));
                }

                // Check for pending macros - insert when clinical history becomes visible
                // This handles the case where Mosaic auto-processes the report on study open
                if (!string.IsNullOrEmpty(_pendingMacroAccession) &&
                    _pendingMacroAccession == currentAccession &&
                    !string.IsNullOrWhiteSpace(reportText) &&
                    reportText.Contains("CLINICAL HISTORY", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Trace($"Macros: Clinical history now visible, inserting for {_pendingMacroAccession}");
                    InsertMacrosForStudy(_pendingMacroDescription);
                    _pendingMacroAccession = null;
                    _pendingMacroDescription = null;
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

                    // Note: Stroke detection is handled elsewhere (on study change) and sets _strokeDetectedActive

                    // Determine if any alerts are active
                    bool anyAlertActive = newTemplateMismatch || newGenderMismatch || _strokeDetectedActive;

                    // Handle visibility based on always-show vs alerts-only mode
                    if (_config.AlwaysShowClinicalHistory)
                    {
                        // ALWAYS-SHOW MODE: Current behavior - window always visible, show clinical history + border colors

                        // Only update if we have content - don't clear during brief processing gaps
                        if (!string.IsNullOrWhiteSpace(reportText))
                        {
                            _mainForm.Invoke(() => _mainForm.UpdateClinicalHistory(reportText, currentAccession));
                            _mainForm.Invoke(() => _mainForm.UpdateClinicalHistoryTextColor(reportText));
                        }

                        // Update template mismatch state
                        _mainForm.Invoke(() => _mainForm.UpdateClinicalHistoryTemplateMismatch(newTemplateMismatch, templateDescription, templateName));

                        // Update drafted state (green border when drafted) if enabled
                        if (_config.ShowDraftedIndicator)
                        {
                            bool isDrafted = _automationService.LastDraftedState;
                            _mainForm.Invoke(() => _mainForm.UpdateClinicalHistoryDraftedState(isDrafted));
                        }

                        // Update gender check
                        if (_config.GenderCheckEnabled)
                        {
                            _mainForm.Invoke(() => _mainForm.UpdateGenderCheck(reportText, patientGender));
                        }
                        else
                        {
                            _mainForm.Invoke(() => _mainForm.UpdateGenderCheck(null, null));
                        }
                    }
                    else
                    {
                        // ALERTS-ONLY MODE: Window only appears when an alert triggers

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
                                _mainForm.Invoke(() => _mainForm.ToggleClinicalHistory(true));
                                _clinicalHistoryVisible = true;
                            }

                            // Display alert content
                            if (alertToShow == AlertType.GenderMismatch)
                            {
                                // Gender mismatch uses the blinking display
                                _mainForm.Invoke(() => _mainForm.UpdateGenderCheck(reportText, patientGender));
                            }
                            else if (alertToShow.HasValue)
                            {
                                // Clear gender warning if not active
                                _mainForm.Invoke(() => _mainForm.UpdateGenderCheck(null, null));
                                // Show the alert
                                _mainForm.Invoke(() => _mainForm.ShowAlertOnly(alertToShow.Value, alertDetails));
                            }

                            // Also update template mismatch border (for non-gender alerts)
                            if (alertToShow != AlertType.GenderMismatch)
                            {
                                _mainForm.Invoke(() => _mainForm.UpdateClinicalHistoryTemplateMismatch(newTemplateMismatch, templateDescription, templateName));
                            }
                        }
                        else if (_clinicalHistoryVisible)
                        {
                            // No alerts active - hide the notification box
                            _mainForm.Invoke(() => _mainForm.UpdateGenderCheck(null, null));
                            _mainForm.Invoke(() => _mainForm.ClearAlert());
                            _mainForm.Invoke(() => _mainForm.ToggleClinicalHistory(false));
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
                            _mainForm.Invoke(() => _mainForm.UpdateImpression(impression));
                        }
                    }
                    else if (isDrafted && !string.IsNullOrEmpty(impression))
                    {
                        // Auto-show impression when study is drafted (passive mode)
                        // Only show window if not already visible to avoid flashing
                        _mainForm.Invoke(() =>
                        {
                            _mainForm.ShowImpressionWindowIfNotVisible();
                            _mainForm.UpdateImpression(impression);
                        });
                    }
                    else if (!isDrafted)
                    {
                        // Hide impression window when study is not drafted (only for auto-shown)
                        // Don't hide if it was manually triggered by Process Report
                        _mainForm.Invoke(() => _mainForm.HideImpressionWindow());
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Mosaic scrape error: {ex.Message}");
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

    // Keywords for stroke detection in clinical history
    private static readonly string[] StrokeKeywords = new[]
    {
        "stroke", "CVA", "TIA", "hemiparesis", "hemiplegia", "aphasia",
        "dysarthria", "facial droop", "weakness", "numbness", "code stroke",
        "NIH stroke scale", "NIHSS"
    };

    private void PerformStrokeDetection(string? accession, string? reportText)
    {
        bool isStroke = false;

        // Check Clario Priority/Class for stroke protocol
        try
        {
            var priorityData = _automationService.ExtractClarioPriorityAndClass(accession);
            if (priorityData != null)
            {
                Logger.Trace($"Stroke detection: Priority='{priorityData.Priority}', Class='{priorityData.Class}'");
                isStroke = _automationService.IsStrokeStudy;
            }
            else
            {
                Logger.Trace("Stroke detection: Could not extract Clario Priority/Class");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Stroke detection Clario error: {ex.Message}");
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
            _mainForm.Invoke(() =>
            {
                _mainForm.ToggleClinicalHistory(true);
                _clinicalHistoryVisible = true;
                _mainForm.SetStrokeState(true);
            });

            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Stroke Protocol Detected", 4000));
        }
        else
        {
            _mainForm.Invoke(() => _mainForm.SetStrokeState(false));
        }
    }

    private static bool ContainsStrokeKeywords(string text)
    {
        var lowerText = text.ToLowerInvariant();
        foreach (var keyword in StrokeKeywords)
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

        if (_criticalNoteCreatedForAccession == accession)
        {
            Logger.Trace($"CreateStrokeCriticalNote: Note already created for {accession}");
            return false; // Already created
        }

        var success = _automationService.CreateCriticalCommunicationNote();
        if (success)
        {
            _criticalNoteCreatedForAccession = accession;
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Critical note created", 3000));
        }
        return success;
    }

    /// <summary>
    /// Check if a critical note has already been created for the given accession.
    /// </summary>
    public bool HasCriticalNoteForAccession(string? accession)
    {
        return accession != null && _criticalNoteCreatedForAccession == accession;
    }

    #endregion


    public void Dispose()
    {
        _stopThread = true;
        _actionEvent.Set();
        _actionThread.Join(500);
        _syncTimer?.Dispose();
        _scrapeTimer?.Dispose();
        _hidService?.Dispose();
        _keyboardService?.Dispose();
        _automationService.Dispose();
    }
}
