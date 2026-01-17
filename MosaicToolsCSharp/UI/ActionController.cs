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
    
    // PTT (Push-to-talk) state
    private bool _pttBusy = false;
    private DateTime _lastSyncTime = DateTime.Now;
    private readonly System.Threading.Timer? _syncTimer;
    private DateTime _lastManualToggleTime = DateTime.MinValue;
    private int _consecutiveInactiveCount = 0; // For indicator debounce
    
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
                PerformCriticalFindings();
                break;
            case Actions.DebugScrape:
                PerformDebugScrape();
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
                PerformSignReport();
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
    }
    
    private void PerformSignReport()
    {
        Logger.Trace("Sign Report");
        NativeWindows.KeyUpModifiers();
        Thread.Sleep(50);

        NativeWindows.ActivateMosaicForcefully();
        Thread.Sleep(100);
        NativeWindows.SendAltKey('F');
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
            
            // Paste into Mosaic
            ClipboardService.SetText(formatted);
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
    
    private void PerformCriticalFindings()
    {
        _isUserActive = true;
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
            
            // Format
            var formatted = _noteFormatter.FormatNote(rawNote);
            
            // Paste into Mosaic
            ClipboardService.SetText(formatted);
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(200);
            
            NativeWindows.SendHotkey("ctrl+v");
            
            Logger.Trace($"Critical Findings complete: {formatted}");
            _mainForm.Invoke(() => _mainForm.ShowStatusToast("Critical findings inserted"));
        }
        finally
        {
            _isUserActive = false;
        }
    }
    
    private void PerformDebugScrape()
    {
        _isUserActive = true;
        Logger.Trace("Debug Scrape");
        _mainForm.Invoke(() => _mainForm.ShowStatusToast("Debug scraping..."));
        
        try
        {
            var rawNote = _automationService.PerformClarioScrape(msg => 
            {
                _mainForm.Invoke(() => _mainForm.ShowStatusToast(msg));
            });
            var formatted = rawNote != null ? _noteFormatter.FormatNote(rawNote) : "No note found";
            
            _mainForm.ShowDebugResults(rawNote ?? "None", formatted);
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
        _isUserActive = true;
        Logger.Trace("Show Report (Alt-C method)");
        
        // Toggle Logic: If open, close it and return
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed && _currentReportPopup.Visible)
        {
            _mainForm.Invoke(() => 
            {
               try { _currentReportPopup.Close(); } catch {}
               _currentReportPopup = null;
            });
            _isUserActive = false;
            return;
        }

        _mainForm.Invoke(() => _mainForm.ShowStatusToast("Getting report..."));
        
        try
        {
            // 1. Activate Mosaic
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(50); 
            
            // 2. Capture OLD content (Diff Strategy)
            // Clearing the clipboard can cause contention if Mosaic is trying to write.
            // Instead, we just watch for the content to *change*.
            string oldText = ClipboardService.GetText() ?? "";
            
            // 3. Send Alt-C
            NativeWindows.SendAltKey('C');
            
            // 4. Wait for clipboard change
            string? reportText = null;
            for (int i = 0; i < 50; i++) // 500ms max
            {
                Thread.Sleep(10);
                string? current = ClipboardService.GetText();
                
                // Success if we got text AND it's different (or we just accept it if old was empty)
                if (!string.IsNullOrEmpty(current) && (current != oldText || string.IsNullOrEmpty(oldText)))
                {
                    reportText = current;
                    break;
                }
            }
            
            // Edge case: If text didn't change (e.g. same report copied again), 
            // fallback to just using what we have after the wait, 
            // assuming the command worked but the content was identical.
            if (string.IsNullOrEmpty(reportText))
            {
                // Give it one last check
                reportText = ClipboardService.GetText();
            }

            if (string.IsNullOrEmpty(reportText))
            {
                 _mainForm.Invoke(() => _mainForm.ShowStatusToast("No report copied (timed out)"));
                 return;
            }

            _mainForm.Invoke(() =>
            {
                _currentReportPopup = new ReportPopupForm(_config, reportText);
                
                // Handle closure to clear reference
                _currentReportPopup.FormClosed += (s, e) => _currentReportPopup = null;
                
                _currentReportPopup.Show();
            });
        }
        catch (Exception ex)
        {
             Logger.Trace($"ShowReport error: {ex.Message}");
             _mainForm.Invoke(() => _mainForm.ShowStatusToast("Error getting report"));
        }
        finally
        {
            _isUserActive = false;
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
    
    
    public void Dispose()
    {
        _stopThread = true;
        _actionEvent.Set();
        _actionThread.Join(500);
        _syncTimer?.Dispose();
        _hidService?.Dispose();
        _keyboardService?.Dispose();
        _automationService.Dispose();
    }
}
