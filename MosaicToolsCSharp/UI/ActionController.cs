using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
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
    private const int THREAD_MODE_BACKGROUND_BEGIN = 0x00010000;
    private const int THREAD_MODE_BACKGROUND_END = 0x00020000;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetCurrentThread();

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetThreadPriority(IntPtr hThread, int nPriority);

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
    private string? _lastPriorRawText;           // Debug: raw text from InteleViewer clipboard
    private string? _lastPriorFormattedText;      // Debug: formatted text after GetPriorService
    private string? _lastCriticalRawNote;         // Debug: raw Clario note text
    private string? _lastCriticalFormattedText;   // Debug: formatted critical findings text
    private readonly OcrService _ocrService;
    private readonly PipeService _pipeService;
    private readonly TemplateDatabase _templateDatabase;
    private readonly AidocService _aidocService;
    private readonly RadAiService? _radAiService;  // [RadAI] — null if RadAI not installed
    private readonly RecoMdService _recoMdService;
    private CdpService? _cdpService;  // [CDP] — null if CDP not enabled
    private SttService? _sttService;  // [CustomSTT] — null if Custom STT not enabled
    private LlmService? _llmService;  // [LLM] — null if LLM Process not enabled
    private StructuredReport? _llmCapturedTemplate; // [LLM] Template captured on new accession for re-processing
    private string? _llmCapturedTemplateAccession;  // [LLM] Accession the captured template belongs to
    private string? _llmCapturedDocStructure;        // [LLM] Doc structure JSON captured alongside template
    private string? _llmDbTemplateOverride;           // [LLM] Clean template from DB (used for drafted studies)
    private Dictionary<string, KeytermLearningService> _keytermLearningByProvider = new(); // [KeytermLearning] per-provider instances
    private EnsembleMetricsForm? _ensembleMetrics; // [Ensemble] live stats popup
    private SttTranscriptComparisonForm? _sttComparisonForm; // [Ensemble] transcript comparison popup
    private readonly List<CorrectionRecord> _pendingCorrections = new(); // [Ensemble] corrections awaiting validation

    // [Ensemble] Per-provider transcript accumulation for comparison view
    private readonly Dictionary<string, System.Text.StringBuilder> _currentProviderTranscripts = new();
    private readonly Dictionary<string, System.Text.StringBuilder> _currentRawProviderTranscripts = new();
    private Dictionary<string, string> _lastProviderTranscripts = new(); // Snapshot from previous dictation
    private Dictionary<string, string> _lastRawProviderTranscripts = new();
    private string? _lastNonEmptyReport; // [Ensemble] snapshot of last non-empty report for correction validation
    private bool _sttDirectPasteActive; // [CustomSTT] Track whether direct paste is active
    private IntPtr _sttPasteTargetWindow; // [CustomSTT] Window that had focus when recording started
    private readonly object _directPasteLock = new(); // [CustomSTT] Serialize direct paste operations

    public PipeService PipeService => _pipeService;
    public CdpService? CdpService => _cdpService;

    /// <summary>
    /// Safe BeginInvoke that silently no-ops if the MainForm handle isn't created.
    /// Prevents crash loops when timers fire during form creation/destruction.
    /// </summary>
    private void InvokeUI(Action action)
    {
        if (_mainForm.IsHandleCreated && !_mainForm.IsDisposed)
            _mainForm.BeginInvoke(action);
    }

    private static string GrokDisplayName(string? modelId) => modelId switch
    {
        "grok-3-mini" => "Grok 3 Mini",
        "grok-4-1-fast-non-reasoning" => "Grok 4.1 Fast",
        "grok-4-1-fast-reasoning" => "Grok 4.1 Fast",
        "grok-3" => "Grok 3",
        "grok-4-0709" => "Grok 4",
        _ => modelId ?? "Grok"
    };

    private string? QuadApiKeyFor(string modelId) => modelId switch
    {
        _ when modelId.StartsWith("gemini", StringComparison.OrdinalIgnoreCase) => _config.LlmApiKey,
        _ when modelId.StartsWith("gpt-", StringComparison.OrdinalIgnoreCase) => _config.LlmOpenAiApiKey,
        _ when modelId.StartsWith("grok-", StringComparison.OrdinalIgnoreCase) => _config.LlmGrokApiKey,
        _ => _config.LlmGroqApiKey // Groq hosts llama, openai/ prefix models
    };

    private static string QuadDisplayName(string modelId)
    {
        if (modelId.StartsWith("gemini", StringComparison.OrdinalIgnoreCase))
            return "Gemini " + modelId.Replace("gemini-", "").Replace("-preview", "").Replace("-", " ");
        if (modelId.StartsWith("gpt-", StringComparison.OrdinalIgnoreCase))
            return "GPT " + modelId.Replace("gpt-", "").Replace("-mini", " Mini").Replace("-nano", " Nano");
        if (modelId.StartsWith("grok-", StringComparison.OrdinalIgnoreCase))
            return GrokDisplayName(modelId);
        return GroqDisplayName(modelId);
    }

    private static string GroqDisplayName(string? modelId) => modelId switch
    {
        "openai/gpt-oss-120b" => "GPT-OSS 120B",
        "meta-llama/llama-4-scout-17b-16e-instruct" => "Llama Scout",
        "llama-3.3-70b-versatile" => "Llama 3.3 70B",
        _ => modelId ?? "Llama"
    };

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
    private long _scrapeStartedTicks; // Watchdog: when current scrape started (Interlocked)
    private volatile bool _stopThread = false;

    // Impression search state
    private bool _searchingForImpression = false;
    private bool _impressionFromProcessReport = false; // True if opened by Process Report (stays until Sign)
    private DateTime _impressionSearchStartTime; // When we started searching - used for initial delay

    // Impression delete state
    private string? _pendingImpressionReplaceText;
    private volatile bool _impressionDeletePending;
    private int NormalScrapeIntervalMs => Math.Max(_config.ScrapeIntervalSeconds * 1000, 3000); // Floor at 3s — heavy UIA calls cause system lag
    private const int ScrapeTimerIntervalMs = 1000; // Timer always 1s; fast/slow path gating handles throttling
    private const int IdleScrapeIntervalMs = 10_000; // Slow polling when no study is open
    private const int DeepIdleScrapeIntervalMs = 30_000; // Very slow polling after extended idle (reduces UIA load)
    private const int SteadyStateReportScrapeMinIntervalMs = 3000; // Throttle unchanged report polling while study is stable
    private const int DictationStopBurstMs = 3000; // Brief burst for final transcript/render updates
    private const int ProcessReportBurstMs = 20000; // Mosaic editor rebuild takes 15-18s
    private const int ShowReportBurstMs = 5000; // Brief burst after opening report popup
    private const int PopupChangeBurstMs = 5000; // Re-arm burst when report changes while popup visible
    private int _consecutiveIdleScrapes = 0; // Counts scrapes with no report content
    private int _scrapesSinceLastGc = 0; // Counter for periodic GC to release FlaUI COM wrappers
    private int _scrapeHeartbeatCount = 0; // Heartbeat counter for diagnostic logging
    private long _lastFullReportScrapeTick64; // Tick64 timestamp of last full GetFinalReportFast call
    private long _reportBurstUntilTick64; // Burst mode deadline for temporary rapid report scraping

    // Fast/slow path split: fast path (cached reads ~0ms) runs every tick,
    // slow path (full UIA scrape) runs on its own throttle schedule.
    private int _fastReadFailCount; // Consecutive TryFastRead failures; auto-disable after threshold
    private const int FastReadFailThreshold = 3;
    private long _lastSlowPathTick64; // Tick64 timestamp of last slow path execution
    private volatile bool _forceSlowPathOnNextTick; // Set by fast path when accession mismatch detected

    // Slow path dormancy: once all study metadata is populated and fast path is healthy,
    // the slow path goes dormant — zero UIA tree walks. Only wakes on:
    //   - Fast path failure (caches cold/stale)
    //   - Accession change detected by fast path
    //   - Burst mode (Process Report, impression search, etc.)
    //   - Aidoc schedule (if Aidoc enabled, lazy 10s ticks for finding updates)
    //   - Discard dialog window (bounded 10s after Process/Sign)
    private bool _slowPathDormant; // True when all metadata populated and fast path is healthy
    private volatile bool _slowPathEverCompleted; // Gate: fast path waits until first slow path builds caches
    private const int AidocDormantIntervalMs = 10_000; // Aidoc-only slow path during dormancy

    // PTT desired-state model — never drops button events
    private volatile bool _pttButtonDown;     // Physical button state (always updated)
    private volatile bool _pttProcessing;     // Whether a start/stop operation is in flight
    private readonly object _pttLock = new(); // Lightweight lock for state transitions
    private DateTime _lastSyncTime = DateTime.Now;
    private readonly System.Threading.Timer? _syncTimer;
    // Dedicated scrape thread (replaces System.Threading.Timer to avoid ThreadPool saturation)
    private System.Threading.Timer? _scrapeTimer;
    private volatile int _scrapeRunning; // Reentrancy guard (0=idle, 1=running)
    private DateTime _lastManualToggleTime = DateTime.MinValue;
    private int _consecutiveInactiveCount = 0; // For indicator debounce

    // Accession tracking - only track non-empty accessions
    private string? _lastNonEmptyAccession;
    private string? _impFixerEditorAccession; // Track which accession has editor fixer buttons

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
    private string? _processReportLastSeenText; // Last report text seen during hold (for stability detection)
    private DateTime _processReportTextStableSince; // When the held text last changed
    private bool _draftedAutoProcessDetected = false; // Set true when drafted study's report changes from baseline
    private bool _autoShowReportDoneForAccession = false; // Only auto-show report overlay once per accession
    private bool _needsBaselineCapture = false; // Set true on new study, cleared when baseline captured
    private bool _baselineIsFromTemplateDb = false; // True when baseline came from template DB fallback (section-only diffing)
    private int _baselineCaptureAttempts = 0; // Tick counter for template DB fallback timing
    private bool _templateRecordedForStudy = false; // Only record template once per study, before Process Report
    private string? _lastSeenTemplateName; // Track template name for rainbow baseline reset on template change
    private string? _lastPopupReportText; // Track what's currently displayed in popup for change detection
    private DateTime _staleSetTime; // When SetStaleState(true) was last triggered
    private DateTime _staleTextStableTime; // When text first matched _lastPopupReportText after going stale

    // RVUCounter integration - track if discard dialog was shown for current accession
    // If dialog was shown while on accession X, and then X changes, X was discarded
    private bool _discardDialogShownForCurrentAccession = false;
    // Discard dialog check throttle — only check during a 10s window after Process/Sign Report
    private long _discardDialogCheckUntilTick64;

    // Pending macro insertion - wait for clinical history to be visible before inserting
    private string? _pendingMacroAccession;
    private string? _pendingMacroDescription;

    // Alert state tracking for alerts-only mode
    private bool _patientMismatchActive = false;
    private bool _templateMismatchActive = false;
    private int _templateCorrectionAttempts = 0;
    private long _templateCorrectionNextRetryTick64 = 0;
    private bool _genderMismatchActive = false;
    private bool _fimMismatchActive = false;
    private string? _lastMismatchCheckReportText;
    private List<MismatchResult>? _lastFimMismatches;
    private bool _consistencyMismatchActive = false;
    private string? _lastConsistencyCheckReportText;
    private string? _lastConsistencyCheckDescription;
    private List<ConsistencyResult>? _lastConsistencyResults;
    private bool _strokeDetectedActive = false;
    private bool _pendingClarioPriorityRetry = false;
    private int _clarioPriorityRetryCount = 0;
    private long _nextClarioPriorityRetryTick64 = 0;
    private bool _clinicalHistoryVisible = false;

    // Aidoc state tracking
    private string? _lastAidocFindings; // comma-joined relevant finding types
    private List<string>? _lastAidocRelevantList; // preserved for verification on subsequent ticks
    private bool _lastAidocRelevant = false;
    private bool _aidocDoneForStudy = false; // Once scraped successfully, don't re-scrape
    private long _lastAidocScrapeTick64; // Throttle: minimum 5s between Aidoc scrapes
    private long _lastColumnRatioReadTick64; // Throttle: read back column ratio every ~30s
    private bool _scrollDiagLogged; // One-shot: dump DOM after scroll fix for debugging
    private bool _cdpAlertFlashApplied; // Tracks whether flashing classes are currently applied
    private string? _cdpAlertFlashSignature; // Last terms/report signature pushed to CDP
    private long _lastCdpAlertFlashApplyTick64; // Periodic refresh in case Mosaic re-renders editor DOM

    // Highlight mode buttons (OFF / STT / RAINBOW) in editor
    private string _currentHighlightMode = "regular";
    private bool _highlightModeAppliedForStudy; // Reset on accession change to re-apply mode effects
    private string? _lastRainbowReportText; // Track report text to recompute rainbow on changes
    private long _rainbowTextChangedTick64; // Debounce: tick when report text last changed in rainbow mode

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
        _aidocService = new AidocService(_automationService);
        _radAiService = RadAiService.TryCreate();  // [RadAI]

        // [CDP] Initialize Chrome DevTools Protocol service
        if (_config.CdpEnabled)
        {
            _cdpService = new CdpService();
            _cdpService.ColumnRatio = _config.CdpColumnRatio;
            _cdpService.AutoScrollEnabled = _config.CdpAutoScrollEnabled;
            _cdpService.HideDragHandles = _config.CdpHideDragHandles;
            _cdpService.VisualEnhancements = _config.CdpVisualEnhancements;
            if (!_config.CdpEnvVarSet)
            {
                CdpService.EnsureEnvVar();
                _config.CdpEnvVarSet = true;
                _config.Save();
            }
        }

        // [LLM] Initialize LLM service for Custom Process Report
        if (_config.LlmProcessEnabled && !string.IsNullOrEmpty(_config.LlmApiKey))
        {
            _llmService = new LlmService();
            _llmService.Configure(_config.LlmApiKey, _config.LlmModel, _config.LlmProvider, _config.LlmOpenAiApiKey, _config.LlmGroqApiKey, _config.LlmGrokApiKey);
        }

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

        // Ruler overlay: start if enabled (its own timer handles detection)
        TryShowRulerOverlay();

        // [CDP] Toast only on first-time enable (env var just set in constructor above)
        // Don't nag every startup — the scrape loop will silently connect when Mosaic is ready

        // Load template database and prune stale entries
        if (_config.TemplateDatabaseEnabled)
        {
            _templateDatabase.Load();
            _templateDatabase.Cleanup();
        }

        // [KeytermLearning] Initialize per-provider keyterm auto-learning
        if (_config.SttKeytermLearningEnabled && _config.CustomSttEnabled)
        {
            if (_config.SttEnsembleEnabled)
            {
                // Ensemble: learning instances for active providers
                foreach (var prov in new[] { "deepgram", _config.SttEnsembleSecondary1, _config.SttEnsembleSecondary2 })
                {
                    if (prov == "none") continue;
                    var svc = new KeytermLearningService(prov);
                    svc.Load();
                    _keytermLearningByProvider[prov] = svc;
                }
            }
            else
            {
                // Single-provider: one instance for the active provider
                var provName = _config.SttProvider ?? "deepgram";
                var svc = new KeytermLearningService(provName);
                svc.Load();
                _keytermLearningByProvider[provName] = svc;
            }
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

        // [CDP] Re-initialize on settings change
        // Check if CSS-affecting settings changed BEFORE updating properties (old != new triggers re-inject)
        bool needsScrollReinject = _cdpService != null && _cdpService.ScrollFixActive &&
            (_cdpService.VisualEnhancements != _config.CdpVisualEnhancements
            || _cdpService.HideDragHandles != _config.CdpHideDragHandles);

        if (_config.CdpEnabled)
        {
            if (_cdpService == null)
            {
                _cdpService = new CdpService();
                if (!_config.CdpEnvVarSet)
                {
                    CdpService.EnsureEnvVar();
                    _config.CdpEnvVarSet = true;
                    _config.Save();
                    InvokeUI(() => _mainForm.ShowStatusToast("CDP: Restart Mosaic to enable direct DOM access.", 8000));
                }
            }
            _cdpService.ColumnRatio = _config.CdpColumnRatio;
            _cdpService.AutoScrollEnabled = _config.CdpAutoScrollEnabled;
            _cdpService.HideDragHandles = _config.CdpHideDragHandles;
            _cdpService.VisualEnhancements = _config.CdpVisualEnhancements;
            if (!_config.CdpFlashingAlertText)
                ClearCdpAlertTextFlashing();
        }
        else if (_cdpService != null)
        {
            ClearCdpAlertTextFlashing();
            _cdpService.Dispose();
            _cdpService = null;
            Logger.Trace("CDP: Disabled and disposed");
        }

        // [CDP] Remove scroll fix if setting was toggled off, or force re-inject on visual/drag-handle changes
        if (_cdpService != null && _cdpService.ScrollFixActive)
        {
            if (!_config.CdpIndependentScrolling)
                _cdpService.RemoveScrollFix();
            else if (needsScrollReinject)
                _cdpService.RemoveScrollFix(); // Force re-inject with new CSS on next tick
        }

        // [KeytermLearning] Re-initialize on settings change
        if (_config.SttKeytermLearningEnabled && _config.CustomSttEnabled)
        {
            _keytermLearningByProvider.Clear();
            if (_config.SttEnsembleEnabled)
            {
                foreach (var prov in new[] { "deepgram", _config.SttEnsembleSecondary1, _config.SttEnsembleSecondary2 })
                {
                    if (prov == "none") continue;
                    var svc = new KeytermLearningService(prov);
                    svc.Load();
                    _keytermLearningByProvider[prov] = svc;
                }
            }
            else
            {
                var provName = _config.SttProvider ?? "deepgram";
                var svc = new KeytermLearningService(provName);
                svc.Load();
                _keytermLearningByProvider[provName] = svc;
            }
        }
        else
        {
            _keytermLearningByProvider.Clear();
        }

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

        // [LLM] Re-initialize on settings change
        if (_config.LlmProcessEnabled && !string.IsNullOrEmpty(_config.LlmApiKey))
        {
            if (_llmService == null)
                _llmService = new LlmService();
            _llmService.Configure(_config.LlmApiKey, _config.LlmModel, _config.LlmProvider, _config.LlmOpenAiApiKey, _config.LlmGroqApiKey, _config.LlmGrokApiKey);
            Logger.Trace($"LLM: Configured (provider={_config.LlmProvider}, model={(_config.LlmProvider == "openai" ? _config.LlmOpenAiModel : _config.LlmProvider == "groq" ? _config.LlmGroqModel : _config.LlmModel)})");
        }
        else
        {
            _llmService?.Dispose();
            _llmService = null;
        }
    }

    // [CustomSTT] Initialize the custom STT service
    private void InitializeSttService()
    {
        // [KeytermLearning] Merge manual + auto-learned keyterms, per-provider
        string? mergedKeyterms = null;
        Dictionary<string, string>? perProviderKeyterms = null;

        if (_config.SttEnsembleEnabled && _keytermLearningByProvider.Count > 0)
        {
            // Ensemble mode: each provider gets its own merged keyterms
            perProviderKeyterms = new Dictionary<string, string>();
            foreach (var prov in new[] { "deepgram", _config.SttEnsembleSecondary1, _config.SttEnsembleSecondary2 })
            {
                if (prov == "none") continue;
                if (_keytermLearningByProvider.TryGetValue(prov, out var learning))
                    perProviderKeyterms[prov] = MergeKeyterms(_config.SttDeepgramKeyterms, learning);
                else
                    perProviderKeyterms[prov] = _config.SttDeepgramKeyterms;
            }
        }
        else if (_keytermLearningByProvider.Count > 0)
        {
            // Single-provider mode: one merged keyterm string
            var provName = _config.SttProvider ?? "deepgram";
            if (_keytermLearningByProvider.TryGetValue(provName, out var learning))
                mergedKeyterms = MergeKeyterms(_config.SttDeepgramKeyterms, learning);
        }

        _sttService = new SttService(_config, mergedKeyterms, perProviderKeyterms);
        var error = _sttService.Initialize();
        if (error != null)
        {
            InvokeUI(() => _mainForm.ShowStatusToast(error));
            Logger.Trace($"SttService: Init error: {error}");
            _sttService.Dispose();
            _sttService = null;
            return;
        }

        // [KeytermLearning] In ensemble mode, subscribe to per-provider raw finals for learning
        _sttService.RawProviderFinalReceived += result =>
        {
            if (result.Words?.Length > 0 && !string.IsNullOrEmpty(result.ProviderName) &&
                _keytermLearningByProvider.TryGetValue(result.ProviderName, out var learning))
            {
                learning.CollectLowConfidenceWords(result.Words,
                    (float)_config.SttKeytermLearningConfidenceThreshold);
            }

            // [Ensemble] Accumulate per-provider transcript for comparison view
            if (result.IsFinal && !string.IsNullOrEmpty(result.Transcript) && !string.IsNullOrEmpty(result.ProviderName))
            {
                lock (_currentRawProviderTranscripts)
                {
                    if (!_currentRawProviderTranscripts.TryGetValue(result.ProviderName, out var sb))
                    {
                        sb = new System.Text.StringBuilder();
                        _currentRawProviderTranscripts[result.ProviderName] = sb;
                    }
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(result.Transcript);
                }
            }
        };

        // [Ensemble] Live metrics popup
        _sttService.EnsembleStatsUpdated += stats =>
        {
            _ensembleMetrics?.UpdateStats(stats);
        };

        // [Ensemble] Collect corrections for validation against signed report
        _sttService.EnsembleCorrectionsEmitted += corrections =>
        {
            lock (_pendingCorrections)
            {
                _pendingCorrections.AddRange(corrections);
                Logger.Trace($"Ensemble: collected {corrections.Count} corrections (total pending: {_pendingCorrections.Count})");
            }
        };

        // [Ensemble] Show/hide metrics popup and track recording state
        _sttService.RecordingStateChanged += recording =>
        {
            if (_config.SttEnsembleEnabled && _config.SttEnsembleShowMetrics)
            {
                InvokeUI(() =>
                {
                    if (_ensembleMetrics == null)
                    {
                        _ensembleMetrics = new EnsembleMetricsForm(_config, _config.SttEnsembleSecondary1, _config.SttEnsembleSecondary2, _config.SttEnsembleAnchor);
                        _ensembleMetrics.SetShowTranscriptCallback(ShowSttTranscriptComparison);
                    }
                    _ensembleMetrics.SetRecording(recording);
                    if (recording && !_ensembleMetrics.Visible)
                        _ensembleMetrics.Show();
                });
            }

            // [Ensemble] Rotate current → last transcripts when a new recording starts
            if (recording)
            {
                lock (_currentProviderTranscripts)
                {
                    // Only rotate if there's actual content
                    if (_currentProviderTranscripts.Count > 0 || _currentRawProviderTranscripts.Count > 0)
                    {
                        _lastProviderTranscripts = new Dictionary<string, string>();
                        foreach (var kv in _currentProviderTranscripts)
                            _lastProviderTranscripts[kv.Key] = kv.Value.ToString();
                        _lastRawProviderTranscripts = new Dictionary<string, string>();
                        lock (_currentRawProviderTranscripts)
                        {
                            foreach (var kv in _currentRawProviderTranscripts)
                                _lastRawProviderTranscripts[kv.Key] = kv.Value.ToString();
                            _currentRawProviderTranscripts.Clear();
                        }
                        _currentProviderTranscripts.Clear();
                    }
                }
            }
        };

        _sttService.TranscriptionReceived += result =>
        {
            // Always update TranscriptionForm
            InvokeUI(() => _mainForm.OnSttTranscriptionReceived(result));

            // [KeytermLearning] In single-provider mode, collect from main transcription events
            if (result.ProviderName != "ensemble" && result.IsFinal && result.Words?.Length > 0)
            {
                var provName = !string.IsNullOrEmpty(result.ProviderName) ? result.ProviderName : (_config.SttProvider ?? "deepgram");
                if (_keytermLearningByProvider.TryGetValue(provName, out var learning))
                    learning.CollectLowConfidenceWords(result.Words,
                        (float)_config.SttKeytermLearningConfidenceThreshold);
            }

            // [Ensemble] Accumulate ensemble transcript for comparison view
            if (result.IsFinal && !string.IsNullOrEmpty(result.Transcript) && result.ProviderName == "ensemble")
            {
                lock (_currentProviderTranscripts)
                {
                    if (!_currentProviderTranscripts.TryGetValue("ensemble", out var sb))
                    {
                        sb = new System.Text.StringBuilder();
                        _currentProviderTranscripts["ensemble"] = sb;
                    }
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(result.Transcript);
                }
            }

            // [CustomSTT] Direct paste: final results go straight into Mosaic's transcript box.
            if (result.IsFinal && result.Words?.Length > 0 && (string.IsNullOrEmpty(result.Transcript) || !_sttDirectPasteActive))
                Logger.Trace($"CustomSTT: SKIPPED paste — transcript={(string.IsNullOrEmpty(result.Transcript) ? "EMPTY" : $"\"{result.Transcript}\"")} pasteActive={_sttDirectPasteActive} words={result.Words.Length}");
            if (result.IsFinal && !string.IsNullOrEmpty(result.Transcript) && _sttDirectPasteActive)
            {
                // [CustomSTT] Check for voice commands/macros on raw transcript BEFORE
                // spoken punctuation (so "process report" isn't mangled by punctuation rules).
                var voiceTrigger = CheckVoiceTrigger(result.Transcript);

                // Use the prefix text (everything before the trigger) for normal processing
                var rawText = voiceTrigger.Kind != VoiceTriggerKind.None
                    ? voiceTrigger.PrefixText
                    : result.Transcript;

                // Apply all STT text transforms — skip for ensemble results (already post-processed by merger)
                var transcript = result.ProviderName == "ensemble"
                    ? rawText
                    : SttTextProcessor.ProcessTranscript(rawText, _config);

                var hasTextToPaste = transcript.Length > 0;
                // Smart spacing is handled by ApplyCdpSmartInsert when CDP is active;
                // space-wrap only needed for standard paste fallback (non-CDP path)
                var textToPaste = (_cdpService?.IsIframeConnected == true)
                    ? transcript
                    : " " + transcript + " ";

                // Capture trigger info and uncertain word indices for use on STA thread
                var triggerKind = voiceTrigger.Kind;
                var triggerAction = voiceTrigger.ActionName;
                var triggerMacro = voiceTrigger.Macro;
                // Only pass confidence indices for full transcript (not voice trigger prefixes)
                var mediumIdx = voiceTrigger.Kind == VoiceTriggerKind.None
                    ? result.MediumConfWordIndices : null;
                var lowIdx = voiceTrigger.Kind == VoiceTriggerKind.None
                    ? result.LowConfWordIndices : null;

                var svc = _automationService;
                var t = new Thread(() =>
                {
                    lock (_directPasteLock)
                    {
                        // [CDP] Skip window activation — CDP inserts via DOM, no focus needed
                        if (_cdpService?.IsIframeConnected != true)
                        {
                            NativeWindows.ActivateMosaicForcefully();
                            Thread.Sleep(100);
                        }

                        // Paste prefix text (if any)
                        if (hasTextToPaste)
                        {
                            InsertTextToFocusedEditor(textToPaste, mediumConfIndices: mediumIdx, lowConfIndices: lowIdx);
                            Thread.Sleep(50);
                            Logger.Trace($"CustomSTT: Direct paste ({textToPaste.Length} chars): \"{(textToPaste.Length > 40 ? textToPaste[..40] + "..." : textToPaste)}\"");
                        }

                        // Execute voice trigger after pasting prefix
                        if (triggerKind == VoiceTriggerKind.Command)
                        {
                            Logger.Trace($"CustomSTT: Voice command detected: \"{triggerAction}\"");
                            TriggerAction(triggerAction!, "Voice");
                            return; // Command takes over — don't restore focus
                        }
                        else if (triggerKind == VoiceTriggerKind.Macro && triggerMacro != null)
                        {
                            Logger.Trace($"CustomSTT: Voice macro triggered: \"{triggerMacro.Name}\"");
                            // Macros insert into transcript (editor 0) — same target as STT dictation
                            InsertTextToFocusedEditor(PrepareTextForPaste(triggerMacro.Text), cdpEditorIndex: 0);
                            Thread.Sleep(50);
                        }

                        // Restore focus to the app the user was in before paste.
                        // Skip restoration while still recording — avoids focus ping-pong
                        // (Mosaic→IV→Mosaic→IV) on every endpointing final, which creates
                        // windows for paste failure. Focus restores after recording stops.
                        if (!_sttDirectPasteActive)
                        {
                            var restoreHwnd = _sttPasteTargetWindow;
                            if (restoreHwnd != IntPtr.Zero && NativeWindows.IsWindow(restoreHwnd))
                            {
                                Thread.Sleep(50);
                                NativeWindows.ActivateWindow(restoreHwnd, 500);
                            }
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

    // [KeytermLearning] Merge manual keyterms with auto-learned ones (manual takes priority)
    private static string MergeKeyterms(string manual, KeytermLearningService learning)
    {
        var manualTerms = manual.Split(new[] { '\n', '\r', ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(t => t.Trim()).Where(t => t.Length > 0).ToList();
        var manualSet = new HashSet<string>(manualTerms, StringComparer.OrdinalIgnoreCase);
        var autoSlots = Math.Max(0, 100 - manualTerms.Count);
        var autoTerms = learning.GetTopKeyterms(autoSlots)
            .Where(t => !manualSet.Contains(t)).Take(autoSlots).ToList();
        if (autoTerms.Count > 0)
            Logger.Trace($"KeytermLearning: Merged {manualTerms.Count} manual + {autoTerms.Count} auto keyterms");
        return string.Join("\n", manualTerms.Concat(autoTerms));
    }

    public void RefreshFloatingToolbar() =>
        InvokeUI(() => _mainForm.RefreshFloatingToolbar(_config.FloatingButtons));

    /// <summary>
    /// Get the name of the currently connected microphone, or null if not connected.
    /// </summary>
    public string? GetConnectedMicrophoneName() => _hidService.ConnectedDeviceName;
    public OcrService GetOcrService() => _ocrService;
    
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
                ThreadPool.QueueUserWorkItem(_ => PerformToggleRecord(null, sendKey: true, restoreFocus: true));
            }
            return;
        }

        // 1. Dead Man's Switch (Push-to-Talk) — desired-state model
        if (_config.DeadManSwitch)
        {
            _pttButtonDown = isDown;
            _lastManualToggleTime = DateTime.Now; // Lock out sync timer immediately
            PttReconcile();
            return;
        }
        
        // 2. Passive Monitoring - Disabled to prevent conflict with mapped actions.
        // State sync is handled by manual actions or the background reality check.
    }

    /// <summary>
    /// Reconciles physical button state with dictation state.
    /// If an operation is already in flight, the finally block will re-check
    /// so no button event is ever lost.
    /// </summary>
    private void PttReconcile()
    {
        bool shouldRecord;
        lock (_pttLock)
        {
            shouldRecord = _pttButtonDown;

            // Already in correct state? Nothing to do.
            if (shouldRecord == _dictationActive) return;

            // Already processing? The finally block will re-check.
            if (_pttProcessing) return;

            _pttProcessing = true;
        }

        ThreadPool.QueueUserWorkItem(_ =>
        {
            try
            {
                if (!shouldRecord)
                    Thread.Sleep(50); // Brief delay on release to absorb accidental bounces
                PerformToggleRecord(shouldRecord, sendKey: true, restoreFocus: true);
            }
            finally
            {
                lock (_pttLock)
                {
                    _pttProcessing = false;
                }
                // Re-check: if button state changed while we were busy, reconcile again
                PttReconcile();
            }
        });
    }


    public void TriggerAction(string action, string source = "Manual")
    {
        Logger.Trace($"TriggerAction Queued: {action} (Source: {source})");

        // Wake up scrape from idle/dormancy — user activity means state may change
        if (_consecutiveIdleScrapes >= 3 || _slowPathDormant)
        {
            _consecutiveIdleScrapes = 0;
            _slowPathDormant = false;
            _forceSlowPathOnNextTick = true;
            Logger.Trace("Scrape wake-up: user action, forcing slow path on next tick");
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
                // Record Button (PowerMic) and Record (SpeechMike) are hardcoded in Mosaic to toggle dictation.
                // Only send Alt+R if triggered by other sources (hotkey, non-Record mic button).
                // Sending Alt+R when Mosaic already handled it causes a double-toggle and visible focus bounce.
                bool isNativeRecordButton = req.Source == "Record Button" || req.Source == "Record";
                PerformToggleRecord(null, sendKey: !isNativeRecordButton);
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
            case Actions.CustomProcessReport:
                PerformCustomProcessReport(req.Source);
                break;
            case Actions.AutoMeasure:
                PerformAutoMeasure();
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
            case "__ManualCorrectTemplate__":
                PerformManualCorrectTemplate();
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

    /// <summary>
    /// Context-aware STT text insertion via CDP. Adjusts capitalization and spacing
    /// based on characters surrounding the cursor in the ProseMirror editor.
    /// </summary>
    private bool ApplyCdpSmartInsert(string text, int editorIndex, int[]? mediumConfIndices = null, int[]? lowConfIndices = null)
    {
        if (_cdpService == null) return false;
        var ctx = _cdpService.GetCursorContext(editorIndex);
        if (ctx == null)
            return _cdpService.InsertContent(editorIndex, text, highlight: _config.SttHighlightDictated, mediumConfIndices: mediumConfIndices, lowConfIndices: lowConfIndices); // fallback to plain insert

        var (before, after, selectedText) = ctx.Value;
        bool hasSelection = selectedText.Length > 0;

        // Smart casing: match the replaced text's casing, or infer from context
        if (text.Length > 0 && char.IsUpper(text[0]))
        {
            if (hasSelection)
            {
                // Match the first char's casing of the text being replaced
                if (char.IsLower(selectedText[0]))
                    text = char.ToLower(text[0]) + text[1..];
            }
            else
            {
                // No selection — lowercase if mid-sentence
                bool startOfSentence = before == '.' || before == '!' || before == '?'
                                    || before == '\n' || before == '\0';
                if (!startOfSentence)
                    text = char.ToLower(text[0]) + text[1..];
            }
        }

        // Smart trailing period: Deepgram adds periods to most utterances.
        // Keep if: at end of content, or replacing text that ended with a period.
        // Strip if: mid-sentence with more text after cursor.
        bool atEndOfContent = after == '\0' || after == '\n';
        bool selectedHadPeriod = hasSelection && selectedText.TrimEnd().EndsWith('.');
        if (text.TrimEnd().EndsWith('.') && !atEndOfContent && !selectedHadPeriod)
            text = text.TrimEnd()[..^1];

        // Smart spacing
        bool beforeIsSpace = before == ' ' || before == '\n' || before == '\0';
        bool afterIsSpace = after == ' ' || after == '\n' || after == '\0';

        // Trim existing spaces from text edges before re-adding
        text = text.Trim();

        // Add leading space if needed (char before is a letter/digit, not already a space)
        if (!beforeIsSpace && (char.IsLetterOrDigit(before) || char.IsPunctuation(before)))
            text = " " + text;

        // Add trailing space if needed (char after is a letter/digit, not already a space)
        if (!afterIsSpace && char.IsLetterOrDigit(after))
            text = text + " ";

        return _cdpService.InsertContent(editorIndex, text, highlight: _config.SttHighlightDictated, mediumConfIndices: mediumConfIndices, lowConfIndices: lowConfIndices);
    }

    private void InsertTextToFocusedEditor(string text, int cdpEditorIndex = -1, int[]? mediumConfIndices = null, int[]? lowConfIndices = null)
    {
        // [CDP] Direct DOM insertion — no clipboard, no focus management
        if (_cdpService?.IsIframeConnected == true)
        {
            int idx = cdpEditorIndex;
            bool isSttPaste = cdpEditorIndex < 0; // auto-detect = STT
            if (idx < 0)
            {
                // Auto-detect: live DOM check, falls back to last-known from scrape loop
                idx = _cdpService.GetFocusedEditorIndex();
                if (idx < 0)
                {
                    Logger.Trace("CDP: No focused editor ever detected, using standard paste");
                    goto standardPaste;
                }
                Logger.Trace($"CDP: Targeting editor {idx} (0=transcript, 1=report)");
            }
            if (isSttPaste)
            {
                // Strip auto-punctuation for transcript when only final-report punctuation is enabled
                if (idx == 0 && !_config.SttAutoPunctuate && _config.SttAutoPunctuateFinalReport)
                {
                    text = SttTextProcessor.StripAutoPunctuation(text);
                    if (string.IsNullOrEmpty(text)) return; // Standalone punctuation stripped entirely
                }

                if (ApplyCdpSmartInsert(text, idx, mediumConfIndices, lowConfIndices))
                    return;
            }
            else
            {
                if (_cdpService.InsertContent(idx, text))
                    return;
            }
            Logger.Trace("CDP: InsertContent failed, falling through to standard insert");
        }
        standardPaste:

        if (_config.ExperimentalUseSendInputInsert)
        {
            var ok = NativeWindows.SendUnicodeText(text);
            if (!ok)
                Logger.Trace($"InsertTextToFocusedEditor: SendInput failed ({text.Length} chars)");
            return;
        }

        // Set clipboard with verification — if the clipboard is locked by another app,
        // SetText fails silently and Ctrl+V pastes stale/wrong content.
        bool clipOk = ClipboardService.SetText(text);
        if (!clipOk)
        {
            Logger.Trace($"InsertTextToFocusedEditor: Clipboard.SetText FAILED after retries ({text.Length} chars). Skipping paste.");
            return;
        }

        // Verify Mosaic is still foreground before sending keystrokes
        if (!NativeWindows.IsMosaicForeground())
        {
            Logger.Trace("InsertTextToFocusedEditor: Mosaic lost foreground before Ctrl+V. Re-activating...");
            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(50);
        }

        Thread.Sleep(50);
        NativeWindows.SendHotkey("ctrl+v");
    }

    /// <summary>
    /// Insert text directly into the transcript editor (index 0) via CDP.
    /// Returns true if successful, false if CDP is unavailable.
    /// </summary>
    public bool InsertToTranscript(string text)
    {
        if (_cdpService?.IsIframeConnected != true) return false;
        return _cdpService.InsertContent(0, text);
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

            // Check registry (UIA fallback disabled — creating UIA3Automation per call is extremely heavy)
            bool? isRealActive = NativeWindows.IsMicrophoneActiveFromRegistry();

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
            // Skip when Custom STT is active — registry reflects Mosaic's native dictation,
            // which is independent of Custom STT. _dictationActive is managed by PerformToggleRecordStt.
            if (!_config.CustomSttEnabled && (DateTime.Now - _lastManualToggleTime).TotalMilliseconds > 500)
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

    private void PerformToggleRecord(bool? desiredState = null, bool sendKey = true, bool restoreFocus = false)
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
            if (restoreFocus)
                NativeWindows.SavePreviousFocus();

            NativeWindows.ActivateMosaicForcefully();
            Thread.Sleep(100);
            NativeWindows.SendAltKey('R');

            // Restore focus immediately so dictation toggle doesn't steal focus
            // from InteleViewer. Only needed for PTT/Record button path which
            // bypasses ActionLoop's own save/restore.
            if (restoreFocus)
            {
                Thread.Sleep(50);
                NativeWindows.RestorePreviousFocus(0);
            }
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
        if (!_dictationActive)
            RequestReportScrapeBurst(DictationStopBurstMs, "Dictation stopped");

        // [CDP] Manage auto-scroll watcher for built-in dictation (Dragon)
        if (_cdpService != null && _config.CdpAutoScrollEnabled)
        {
            if (_dictationActive)
                ThreadPool.QueueUserWorkItem(_ => { try { _cdpService.InjectAutoScrollWatcher(); } catch { } });
            else
                ThreadPool.QueueUserWorkItem(_ => { try { _cdpService.RemoveAutoScrollWatcher(); } catch { } });
        }
    }

    // [Ensemble] Show transcript comparison popup
    private void ShowSttTranscriptComparison(bool showLast)
    {
        Dictionary<string, string> ensemble, raw;
        if (showLast)
        {
            ensemble = _lastProviderTranscripts;
            raw = _lastRawProviderTranscripts;
        }
        else
        {
            ensemble = new Dictionary<string, string>();
            lock (_currentProviderTranscripts)
                foreach (var kv in _currentProviderTranscripts)
                    ensemble[kv.Key] = kv.Value.ToString();
            raw = new Dictionary<string, string>();
            lock (_currentRawProviderTranscripts)
                foreach (var kv in _currentRawProviderTranscripts)
                    raw[kv.Key] = kv.Value.ToString();
        }

        var dgText = raw.GetValueOrDefault("deepgram", "");
        var s1Name = _config.SttEnsembleSecondary1 ?? "soniox";
        var s2Name = _config.SttEnsembleSecondary2 ?? "speechmatics";
        var s1Text = raw.GetValueOrDefault(s1Name, "");
        var s2Text = raw.GetValueOrDefault(s2Name, "");
        var ensText = ensemble.GetValueOrDefault("ensemble", "");

        var s1Display = EnsembleMetricsForm.DisplayNameStatic(s1Name);
        var s2Display = EnsembleMetricsForm.DisplayNameStatic(s2Name);

        InvokeUI(() =>
        {
            if (_sttComparisonForm == null || _sttComparisonForm.IsDisposed)
                _sttComparisonForm = new SttTranscriptComparisonForm(_config);
            _sttComparisonForm.ShowTranscripts(dgText, s1Text, s2Text, ensText, s1Display, s2Display);
        });
    }

    // [CustomSTT] Toggle recording via SttService instead of Mosaic's built-in dictation
    private void PerformToggleRecordStt(bool? desiredState)
    {
        // If a stop is still in progress (ensemble takes ~2.4s), wait for it to finish
        // before checking state. Without this, _recording may still be true from the
        // incomplete stop, causing the early-exit below to think we're already recording.
        if (_sttService!.IsStopping)
        {
            Logger.Trace("CustomSTT: Stop still in progress, waiting...");
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (_sttService.IsStopping && sw.ElapsedMilliseconds < 3000)
                Thread.Sleep(50);
            Logger.Trace($"CustomSTT: Stop completed after {sw.ElapsedMilliseconds}ms wait");
        }

        bool isRecording = _sttService!.IsRecording;

        // If we have a desired state and are already there, correct _dictationActive and skip
        if (desiredState.HasValue && desiredState.Value == isRecording)
        {
            Logger.Trace($"CustomSTT: Already in desired state ({desiredState}). Skipping.");
            _dictationActive = isRecording; // Correct any stale state from registry sync
            return;
        }

        bool newState = desiredState ?? !isRecording;

        if (newState)
        {
            // Immediate audio feedback — beep confirms the button press was registered
            if (_config.SttStartBeepEnabled)
                AudioService.PlayBeepAsync(1000, 200, _config.SttStartBeepVolume);

            // Start recording
            _sttPasteTargetWindow = NativeWindows.GetForegroundWindow(); // [CustomSTT] Save target window before showing TranscriptionForm
            _sttDirectPasteActive = true; // [CustomSTT] Enable direct paste to foreground window
            if (_config.SttShowIndicator) // [CustomSTT] Only show indicator if enabled in settings
                InvokeUI(() => _mainForm.ShowTranscriptionForm());
            Task.Run(async () => await _sttService!.StartRecordingAsync()).Wait(2000);
            _dictationActive = true;
        }
        else
        {
            // Immediate audio feedback — beep confirms stop was registered
            if (_config.SttStopBeepEnabled)
                AudioService.PlayBeepAsync(500, 200, _config.SttStopBeepVolume);

            // Stop recording — primary finalize is fast (~500ms), secondary cleanup is background
            Task.Run(async () => await _sttService!.StopRecordingAsync()).Wait(2000);
            _dictationActive = false;
            // Keep _sttDirectPasteActive true — ensemble merge timer (500ms) may still be running.
            // Clear it in the background task after giving the timer time to fire.

            // Drain any in-flight paste and restore focus in background so the STA
            // action thread is freed up quickly for the next button press.
            var restoreHwnd = _sttPasteTargetWindow;
            var stopTimestamp = Environment.TickCount64;
            _ = Task.Run(() =>
            {
                Thread.Sleep(600); // Wait for ensemble merge timer to fire
                lock (_directPasteLock) { } // drain any in-flight paste
                // Guard: don't clobber _sttDirectPasteActive if a NEW recording started
                // after this stop (user quickly pressed record again)
                if (_dictationActive)
                {
                    Logger.Trace($"CustomSTT: Cleanup task skipped — new recording is active (stop@{stopTimestamp})");
                    return;
                }
                _sttDirectPasteActive = false;
                Logger.Trace($"CustomSTT: Cleanup task set pasteActive=false (stop@{stopTimestamp})");
                InvokeUI(() => _mainForm.HideTranscriptionForm());
                if (restoreHwnd != IntPtr.Zero && NativeWindows.IsWindow(restoreHwnd))
                {
                    Thread.Sleep(50);
                    NativeWindows.ActivateWindow(restoreHwnd, 200);
                }
            });
        }

        Logger.Trace($"CustomSTT: Recording toggled to {newState}");
        if (!newState)
            RequestReportScrapeBurst(DictationStopBurstMs, "CustomSTT stopped");
    }

    private void PerformProcessReport(string source = "Manual")
    {
        Logger.Trace($"Process Report (Source: {source})");
        RequestReportScrapeBurst(ProcessReportBurstMs, "Process Report");
        _discardDialogCheckUntilTick64 = Environment.TickCount64 + 10_000;

        // Mark that Process Report was pressed for this accession (for diff highlighting)
        _processReportPressedForCurrentAccession = true;
        _processReportLastSeenText = null;
        _processReportTextStableSince = DateTime.UtcNow;

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

            // Send Process Report — CDP button click or Alt+P
            if (!(_cdpService?.ClickProcessReport() == true))
            {
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(100);
                NativeWindows.SendAltKey('P');
            }

            // Auto-restart STT after process report
            if (_config.SttAutoStartOnCase)
            {
                Thread.Sleep(300);
                PerformToggleRecordStt(true);
            }
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

            // 2. Send Process Report — CDP button click bypasses all Alt+P logic
            if (_cdpService?.ClickProcessReport() == true)
            {
                Logger.Trace("Process Report: Sent via CDP button click");
            }
            else
            {
                // Conditional Alt+P logic for hardcoded mic buttons
                // PowerMic: Skip Back is hardcoded to Process Report
                // SpeechMike: Ins/Ovr is hardcoded to Process Report
                bool isHardcodedProcessButton = (source == "Skip Back") || (source == "Ins/Ovr");
                var currentMappings = GetCurrentMicMappings();
                bool isButtonMappedToProcess =
                    currentMappings.GetValueOrDefault(Actions.ProcessReport)?.MicButton == "Skip Back" ||
                    currentMappings.GetValueOrDefault(Actions.ProcessReport)?.MicButton == "Ins/Ovr";

                if (isHardcodedProcessButton && isButtonMappedToProcess)
                {
                    if (dictationWasActive)
                    {
                        Logger.Trace($"Process Report: Dictation was ON + {source}. Sending Alt+P manually.");
                        NativeWindows.ActivateMosaicForcefully();
                        Thread.Sleep(100);
                        NativeWindows.SendAltKey('P');
                    }
                    else
                    {
                        Logger.Trace($"Process Report: Dictation was OFF + {source}. Skipping redundant Alt+P.");
                    }
                }
                else
                {
                    NativeWindows.ActivateMosaicForcefully();
                    Thread.Sleep(100);
                    NativeWindows.SendAltKey('P');
                }
            }
        }

        // Invalidate the cached ProseMirror editor element. After Alt+P, Mosaic rebuilds
        // the editor — the old cached element returns stale pre-process text for 10-15s.
        _mosaicReader.InvalidateEditorCache();

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
                _staleSetTime = DateTime.UtcNow;
                _staleTextStableTime = default;
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
                _staleSetTime = DateTime.UtcNow;
                _staleTextStableTime = default;
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

    // [LLM] Custom Process Report — reads transcript + template via CDP, sends to Gemini, writes result back
    private void PerformCustomProcessReport(string source = "Manual")
    {
        Logger.Trace($"Custom Process Report (Source: {source})");
        var sw = System.Diagnostics.Stopwatch.StartNew();

        // Validate prerequisites
        if (_llmService == null || !_llmService.IsConfigured)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Custom Process: LLM not configured"));
            Logger.Trace("Custom Process Report: LLM service not configured");
            return;
        }

        if (_cdpService == null)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Custom Process: CDP not enabled"));
            Logger.Trace("Custom Process Report: CDP service not available");
            return;
        }

        // Clear any previous model badges
        _cdpService.SetEditorModelBadge(1, null);
        _cdpService.SetEditorModelBadges(1, null!, null);

        // Stop STT recording if active (same preamble as PerformProcessReport)
        if (_config.CustomSttEnabled && _sttService != null && _sttService.IsRecording)
        {
            Logger.Trace("Custom Process Report: Stopping STT recording...");
            Task.Run(async () => await _sttService.StopRecordingAsync()).Wait(4000);
            _dictationActive = false;
            if (_config.SttStopBeepEnabled)
                AudioService.PlayBeepAsync(500, 200, _config.SttStopBeepVolume);
            Thread.Sleep(500);
            _sttDirectPasteActive = false;
            InvokeUI(() => _mainForm.HideTranscriptionForm());
        }

        // 1. Read transcript from editor 0
        var transcript = _cdpService.GetEditorText(0);
        if (string.IsNullOrWhiteSpace(transcript))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Custom Process: No transcript in editor"));
            Logger.Trace("Custom Process Report: Empty transcript");
            return;
        }

        // 2. Use eagerly captured template (from scrape loop). Fallback: capture now if scrape missed it.
        //    On re-process, reuse the captured template so the LLM always works from the original.
        var currentAccession = _mosaicReader.LastAccession;
        if (_llmCapturedTemplate == null
            || (!string.IsNullOrEmpty(currentAccession) && _llmCapturedTemplateAccession != currentAccession))
        {
            // Scrape loop didn't capture yet — try now as fallback
            // Guard: skip if currentAccession is null (transient during study transition) and we have a good capture
            var freshStructured = _cdpService.GetStructuredReport();
            if (freshStructured == null || freshStructured.Sections.Count == 0)
            {
                InvokeUI(() => _mainForm.ShowStatusToast("Custom Process: No template in report editor"));
                Logger.Trace("Custom Process Report: No structured report to capture");
                return;
            }
            _llmCapturedTemplate = freshStructured;
            _llmCapturedTemplateAccession = currentAccession;
            _llmCapturedDocStructure = _cdpService.GetDocStructure(1);
            Logger.Trace($"Custom Process Report: Fallback template capture for {currentAccession} ({freshStructured.Sections.Count} sections)");
        }
        else
        {
            Logger.Trace("Custom Process Report: Using eagerly captured template");
        }

        var structured = _llmCapturedTemplate;
        var docStructure = _llmCapturedDocStructure;

        // 3. Build template for LLM (FINDINGS subsections + IMPRESSION only)
        //    Prefix sections (EXAM, TECHNIQUE, etc.) are read LIVE from editor 1 — user may have edited them.
        //    For drafted studies: prefer clean default template from TemplateDatabase over captured report.

        // Sections the LLM should NOT touch
        var preserveSections = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "EXAM", "EXAMINATION", "TECHNIQUE", "CONTRAST", "RADIATION DOSE", "COMPARISON", "CLINICAL HISTORY", "CLINICAL INFORMATION", "HISTORY", "INDICATION", "INDICATIONS" };

        string builtTemplate;
        if (!string.IsNullOrEmpty(_llmDbTemplateOverride))
        {
            // Drafted study: use clean default template from TemplateDatabase
            builtTemplate = _llmDbTemplateOverride;
            Logger.Trace("Custom Process Report: Using DB template override (drafted study)");
        }
        else
        {
            var templateBuilder = new System.Text.StringBuilder();
            foreach (var section in structured.Sections)
            {
                if (preserveSections.Contains(section.Name))
                    continue; // Skip — these are stitched from live editor later

                // Send to LLM: FINDINGS (with subsections), IMPRESSION, etc.
                templateBuilder.AppendLine(section.Name + ":");

                if (section.Subsections.Count > 0)
                {
                    foreach (var sub in section.Subsections)
                    {
                        templateBuilder.AppendLine();
                        templateBuilder.AppendLine(sub.Name + ":");
                        var subText = !string.IsNullOrWhiteSpace(sub.TemplateText) ? sub.TemplateText : sub.FullText;
                        templateBuilder.AppendLine(subText);
                    }
                }
                else
                {
                    var text = !string.IsNullOrWhiteSpace(section.TemplateText) ? section.TemplateText : section.FullText;
                    templateBuilder.AppendLine(text);
                }
                templateBuilder.AppendLine();
            }
            builtTemplate = templateBuilder.ToString();
        }

        // Read prefix sections LIVE from current editor 1 (user may have edited comparison, etc.)
        var prefixBuilder = new System.Text.StringBuilder();
        var liveStructured = _cdpService.GetStructuredReport();
        // Fall back to captured template if live read fails or returns empty (e.g., editor was cleared)
        var liveSource = (liveStructured != null && liveStructured.Sections.Count > 0) ? liveStructured : structured;
        foreach (var section in liveSource.Sections)
        {
            if (!preserveSections.Contains(section.Name)) continue;
            prefixBuilder.AppendLine(section.Name + ":");
            var fullText = !string.IsNullOrWhiteSpace(section.FullText) ? section.FullText : section.TemplateText;
            // Each line must be its own paragraph — blank lines between them signal paragraph
            // breaks to collectBody() in ReplaceEditorContent (e.g., EXAM: study name + date)
            foreach (var line in fullText.Split('\n'))
            {
                prefixBuilder.AppendLine(line.TrimEnd());
                prefixBuilder.AppendLine();
            }
        }

        var safeTemplate = builtTemplate.Trim();
        var scrubbedTranscript = LlmService.ScrubPhi(transcript);
        var preservedPrefix = prefixBuilder.ToString().TrimEnd();

        // Extract clinical history for LLM context (scrub PHI — only the reason/question matters)
        string? clinicalHistory = null;
        var histSection = liveSource.GetSection("CLINICAL HISTORY") ?? liveSource.GetSection("CLINICAL INFORMATION") ?? liveSource.GetSection("INDICATION");
        if (histSection != null && !string.IsNullOrWhiteSpace(histSection.FullText))
            clinicalHistory = LlmService.ScrubPhi(histSection.FullText);

        // Extract prefix sections (clinical history, comparison) from transcript if radiologist dictated them.
        // Replace the corresponding template sections in the preserved prefix.
        var prefixExtractions = new[]
        {
            (rx: @"(?:^|\n)\s*(?:clinical\s+(?:history|information)|indications?)\s*[:\.]?\s*(.+)",
             headers: new[] { "CLINICAL HISTORY:", "CLINICAL INFORMATION:", "INDICATION:", "INDICATIONS:" },
             label: "Clinical history"),
            (rx: @"(?:^|\n)\s*-*\s*comparison\s*[:\.]?\s*(.+)",
             headers: new[] { "COMPARISON:" },
             label: "Comparison"),
            // Contrast: "contrast 100 mL Omnipaque 300", "contrast: 95 mL of Isovue 370", "IV contrast 100 mL..."
            // Allow period boundary (not just line start) — STT may produce "...technique. Contrast: 100 mL..." on one line
            // Lookbehind for period so removal doesn't eat the preceding sentence's period
            (rx: @"(?:(?:^|\n)\s*|(?<=\.)\s*)(?:iv\s+)?contrast\s*[:\.]?\s*(.+)",
             headers: new[] { "CONTRAST:" },
             label: "Contrast"),
            // Radiation dose: "radiation dose DLP 450 mGy cm effective dose 6.3 mSv", "dose DLP 320"
            (rx: @"(?:(?:^|\n)\s*|(?<=\.)\s*)(?:radiation\s+dose|dose)\s*[:\.]?\s*(.+)",
             headers: new[] { "RADIATION DOSE:" },
             label: "Radiation dose"),
        };

        foreach (var (rx, headers, label) in prefixExtractions)
        {
            var regex = new System.Text.RegularExpressions.Regex(rx, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // Collect ALL matches (e.g. multiple comparison priors)
            var matches = regex.Matches(transcript);
            if (matches.Count == 0) continue;

            var parts = new List<string>();
            foreach (System.Text.RegularExpressions.Match m in matches)
            {
                var val = m.Groups[1].Value.Trim();
                if (string.IsNullOrWhiteSpace(val)) continue;
                // For comparison: skip "none", "none available", etc. — only keep entries with actual dates
                if (label == "Comparison" && System.Text.RegularExpressions.Regex.IsMatch(val, @"^none\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    continue;
                parts.Add(val);
            }
            if (parts.Count == 0) continue;

            var rawText = string.Join("\n", parts);
            Logger.Trace($"Custom Process Report: {label} from transcript ({parts.Count} match{(parts.Count > 1 ? "es" : "")}): {rawText}");

            // Remove ALL matches from scrubbed transcript
            var scrubbedResult = scrubbedTranscript;
            var scrubbedMatches = regex.Matches(scrubbedResult);
            for (int i = scrubbedMatches.Count - 1; i >= 0; i--) // reverse to preserve indices
                scrubbedResult = scrubbedResult.Remove(scrubbedMatches[i].Index, scrubbedMatches[i].Length);
            scrubbedTranscript = scrubbedResult.Trim();

            // Replace section content in preserved prefix
            string? foundHeader = null;
            int headerIdx = -1;
            foreach (var hdr in headers)
            {
                headerIdx = preservedPrefix.IndexOf(hdr, StringComparison.OrdinalIgnoreCase);
                if (headerIdx >= 0) { foundHeader = hdr; break; }
            }
            if (headerIdx >= 0 && foundHeader != null)
            {
                int contentStart = headerIdx + foundHeader.Length;
                while (contentStart < preservedPrefix.Length && preservedPrefix[contentStart] is '\r' or '\n')
                    contentStart++;

                int nextSection = preservedPrefix.Length;
                foreach (var s in preserveSections)
                {
                    int idx = preservedPrefix.IndexOf(s + ":", contentStart, StringComparison.OrdinalIgnoreCase);
                    if (idx >= 0 && idx < nextSection) nextSection = idx;
                }

                preservedPrefix = preservedPrefix[..contentStart] + rawText + "\n\n" + preservedPrefix[nextSection..];
                preservedPrefix = preservedPrefix.TrimEnd();
            }
            else if (label == "Contrast" || label == "Radiation dose")
            {
                // Section doesn't exist in prefix — insert after TECHNIQUE
                var sectionHeader = headers[0]; // "CONTRAST:" or "RADIATION DOSE:"
                var techIdx = preservedPrefix.IndexOf("TECHNIQUE:", StringComparison.OrdinalIgnoreCase);
                if (techIdx >= 0)
                {
                    // Find the end of TECHNIQUE section (next section header)
                    int insertPos = preservedPrefix.Length;
                    foreach (var s in preserveSections)
                    {
                        int idx = preservedPrefix.IndexOf(s + ":", techIdx + 10, StringComparison.OrdinalIgnoreCase);
                        if (idx >= 0 && idx < insertPos) insertPos = idx;
                    }
                    preservedPrefix = preservedPrefix[..insertPos].TrimEnd() + "\n\n" + sectionHeader + "\n" + rawText + "\n\n" + preservedPrefix[insertPos..];
                    preservedPrefix = preservedPrefix.TrimEnd();
                }
                else
                {
                    // No TECHNIQUE section — append at end
                    preservedPrefix = preservedPrefix.TrimEnd() + "\n\n" + sectionHeader + "\n" + rawText;
                }
                Logger.Trace($"Custom Process Report: Inserted {sectionHeader} section into prefix");
            }

            // Use transcript clinical history as LLM context
            if (label == "Clinical history")
                clinicalHistory = LlmService.ScrubPhi(rawText);
        }

        Logger.Trace($"Custom Process Report: Transcript {scrubbedTranscript.Length} chars, Template {safeTemplate.Length} chars, Prefix {preservedPrefix.Length} chars");

        string? result;
        string? resultModel = null;
        Task<string?>? pendingFlash = null; // Flash task still running after lite result
        var provider = _config.LlmProvider ?? "gemini";

        // Track all model results for badge display: (name, status) where status is time like "0.4s", "failed", or "…"
        var modelResults = new List<(string Name, string Status, bool IsWinner)>();
        var modelTimes = new System.Collections.Concurrent.ConcurrentDictionary<string, double>();
        // Store each model's LLM output (findings+impression) for badge click-to-swap
        var modelOutputs = new System.Collections.Concurrent.ConcurrentDictionary<string, string>();

        if (string.IsNullOrWhiteSpace(scrubbedTranscript))
        {
            // Transcript was only prefix sections (comparison, clinical history) — no findings to process.
            // Use template as-is, just stitch with updated prefix.
            Logger.Trace("Custom Process Report: Empty transcript after extraction — using template as-is");
            result = safeTemplate;
        }
        else
        {
            // 4. Call LLM — dim the report editor while processing
            _cdpService.SetEditorDim(1, true);
            result = null;
            var mode = provider switch
            {
                "openai" => _config.LlmOpenAiProcessMode ?? "single",
                "groq" => _config.LlmGroqProcessMode ?? "single",
                "grok" => _config.LlmGrokProcessMode ?? "single",
                _ => _config.LlmProcessMode ?? "single"
            };
            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));

                if (provider == "openai" && mode == "triple")
                {
                    // GPT triple: 4.1 mini + 5 nano + 5 mini — first success wins, 5 mini upgrades in background
                    var gpt41 = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, cts.Token, "gpt-4.1-mini");
                    var gpt5nano = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, cts.Token, "gpt-5-nano");
                    var gpt5mini = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, CancellationToken.None, "gpt-5-mini");
                    gpt41.ContinueWith(t => { modelTimes["GPT 4.1 Mini"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["GPT 4.1 Mini"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);
                    gpt5nano.ContinueWith(t => { modelTimes["GPT 5 Nano"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["GPT 5 Nano"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);
                    gpt5mini.ContinueWith(t => { modelTimes["GPT 5 Mini"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["GPT 5 Mini"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);

                    (result, resultModel) = Task.Run(async () =>
                    {
                        string? best = null;
                        string? bestModel = null;
                        var all = new[] { gpt41, gpt5nano, gpt5mini };

                        // Take first successful non-null result — no grace period
                        while (all.Length > 0)
                        {
                            var done = await Task.WhenAny(all);
                            var name = done == gpt5mini ? "GPT 5 Mini" : done == gpt41 ? "GPT 4.1 Mini" : "GPT 5 Nano";
                            if (done.IsCompletedSuccessfully && done.Result != null)
                            {
                                best = done.Result;
                                bestModel = name;
                                break;
                            }
                            var reason = done.IsFaulted ? $"error: {done.Exception?.InnerException?.Message}" : "empty response";
                            Logger.Trace($"LLM: {name} failed ({reason})");
                            all = all.Where(t => t != done).ToArray();
                        }

                        return (best, bestModel);
                    }, cts.Token).GetAwaiter().GetResult();

                    // Track all model statuses for badge
                    foreach (var (task, name, isWinner) in new[] { (gpt41, "GPT 4.1 Mini", resultModel == "GPT 4.1 Mini"), (gpt5nano, "GPT 5 Nano", resultModel == "GPT 5 Nano"), (gpt5mini, "GPT 5 Mini", resultModel == "GPT 5 Mini") })
                    {
                        if (task.IsCompletedSuccessfully && task.Result != null)
                        {
                            var timeStr = modelTimes.TryGetValue(name, out var secs) ? $"{secs:F1}s" : "";
                            modelResults.Add((name, timeStr, isWinner));
                        }
                        else if (!task.IsCompleted)
                        {
                            Logger.Trace($"LLM: {name} still running (will upgrade or timeout)");
                            modelResults.Add((name, "\u2026", false));
                        }
                        else
                        {
                            Logger.Trace($"LLM: {name} failed");
                            modelResults.Add((name, "failed", false));
                        }
                    }

                    // If 5 Mini is still running, keep reference for background upgrade
                    if (!gpt5mini.IsCompleted && result != null)
                        pendingFlash = gpt5mini;
                }
                else if (provider == "gemini" && (mode == "dual" || mode == "triple"))
                {
                    // Gemini dual/triple: 2.5 Lite + 3.1 Lite, optional 3.0 Flash
                    var lite25 = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, cts.Token, "gemini-2.5-flash-lite");
                    var lite31 = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, cts.Token, "gemini-3.1-flash-lite-preview");
                    Task<string?>? flash = mode == "triple"
                        ? _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, CancellationToken.None, "gemini-3-flash-preview")
                        : null;
                    lite25.ContinueWith(t => { modelTimes["Gemini 2.5 Lite"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["Gemini 2.5 Lite"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);
                    lite31.ContinueWith(t => { modelTimes["Gemini 3.1 Lite"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["Gemini 3.1 Lite"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);
                    flash?.ContinueWith(t => { modelTimes["Gemini 3.0 Flash"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["Gemini 3.0 Flash"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);

                    (result, resultModel) = Task.Run(async () =>
                    {
                        string? best = null;
                        string? bestModel = null;

                        var first = await Task.WhenAny(lite25, lite31);

                        if (first == lite31 && lite31.IsCompletedSuccessfully && lite31.Result != null)
                        {
                            best = lite31.Result;
                            bestModel = "Gemini 3.1 Lite";
                        }
                        else if (first == lite25 && lite25.IsCompletedSuccessfully && lite25.Result != null)
                        {
                            try { await Task.WhenAny(lite31, Task.Delay(1000)); } catch { }
                            if (lite31.IsCompletedSuccessfully && lite31.Result != null)
                            {
                                best = lite31.Result;
                                bestModel = "Gemini 3.1 Lite";
                            }
                            else
                            {
                                best = lite25.Result;
                                bestModel = "Gemini 2.5 Lite";
                            }
                        }
                        else
                        {
                            var other = first == lite25 ? lite31 : lite25;
                            try { await Task.WhenAny(other, Task.Delay(5000)); } catch { }
                            if (other.IsCompletedSuccessfully && other.Result != null)
                            {
                                best = other.Result;
                                bestModel = other == lite31 ? "Gemini 3.1 Lite" : "Gemini 2.5 Lite";
                            }
                        }

                        if (flash != null)
                        {
                            if (flash.IsCompletedSuccessfully && flash.Result != null)
                            {
                                best = flash.Result;
                                bestModel = "Gemini 3.0 Flash";
                            }
                            else if (best == null)
                            {
                                try { await Task.WhenAny(flash, Task.Delay(15000)); } catch { }
                                if (flash.IsCompletedSuccessfully && flash.Result != null)
                                {
                                    best = flash.Result;
                                    bestModel = "Gemini 3.0 Flash";
                                }
                            }
                        }

                        return (best, bestModel);
                    }, cts.Token).GetAwaiter().GetResult();

                    if (flash != null && !flash.IsCompleted && result != null)
                        pendingFlash = flash;

                    // Track all model statuses for badge
                    foreach (var (task, name, isWinner) in new[] { (lite25, "Gemini 2.5 Lite", resultModel == "Gemini 2.5 Lite"), (lite31, "Gemini 3.1 Lite", resultModel == "Gemini 3.1 Lite") })
                    {
                        if (task.IsCompletedSuccessfully && task.Result != null)
                        {
                            var timeStr = modelTimes.TryGetValue(name, out var secs) ? $"{secs:F1}s" : "";
                            modelResults.Add((name, timeStr, isWinner));
                        }
                        else
                            modelResults.Add((name, "failed", false));
                    }
                    if (flash != null)
                    {
                        if (flash.IsCompleted && flash.IsCompletedSuccessfully && flash.Result != null)
                        {
                            var timeStr = modelTimes.TryGetValue("Gemini 3.0 Flash", out var secs) ? $"{secs:F1}s" : "";
                            modelResults.Add(("Gemini 3.0 Flash", timeStr, resultModel == "Gemini 3.0 Flash"));
                        }
                        else if (flash.IsCompleted)
                            modelResults.Add(("Gemini 3.0 Flash", "failed", false));
                        else
                            modelResults.Add(("Gemini 3.0 Flash", "\u2026", false)); // still running
                    }
                }
                else if (provider == "groq" && mode == "triple")
                {
                    // Groq triple: GPT-OSS-120B + Scout + 3.3 70B — first wins, 70B upgrades in background
                    var gptOss = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, cts.Token, "openai/gpt-oss-120b");
                    var scout = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, cts.Token, "meta-llama/llama-4-scout-17b-16e-instruct");
                    var llama70b = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, CancellationToken.None, "llama-3.3-70b-versatile");
                    gptOss.ContinueWith(t => { modelTimes["GPT-OSS 120B"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["GPT-OSS 120B"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);
                    scout.ContinueWith(t => { modelTimes["Llama Scout"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["Llama Scout"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);
                    llama70b.ContinueWith(t => { modelTimes["Llama 3.3 70B"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["Llama 3.3 70B"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);

                    (result, resultModel) = Task.Run(async () =>
                    {
                        string? best = null;
                        string? bestModel = null;
                        var all = new[] { gptOss, scout, llama70b };

                        while (all.Length > 0)
                        {
                            var done = await Task.WhenAny(all);
                            var name = done == gptOss ? "GPT-OSS 120B" : done == scout ? "Llama Scout" : "Llama 3.3 70B";
                            if (done.IsCompletedSuccessfully && done.Result != null)
                            {
                                best = done.Result;
                                bestModel = name;
                                break;
                            }
                            var reason = done.IsFaulted ? $"error: {done.Exception?.InnerException?.Message}" : "empty response";
                            Logger.Trace($"LLM: {name} failed ({reason})");
                            all = all.Where(t => t != done).ToArray();
                        }

                        return (best, bestModel);
                    }, cts.Token).GetAwaiter().GetResult();

                    // Track all model statuses for badge
                    foreach (var (task, name, isWinner) in new[] { (gptOss, "GPT-OSS 120B", resultModel == "GPT-OSS 120B"), (scout, "Llama Scout", resultModel == "Llama Scout"), (llama70b, "Llama 3.3 70B", resultModel == "Llama 3.3 70B") })
                    {
                        if (task.IsCompletedSuccessfully && task.Result != null)
                        {
                            var timeStr = modelTimes.TryGetValue(name, out var secs) ? $"{secs:F1}s" : "";
                            modelResults.Add((name, timeStr, isWinner));
                        }
                        else if (!task.IsCompleted)
                        {
                            Logger.Trace($"LLM: {name} still running (will upgrade or timeout)");
                            modelResults.Add((name, "\u2026", false));
                        }
                        else
                        {
                            Logger.Trace($"LLM: {name} failed");
                            modelResults.Add((name, "failed", false));
                        }
                    }

                    // If 70B is still running, keep reference for background upgrade
                    if (!llama70b.IsCompleted && result != null)
                        pendingFlash = llama70b;
                }
                else if (provider == "grok" && mode == "triple")
                {
                    // Grok triple: 3-mini + 4.1 Fast — first wins, second upgrades
                    var grokMini = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, cts.Token, "grok-3-mini");
                    var grok41 = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, CancellationToken.None, "grok-4-1-fast-non-reasoning");
                    grokMini.ContinueWith(t => { modelTimes["Grok 3 Mini"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["Grok 3 Mini"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);
                    grok41.ContinueWith(t => { modelTimes["Grok 4.1 Fast"] = sw.Elapsed.TotalSeconds; if (t.IsCompletedSuccessfully && t.Result != null) modelOutputs["Grok 4.1 Fast"] = t.Result; }, TaskContinuationOptions.ExecuteSynchronously);

                    (result, resultModel) = Task.Run(async () =>
                    {
                        string? best = null;
                        string? bestModel = null;
                        var all = new[] { grokMini, grok41 };

                        while (all.Length > 0)
                        {
                            var done = await Task.WhenAny(all);
                            var name = done == grokMini ? "Grok 3 Mini" : "Grok 4.1 Fast";
                            if (done.IsCompletedSuccessfully && done.Result != null)
                            {
                                best = done.Result;
                                bestModel = name;
                                break;
                            }
                            var reason = done.IsFaulted ? $"error: {done.Exception?.InnerException?.Message}" : "empty response";
                            Logger.Trace($"LLM: {name} failed ({reason})");
                            all = all.Where(t => t != done).ToArray();
                        }

                        return (best, bestModel);
                    }, cts.Token).GetAwaiter().GetResult();

                    // Track all model statuses for badge
                    foreach (var (task, name, isWinner) in new[] { (grokMini, "Grok 3 Mini", resultModel == "Grok 3 Mini"), (grok41, "Grok 4.1 Fast", resultModel == "Grok 4.1 Fast") })
                    {
                        if (task.IsCompletedSuccessfully && task.Result != null)
                        {
                            var timeStr = modelTimes.TryGetValue(name, out var secs) ? $"{secs:F1}s" : "";
                            modelResults.Add((name, timeStr, isWinner));
                        }
                        else if (!task.IsCompleted)
                        {
                            Logger.Trace($"LLM: {name} still running (will upgrade or timeout)");
                            modelResults.Add((name, "\u2026", false));
                        }
                        else
                        {
                            Logger.Trace($"LLM: {name} failed");
                            modelResults.Add((name, "failed", false));
                        }
                    }

                    // If 4.1 Fast is still running, keep reference for background upgrade
                    if (!grok41.IsCompleted && result != null)
                        pendingFlash = grok41;
                }
                else if (provider == "quad")
                {
                    // Quad Compare: fire one model per slot (up to 4), skip slots without API keys
                    var quadSlots = new (string ModelId, string? ApiKey)[]
                    {
                        (_config.LlmQuadModel1, QuadApiKeyFor(_config.LlmQuadModel1)),
                        (_config.LlmQuadModel2, QuadApiKeyFor(_config.LlmQuadModel2)),
                        (_config.LlmQuadModel3, QuadApiKeyFor(_config.LlmQuadModel3)),
                        (_config.LlmQuadModel4, QuadApiKeyFor(_config.LlmQuadModel4)),
                    };
                    var quadModels = quadSlots
                        .Where(s => !string.IsNullOrEmpty(s.ModelId) && !string.IsNullOrEmpty(s.ApiKey))
                        .Select(s => (s.ModelId, DisplayName: QuadDisplayName(s.ModelId)))
                        .ToList();

                    if (quadModels.Count == 0)
                    {
                        InvokeUI(() => _mainForm.ShowStatusToast("Quad Compare: No API keys configured"));
                        return;
                    }

                    Logger.Trace($"LLM Quad: Firing {quadModels.Count} models: {string.Join(", ", quadModels.Select(m => m.DisplayName))}");

                    // Fire all in parallel
                    var quadTasks = quadModels.Select(m =>
                    {
                        var task = _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, cts.Token, m.ModelId);
                        task.ContinueWith(t =>
                        {
                            modelTimes[m.DisplayName] = sw.Elapsed.TotalSeconds;
                            if (t.IsCompletedSuccessfully && t.Result != null)
                                modelOutputs[m.DisplayName] = t.Result;
                        }, TaskContinuationOptions.ExecuteSynchronously);
                        return (Task: task, m.DisplayName);
                    }).ToList();

                    // Race: first successful result wins
                    (result, resultModel) = Task.Run(async () =>
                    {
                        string? best = null;
                        string? bestModel = null;
                        var remaining = quadTasks.Select(q => q.Task).ToList();

                        while (remaining.Count > 0)
                        {
                            var done = await Task.WhenAny(remaining);
                            var name = quadTasks.First(q => q.Task == done).DisplayName;
                            if (done.IsCompletedSuccessfully && done.Result != null)
                            {
                                best = done.Result;
                                bestModel = name;
                                break;
                            }
                            var reason = done.IsFaulted ? $"error: {done.Exception?.InnerException?.Message}" : "empty response";
                            Logger.Trace($"LLM Quad: {name} failed ({reason})");
                            remaining.Remove(done);
                        }

                        return (best, bestModel);
                    }, cts.Token).GetAwaiter().GetResult();

                    // Wait a brief moment for any stragglers to finish (for badge display)
                    try { Task.WhenAll(quadTasks.Select(q => q.Task)).Wait(2000); } catch { }

                    // Track all model statuses for badge
                    foreach (var (task, name) in quadTasks)
                    {
                        bool isWinner = name == resultModel;
                        if (task.IsCompletedSuccessfully && task.Result != null)
                        {
                            var timeStr = modelTimes.TryGetValue(name, out var secs) ? $"{secs:F1}s" : "";
                            modelResults.Add((name, timeStr, isWinner));
                        }
                        else if (!task.IsCompleted)
                        {
                            modelResults.Add((name, "\u2026", false));
                        }
                        else
                        {
                            Logger.Trace($"LLM Quad: {name} failed");
                            modelResults.Add((name, "failed", false));
                        }
                    }
                }
                else
                {
                    // Single mode — use configured model
                    var singleModel = provider switch
                    {
                        "openai" => _config.LlmOpenAiModel,
                        "groq" => _config.LlmGroqModel,
                        "grok" => _config.LlmGrokModel,
                        _ => (string?)null
                    };
                    result = Task.Run(async () =>
                        await _llmService.ProcessReportAsync(scrubbedTranscript, safeTemplate, clinicalHistory, cts.Token, singleModel),
                        cts.Token).GetAwaiter().GetResult();
                    if (result != null)
                    {
                        resultModel = provider switch
                        {
                            "openai" => "GPT " + (_config.LlmOpenAiModel ?? "").Replace("gpt-", "").Replace("-mini", " Mini").Replace("-nano", " Nano"),
                            "groq" => GroqDisplayName(_config.LlmGroqModel),
                            "grok" => GrokDisplayName(_config.LlmGrokModel),
                            _ => "Gemini " + (_config.LlmModel ?? "").Replace("gemini-", "").Replace("-preview", "").Replace("-", " ")
                        };
                        // Track single model for badge
                        var singleShort = resultModel ?? "LLM";
                        modelResults.Add((singleShort, $"{sw.Elapsed.TotalSeconds:F1}s", true));
                        modelOutputs[singleShort] = result;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                Logger.Trace($"Custom Process Report: Timed out (10s, provider={provider}, mode={mode})");
            }
            catch (Exception ex)
            {
                Logger.Trace($"Custom Process Report: Error — {ex.Message}");
            }
            finally
            {
                _cdpService.SetEditorDim(1, false);
            }

            if (string.IsNullOrWhiteSpace(result))
            {
                var elapsed = sw.Elapsed.TotalSeconds;
                InvokeUI(() => _mainForm.ShowStatusToast($"Custom Process: LLM returned no result ({elapsed:F1}s)"));
                Logger.Trace("Custom Process Report: LLM returned null/empty");
                return;
            }

            if (resultModel != null)
                Logger.Trace($"Custom Process Report: Using {resultModel} result");
        }

        // 5. Stitch preserved prefix sections + LLM result, write into editor 1
        var fullReport = string.IsNullOrEmpty(preservedPrefix)
            ? result
            : preservedPrefix + "\n\n" + result;
        if (!_cdpService.ReplaceEditorContent(1, fullReport, docStructure))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Custom Process: Failed to write to editor"));
            Logger.Trace("Custom Process Report: ReplaceEditorContent failed");
            return;
        }

        sw.Stop();
        var elapsedSec = sw.Elapsed.TotalSeconds;
        var modelSuffix = resultModel != null ? $", {resultModel}" : "";
        Logger.Trace($"Custom Process Report: Success in {elapsedSec:F1}s{modelSuffix}");

        InvokeUI(() => _mainForm.ShowStatusToast($"Custom Report processed ({elapsedSec:F1}s)"));

        // Validate each model output against sent transcript for dropped findings
        var droppedMap = new Dictionary<string, (bool HasDropped, List<string> Clauses)>();
        foreach (var kv in modelOutputs)
        {
            var validation = LlmService.ValidateTranscriptCoverage(scrubbedTranscript, kv.Value);
            droppedMap[kv.Key] = (validation.HasDroppedFindings, validation.DroppedClauses);
            if (validation.HasDroppedFindings)
                Logger.Trace($"Dropped findings [{kv.Key}]: {string.Join(" | ", validation.DroppedClauses)}");
        }

        // Helper: build badge list from modelResults and modelOutputs
        var badgePrefix = provider == "gemini" ? "Gemini " : provider == "groq" ? "Llama " : "";
        List<(string Id, string Display, string? Output, bool Failed, bool Active, bool DroppedFindings, List<string>? DroppedClauses)> BuildBadges()
        {
            var badges = new List<(string Id, string Display, string? Output, bool Failed, bool Active, bool DroppedFindings, List<string>? DroppedClauses)>();
            foreach (var mr in modelResults)
            {
                if (mr.Status == "\u2026") continue; // skip pending
                var display = $"{badgePrefix}{mr.Name} \u00b7 {(mr.Status == "failed" ? "failed" : mr.Status)}";
                string? output = null;
                if (mr.Status != "failed" && modelOutputs.TryGetValue(mr.Name, out var rawOutput))
                    output = string.IsNullOrEmpty(preservedPrefix) ? rawOutput : preservedPrefix + "\n\n" + rawOutput;
                var hasDropped = droppedMap.TryGetValue(mr.Name, out var dv) && dv.HasDropped;
                var droppedClauses = hasDropped ? dv.Clauses : null;
                badges.Add((mr.Name, display, output, mr.Status == "failed", mr.IsWinner, hasDropped, droppedClauses));
            }
            return badges;
        }

        // Show model badges in "Final Report" title bar
        if (resultModel != null)
            _cdpService.SetEditorModelBadges(1, BuildBadges(), docStructure);

        // Triple mode: Flash upgrade — wait up to 8s for background flash task
        if (pendingFlash != null)
        {
            Logger.Trace("Custom Process Report: Waiting for Flash upgrade...");
            try { Task.WhenAny(pendingFlash, Task.Delay(8000)).GetAwaiter().GetResult(); } catch { }
            if (pendingFlash.IsCompletedSuccessfully && pendingFlash.Result != null)
            {
                var flashFull = string.IsNullOrEmpty(preservedPrefix)
                    ? pendingFlash.Result
                    : preservedPrefix + "\n\n" + pendingFlash.Result;
                if (_cdpService.ReplaceEditorContent(1, flashFull, docStructure))
                {
                    var totalSec = sw.Elapsed.TotalSeconds;
                    var upgradeName = provider switch { "openai" => "GPT 5 Mini", "groq" => "Llama 3.3 70B", "grok" => "Grok 4.1 Fast", _ => "Gemini 3.0 Flash" };
                    var upgradeShort = provider switch { "openai" => "GPT 5 Mini", "groq" => "Llama 3.3 70B", "grok" => "Grok 4.1 Fast", _ => "Gemini 3.0 Flash" };
                    // Validate flash upgrade output
                    var flashValidation = LlmService.ValidateTranscriptCoverage(scrubbedTranscript, pendingFlash.Result!);
                    droppedMap[upgradeShort] = (flashValidation.HasDroppedFindings, flashValidation.DroppedClauses);
                    if (flashValidation.HasDroppedFindings)
                        Logger.Trace($"Dropped findings [{upgradeShort}]: {string.Join(" | ", flashValidation.DroppedClauses)}");
                    // Mark upgrade model as winner, demote previous winner; resolve pending times
                    for (int i = 0; i < modelResults.Count; i++)
                    {
                        var mr = modelResults[i];
                        if (mr.Name == upgradeShort)
                            modelResults[i] = (mr.Name, $"{totalSec:F1}s", true);
                        else if (mr.IsWinner)
                            modelResults[i] = (mr.Name, mr.Status, false);
                        else if (mr.Status == "\u2026" && modelTimes.TryGetValue(mr.Name, out var pendingSecs))
                            modelResults[i] = (mr.Name, $"{pendingSecs:F1}s", false);
                    }
                    if (!modelResults.Any(m => m.Name == upgradeShort))
                        modelResults.Add((upgradeShort, $"{totalSec:F1}s", true));
                    Logger.Trace($"Custom Process Report: Upgraded to {upgradeName} ({totalSec:F1}s)");
                    _cdpService.SetEditorModelBadges(1, BuildBadges(), docStructure);
                    _cdpService.FlashEditor(1);
                }
            }
            else
            {
                Logger.Trace("Custom Process Report: Flash upgrade timed out or failed");
                var timedOutShort = provider switch { "openai" => "GPT 5 Mini", "groq" => "Llama 3.3 70B", "grok" => "Grok 4.1 Fast", _ => "Gemini 3.0 Flash" };
                for (int i = 0; i < modelResults.Count; i++)
                {
                    var mr = modelResults[i];
                    if (mr.Name == timedOutShort)
                        modelResults[i] = (timedOutShort, "failed", false);
                    else if (mr.Status == "\u2026" && modelTimes.TryGetValue(mr.Name, out var pendingSecs))
                        modelResults[i] = (mr.Name, $"{pendingSecs:F1}s", false);
                }
                if (!modelResults.Any(m => m.Name == timedOutShort))
                    modelResults.Add((timedOutShort, "failed", false));
                if (resultModel != null)
                    _cdpService.SetEditorModelBadges(1, BuildBadges(), docStructure);
            }
        }

        // Background badge update: watch for pending models to complete and refresh badges
        if (resultModel != null && modelResults.Any(m => m.Status == "\u2026"))
        {
            var capturedResults = modelResults;
            var capturedTimes = modelTimes;
            var capturedOutputs = modelOutputs;
            var capturedSw = sw;
            var capturedCdp = _cdpService;
            var capturedPrefix = preservedPrefix;
            var capturedBadgePrefix = badgePrefix;
            var capturedDocStructure = docStructure;
            var capturedDroppedMap = droppedMap;
            var capturedTranscript = scrubbedTranscript;
            var timeoutSec = 10.0;
            Task.Run(async () =>
            {
                try
                {
                    while (capturedResults.Any(m => m.Status == "\u2026") && capturedSw.Elapsed.TotalSeconds < timeoutSec)
                    {
                        await Task.Delay(200);
                        bool changed = false;
                        for (int i = 0; i < capturedResults.Count; i++)
                        {
                            if (capturedResults[i].Status == "\u2026" && capturedTimes.TryGetValue(capturedResults[i].Name, out var secs))
                            {
                                capturedResults[i] = (capturedResults[i].Name, $"{secs:F1}s", false);
                                changed = true;
                                // Validate newly completed model output
                                if (capturedOutputs.TryGetValue(capturedResults[i].Name, out var newOut))
                                {
                                    var v = LlmService.ValidateTranscriptCoverage(capturedTranscript, newOut);
                                    capturedDroppedMap[capturedResults[i].Name] = (v.HasDroppedFindings, v.DroppedClauses);
                                    if (v.HasDroppedFindings)
                                        Logger.Trace($"Dropped findings [{capturedResults[i].Name}]: {string.Join(" | ", v.DroppedClauses)}");
                                }
                            }
                        }
                        if (changed)
                        {
                            var badges = new List<(string Id, string Display, string? Output, bool Failed, bool Active, bool DroppedFindings, List<string>? DroppedClauses)>();
                            foreach (var mr in capturedResults)
                            {
                                if (mr.Status == "\u2026") continue;
                                var display = $"{capturedBadgePrefix}{mr.Name} \u00b7 {(mr.Status == "failed" ? "failed" : mr.Status)}";
                                string? output = null;
                                if (mr.Status != "failed" && capturedOutputs.TryGetValue(mr.Name, out var rawOut))
                                    output = string.IsNullOrEmpty(capturedPrefix) ? rawOut : capturedPrefix + "\n\n" + rawOut;
                                var hasDropped = capturedDroppedMap.TryGetValue(mr.Name, out var dv) && dv.HasDropped;
                                var droppedClauses = hasDropped ? dv.Clauses : null;
                                badges.Add((mr.Name, display, output, mr.Status == "failed", mr.IsWinner, hasDropped, droppedClauses));
                            }
                            capturedCdp.SetEditorModelBadges(1, badges, capturedDocStructure);
                        }
                    }
                }
                catch { /* best-effort */ }
            });
        }

        // Mark process report pressed so change highlighting works
        _processReportPressedForCurrentAccession = true;

        // Auto-show report overlay if enabled (same logic as PerformProcessReport)
        if (_config.ShowReportAfterProcess && !_autoShowReportDoneForAccession)
        {
            _autoShowReportDoneForAccession = true;
            var popupRef = _currentReportPopup;
            bool popupAlreadyOpen = popupRef != null && !popupRef.IsDisposed && popupRef.Visible;
            if (popupAlreadyOpen)
            {
                Logger.Trace("Custom Process Report: Skipping auto-show (popup already open), marking as stale");
                _staleSetTime = DateTime.UtcNow;
                _staleTextStableTime = default;
                InvokeUI(() => { if (popupRef != null && !popupRef.IsDisposed) popupRef.SetStaleState(true); });
            }
            else
            {
                Logger.Trace("Custom Process Report: Auto-showing report overlay");
                PerformShowReport();
            }
        }
        else
        {
            var popupRef2 = _currentReportPopup;
            if (popupRef2 != null && !popupRef2.IsDisposed && popupRef2.Visible)
            {
                Logger.Trace("Custom Process Report: Marking popup as stale during processing");
                _staleSetTime = DateTime.UtcNow;
                _staleTextStableTime = default;
                InvokeUI(() => { if (popupRef2 != null && !popupRef2.IsDisposed) popupRef2.SetStaleState(true); });
            }
        }

        // Request burst for UI updates
        RequestReportScrapeBurst(5000, "Custom Process Report");

        // Auto-restart STT if configured
        if (_config.CustomSttEnabled && _config.SttAutoStartOnCase && _sttService != null)
        {
            Thread.Sleep(300);
            PerformToggleRecordStt(true);
        }
    }

    private void StartImpressionSearch()
    {
        Logger.Trace("Starting impression search (fast scrape mode)...");
        _searchingForImpression = true;
        _impressionFromProcessReport = true; // Mark as manually triggered - stays until Sign Report
        _impressionSearchStartTime = DateTime.Now; // Track when we started - wait 2s before showing

        // Show the impression window with waiting message
        InvokeUI(() => _mainForm.ShowImpressionWindow());

        // Force slow path on next tick for responsive impression detection
        _forceSlowPathOnNextTick = true;
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
        // No-op: timer always runs at 1s. Fast/slow path gating handles all throttling.
        // Kept as a method to avoid changing dozens of call sites. The slow path throttle
        // and burst mechanism (RequestReportScrapeBurst) replace dynamic timer intervals.
    }

    private void RequestReportScrapeBurst(int durationMs, string reason)
    {
        if (durationMs <= 0) return;

        _slowPathDormant = false; // Wake slow path — burst means something important happened

        long nowTick64 = Environment.TickCount64;
        long requestedUntil = nowTick64 + durationMs;
        while (true)
        {
            long current = Interlocked.Read(ref _reportBurstUntilTick64);
            if (current >= requestedUntil) return;
            if (Interlocked.CompareExchange(ref _reportBurstUntilTick64, requestedUntil, current) == current)
            {
                Logger.Trace($"Report scrape burst requested: {reason} ({durationMs}ms)");
                return;
            }
        }
    }

    private bool IsReportPopupVisible()
    {
        var popup = _currentReportPopup;
        return popup != null && !popup.IsDisposed && popup.Visible;
    }

    private bool IsReportBurstModeActive(long nowTick64)
    {
        if (_searchingForImpression) return true;
        if (_needsBaselineCapture) return true;
        if (_radAiAutoInsertPending) return true;
        // Note: _dictationActive and IsReportPopupVisible deliberately excluded —
        // they were keeping burst mode permanently active during normal reading,
        // defeating the adaptive throttle and hammering UIA every 1s.
        // Dictation/popup updates use timed bursts (RequestReportScrapeBurst) instead.
        return nowTick64 < Interlocked.Read(ref _reportBurstUntilTick64);
    }

    // ShouldThrottleFullReportScrape removed — replaced by inline slow path throttle
    // in ScrapeTimerCallback with _lastSlowPathTick64 and idle-aware intervals.

    private void PerformSignReport(string source = "Manual")
    {
        Logger.Trace($"Sign Report (Source: {source})");
        _discardDialogCheckUntilTick64 = Environment.TickCount64 + 10_000;

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

        // Auto-stop custom STT on sign (full stop, no restart)
        if (_config.SttAutoStartOnCase && _config.CustomSttEnabled
            && _sttService != null && _sttService.IsRecording)
        {
            PerformToggleRecordStt(false);
        }

        // [CDP] Sign via button click — bypasses all Alt+F / hardware button logic
        if (_cdpService?.ClickSignReport() == true)
        {
            Logger.Trace("Sign Report: Sent via CDP button click");
        }
        // [CustomSTT] When Custom STT is enabled, always send Alt+F (Mosaic doesn't have the PowerMic)
        else if (_config.CustomSttEnabled)
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

        // Hide clinical history immediately — stale data from the signed study
        // shouldn't linger while waiting for the next study's slow path scrape.
        if (_config.ShowClinicalHistory)
        {
            InvokeUI(() => _mainForm.ToggleClinicalHistory(false));
            _clinicalHistoryVisible = false;
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

        // [Ensemble] Clear dictation data — discarded study has no valid report to validate against
        if (_sttService?.Merger is { } discardMerger)
        {
            discardMerger.ResetStudyCounters();
            lock (_pendingCorrections)
                _pendingCorrections.Clear();
        }
        _lastNonEmptyReport = null;

        // [CDP] Discard via button click, fall back to FlaUI
        bool success = _cdpService?.ClickDiscardStudy() == true;
        if (success)
            Logger.Trace("Discard Study: Sent via CDP button click");
        else
            success = _mosaicCommander.ClickDiscardStudy();

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
        Logger.Trace("Create Impression");

        // [CDP] Try button click first, fall back to FlaUI
        bool success = _cdpService?.ClickCreateImpression() == true;
        if (success)
            Logger.Trace("Create Impression: Sent via CDP button click");
        else
            success = _mosaicCommander.ClickCreateImpression();

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

        // Find all matching enabled macros (exclude voice-only macros — they trigger via speech, not study match)
        var matchingMacros = _config.Macros
            .Where(m => m.Enabled && !m.Voice && m.MatchesStudy(studyDescription))
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
                // [CDP] Direct DOM insertion — no window activation or focus needed
                if (_cdpService?.IsIframeConnected != true)
                {
                    NativeWindows.ActivateMosaicForcefully();
                    Thread.Sleep(50);
                    _mosaicCommander.FocusTranscriptBox();
                    Thread.Sleep(50);
                }

                InsertTextToFocusedEditor(PrepareTextForPaste(text), cdpEditorIndex: 0);
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
                // [CDP] Direct DOM insertion — no window activation or focus needed
                if (_cdpService?.IsIframeConnected != true)
                {
                    NativeWindows.ActivateMosaicForcefully();
                    Thread.Sleep(50);
                    _mosaicCommander.FocusTranscriptBox();
                    Thread.Sleep(50);
                }

                InsertTextToFocusedEditor(PrepareTextForPaste(text), cdpEditorIndex: 0);
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
            Thread.Sleep(100);

            // Send configured hotkey (e.g., "v" or "ctrl+shift+r")
            var hotkey = _config.IvReportHotkey;
            if (string.IsNullOrEmpty(hotkey)) hotkey = "v";

            Logger.Trace($"GetPrior: Sending hotkey '{hotkey}'");
            NativeWindows.SendHotkey(hotkey);

            // Wait for InteleViewer to start loading the prior report
            Thread.Sleep(150);

            // Copy Attempt Loop — send Ctrl+C and poll clipboard.
            // Attempt 1: generous wait (report may still be loading).
            // Attempt 2: shorter (report should be loaded by now).
            // TryGetText() is a single no-sleep clipboard check for tight polling.
            string? rawText = null;
            int[] pollCounts = { 30, 15 }; // ~1.5s, ~0.75s at 50ms intervals
            for (int attempt = 0; attempt < pollCounts.Length; attempt++)
            {
                Logger.Trace($"GetPrior: Copy Attempt {attempt + 1}...");
                ClipboardService.Clear();
                Thread.Sleep(30);

                NativeWindows.SendHotkey("ctrl+c");

                for (int i = 0; i < pollCounts[attempt]; i++)
                {
                    Thread.Sleep(50);
                    rawText = ClipboardService.TryGetText();
                    if (!string.IsNullOrEmpty(rawText) && rawText.Length >= 5)
                        break;
                }

                if (!string.IsNullOrEmpty(rawText) && rawText.Length >= 5)
                    break;
            }
            
            if (string.IsNullOrEmpty(rawText) || rawText.Length < 5)
            {
                Logger.Trace("GetPrior: Failed to retrieve text after 3 attempts.");
                InvokeUI(() => _mainForm.ShowStatusToast("No text retrieved"));
                return;
            }
            
            Logger.Trace($"Raw prior text: {rawText.Substring(0, Math.Min(100, rawText.Length))}...");
            _lastPriorRawText = rawText;

            // Process
            var formatted = _getPriorService.ProcessPriorText(rawText);
            _lastPriorFormattedText = formatted;
            if (string.IsNullOrEmpty(formatted))
            {
                InvokeUI(() => _mainForm.ShowStatusToast("Could not parse prior"));
                return;
            }
            
            // Paste into Mosaic transcript
            // [CDP] Direct DOM insertion — no window activation or focus needed
            if (_cdpService?.IsIframeConnected != true)
            {
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(100);
                _mosaicCommander.FocusTranscriptBox();
                Thread.Sleep(100);
            }

            InsertTextToFocusedEditor(PrepareTextForPaste(formatted + "\n"), cdpEditorIndex: 0);

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
                var rawNote = ScrapeClarioExamNote(msg =>
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
            var rawNote = ScrapeClarioExamNote(msg =>
            {
                InvokeUI(() => _mainForm.ShowStatusToast(msg));
            });

            _lastCriticalRawNote = rawNote;
            if (string.IsNullOrEmpty(rawNote))
            {
                _lastCriticalFormattedText = null;
                InvokeUI(() => _mainForm.ShowStatusToast("No EXAM NOTE found"));
                return;
            }

            // Update clinical history window if visible
            InvokeUI(() => _mainForm.UpdateClinicalHistory(rawNote));

            // Format
            var formatted = _noteFormatter.FormatNote(rawNote);
            _lastCriticalFormattedText = formatted;

            // Insert into Mosaic final report box (not transcript)
            // [CDP] Direct DOM insertion — no window activation or focus needed
            if (_cdpService?.IsIframeConnected != true)
            {
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(200);
                _mosaicCommander.FocusFinalReportBox();
                Thread.Sleep(100);
            }

            InsertTextToFocusedEditor(formatted, cdpEditorIndex: 1);

            Logger.Trace($"Critical Findings complete: {formatted}");

            // Remove from critical studies tracker (user has dealt with this study)
            UntrackCriticalStudy();

            InvokeUI(() => _mainForm.ShowStatusToast(
                "Critical findings inserted.\nHold Win key and trigger again to debug.", 20000));
        }
    }

    /// <summary>
    /// Scrape Clario exam note — tries CDP first if enabled, falls back to FlaUI.
    /// </summary>
    private string? ScrapeClarioExamNote(Action<string>? statusCallback = null)
    {
        // [Clario CDP] Try CDP first — instant JS eval
        if (_config.ClarioCdpEnabled && _cdpService?.IsClarioConnected == true)
        {
            var cdpNote = _cdpService.ClarioScrapeExamNote();
            if (cdpNote != null)
            {
                Logger.Trace($"Clario exam note via CDP (len={cdpNote.Length})");
                return cdpNote;
            }
            Logger.Trace("Clario CDP exam note returned null, falling back to FlaUI");
        }

        // FlaUI fallback
        return _automationService.PerformClarioScrape(statusCallback);
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
        RequestReportScrapeBurst(ShowReportBurstMs, "Show Report");

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
                        ApplyImpressionFixersToPopup(_currentReportPopup);
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
                var sr = (_cdpService?.IsIframeConnected == true) ? _cdpService.GetStructuredReport() : null;
                _currentReportPopup = new ReportPopupForm(_config, reportText, baselineForDiff,
                    changesEnabled: _config.ShowReportChanges,
                    correlationEnabled: _config.CorrelationEnabled,
                    baselineIsSectionOnly: baselineForDiff != null && _baselineIsFromTemplateDb,
                    accession: _mosaicReader.LastAccession,
                    structuredReport: sr);
                _lastPopupReportText = reportText;

                _currentReportPopup.ImpressionDeleteRequested += OnImpressionDeleteRequested;
                // Handle closure to clear references
                _currentReportPopup.FormClosed += (s, e) =>
                {
                    _currentReportPopup = null;
                    _lastPopupReportText = null;
                };

                _currentReportPopup.Show();
                ApplyImpressionFixersToPopup(_currentReportPopup);

                // If Process Report was just pressed, show "Updating..." indicator immediately
                // (report is being processed, so current content may be stale)
                if (_processReportPressedForCurrentAccession)
                {
                    _staleSetTime = DateTime.UtcNow;
                    _staleTextStableTime = default;
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

    private void ApplyImpressionFixersToPopup(ReportPopupForm popup)
    {
        if (!_config.ImpressionFixerEnabled)
        {
            popup.SetImpressionFixers(new List<ImpressionFixerEntry>());
            return;
        }

        var studyDesc = _mosaicReader.LastDescription;
        var matching = _config.ImpressionFixers
            .Where(f => f.Enabled && f.MatchesStudy(studyDesc))
            .ToList();

        // Specificity dedup: if two entries share the same blurb+mode and one has
        // study criteria while the other doesn't, the specific one wins.
        var deduped = new List<ImpressionFixerEntry>();
        var seen = new Dictionary<(string, bool), ImpressionFixerEntry>(
            EqualityComparer<(string, bool)>.Default);
        foreach (var entry in matching)
        {
            var key = (entry.Blurb.ToUpperInvariant(), entry.ReplaceMode);
            bool hasSpecificCriteria = !string.IsNullOrWhiteSpace(entry.CriteriaRequired)
                || !string.IsNullOrWhiteSpace(entry.CriteriaAnyOf)
                || !string.IsNullOrWhiteSpace(entry.CriteriaExclude);
            if (seen.TryGetValue(key, out var existing))
            {
                bool existingIsSpecific = !string.IsNullOrWhiteSpace(existing.CriteriaRequired)
                    || !string.IsNullOrWhiteSpace(existing.CriteriaAnyOf)
                    || !string.IsNullOrWhiteSpace(existing.CriteriaExclude);
                if (hasSpecificCriteria && !existingIsSpecific)
                {
                    // Replace generic with specific
                    deduped[deduped.IndexOf(existing)] = entry;
                    seen[key] = entry;
                }
                // else: keep existing (first specific wins, or both generic keeps first)
            }
            else
            {
                deduped.Add(entry);
                seen[key] = entry;
            }
        }
        popup.SetImpressionFixers(deduped);
    }

    /// <summary>
    /// Inject impression fixer buttons into the Mosaic report editor via CDP.
    /// Filters entries by study description and comparison requirements,
    /// same logic as ApplyImpressionFixersToPopup.
    /// </summary>
    private void ApplyImpressionFixersToEditor()
    {
        if (_cdpService?.IsIframeConnected != true) return;
        if (!_config.ImpressionFixerEnabled || !_config.ImpressionFixerInEditor)
        {
            _cdpService.SetImpressionFixerButtons(null);
            return;
        }

        var studyDesc = _mosaicReader.LastDescription;
        var matching = _config.ImpressionFixers
            .Where(f => f.Enabled && f.MatchesStudy(studyDesc))
            .ToList();
        Logger.Trace($"ImpressionFixerEditor: {matching.Count} matching entries for '{studyDesc}'");

        // Specificity dedup (same as popup logic)
        var deduped = new List<ImpressionFixerEntry>();
        var seen = new Dictionary<(string, bool), ImpressionFixerEntry>(
            EqualityComparer<(string, bool)>.Default);
        foreach (var entry in matching)
        {
            var key = (entry.Blurb.ToUpperInvariant(), entry.ReplaceMode);
            bool hasSpecificCriteria = !string.IsNullOrWhiteSpace(entry.CriteriaRequired)
                || !string.IsNullOrWhiteSpace(entry.CriteriaAnyOf)
                || !string.IsNullOrWhiteSpace(entry.CriteriaExclude);
            if (seen.TryGetValue(key, out var existing))
            {
                bool existingIsSpecific = !string.IsNullOrWhiteSpace(existing.CriteriaRequired)
                    || !string.IsNullOrWhiteSpace(existing.CriteriaAnyOf)
                    || !string.IsNullOrWhiteSpace(existing.CriteriaExclude);
                if (hasSpecificCriteria && !existingIsSpecific)
                {
                    deduped[deduped.IndexOf(existing)] = entry;
                    seen[key] = entry;
                }
            }
            else
            {
                deduped.Add(entry);
                seen[key] = entry;
            }
        }

        _cdpService.SetImpressionFixerButtons(deduped);
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
            // [CDP] Direct DOM insertion — no window activation needed
            if (_cdpService?.IsIframeConnected != true)
            {
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(200);
            }
            InsertTextToFocusedEditor(result);
            
            InvokeUI(() => _mainForm.ShowStatusToast($"Inserted: {result}"));
        }
        catch (Exception ex)
        {
            Logger.Trace($"CaptureSeries error: {ex.Message}");
            InvokeUI(() => _mainForm.ShowStatusToast($"OCR Error: {ex.Message}"));
        }
    }

    private void TryShowRulerOverlay()
    {
        if (!_config.RulerOverlayEnabled || RulerOverlayForm.IsOpen) return;
        RulerOverlayForm.HideOtherOverlays = hide =>
            InvokeUI(() => _mainForm.SetFloatingToolbarVisible(!hide));
        RulerOverlayForm.GetToolbarBounds = () => _mainForm.GetFloatingToolbarBounds();
        UpdateCachedToolbarBounds();
        InvokeUI(() => RulerOverlayForm.Show(_ocrService));
    }

    private void UpdateCachedToolbarBounds()
    {
        InvokeUI(() => RulerOverlayForm.CachedToolbarBounds = _mainForm.GetFloatingToolbarBounds());
    }

    private void PerformAutoMeasure()
    {
        // Only works if InteleViewer is the active window
        var foreground = NativeWindows.GetForegroundWindow();
        var title = NativeWindows.GetWindowTitle(foreground);
        if (!title.Contains("InteleViewer", StringComparison.OrdinalIgnoreCase))
        {
            Logger.Trace($"AutoMeasure: InteleViewer not active (title='{title.Substring(0, Math.Min(60, title.Length))}')");
            return;
        }

        NativeWindows.GetCursorPos(out NativeWindows.POINT cursorPt);
        var cursor = new Point(cursorPt.X, cursorPt.Y);

        var viewport = _ocrService.FindYellowTarget();
        if (!viewport.HasValue)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No viewport detected"));
            return;
        }

        if (!viewport.Value.Contains(cursor))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Cursor outside viewport"));
            return;
        }

        var scale = _ocrService.CalculatePixelsPerCm(viewport.Value);
        if (scale == null)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("Cannot calibrate ruler"));
            return;
        }

        Logger.Trace($"AutoMeasure: scale={scale.Value.vertical:F1}v/{scale.Value.horizontal:F1}h px/cm at cursor=({cursor.X},{cursor.Y})");

        var result = _ocrService.MeasureObjectAtPoint(cursor, viewport.Value, scale.Value.horizontal, scale.Value.vertical);
        if (result == null)
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No object detected at cursor"));
            return;
        }

        Logger.Trace($"AutoMeasure: {result.MajorAxisCm:F1} \u00d7 {result.MinorAxisCm:F1} cm");

        InvokeUI(() =>
        {
            MeasurementOverlayForm.ShowMeasurement(result);
        });
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

        bool success = CreateClarioCriticalNoteWithFallback();
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
        // [CDP] Read directly from DOM — skip UIA scrape
        var reportText = _cdpService?.IsIframeConnected == true
            ? _cdpService.GetEditorText(1)
            : null;
        if (string.IsNullOrEmpty(reportText))
        {
            _mosaicReader.GetFinalReportFast();
            reportText = _mosaicReader.LastFinalReport;
        }
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
        // [CDP] Read directly from DOM
        var reportText = _cdpService?.IsIframeConnected == true
            ? _cdpService.GetEditorText(1)
            : null;
        if (string.IsNullOrEmpty(reportText))
        {
            reportText = _mosaicReader.LastFinalReport;
            if (string.IsNullOrEmpty(reportText))
            {
                Logger.Trace("RecoMD: No scraped report available, attempting fresh scrape");
                reportText = _mosaicReader.GetFinalReportFast();
            }
        }

        string existingImpression = "";
        if (!string.IsNullOrEmpty(reportText))
        {
            // CDP-first: structured report gives instant section extraction
            var (_, impression) = (_cdpService?.IsIframeConnected == true)
                ? Services.CorrelationService.ExtractSections(_cdpService.GetStructuredReport())
                : Services.CorrelationService.ExtractSections(reportText);
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
            // [CDP] Direct DOM: select impression via JS, then replace via insertContent
            if (_cdpService?.IsIframeConnected == true)
            {
                bool selected = _cdpService.SelectImpressionContent();
                if (!selected)
                {
                    Logger.Trace("RecoMD: CDP SelectImpressionContent failed");
                    InvokeUI(() => _mainForm.ShowStatusToast("Could not find IMPRESSION section", 3000));
                    return;
                }
                Thread.Sleep(50);
                // Format as HTML ordered list so TipTap creates proper <li> nodes
                var recoLines = combined.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                var recoHtml = "<ol>" + string.Join("", recoLines.Select(l =>
                    $"<li><p>{System.Net.WebUtility.HtmlEncode(l)}</p></li>")) + "</ol>";
                if (!_cdpService.InsertContent(1, recoHtml))
                {
                    Logger.Trace("RecoMD: CDP InsertContent failed");
                    InvokeUI(() => _mainForm.ShowStatusToast("CDP: Failed to paste recommendations", 3000));
                    return;
                }
            }
            else
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

                InsertTextToFocusedEditor(combined, cdpEditorIndex: 1);
                Thread.Sleep(100);
            }
        }

        Logger.Trace("RecoMD: Pasted recommendations into Mosaic impression");
        RequestReportScrapeBurst(5000, "RecoMD Paste");
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
            // [CDP] Direct DOM: select impression via JS, then replace via insertContent
            if (_cdpService?.IsIframeConnected == true)
            {
                bool selected = _cdpService.SelectImpressionContent();
                if (!selected)
                {
                    Logger.Trace("RadAI Insert: CDP SelectImpressionContent failed");
                    InvokeUI(() => _mainForm.ShowStatusToast("Could not find IMPRESSION section in editor", 3000));
                    return;
                }
                Thread.Sleep(50);
                // Format as HTML ordered list so TipTap creates proper <li> nodes
                var htmlItems = string.Join("", items.Select(i =>
                    $"<li><p>{System.Net.WebUtility.HtmlEncode(i)}</p></li>"));
                var htmlImpression = $"<ol>{htmlItems}</ol>";
                if (!_cdpService.InsertContent(1, htmlImpression))
                {
                    Logger.Trace("RadAI Insert: CDP InsertContent failed");
                    InvokeUI(() => _mainForm.ShowStatusToast("CDP: Failed to insert impression", 3000));
                    return;
                }
            }
            else
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
                InsertTextToFocusedEditor(newImpression, cdpEditorIndex: 1);
                Thread.Sleep(100);
            }
        }

        // Kick off burst scraping so popup catches the update quickly
        RequestReportScrapeBurst(5000, "RadAI Insert");

        // Overlay mode: verify the paste succeeded, then auto-close
        if (_currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed)
        {
            // Verify with retries — Mosaic needs time to process the paste and
            // the UIA cached element may not reflect the new content immediately.
            bool verified = false;
            string updatedReport = "";
            var checkText = items.Length > 0 ? items[0].Trim() : "";
            if (checkText.Length > 30) checkText = checkText.Substring(0, 30);

            for (int attempt = 0; attempt < 3 && !verified; attempt++)
            {
                Thread.Sleep(attempt == 0 ? 500 : 1000);
                // [CDP] Read directly from DOM for verification
                if (_cdpService?.IsIframeConnected == true)
                    updatedReport = _cdpService.GetEditorText(1) ?? "";
                else
                {
                    _mosaicReader.GetFinalReportFast();
                    updatedReport = _mosaicReader.LastFinalReport ?? "";
                }
                if (!string.IsNullOrEmpty(checkText) && !string.IsNullOrEmpty(updatedReport))
                    verified = updatedReport.Contains(checkText, StringComparison.OrdinalIgnoreCase);
                if (!verified)
                    Logger.Trace($"RadAI Insert: Verification attempt {attempt + 1}/3 failed");
            }

            if (verified)
            {
                Logger.Trace("RadAI Insert: Verified — impression found in report");
                _pendingRadAiImpressionItems = null;
                // Push updated report to popup immediately (bypass BatchUI for instant update)
                UpdateReportPopup(updatedReport, immediate: true);
                InvokeUI(() =>
                {
                    if (_currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed)
                    {
                        try { _currentRadAiOverlay.Close(); } catch { }
                        _currentRadAiOverlay = null;
                    }
                    if (_currentReportPopup is { IsDisposed: false })
                        _currentReportPopup.SetRadAiImpressionActive(true);
                    _mainForm.ShowStatusToast("RadAI impression inserted", 3000);
                });
            }
            else
            {
                Logger.Trace("RadAI Insert: Verification failed after 3 attempts — proceeding anyway");
                // Proceed with best-effort: close overlay and mark as inserted
                _pendingRadAiImpressionItems = null;
                if (!string.IsNullOrEmpty(updatedReport))
                    UpdateReportPopup(updatedReport, immediate: true);
                InvokeUI(() =>
                {
                    if (_currentRadAiOverlay != null && !_currentRadAiOverlay.IsDisposed)
                    {
                        try { _currentRadAiOverlay.Close(); } catch { }
                        _currentRadAiOverlay = null;
                    }
                    if (_currentReportPopup is { IsDisposed: false })
                        _currentReportPopup.SetRadAiImpressionActive(true);
                    _mainForm.ShowStatusToast("RadAI impression inserted", 3000);
                });
            }
        }
        else
        {
            // Classic popup mode: scrape and push to popup immediately
            Thread.Sleep(500);
            // [CDP] Read directly from DOM
            string? updatedReport;
            if (_cdpService?.IsIframeConnected == true)
                updatedReport = _cdpService.GetEditorText(1);
            else
            {
                _mosaicReader.GetFinalReportFast();
                updatedReport = _mosaicReader.LastFinalReport;
            }
            if (!string.IsNullOrEmpty(updatedReport))
                UpdateReportPopup(updatedReport, immediate: true);
            _pendingRadAiImpressionItems = null;
            Logger.Trace("RadAI Insert: Impression replaced in report");
            InvokeUI(() =>
            {
                if (_currentRadAiPopup != null && !_currentRadAiPopup.IsDisposed)
                {
                    try { _currentRadAiPopup.Close(); } catch { }
                    _currentRadAiPopup = null;
                }
                if (_currentReportPopup is { IsDisposed: false })
                    _currentReportPopup.SetRadAiImpressionActive(true);
                _mainForm.ShowStatusToast("RadAI impression inserted", 3000);
            });
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
            // [CDP] Direct DOM: select impression via JS, then replace
            if (_cdpService?.IsIframeConnected == true)
            {
                bool selected = _cdpService.SelectImpressionContent();
                if (!selected)
                {
                    Logger.Trace("PerformReplaceImpression: CDP SelectImpressionContent failed");
                    InvokeUI(() => _mainForm.ShowStatusToast("Could not find IMPRESSION section", 3000));
                    _impressionDeletePending = false;
                    _pendingImpressionReplaceText = null;
                    InvokeUI(() => _currentReportPopup?.ClearDeletePending());
                    return;
                }
                Thread.Sleep(50);

                if (string.IsNullOrEmpty(newText))
                {
                    // All points deleted — delete selection via TipTap
                    _cdpService.InsertContent(1, "");
                }
                else
                {
                    // Format as HTML ordered list for proper numbering
                    var lines = newText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    var html = "<ol>" + string.Join("", lines.Select(l =>
                        $"<li><p>{System.Net.WebUtility.HtmlEncode(l)}</p></li>")) + "</ol>";
                    _cdpService.InsertContent(1, html);
                }
                Thread.Sleep(100);
            }
            else
            {
                NativeWindows.ActivateMosaicForcefully();
                Thread.Sleep(100);

                NativeWindows.SendHotkey("ctrl+end");
                Thread.Sleep(100);

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
                    const byte VK_DELETE = 0x2E;
                    NativeWindows.keybd_event(VK_DELETE, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(10);
                    NativeWindows.keybd_event(VK_DELETE, 0, NativeWindows.KEYEVENTF_KEYUP, UIntPtr.Zero);
                }
                else
                {
                    InsertTextToFocusedEditor(newText, cdpEditorIndex: 1);
                }
                Thread.Sleep(100);
            }
        }

        // Verify: re-scrape and update report popup display
        Thread.Sleep(500);
        // [CDP] Read directly from DOM
        string updatedReport;
        if (_cdpService?.IsIframeConnected == true)
            updatedReport = _cdpService.GetEditorText(1) ?? "";
        else
        {
            _mosaicReader.GetFinalReportFast();
            updatedReport = _mosaicReader.LastFinalReport ?? "";
        }

        _pendingImpressionReplaceText = null;
        _impressionDeletePending = false;

        InvokeUI(() =>
        {
            _currentReportPopup?.ClearDeletePending();
            // Update the report popup with the re-scraped content
            if (!string.IsNullOrEmpty(updatedReport) && _currentReportPopup != null && !_currentReportPopup.IsDisposed)
            {
                var sr = (_cdpService?.IsIframeConnected == true) ? _cdpService.GetStructuredReport() : null;
                _currentReportPopup.UpdateReport(updatedReport, structuredReport: sr, forceUpdate: true);
                _lastPopupReportText = updatedReport;
            }
            // Also update impression window if visible (CDP-first)
            var updatedImpression = (_cdpService?.IsIframeConnected == true)
                ? ImpressionForm.ExtractImpression(_cdpService.GetStructuredReport())
                  ?? ImpressionForm.ExtractImpression(updatedReport)
                : ImpressionForm.ExtractImpression(updatedReport);
            if (!string.IsNullOrEmpty(updatedImpression))
                _mainForm.UpdateImpression(updatedImpression);
        });

        Logger.Trace("PerformReplaceImpression completed");
    }

    private void PerformManualCorrectTemplate()
    {
        var description = _mosaicReader.LastDescription;
        if (string.IsNullOrWhiteSpace(description))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No study description available"));
            return;
        }

        var searchText = AutomationService.BuildTemplateSearchText(description);
        Logger.Trace($"ManualCorrectTemplate: correcting to '{searchText}'");
        bool corrected = _cdpService?.SetStudyType(searchText, description) == true
            || _automationService.AttemptCorrectTemplate(description);
        InvokeUI(() => _mainForm.ShowStatusToast(
            corrected ? "Template correction attempted" : "Template correction failed"));
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
        if (_scrapeTimer != null) return;

        Logger.Trace($"Starting Mosaic scrape timer (1s fast path, {_config.ScrapeIntervalSeconds}s slow path)...");
        _scrapeTimer = new System.Threading.Timer(ScrapeTimerCallback, null, ScrapeTimerIntervalMs, ScrapeTimerIntervalMs);
    }

    private void ScrapeTimerCallback(object? state)
    {
        // Reentrancy guard — if previous tick is still running, skip this one
        if (Interlocked.CompareExchange(ref _scrapeRunning, 1, 0) != 0)
        {
            // Watchdog: if a scrape has been running for >30s, it's probably hung in a COM call.
            var startTicks = Interlocked.Read(ref _scrapeStartedTicks);
            if (startTicks > 0)
            {
                var elapsed = (DateTime.UtcNow.Ticks - startTicks) / (double)TimeSpan.TicksPerSecond;
                if (elapsed > 30)
                {
                    Logger.Trace($"Scrape watchdog: force-releasing hung scrape after {elapsed:F0}s");
                    Interlocked.Exchange(ref _scrapeRunning, 0);
                    Interlocked.Exchange(ref _scrapeStartedTicks, 0);
                }
            }
            return;
        }

        if (_isUserActive) { Interlocked.Exchange(ref _scrapeRunning, 0); return; }

        // Enter background mode for this callback only (lower CPU/IO/memory priority)
        bool enteredBackground = false;
        try { enteredBackground = SetThreadPriority(GetCurrentThread(), THREAD_MODE_BACKGROUND_BEGIN); } catch { }

        Interlocked.Exchange(ref _scrapeStartedTicks, DateTime.UtcNow.Ticks);
        long tickStartTick64 = Environment.TickCount64;

        try
        {
            long nowTick64 = Environment.TickCount64;
            string? fastReportText = null;

            // ═══════ CDP PATH (replaces both fast + slow UIA paths when connected) ═══════
            bool cdpHandled = false;
            if (_cdpService != null)
            {
                if (!_cdpService.IsConnected)
                    try { _cdpService.TryConnect(); } catch { }

                // [Clario CDP] Try to connect to Clario's Chrome instance
                if (_config.ClarioCdpEnabled && !_cdpService.IsClarioConnected)
                    try { _cdpService.TryConnectClario(); } catch { }

                if (_cdpService.IsConnected)
                {
                    // [CDP] Inject independent column scrolling CSS if enabled
                    if (_config.CdpIndependentScrolling && _cdpService.IsIframeConnected)
                    {
                        try { _cdpService.InjectScrollFix(); }
                        catch (Exception ex) { Logger.Trace($"CDP: InjectScrollFix exception: {ex.Message}"); }

                        // One-shot diagnostic: dump DOM structure after scroll fix succeeds
                        if (_cdpService.ScrollFixActive && !_scrollDiagLogged)
                        {
                            _scrollDiagLogged = true;
                            try
                            {
                                var diag = _cdpService.DumpScrollDiagnostic();
                                if (diag != null)
                                    Logger.Trace($"CDP: Scroll diagnostic (fixActive={_cdpService.ScrollFixActive}):\n{diag}");
                                var vdiag = _cdpService.DumpVerticalLayoutDiag();
                                if (vdiag != null)
                                    Logger.Trace($"CDP: Vertical layout diagnostic:\n{vdiag}");
                                var blankDiag = _cdpService.DumpBlankLines();
                                if (blankDiag != null)
                                    Logger.Trace($"CDP: Blank line structure:\n{blankDiag}");
                            }
                            catch { }
                        }

                        // Read back column ratio after user drag (every ~30s to avoid overhead)
                        if (_cdpService.ScrollFixActive && nowTick64 - _lastColumnRatioReadTick64 > 30_000)
                        {
                            _lastColumnRatioReadTick64 = nowTick64;
                            try
                            {
                                var ratio = _cdpService.ReadColumnRatio();
                                if (ratio.HasValue && Math.Abs(ratio.Value - _config.CdpColumnRatio) > 0.005)
                                {
                                    _config.CdpColumnRatio = ratio.Value;
                                    _config.Save();
                                }
                            }
                            catch { }
                        }
                    }

                    // [Highlight Mode] Inject buttons and poll mode when iframe is connected
                    if (_cdpService.IsIframeConnected)
                    {
                        try { _cdpService.InjectHighlightModeButtons(_currentHighlightMode); } catch { }

                        // Check if user clicked a highlight mode button (consume-once flag)
                        try
                        {
                            var userMode = _cdpService.ConsumeHighlightModeChange();
                            if (userMode != null && userMode != _currentHighlightMode)
                            {
                                Logger.Trace($"Highlight mode changed by user: {_currentHighlightMode} -> {userMode}");
                                _currentHighlightMode = userMode;
                                _highlightModeAppliedForStudy = false; // Force re-apply
                            }
                        }
                        catch { }

                        // Apply mode effects when needed (new study or user changed mode)
                        if (!_highlightModeAppliedForStudy)
                        {
                            try
                            {
                                _lastRainbowReportText = null;
                                switch (_currentHighlightMode)
                                {
                                    case "none":
                                        _cdpService.HideAllHighlights();
                                        _highlightModeAppliedForStudy = true;
                                        break;
                                    case "regular":
                                        _cdpService.ClearRainbowHighlights();
                                        _cdpService.SetRegularHighlightsVisible(true);
                                        _highlightModeAppliedForStudy = true;
                                        break;
                                    case "rainbow":
                                        _cdpService.SetRegularHighlightsVisible(false);
                                        var currentReport = _mosaicReader.LastFinalReport;
                                        if (!string.IsNullOrEmpty(currentReport))
                                        {
                                            try { ComputeAndApplyRainbow(currentReport); }
                                            catch { }
                                            _lastRainbowReportText = currentReport;
                                            _rainbowTextChangedTick64 = 0;
                                            _highlightModeAppliedForStudy = true;
                                        }
                                        break;
                                }
                            }
                            catch { }
                        }
                    }

                    try
                    {
                        var cdp = _cdpService.Scrape();
                        if (cdp != null)
                        {
                            _automationService.PopulateFromCdpData(cdp);
                            fastReportText = cdp.ReportText;
                            // Only mark CDP as authoritative if it returned meaningful data.
                            // Empty accession means the iframe DOM read failed — fall through to FlaUI.
                            cdpHandled = !string.IsNullOrEmpty(cdp.Accession);
                            if (cdpHandled)
                            {
                                _slowPathEverCompleted = true;
                                _fastReadFailCount = 0;
                            }
                            else
                            {
                                Logger.Trace("CDP scrape returned empty accession — falling through to FlaUI");
                            }
                            if (!string.IsNullOrEmpty(fastReportText))
                            {
                                UpdateReportPopup(fastReportText);
                                UpdateRecoMd(_lastNonEmptyAccession, fastReportText);

                                // [Highlight Mode] Apply rainbow highlights with debounce
                                // Wait for text to stabilize (2s) before recomputing — avoids
                                // clearing+rebuilding highlights while user is actively typing
                                if (_currentHighlightMode == "rainbow")
                                {
                                    if (fastReportText != _lastRainbowReportText)
                                    {
                                        _rainbowTextChangedTick64 = nowTick64;
                                        _lastRainbowReportText = fastReportText;
                                    }
                                    if (_rainbowTextChangedTick64 > 0 && nowTick64 - _rainbowTextChangedTick64 >= 2000)
                                    {
                                        _rainbowTextChangedTick64 = 0;
                                        try { ComputeAndApplyRainbow(fastReportText); }
                                        catch (Exception ex) { Logger.Trace($"Rainbow highlight error: {ex.Message}"); }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex) { Logger.Trace($"CDP scrape failed: {ex.Message}"); }
                }
            }

            // ═══════ FAST PATH (every tick, ~0ms) ═══════
            // Reads cached report editor + cached accession element. No tree walks.
            // Gated on: first slow path has completed (caches built) AND not too many consecutive failures.
            if (!cdpHandled && _slowPathEverCompleted && _fastReadFailCount < FastReadFailThreshold)
            {
                if (_mosaicReader.TryFastRead(out fastReportText, out var fastAccession))
                {
                    _fastReadFailCount = 0;

                    // If report changed → update popup immediately (~0ms on UI thread)
                    if (!string.IsNullOrEmpty(fastReportText))
                    {
                        UpdateReportPopup(fastReportText);
                        UpdateRecoMd(_lastNonEmptyAccession, fastReportText);
                    }
                    // Editor cache cold is OK — dormant slow path still runs report reads
                    // every 3s and will rebuild the cache via ProseMirror search.

                    // If accession differs from last known → wake slow path
                    if (fastAccession != null && fastAccession != _lastNonEmptyAccession
                        && !string.IsNullOrEmpty(_lastNonEmptyAccession))
                    {
                        _forceSlowPathOnNextTick = true;
                        _slowPathDormant = false;
                    }
                }
                else
                {
                    _fastReadFailCount++;
                    if (_fastReadFailCount >= FastReadFailThreshold)
                    {
                        // Caches cold — wake slow path to rebuild them
                        _slowPathDormant = false;
                        Logger.Trace("Fast path: disabled after consecutive failures (caches cold), waking slow path");
                    }
                }
            }

            // Re-assert topmost on all tool windows periodically (cheap, every tick)
            BatchUI(() => _mainForm.EnsureWindowsOnTop());

            // Patient mismatch check: compare Mosaic patient name vs topmost InteleViewer window
            // Cheap Win32 calls only (EnumWindows + GetWindowText), safe to run every tick
            if (_config.PatientMismatchEnabled && _config.ShowClinicalHistory)
                CheckPatientMismatch();

            // Heartbeat: log every ~240 ticks (~4 min at 1s) to confirm timer is alive
            _scrapeHeartbeatCount++;
            if (_scrapeHeartbeatCount >= 240)
            {
                _scrapeHeartbeatCount = 0;
                var acc = _mosaicReader.LastAccession ?? "(none)";
                var proc = Process.GetCurrentProcess();
                long managedMb = GC.GetTotalMemory(false) / (1024 * 1024);
                long workingSetMb = proc.WorkingSet64 / (1024 * 1024);
                long privateMb = proc.PrivateMemorySize64 / (1024 * 1024);
                int handles = proc.HandleCount;
                int threads = proc.Threads.Count;
                int gen0 = GC.CollectionCount(0);
                int gen1 = GC.CollectionCount(1);
                int gen2 = GC.CollectionCount(2);
                Logger.Trace($"Scrape heartbeat: acc={acc}, idle={_consecutiveIdleScrapes}, fastFail={_fastReadFailCount}, dormant={_slowPathDormant}, clinHist={_clinicalHistoryVisible}");
                long uiaCalls = _automationService.UiaCallCount;
                Logger.Trace($"  Memory: managed={managedMb}MB, workingSet={workingSetMb}MB, private={privateMb}MB, handles={handles}, threads={threads}, GC={gen0}/{gen1}/{gen2}, uiaCalls={uiaCalls}");
            }

            // Periodic UIA reset (time-gated, runs in slow path guard to avoid duplicate work)
            _automationService.CheckPeriodicUiaReset();

            // Periodic GC: time-based safety net for COM wrappers that escape deterministic release.
            _scrapesSinceLastGc++;
            if (_scrapesSinceLastGc >= 240) // ~4 min at 1s tick
            {
                _scrapesSinceLastGc = 0;
                GC.Collect(2, GCCollectionMode.Forced);
                GC.WaitForPendingFinalizers();
            }

            // ═══════ SLOW PATH THROTTLE ═══════
            bool shouldRunSlowPath = _forceSlowPathOnNextTick;
            _forceSlowPathOnNextTick = false;

            if (!shouldRunSlowPath)
            {
                long elapsed = nowTick64 - _lastSlowPathTick64;

                if (IsReportBurstModeActive(nowTick64))
                {
                    // Burst: slow path every 1s (every tick)
                    shouldRunSlowPath = elapsed >= ScrapeTimerIntervalMs;
                }
                else if (nowTick64 < _discardDialogCheckUntilTick64)
                {
                    // Brief discard dialog detection window — keep slow path active at normal rate
                    shouldRunSlowPath = elapsed >= SteadyStateReportScrapeMinIntervalMs;
                }
                else if (_slowPathDormant)
                {
                    // Dormant: still read report for popup updates, skip expensive Clario/metadata.
                    // Report read is 0ms (cache hit) or 300ms (ProseMirror search when cache cold).
                    // Run every tick (1s) since it's lightweight without Clario work.
                    shouldRunSlowPath = elapsed >= ScrapeTimerIntervalMs;
                }
                else if (_consecutiveIdleScrapes >= 10)
                {
                    shouldRunSlowPath = elapsed >= DeepIdleScrapeIntervalMs;
                }
                else if (_consecutiveIdleScrapes >= 3)
                {
                    shouldRunSlowPath = elapsed >= IdleScrapeIntervalMs;
                }
                else
                {
                    int minInterval = Math.Max(NormalScrapeIntervalMs, SteadyStateReportScrapeMinIntervalMs);
                    shouldRunSlowPath = elapsed >= minInterval;
                }
            }

            if (!shouldRunSlowPath)
            {
                // Fast path only this tick
                return;
            }

            // ═══════ SLOW PATH (full UIA scrape) ═══════
            _lastSlowPathTick64 = nowTick64;
            Interlocked.Exchange(ref _lastFullReportScrapeTick64, nowTick64);

            // Tell AutomationService whether burst mode is active so it can relax the
            // ProseMirror search throttle (3s instead of 10s) for faster report pickup.
            _automationService.IsBurstModeActive = IsReportBurstModeActive(nowTick64);

            // Scrape Mosaic for report data (skip UIA if CDP already handled it)
            var reportText = cdpHandled ? _mosaicReader.LastFinalReport : _mosaicReader.GetFinalReportFast();

            // [Ensemble] Snapshot non-empty report for correction validation at study change.
            // By the time study-change fires, the scrape may have already picked up the new (empty) study.
            if (!string.IsNullOrEmpty(reportText))
                _lastNonEmptyReport = reportText;

            // Bail out if user action started during the scrape
            if (_isUserActive) return;

            // Idle tracking: count consecutive scrapes with no report content.
            bool suppressIdleBackoff = IsReportBurstModeActive(nowTick64) || _pendingCloseAccession != null;
            if (string.IsNullOrEmpty(reportText))
            {
                if (!suppressIdleBackoff)
                {
                    _consecutiveIdleScrapes++;
                    if (_consecutiveIdleScrapes == 3)
                        Logger.Trace("Scrape idle backoff: slowing slow path to 10s (no report content)");
                    else if (_consecutiveIdleScrapes == 10)
                        Logger.Trace("Scrape deep idle: slowing slow path to 30s (extended no report content)");
                }
            }
            else if (_consecutiveIdleScrapes > 0)
            {
                _consecutiveIdleScrapes = 0;
            }

            // Slow path rebuilt the caches — re-enable fast path
            if (!string.IsNullOrEmpty(reportText))
            {
                _fastReadFailCount = 0;
                _slowPathEverCompleted = true;
            }

            // [RadAI] Auto-insert: trigger when scrape succeeds after Process Report.
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
                Description: _mosaicReader.LastDescription ?? _mosaicReader.LastTemplateName,
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
            if (!string.IsNullOrEmpty(currentAccession)
                && nowTick64 < _discardDialogCheckUntilTick64
                && _mosaicReader.IsDiscardDialogVisible())
            {
                _discardDialogShownForCurrentAccession = true;
                Logger.Trace("RVUCounter: Discard dialog detected for current accession");
            }

            // Detect accession change with flap debounce
            var (accessionChanged, studyClosed) = DetectAccessionChange(currentAccession);

            if (accessionChanged)
            {
                _slowPathDormant = false; // Wake slow path for new study metadata
                _slowPathEverCompleted = false; // Fast path waits until slow path rebuilds caches
                OnStudyChanged(currentAccession, reportText, studyClosed);
            }

            // Mark popup as stale when accession goes empty (study may be transitioning).
            // Don't close yet — accession flaps (brief empty during transitions) would kill
            // the popup unnecessarily. OnStudyChanged closes it after debounce confirms.
            if (_pendingCloseAccession != null && _pendingCloseTickCount == 1
                && _currentReportPopup != null && !_currentReportPopup.IsDisposed)
            {
                var popup = _currentReportPopup;
                BatchUI(() => { if (!popup.IsDisposed) popup.SetStaleState(true); });
            }

            // Detect template change within the same accession (e.g., wrong template → corrected).
            // Reset rainbow baseline so diff highlights the new template's actual changes, not boilerplate.
            var currentTemplateName = _mosaicReader.LastTemplateName;
            if (!string.IsNullOrEmpty(currentTemplateName) && currentTemplateName != _lastSeenTemplateName)
            {
                if (_lastSeenTemplateName != null)
                {
                    Logger.Trace($"Template changed on same accession: \"{_lastSeenTemplateName}\" → \"{currentTemplateName}\". Resetting rainbow baseline.");
                    _baselineReport = null;
                    _baselineIsFromTemplateDb = false;
                    _baselineCaptureAttempts = 0;
                    _needsBaselineCapture = _config.ShowReportChanges || _config.CorrelationEnabled;
                    _processReportPressedForCurrentAccession = false;
                    _draftedAutoProcessDetected = false;
                    _llmCapturedTemplate = null; // [LLM] Re-capture template after template switch
                    _llmCapturedTemplateAccession = null;
                    _llmCapturedDocStructure = null;
                    _llmDbTemplateOverride = null;
                    Logger.Trace("LLM: Template changed — will re-capture on next scrape");
                }
                _lastSeenTemplateName = currentTemplateName;
            }

            // [LLM] Eager template capture — grab structured report + doc structure on new accession.
            // For drafted studies, the captured report already has dictated content, so prefer
            // the clean default template from TemplateDatabase instead for the LLM.
            if (_config.LlmProcessEnabled && _cdpService != null
                && _llmCapturedTemplate == null && !string.IsNullOrEmpty(currentAccession))
            {
                try
                {
                    var earlyStructured = _cdpService.GetStructuredReport();
                    if (earlyStructured != null && earlyStructured.Sections.Count > 0)
                    {
                        _llmCapturedTemplate = earlyStructured;
                        _llmCapturedTemplateAccession = currentAccession;
                        _llmCapturedDocStructure = _cdpService.GetDocStructure(1);
                        Logger.Trace($"LLM: Eager template capture for {currentAccession} ({earlyStructured.Sections.Count} sections)");

                        // Drafted studies: the captured report has prior dictation, not a clean template.
                        // Try the TemplateDatabase for the real default template.
                        if (_mosaicReader.LastDraftedState)
                        {
                            var desc = _mosaicReader.LastDescription;
                            if (!string.IsNullOrEmpty(desc))
                            {
                                var dbTemplate = _templateDatabase.GetFallbackTemplate(desc);
                                if (dbTemplate != null)
                                {
                                    _llmDbTemplateOverride = dbTemplate;
                                    Logger.Trace($"LLM: Drafted study — using DB template for '{desc}' ({dbTemplate.Length} chars)");
                                }
                                else
                                {
                                    Logger.Trace($"LLM: Drafted study — no confident DB template for '{desc}', will use captured report");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Trace($"LLM: Eager template capture failed — {ex.Message}");
                }
            }

            // These always run (cheap, needed for popup updates):
            UpdateRecoMd(currentAccession, reportText);
            CaptureBaselineReport(reportText);
            // Skip popup update if fast path already pushed the same text this tick
            if (fastReportText == null || reportText != fastReportText)
                UpdateReportPopup(reportText);

            // Inject impression fixer buttons into editor (once per accession — JS interval handles persistence)
            if (_config.ImpressionFixerEnabled && _config.ImpressionFixerInEditor
                && !string.IsNullOrEmpty(currentAccession) && currentAccession != _impFixerEditorAccession
                && _cdpService?.IsIframeConnected == true)
            {
                _impFixerEditorAccession = currentAccession;
                ApplyImpressionFixersToEditor();
            }

            // Flush popup update to UI now, before expensive Clario/metadata work below
            FlushUI();

            if (_slowPathDormant)
            {
                // Dormant: report read + popup update done above. Still run alerts
                // (template mismatch, gender check, Aidoc) — all cheap, no FlaUI tree walks.
                UpdateClinicalHistoryAndAlerts(currentAccession, reportText, tickStartTick64);
                // Flush batched UI updates and re-assert z-order after Aidoc FlaUI interactions
                BatchUI(() => _mainForm.EnsureWindowsOnTop());
                FlushUI();
                return;
            }

            // Full slow path: expensive metadata/Clario work
            RecordTemplateIfNeeded(reportText);
            InsertPendingMacros(currentAccession, reportText);
            // Skip clinical history + impression updates on the same tick as a study change.
            // reportText was read BEFORE the change was detected, so it may contain stale data
            // from the previous study. Next tick will have fresh data from the new study.
            if (!accessionChanged)
            {
                UpdateClinicalHistoryAndAlerts(currentAccession, reportText, tickStartTick64);
                UpdateImpressionDisplay(reportText);
            }
            // Flush batched UI updates and re-assert z-order after Aidoc FlaUI interactions
            BatchUI(() => _mainForm.EnsureWindowsOnTop());
            FlushUI();

            // Check if all study metadata is populated — if so, go dormant.
            // Dormant slow path still reads the report + updates popup, just skips Clario/metadata.
            if (!string.IsNullOrEmpty(currentAccession) && !string.IsNullOrEmpty(reportText)
                && !_needsBaselineCapture && !_searchingForImpression
                && !_pendingClarioPriorityRetry
                && string.IsNullOrEmpty(_pendingMacroAccession)
                && !IsReportBurstModeActive(nowTick64))
            {
                _slowPathDormant = true;
                Logger.Trace("Slow path: entering dormancy (skipping Clario/metadata, report read continues)");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Mosaic scrape error: {ex.Message}");
        }
        finally
        {
            // Exit background mode before flushing UI
            if (enteredBackground)
                try { SetThreadPriority(GetCurrentThread(), THREAD_MODE_BACKGROUND_END); } catch { }

            // Flush all batched UI updates in a single BeginInvoke call.
            FlushUI();
            Interlocked.Exchange(ref _scrapeStartedTicks, 0);

            // Log total tick duration (UIA + processing + UI flush) for degradation tracking
            long tickMs = Environment.TickCount64 - tickStartTick64;
            if (tickMs > 500)
                Logger.Trace($"Scrape tick: {tickMs}ms total");

            Interlocked.Exchange(ref _scrapeRunning, 0);
        }
    }

    private void StopMosaicScrapeTimer()
    {
        _scrapeTimer?.Dispose();
        _scrapeTimer = null;
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
        _lastSeenTemplateName = null;
        _lastPopupReportText = null;
        // _currentHighlightMode intentionally NOT reset — persists across accession changes
        _highlightModeAppliedForStudy = false; // Re-apply mode effects on new study
        _lastRainbowReportText = null;
        _llmCapturedTemplate = null; // [LLM] Reset template for new study
        _llmCapturedTemplateAccession = null;
        _llmCapturedDocStructure = null;
        _llmDbTemplateOverride = null;

        // [KeytermLearning] Verify collected words against final report before clearing
        if (_keytermLearningByProvider.Count > 0 && !_discardDialogShownForCurrentAccession)
        {
            var reportForVerify = _mosaicReader.LastFinalReport;
            if (!string.IsNullOrEmpty(reportForVerify))
            {
                foreach (var learning in _keytermLearningByProvider.Values)
                    learning.VerifyAgainstReport(reportForVerify);
            }
        }
        foreach (var learning in _keytermLearningByProvider.Values)
            learning.ClearSessionBuffer();

        // [Ensemble] Validate corrections and commit all-time stats (only for signed reports)
        if (_sttService?.Merger is { } merger)
        {
            if (!_discardDialogShownForCurrentAccession)
            {
                // Commit study stats to all-time
                merger.CommitStudyToAlltime();

                // Validate corrections against final report
                List<CorrectionRecord> correctionsToValidate;
                lock (_pendingCorrections)
                {
                    correctionsToValidate = _pendingCorrections.Count > 0
                        ? new List<CorrectionRecord>(_pendingCorrections)
                        : new List<CorrectionRecord>();
                }
                var validationReport = _lastNonEmptyReport;
                Logger.Trace($"Ensemble validation: {correctionsToValidate.Count} pending corrections, report={validationReport?.Length ?? 0} chars");
                if (correctionsToValidate.Count > 0)
                {
                    if (!string.IsNullOrEmpty(validationReport))
                    {
                        int validated = 0, rejected = 0;
                        foreach (var c in correctionsToValidate)
                        {
                            var replacementFound = System.Text.RegularExpressions.Regex.IsMatch(
                                validationReport, @"\b" + System.Text.RegularExpressions.Regex.Escape(c.Replacement) + @"\b",
                                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                            var originalFound = System.Text.RegularExpressions.Regex.IsMatch(
                                validationReport, @"\b" + System.Text.RegularExpressions.Regex.Escape(c.Original) + @"\b",
                                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                            if (replacementFound)
                                validated++;
                            else if (originalFound)
                                rejected++;
                            // else: inconclusive (neither found — user rewrote the section)
                        }
                        if (validated > 0 || rejected > 0)
                        {
                            merger.SessionValidated += validated;
                            merger.SessionRejected += rejected;
                            _config.SttEnsembleAlltimeValidated += validated;
                            _config.SttEnsembleAlltimeRejected += rejected;
                            _config.Save();
                            Logger.Trace($"Ensemble validation: {validated} validated, {rejected} rejected (of {correctionsToValidate.Count} corrections, {correctionsToValidate.Count - validated - rejected} inconclusive)");
                        }
                    }
                }

                // Validate per-provider accuracy against signed report
                if (!string.IsNullOrEmpty(validationReport))
                {
                    var (ensAcc, dgAcc, s1Acc, s2Acc) = merger.ValidateAccuracyAgainstReport(validationReport);
                    if (ensAcc.Total > 0)
                    {
                        merger.AddSessionAccuracy(ensAcc, dgAcc, s1Acc, s2Acc);
                        _config.SttEnsembleAlltimeAccEnsembleMatched += ensAcc.Matched;
                        _config.SttEnsembleAlltimeAccEnsembleTotal += ensAcc.Total;
                        _config.SttEnsembleAlltimeAccDgMatched += dgAcc.Matched;
                        _config.SttEnsembleAlltimeAccDgTotal += dgAcc.Total;
                        _config.SttEnsembleAlltimeAccS1Matched += s1Acc.Matched;
                        _config.SttEnsembleAlltimeAccS1Total += s1Acc.Total;
                        _config.SttEnsembleAlltimeAccS2Matched += s2Acc.Matched;
                        _config.SttEnsembleAlltimeAccS2Total += s2Acc.Total;
                        _config.Save();
                        Logger.Trace($"Ensemble accuracy: ens={ensAcc.Matched}/{ensAcc.Total} dg={dgAcc.Matched}/{dgAcc.Total} s1={s1Acc.Matched}/{s1Acc.Total} s2={s2Acc.Matched}/{s2Acc.Total}");
                    }
                }

                merger.FireStatsUpdated();
            }
            // Always reset study counters and clear pending corrections for next study
            merger.ResetStudyCounters();
            lock (_pendingCorrections)
                _pendingCorrections.Clear();
        }

        _lastNonEmptyReport = null;
        _mosaicReader.ClearLastReport();

        // UIA reset moved to periodic timer (CheckPeriodicUiaReset, every 10 min).
        // Per-study reset was adding 3-6s cold ProseMirror search delay on every study
        // transition, blocking clinical history and report popup from updating.

        // Close stale report popup from prior study (immediate — don't defer via BatchUI)
        if (_currentReportPopup != null && !_currentReportPopup.IsDisposed)
        {
            var stalePopup = _currentReportPopup;
            _currentReportPopup = null;
            InvokeUI(() => { try { stalePopup.Close(); } catch { } });
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
        _templateCorrectionAttempts = 0;
        _templateCorrectionNextRetryTick64 = 0;
        _genderMismatchActive = false;
        _fimMismatchActive = false;
        _consistencyMismatchActive = false;
        ClearCdpAlertTextFlashing();
        _lastMismatchCheckReportText = null;
        _lastFimMismatches = null;
        _lastConsistencyCheckReportText = null;
        _lastConsistencyCheckDescription = null;
        _lastConsistencyResults = null;
        _strokeDetectedActive = false;
        _pendingClarioPriorityRetry = false;
        _lastAidocFindings = null;
        _lastAidocRelevantList = null;
        _lastAidocRelevant = false;
        _aidocDoneForStudy = false;

        // Reset Ignore Inpatient Drafted state
        _macrosCompleteForCurrentAccession = false;
        _autoFixCompleteForCurrentAccession = false;
        _ctrlASentForCurrentAccession = false;

        // Update tracking - only update to new non-empty accession
        if (!studyClosed)
        {
            _lastNonEmptyAccession = currentAccession;
            _consecutiveIdleScrapes = 0; // New study opened — prevent idle backoff

            _needsBaselineCapture = _config.ShowReportChanges || _config.CorrelationEnabled;
            // Force slow path on next tick to catch template flash before auto-processing
            if (_needsBaselineCapture && !_searchingForImpression)
                _forceSlowPathOnNextTick = true;

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

            // Inject impression fixer buttons into editor for new study
            ApplyImpressionFixersToEditor();

            // Reset clinical history state on study change
            BatchUI(() => _mainForm.OnClinicalHistoryStudyChanged());
            // Hide impression window on new study
            BatchUI(() => _mainForm.HideImpressionWindow());
            _searchingForImpression = false;
            _impressionFromProcessReport = false;

            // Defer Clario Priority/Class extraction to subsequent ticks.
            // UIA connection was just reset — doing a depth-12 Chrome traversal on
            // the same tick as GetFinalReportFast causes 15+ second mega-ticks.
            // The retry mechanism handles it on the next tick(s).
            _pendingClarioPriorityRetry = true;
            _clarioPriorityRetryCount = 0;
            _nextClarioPriorityRetryTick64 = 0;

            if (_config.StrokeDetectionEnabled)
            {
                PerformStrokeDetection(currentAccession, reportText);
            }
            else
            {
                InvokeUI(() => _mainForm.SetStrokeState(false));
            }

            // Auto-start custom STT on new case
            if (_config.SttAutoStartOnCase && _config.CustomSttEnabled
                && _sttService != null && !_sttService.IsRecording)
            {
                PerformToggleRecordStt(true);
            }

        }
        else
        {
            // Study closed without new one opening - clear the tracked accession
            _lastNonEmptyAccession = null;
            _needsBaselineCapture = false;
            Logger.Trace("Study closed, no new study opened");

            // Clear impression fixer buttons from editor
            _impFixerEditorAccession = null;
            _cdpService?.SetImpressionFixerButtons(null);

            // Keep normal scrape rate after study close — a new study typically opens within
            // seconds. The idle backoff (3 empty scrapes → 10s, 10 → 30s) will kick in
            // naturally if nothing opens. Don't force 10s immediately.
            if (!_searchingForImpression)
                RestartScrapeTimer(NormalScrapeIntervalMs);

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

            // Hide impression window when no study is open
            if (_config.ShowImpression)
            {
                _searchingForImpression = false;
                _impressionFromProcessReport = false;
                BatchUI(() => _mainForm.HideImpressionWindow());
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
                var impression = (_cdpService?.IsIframeConnected == true)
                    ? ImpressionForm.ExtractImpression(_cdpService.GetStructuredReport())
                      ?? ImpressionForm.ExtractImpression(reportText)
                    : ImpressionForm.ExtractImpression(reportText);
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
    /// <summary>
    /// Compute correlation highlights and apply them as rainbow CSS highlights in the editor.
    /// Uses raw palette colors — CSS handles alpha blending against the dark editor background.
    /// </summary>
    private void ComputeAndApplyRainbow(string reportText)
    {
        var seed = _lastNonEmptyAccession?.GetHashCode();
        var result = CorrelationService.CorrelateReversed(reportText, null, seed);
        if (result.Items.Count == 0) return;

        // Find section boundaries
        int findingsIdx = reportText.IndexOf("FINDINGS:", StringComparison.OrdinalIgnoreCase);
        int impressionIdx = reportText.IndexOf("IMPRESSION:", StringComparison.OrdinalIgnoreCase);
        if (impressionIdx < 0)
            impressionIdx = reportText.IndexOf("RadAI IMPRESSION:", StringComparison.OrdinalIgnoreCase);

        var entries = new List<(string Text, string CssColor, string Section)>();

        foreach (var item in result.Items)
        {
            // Boost colors for dark editor background — raw palette colors look washed out
            // at 50% opacity on dark grey. Max out saturation and push lightness up.
            var color = item.HighlightColor ?? CorrelationService.Palette[item.ColorIndex];
            var boosted = CorrelationService.BoostForDarkBg(color);
            var cssColor = $"{boosted.R},{boosted.G},{boosted.B}";

            // Impression text — highlight for matched groups (non-orphans)
            if (!string.IsNullOrWhiteSpace(item.ImpressionText))
            {
                // Strip impression fixer button text (e.g. "+no change=no change") that walkEditor
                // captures as part of the impression item. These are injected DOM elements whose
                // text bleeds into the reportText but appears at different positions in textContent.
                var impText = item.ImpressionText;
                var fixerIdx = impText.IndexOf("+", StringComparison.Ordinal);
                if (fixerIdx > 0 && impText.IndexOf("change", fixerIdx, StringComparison.OrdinalIgnoreCase) > 0)
                    impText = impText[..fixerIdx].TrimEnd();
                // Also try trimming to last sentence-ending period if fixer text wasn't detected
                if (impText == item.ImpressionText)
                {
                    var lastDot = impText.LastIndexOf('.');
                    if (lastDot > 0 && lastDot < impText.Length - 1)
                        impText = impText[..(lastDot + 1)];
                }

                int searchFrom = impressionIdx >= 0 ? impressionIdx : 0;
                if (reportText.IndexOf(impText, searchFrom, StringComparison.Ordinal) >= 0)
                    entries.Add((impText, cssColor, "impression"));
            }

            // Matched findings — search only within FINDINGS section
            foreach (var finding in item.MatchedFindings)
            {
                if (string.IsNullOrWhiteSpace(finding)) continue;
                int searchFrom = findingsIdx >= 0 ? findingsIdx : 0;
                int searchEnd = impressionIdx >= 0 ? impressionIdx : reportText.Length;
                if (searchFrom < searchEnd && reportText.IndexOf(finding, searchFrom, searchEnd - searchFrom, StringComparison.Ordinal) >= 0)
                    entries.Add((finding, cssColor, "findings"));
            }
        }

        if (entries.Count > 0)
        {
            _cdpService!.ApplyRainbowHighlights(entries);
            foreach (var e in entries)
                Logger.Trace($"  Rainbow entry [{e.Section}]: \"{(e.Text.Length > 60 ? e.Text[..60] + "..." : e.Text)}\" color=({e.CssColor})");
        }

        Logger.Trace($"Rainbow in-editor: {entries.Count} highlight entries from {result.Items.Count} correlations");
    }

    private void UpdateReportPopup(string? reportText, bool immediate = false)
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
            // After Process Report, Mosaic rebuilds the editor over several seconds.
            // Suppress ALL popup updates until the report text stops changing (stable
            // for 2s), so intermediate/partial renders don't flash in the popup.
            if (_processReportPressedForCurrentAccession)
            {
                if (!string.IsNullOrEmpty(reportText) && reportText != _processReportLastSeenText)
                {
                    // Text changed — reset stability timer
                    _processReportLastSeenText = reportText;
                    _processReportTextStableSince = DateTime.UtcNow;
                }

                // Wait for report to have IMPRESSION section AND be stable for 2s.
                // Mosaic rebuilds in stages: template first, then impression — don't release early.
                bool hasImpression = reportText != null
                    && reportText.Contains("IMPRESSION", StringComparison.OrdinalIgnoreCase);
                bool stable = _processReportLastSeenText != null && hasImpression
                    && (DateTime.UtcNow - _processReportTextStableSince).TotalSeconds >= 2;

                bool timedOut = _staleSetTime != default
                    && (DateTime.UtcNow - _staleSetTime).TotalSeconds > 15;

                if (stable || timedOut)
                {
                    _processReportPressedForCurrentAccession = false;
                    Logger.Trace($"Process Report popup hold released: stable={stable}, timedOut={timedOut}, hasImpression={hasImpression}, len={reportText?.Length ?? 0}");
                }
                else
                {
                    if (string.IsNullOrEmpty(reportText) && !string.IsNullOrEmpty(_lastPopupReportText))
                    {
                        _staleSetTime = DateTime.UtcNow;
                        _staleTextStableTime = default;
                        BatchUI(() => { if (!popup.IsDisposed) popup.SetStaleState(true); });
                    }
                    return;
                }
            }

            if (!string.IsNullOrEmpty(reportText) && reportText != _lastPopupReportText)
            {
                Logger.Trace($"Auto-updating popup: report changed ({reportText.Length} chars vs {_lastPopupReportText?.Length ?? 0} chars), baseline={_baselineReport?.Length ?? 0} chars");
                _lastPopupReportText = reportText;
                RequestReportScrapeBurst(PopupChangeBurstMs, "Report changed while popup visible");
                _staleTextStableTime = default;
                // CDP: pass structured report for instant section parsing in diff/rainbow
                var sr = (_cdpService?.IsIframeConnected == true) ? _cdpService.GetStructuredReport() : null;
                if (immediate)
                    InvokeUI(() => { if (!popup.IsDisposed) popup.UpdateReport(reportText, _baselineReport, _baselineIsFromTemplateDb, sr); });
                else
                    BatchUI(() => { if (!popup.IsDisposed) popup.UpdateReport(reportText, _baselineReport, _baselineIsFromTemplateDb, sr); });
            }
            else if (string.IsNullOrEmpty(reportText) && !string.IsNullOrEmpty(_lastPopupReportText))
            {
                Logger.Trace("Popup showing stale content - report being updated");
                _staleSetTime = DateTime.UtcNow;
                _staleTextStableTime = default;
                BatchUI(() => { if (!popup.IsDisposed) popup.SetStaleState(true); });
            }
            else if (!string.IsNullOrEmpty(reportText))
            {
                // Report text matches what's displayed — clear stale indicator
                if (_staleTextStableTime == default)
                    _staleTextStableTime = DateTime.UtcNow;

                bool canClear = (DateTime.UtcNow - _staleTextStableTime).TotalSeconds > 1
                    || (DateTime.UtcNow - _staleSetTime).TotalSeconds > 15;
                if (canClear)
                {
                    BatchUI(() => { if (!popup.IsDisposed) popup.SetStaleState(false); });
                }
            }
        }
    }

    /// <summary>
    /// Update clinical history display, template/gender/stroke/Aidoc alerts.
    /// Future API: replaced by study.opened + report.changed event handlers.
    /// </summary>
    private void UpdateClinicalHistoryAndAlerts(string? currentAccession, string? reportText, long tickStartTick64 = 0)
    {
        if (!_config.ShowClinicalHistory && (!_config.CdpEnabled || !_config.CdpFlashingAlertText))
        {
            ClearCdpAlertTextFlashing();
            return;
        }


        // First, evaluate all alert conditions
        bool newTemplateMismatch = false;
        bool newGenderMismatch = false;
        bool newFimMismatch = false;
        bool newConsistencyMismatch = false;
        List<MismatchResult>? fimMismatches = null;
        List<ConsistencyResult>? consistencyResults = null;
        string? templateDescription = null;
        string? templateName = null;
        string? patientGender = _mosaicReader.LastPatientGender;
        int? patientAge = _mosaicReader.LastPatientAge;
        List<string>? genderMismatches = null;

        // Check template matching (red border when mismatch) if enabled
        // Two checks: (1) study type dropdown vs description, (2) EXAM: line in report vs description.
        // Edge case: dropdown can show correct study type but the editor still has the wrong template loaded.
        if (_config.ShowTemplateMismatch)
        {
            templateDescription = _mosaicReader.LastDescription;
            templateName = _mosaicReader.LastTemplateName;
            bool dropdownMatch = AutomationService.DoBodyPartsMatch(templateDescription, templateName);

            // Also check the EXAM: line from the actual report text
            var examTemplateName = AutomationService.ExtractTemplateName(reportText);
            bool examMatch = AutomationService.DoBodyPartsMatch(templateDescription, examTemplateName);

            newTemplateMismatch = !dropdownMatch || !examMatch;
            // Use whichever name is mismatched for display/correction
            if (!examMatch && dropdownMatch)
                templateName = examTemplateName; // dropdown looks correct but editor has wrong template

            // Auto-correct template if mismatch detected (max 2 attempts: immediate + retry after 15s)
            if (newTemplateMismatch && _config.AttemptCorrectTemplate
                && _templateCorrectionAttempts < 1
                && !string.IsNullOrWhiteSpace(templateDescription)
                && Environment.TickCount64 >= _templateCorrectionNextRetryTick64)
            {
                _templateCorrectionAttempts++;
                _templateCorrectionNextRetryTick64 = Environment.TickCount64 + 15_000; // retry eligible in 15s
                var searchText = AutomationService.BuildTemplateSearchText(templateDescription);
                Logger.Trace($"AttemptCorrectTemplate: Attempt #{_templateCorrectionAttempts} — correcting from '{templateName}' to '{searchText}'");
                bool corrected = _cdpService?.SetStudyType(searchText, templateDescription) == true
                    || _automationService.AttemptCorrectTemplate(templateDescription);
                if (corrected)
                {
                    InvokeUI(() => _mainForm.ShowStatusToast("Template correction attempted"));
                    // Next scrape tick will re-evaluate and clear the red border if fixed
                }
                else
                {
                    InvokeUI(() => _mainForm.ShowStatusToast("Template correction failed"));
                }
            }
        }

        // Check for gender mismatch
        if (_config.GenderCheckEnabled && !string.IsNullOrWhiteSpace(reportText))
        {
            genderMismatches = ClinicalHistoryForm.CheckGenderMismatch(reportText, patientGender);
            newGenderMismatch = genderMismatches.Count > 0;
        }

        // Shared structured report for FIM + consistency checks (one CDP call instead of two)
        var currentDescription = _mosaicReader.LastDescription;
        bool reportChanged = !string.Equals(reportText, _lastMismatchCheckReportText, StringComparison.Ordinal);
        bool descriptionChanged = !string.Equals(currentDescription, _lastConsistencyCheckDescription, StringComparison.Ordinal);
        StructuredReport? sharedStructured = null;
        string? sharedFindings = null, sharedImpression = null;
        bool needsSections = ((_config.FindingsImpressionMismatchEnabled && reportChanged)
                           || (_config.ConsistencyCheckEnabled && (reportChanged || descriptionChanged)))
                           && !string.IsNullOrWhiteSpace(reportText);
        if (needsSections)
        {
            sharedStructured = (_cdpService?.IsIframeConnected == true) ? _cdpService.GetStructuredReport() : null;
            var (f, i) = sharedStructured != null
                ? CorrelationService.ExtractSections(sharedStructured)
                : CorrelationService.ExtractSections(reportText);
            sharedFindings = f;
            sharedImpression = i;
        }

        // Check for findings/impression mismatch
        if (_config.FindingsImpressionMismatchEnabled && !string.IsNullOrWhiteSpace(reportText) && reportChanged)
        {
            try
            {
                fimMismatches = FindingsImpressionChecker.Check(sharedFindings!, sharedImpression!);
                _lastFimMismatches = fimMismatches;
                _lastMismatchCheckReportText = reportText;
                newFimMismatch = fimMismatches.Count > 0;
                if (newFimMismatch != _fimMismatchActive || newFimMismatch)
                    Logger.Trace($"FIM: {(fimMismatches.Count > 0 ? string.Join(", ", fimMismatches.Select(m => m.DisplayName)) : "clear")} (findings={sharedFindings!.Length}c, impression={sharedImpression!.Length}c)");
            }
            catch (Exception ex)
            {
                Logger.Trace($"FIM: error: {ex.Message}");
                _lastFimMismatches = null; // Don't cache stale results on error
            }
        }
        else if (_config.FindingsImpressionMismatchEnabled && _lastFimMismatches != null)
        {
            fimMismatches = _lastFimMismatches;
            newFimMismatch = fimMismatches.Count > 0;
        }

        // Check for measurement/laterality consistency
        if (_config.ConsistencyCheckEnabled && !string.IsNullOrWhiteSpace(reportText)
            && (reportChanged || descriptionChanged))
        {
            try
            {
                consistencyResults = ConsistencyChecker.Check(sharedFindings!, sharedImpression!, currentDescription);
                _lastConsistencyResults = consistencyResults;
                _lastConsistencyCheckReportText = reportText;
                _lastConsistencyCheckDescription = currentDescription;
                newConsistencyMismatch = consistencyResults.Count > 0;
                if (newConsistencyMismatch != _consistencyMismatchActive || newConsistencyMismatch)
                    Logger.Trace($"Consistency: {(consistencyResults.Count > 0 ? string.Join(", ", consistencyResults.Select(r => r.DisplayName)) : "clear")}");
            }
            catch (Exception ex)
            {
                Logger.Trace($"Consistency: error: {ex.Message}");
                _lastConsistencyResults = null; // Don't cache stale results on error
            }
        }
        else if (_config.ConsistencyCheckEnabled && _lastConsistencyResults != null)
        {
            consistencyResults = _lastConsistencyResults;
            newConsistencyMismatch = consistencyResults.Count > 0;
        }

        UpdateCdpAlertTextFlashing(reportText, genderMismatches, fimMismatches, consistencyResults);

        // CDP flashing can run even when clinical history UI is hidden.
        // Skip the rest of the clinical-history/Aidoc UI pipeline in that case.
        if (!_config.ShowClinicalHistory)
        {
            _genderMismatchActive = newGenderMismatch;
            _fimMismatchActive = newFimMismatch;
            _consistencyMismatchActive = newConsistencyMismatch;
            return;
        }

        // Tick budget: skip heavy cross-process UIA operations (Clario traversal, Aidoc scrape)
        // if this tick has already spent 3+ seconds. Deferred work runs on the next tick.
        // Prevents 15+ second mega-ticks on study change when everything fires at once.
        long clarioNow = Environment.TickCount64;
        bool tickBudgetExceeded = tickStartTick64 > 0 && (clarioNow - tickStartTick64) > 3000;

        // Retry Clario priority extraction if it failed on accession change
        // (Clario may not have updated to the new study yet)
        // Capped at 5 retries with exponential backoff (3s, 6s, 12s, 24s, 48s) to prevent
        // the recursive depth-25 Clario traversal from running every tick indefinitely.
        if (!tickBudgetExceeded && _pendingClarioPriorityRetry && !string.IsNullOrEmpty(currentAccession)
            && _clarioPriorityRetryCount < 5 && clarioNow >= _nextClarioPriorityRetryTick64)
        {
            _clarioPriorityRetryCount++;
            _nextClarioPriorityRetryTick64 = clarioNow + (3000L << (_clarioPriorityRetryCount - 1)); // 3s, 6s, 12s, 24s, 48s
            ExtractClarioPriorityAndClass(currentAccession);
            if (!string.IsNullOrEmpty(_automationService.LastClarioPriority))
            {
                _pendingClarioPriorityRetry = false;
                Logger.Trace($"Clario priority retry #{_clarioPriorityRetryCount} succeeded: '{_automationService.LastClarioPriority}'");
                if (_config.StrokeDetectionEnabled)
                {
                    PerformStrokeDetection(currentAccession, reportText);
                }
            }
            else if (_clarioPriorityRetryCount >= 5)
            {
                _pendingClarioPriorityRetry = false;
                Logger.Trace("Clario priority retry exhausted (5 attempts), giving up");
            }
        }

        // Aidoc scraping - check for AI-detected findings
        // Aidoc: single scrape per study. User must disable pulse animation in Aidoc settings
        // (Contextual summary display → No animation) to avoid false positives from color bleed.
        bool aidocAlertActive = false;
        bool aidocScrapedThisTick = false;
        string? aidocFindingText = null;
        List<string>? relevantFindings = null;
        bool prevAidocRelevant = _lastAidocRelevant;
        long aidocNow = Environment.TickCount64;
        if (!tickBudgetExceeded && _config.AidocScrapeEnabled && !_aidocDoneForStudy
            && !string.IsNullOrEmpty(currentAccession)
            && (aidocNow - _lastAidocScrapeTick64) >= 5000)
        {
            _lastAidocScrapeTick64 = aidocNow;
            aidocScrapedThisTick = true;
            try
            {
                var aidocResult = _aidocService.ScrapeShortcutWidget();
                if (aidocResult != null && aidocResult.Findings.Count > 0)
                {
                    _aidocDoneForStudy = true;
                    var studyDescription = _mosaicReader.LastDescription;

                    relevantFindings = aidocResult.Findings
                        .Where(f => f.IsPositive
                            && AidocService.IsRelevantFinding(f.FindingType, studyDescription))
                        .Select(f => f.FindingType)
                        .ToList();

                    if (relevantFindings.Count > 0)
                    {
                        aidocAlertActive = true;
                        aidocFindingText = string.Join(", ", relevantFindings);
                        Logger.Trace($"Aidoc: Relevant findings '{aidocFindingText}' for study '{studyDescription}'");
                        BatchUI(() => _mainForm.ShowStatusToast($"Aidoc: {aidocFindingText} detected", 5000));
                    }

                    _lastAidocFindings = aidocAlertActive ? aidocFindingText : null;
                    _lastAidocRelevantList = aidocAlertActive ? relevantFindings : null;
                    _lastAidocRelevant = aidocAlertActive;
                }
                else if (aidocResult != null)
                {
                    // Widget found but no findings listed — done for this study
                    _aidocDoneForStudy = true;
                    _lastAidocFindings = null;
                    _lastAidocRelevantList = null;
                    _lastAidocRelevant = false;
                }
                // If aidocResult is null (widget not found), don't set _aidocDoneForStudy — retry next tick
            }
            catch (Exception ex)
            {
                Logger.Trace($"Aidoc scrape error: {ex.Message}");
            }
        }

        // Verify Aidoc findings against report text (runs every tick using persisted list,
        // so checkmarks update as the user addresses findings in the report)
        // FlaUI text: only verify against text with U+FFFC (real report editor, not transcript)
        // CDP text: always valid (CDP scrape targets editors[1] directly, never has U+FFFC)
        List<FindingVerification>? aidocVerifications = null;
        var findingsToVerify = relevantFindings ?? _lastAidocRelevantList;
        bool effectiveAidocRelevant = aidocScrapedThisTick ? aidocAlertActive : _lastAidocRelevant;
        if (effectiveAidocRelevant && findingsToVerify != null && !string.IsNullOrEmpty(reportText)
            && (_cdpService?.IsIframeConnected == true || reportText.Contains('\uFFFC')))
        {
            aidocVerifications = AidocFindingVerifier.VerifyFindings(findingsToVerify, reportText);
        }

        // Determine if any alerts are active
        bool anyAlertActive = newTemplateMismatch || newGenderMismatch || newFimMismatch || newConsistencyMismatch || _strokeDetectedActive || effectiveAidocRelevant;

        // Handle visibility based on always-show vs alerts-only mode
        if (_config.AlwaysShowClinicalHistory)
        {
            // ALWAYS-SHOW MODE: Current behavior - window always visible, show clinical history + border colors
            // Uses BatchUI to consolidate into single BeginInvoke (reduces WM_USER spam that kills keyboard hook)

            // Only update if we have content - don't clear during brief processing gaps
            if (!string.IsNullOrWhiteSpace(reportText))
            {
                var sr = (_cdpService?.IsIframeConnected == true) ? _cdpService.GetStructuredReport() : null;
                BatchUI(() => _mainForm.UpdateClinicalHistory(reportText, currentAccession, patientAge, patientGender, sr));
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

            // Update findings/impression mismatch
            if (_config.FindingsImpressionMismatchEnabled && fimMismatches != null)
            {
                var capturedFim = fimMismatches;
                BatchUI(() => _mainForm.UpdateFindingsImpressionMismatch(capturedFim.Count > 0, capturedFim));
            }
            else if (!_config.FindingsImpressionMismatchEnabled && _fimMismatchActive)
            {
                BatchUI(() => _mainForm.UpdateFindingsImpressionMismatch(false, null));
            }

            // Update consistency mismatch
            if (_config.ConsistencyCheckEnabled && consistencyResults != null)
            {
                var capturedConsistency = consistencyResults;
                BatchUI(() => _mainForm.UpdateConsistencyMismatch(capturedConsistency.Count > 0, capturedConsistency));
            }
            else if (!_config.ConsistencyCheckEnabled && _consistencyMismatchActive)
            {
                BatchUI(() => _mainForm.UpdateConsistencyMismatch(false, null));
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
            // Update every tick so verification checkmarks reflect current report text
            if (effectiveAidocRelevant && aidocVerifications != null)
            {
                var captured = aidocVerifications;
                BatchUI(() => _mainForm.SetAidocAppend(captured));
            }
            else if (aidocScrapedThisTick && !aidocAlertActive && prevAidocRelevant)
            {
                // Aidoc scrape ran and found nothing — clear
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
                var sr = (_cdpService?.IsIframeConnected == true) ? _cdpService.GetStructuredReport() : null;
                BatchUI(() => _mainForm.UpdateClinicalHistory(reportText, currentAccession, patientAge, patientGender, sr));
            }

            // Determine highest priority alert to show
            AlertType? alertToShow = null;
            string alertDetails = "";

            if (newConsistencyMismatch && consistencyResults != null)
            {
                alertToShow = AlertType.ConsistencyMismatch;
                alertDetails = string.Join(", ", consistencyResults.Select(r => r.DisplayName));
            }
            else if (newFimMismatch && fimMismatches != null)
            {
                alertToShow = AlertType.FindingsImpressionMismatch;
                alertDetails = string.Join(", ", fimMismatches.Select(m => m.DisplayName));
            }
            else if (newGenderMismatch && genderMismatches != null)
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
            else if (effectiveAidocRelevant && (_lastAidocFindings ?? aidocFindingText) != null)
            {
                alertToShow = AlertType.AidocFinding;
                alertDetails = _lastAidocFindings ?? aidocFindingText ?? "";
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
        _fimMismatchActive = newFimMismatch;
        _consistencyMismatchActive = newConsistencyMismatch;
    }

    private void UpdateCdpAlertTextFlashing(
        string? reportText,
        List<string>? genderMismatches,
        List<MismatchResult>? fimMismatches,
        List<ConsistencyResult>? consistencyResults = null)
    {
        if (_cdpService?.IsIframeConnected != true
            || !_config.CdpEnabled
            || !_config.CdpFlashingAlertText
            || string.IsNullOrWhiteSpace(reportText))
        {
            ClearCdpAlertTextFlashing();
            return;
        }

        var genderTerms = (_config.GenderCheckEnabled ? genderMismatches : null) ?? new List<string>();
        var fimTerms = (_config.FindingsImpressionMismatchEnabled && fimMismatches != null)
            ? fimMismatches.SelectMany(m => m.SearchTerms).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            : new List<string>();

        // Merge consistency search terms into FIM terms (same visual treatment)
        if (_config.ConsistencyCheckEnabled && consistencyResults != null)
        {
            var consistencyTerms = consistencyResults.SelectMany(r => r.SearchTerms)
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Distinct(StringComparer.OrdinalIgnoreCase);
            fimTerms.AddRange(consistencyTerms);
        }

        var normalizedGender = genderTerms
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Select(t => t.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(t => t, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var normalizedFim = fimTerms
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Select(t => t.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(t => t, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (normalizedGender.Length == 0 && normalizedFim.Length == 0)
        {
            ClearCdpAlertTextFlashing();
            return;
        }

        var reportHash = StringComparer.Ordinal.GetHashCode(reportText);
        var signature = $"{reportHash}|g:{string.Join("|", normalizedGender)}|f:{string.Join("|", normalizedFim)}";
        var nowTick64 = Environment.TickCount64;
        if (_cdpAlertFlashApplied
            && string.Equals(signature, _cdpAlertFlashSignature, StringComparison.Ordinal)
            && (nowTick64 - _lastCdpAlertFlashApplyTick64) < 60000)
            return;

        if (_cdpService.UpdateAlertTextFlashing(normalizedGender, normalizedFim))
        {
            _cdpAlertFlashApplied = true;
            _cdpAlertFlashSignature = signature;
            _lastCdpAlertFlashApplyTick64 = nowTick64;
        }
        else
        {
            ClearCdpAlertTextFlashing();
        }
    }

    private void ClearCdpAlertTextFlashing()
    {
        if (_cdpService?.IsIframeConnected == true)
        {
            try { _cdpService.ClearAlertTextFlashing(); }
            catch { }
        }

        _cdpAlertFlashApplied = false;
        _cdpAlertFlashSignature = null;
        _lastCdpAlertFlashApplyTick64 = 0;
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

        // CDP path: structured report gives instant, reliable impression extraction
        var impression = (_cdpService?.IsIframeConnected == true)
            ? ImpressionForm.ExtractImpression(_cdpService.GetStructuredReport())
              ?? ImpressionForm.ExtractImpression(reportText) // fallback if CDP parse fails
            : ImpressionForm.ExtractImpression(reportText);
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
    /// Tries CDP first if enabled and connected, falls back to FlaUI.
    /// </summary>
    private void ExtractClarioPriorityAndClass(string? accession)
    {
        try
        {
            // [Clario CDP] Try CDP first — instant JS eval instead of depth-25 FlaUI tree walk
            if (_config.ClarioCdpEnabled && _cdpService?.IsClarioConnected == true)
            {
                var cdpResult = _cdpService.ClarioExtractPriorityAndClass(accession);
                if (cdpResult != null)
                {
                    // Populate AutomationService properties for downstream consumers
                    _automationService.LastClarioPriority = cdpResult.Priority;
                    _automationService.LastClarioClass = cdpResult.Class;
                    _automationService.IsStrokeStudy =
                        cdpResult.Priority.Contains("Stroke", StringComparison.OrdinalIgnoreCase) ||
                        cdpResult.Class.Contains("Stroke", StringComparison.OrdinalIgnoreCase);
                    Logger.Trace($"Extracted Clario Priority='{cdpResult.Priority}', Class='{cdpResult.Class}' [via CDP]");
                    return;
                }
                Logger.Trace("Clario CDP extraction returned null, falling back to FlaUI");
            }

            // FlaUI fallback
            var priorityData = _automationService.ExtractClarioPriorityAndClass(accession);
            if (priorityData != null)
            {
                Logger.Trace($"Extracted Clario Priority='{priorityData.Priority}', Class='{priorityData.Class}' [via FlaUI]");
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
    private static string ApplySpokenPunctuation(string text, bool expandContractions = true)
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

        if (expandContractions)
            result = ExpandContractions(result);

        // Clean up spaces before punctuation marks (e.g., "word . next" → "word. next")
        result = System.Text.RegularExpressions.Regex.Replace(result, @"\s+([.,;!?])", "$1");

        return result;
    }

    /// <summary>
    /// [CustomSTT] Expand contractions that STT providers favor over spoken full forms.
    /// Used in both auto-punctuation and manual punctuation modes.
    /// </summary>
    private static string ExpandContractions(string text)
    {
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
            text = System.Text.RegularExpressions.Regex.Replace(
                text, pattern,
                m => char.IsUpper(m.Value[0])
                    ? char.ToUpper(replacement[0]) + replacement[1..]
                    : replacement,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }
        return text;
    }

    // ── Voice command/macro detection ──────────────────────────────────────
    // Scans raw transcript for voice commands ("process report", "sign report")
    // and voice macros ("insert/input/macro {name}"). Returns the rightmost match.

    private enum VoiceTriggerKind { None, Command, Macro }

    private readonly record struct VoiceTriggerResult(
        VoiceTriggerKind Kind,
        string PrefixText,     // everything before the trigger phrase
        string? ActionName,    // for Command triggers: the action constant
        MacroConfig? Macro     // for Macro triggers: the matched macro
    );

    /// <summary>Map an original-string index to the closest stripped-string index.</summary>
    private static int OriginalToStrippedIndex(int origIdx, System.Collections.Generic.List<int> indexMap)
    {
        for (int i = 0; i < indexMap.Count; i++)
            if (indexMap[i] >= origIdx) return i;
        return indexMap.Count;
    }

    private VoiceTriggerResult CheckVoiceTrigger(string rawTranscript)
    {
        var lower = rawTranscript.ToLowerInvariant();

        // Fix STT concatenation: Deepgram often merges trigger word with macro name
        // (e.g. "macronochange" instead of "macro no change"). Fix before stripping.
        // Apply to both lower and rawTranscript so index mapping stays consistent.
        var triggerPrefixes = new[] { "insert", "input", "macro" };
        var fixedLower = FixConcatenatedMacroTriggers(lower, triggerPrefixes, _config.Macros,
            _cdpService?.MosaicMacros);
        if (fixedLower != lower)
        {
            // Reconstruct rawTranscript with the same split applied (case-preserving)
            rawTranscript = ApplySameInsertions(rawTranscript, lower, fixedLower);
            lower = fixedLower;
        }

        // Build a punctuation-stripped version for macro matching, with index map back to original.
        // STT often transcribes "macro, no change." — punctuation must not break matching.
        var strippedChars = new System.Collections.Generic.List<char>(lower.Length);
        var indexMap = new System.Collections.Generic.List<int>(lower.Length); // stripped pos → original pos
        for (int i = 0; i < lower.Length; i++)
        {
            char c = lower[i];
            if (char.IsLetterOrDigit(c) || c == ' ')
            {
                strippedChars.Add(c);
                indexMap.Add(i);
            }
        }
        var stripped = new string(strippedChars.ToArray());

        // Track the rightmost match (positions are in stripped string)
        int bestStrippedIndex = -1;
        VoiceTriggerKind bestKind = VoiceTriggerKind.None;
        string? bestAction = null;
        MacroConfig? bestMacro = null;
        int bestMatchLen = 0;

        // Check voice commands (match in both original and stripped)
        var commands = new (string phrase, string action)[]
        {
            ("process report", Actions.ProcessReport),
            ("sign report", Actions.SignReport),
        };

        foreach (var (phrase, action) in commands)
        {
            // Try original first (preserves exact position)
            int idx = lower.LastIndexOf(phrase);
            if (idx >= 0)
            {
                // Convert to stripped position for consistent comparison
                int sIdx = OriginalToStrippedIndex(idx, indexMap);
                if (sIdx > bestStrippedIndex)
                {
                    bestStrippedIndex = sIdx;
                    bestKind = VoiceTriggerKind.Command;
                    bestAction = action;
                    bestMacro = null;
                    bestMatchLen = phrase.Length;
                }
            }
            // Also try stripped (handles "process, report" etc.)
            idx = stripped.LastIndexOf(phrase);
            if (idx > bestStrippedIndex)
            {
                bestStrippedIndex = idx;
                bestKind = VoiceTriggerKind.Command;
                bestAction = action;
                bestMacro = null;
                bestMatchLen = phrase.Length;
            }
        }

        // Check voice macros: "insert/input/macro " + exact macro name
        // MT macros checked first, then Mosaic macros as fallback.
        // Match against stripped string so "macro, no change." matches "macro no change".
        var triggerWords = new[] { "insert ", "input ", "macro " };

        foreach (var macro in _config.Macros)
        {
            if (!macro.Enabled || !macro.Voice || string.IsNullOrWhiteSpace(macro.Name))
                continue;

            var macroNameLower = macro.Name.ToLowerInvariant();
            foreach (var trigger in triggerWords)
            {
                var fullPhrase = trigger + macroNameLower;
                int idx = stripped.LastIndexOf(fullPhrase);
                // Prefer rightmost match; at same position prefer longest (e.g. "no change old" over "no change")
                if (idx >= 0 && (idx > bestStrippedIndex || (idx == bestStrippedIndex && fullPhrase.Length > bestMatchLen)))
                {
                    bestStrippedIndex = idx;
                    bestKind = VoiceTriggerKind.Macro;
                    bestAction = null;
                    bestMacro = macro;
                    bestMatchLen = fullPhrase.Length;
                }
            }
        }

        // Mosaic macros fallback: if no MT macro matched, check Mosaic macros from CDP
        if (bestKind != VoiceTriggerKind.Macro && _cdpService?.MosaicMacros is { Count: > 0 } mosaicMacros)
        {
            foreach (var (name, text) in mosaicMacros)
            {
                var macroNameLower = name.ToLowerInvariant();
                foreach (var trigger in triggerWords)
                {
                    var fullPhrase = trigger + macroNameLower;
                    int idx = stripped.LastIndexOf(fullPhrase);
                    if (idx >= 0 && (idx > bestStrippedIndex || (idx == bestStrippedIndex && fullPhrase.Length > bestMatchLen)))
                    {
                        bestStrippedIndex = idx;
                        bestKind = VoiceTriggerKind.Macro;
                        bestAction = null;
                        // Wrap Mosaic macro as a MacroConfig for uniform handling
                        bestMacro = new MacroConfig { Name = name, Text = text, Voice = true, Enabled = true };
                        bestMatchLen = fullPhrase.Length;
                    }
                }
            }
        }

        if (bestKind == VoiceTriggerKind.None)
            return new VoiceTriggerResult(VoiceTriggerKind.None, rawTranscript, null, null);

        // Map stripped position back to original string for prefix extraction
        int originalIndex = (bestStrippedIndex < indexMap.Count) ? indexMap[bestStrippedIndex] : 0;
        var prefix = rawTranscript[..originalIndex].TrimEnd();
        return new VoiceTriggerResult(bestKind, prefix, bestAction, bestMacro);
    }

    /// <summary>
    /// Apply the same space insertions from lower→fixedLower to the original-case string.
    /// Walks both strings in parallel, inserting spaces where fixedLower has them but lower doesn't.
    /// </summary>
    private static string ApplySameInsertions(string original, string lower, string fixedLower)
    {
        if (fixedLower.Length <= lower.Length) return original;
        var sb = new System.Text.StringBuilder(original.Length + (fixedLower.Length - lower.Length));
        int oi = 0; // position in original/lower
        int fi = 0; // position in fixedLower
        while (fi < fixedLower.Length && oi < lower.Length)
        {
            if (lower[oi] == fixedLower[fi])
            {
                sb.Append(original[oi]);
                oi++;
                fi++;
            }
            else if (fixedLower[fi] == ' ')
            {
                // Space was inserted — add it to output
                sb.Append(' ');
                fi++;
            }
            else
            {
                // Mismatch — shouldn't happen, just copy original
                sb.Append(original[oi]);
                oi++;
                fi++;
            }
        }
        // Append remaining
        while (oi < original.Length)
        {
            sb.Append(original[oi]);
            oi++;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Fix STT concatenation errors where the trigger word is merged with the macro name.
    /// Deepgram often merges "macro" with the next word(s): "macronochange" or "macrono change".
    /// This normalizes such concatenations back to "macro no change" using known macro names.
    /// </summary>
    private static string FixConcatenatedMacroTriggers(string text, string[] triggerPrefixes,
        List<MacroConfig> mtMacros, Dictionary<string, string>? mosaicMacros)
    {
        // Collect all known macro names: original form (with spaces) and collapsed (no spaces)
        var macrosByCollapsed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var m in mtMacros)
        {
            if (m.Enabled && m.Voice && !string.IsNullOrWhiteSpace(m.Name))
            {
                var collapsed = m.Name.ToLowerInvariant().Replace(" ", "");
                macrosByCollapsed[collapsed] = m.Name.ToLowerInvariant();
            }
        }
        if (mosaicMacros != null)
        {
            foreach (var (name, _) in mosaicMacros)
            {
                var collapsed = name.ToLowerInvariant().Replace(" ", "");
                macrosByCollapsed.TryAdd(collapsed, name.ToLowerInvariant());
            }
        }

        if (macrosByCollapsed.Count == 0) return text;

        // For each trigger prefix, check if any part of the text contains
        // the prefix concatenated with a known macro name (spaces removed).
        // Replace the concatenated form with "trigger originalname" (with proper spaces).
        foreach (var prefix in triggerPrefixes)
        {
            foreach (var (collapsed, original) in macrosByCollapsed)
            {
                var concatenated = prefix + collapsed; // e.g. "macronochange"
                int idx = text.IndexOf(concatenated, StringComparison.OrdinalIgnoreCase);
                if (idx < 0) continue;

                // Verify it's at a word boundary (start of string or preceded by space)
                if (idx > 0 && text[idx - 1] != ' ') continue;

                // Replace with properly spaced version: "macro no change"
                var replacement = prefix + " " + original;
                text = text[..idx] + replacement + text[(idx + concatenated.Length)..];
                Logger.Trace($"Voice trigger fix: \"{concatenated}\" → \"{replacement}\"");
                return text;
            }
        }

        return text;
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

    /// <summary>
    /// [CustomSTT] Apply user-defined word replacements from config.
    /// Uses word-boundary matching with case preservation (same logic as ExpandContractions).
    /// </summary>
    private string ApplyCustomReplacements(string text)
    {
        if (_config.SttCustomReplacements.Count == 0) return text;

        foreach (var entry in _config.SttCustomReplacements)
        {
            if (!entry.Enabled || string.IsNullOrWhiteSpace(entry.Find)) continue;
            var pattern = @"\b" + System.Text.RegularExpressions.Regex.Escape(entry.Find) + @"\b";
            var replacement = entry.Replace ?? "";
            text = System.Text.RegularExpressions.Regex.Replace(
                text, pattern,
                m => PreserveCaseReplace(m.Value, replacement),
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }
        return text;
    }

    /// <summary>
    /// Insert a newline after each sentence-ending period.
    /// Skips periods in decimal numbers (e.g. "1.2", "3.5").
    /// </summary>
    private static string InsertNewlineAfterSentences(string text)
    {
        // Match a period followed by a space (sentence boundary),
        // but NOT when the period is between digits (decimal number like 1.2)
        return System.Text.RegularExpressions.Regex.Replace(
            text, @"(?<!\d)\.(\s+)", ".\n");
    }

    /// <summary>
    /// Replace text while preserving the case pattern of the original match.
    /// ALL CAPS → ALL CAPS, Title Case → Title Case, otherwise use replacement as-is.
    /// </summary>
    private static string PreserveCaseReplace(string original, string replacement)
    {
        if (replacement.Length == 0) return replacement;
        if (original.Length > 0 && original == original.ToUpperInvariant())
            return replacement.ToUpperInvariant();
        if (original.Length > 0 && char.IsUpper(original[0]))
            return char.ToUpper(replacement[0]) + replacement[1..];
        return replacement;
    }

    /// <summary>
    /// Compare the patient name in Mosaic with the topmost InteleViewer window.
    /// InteleViewer title: "LASTNAME^FIRSTNAME^MID - RP Cloud ..."
    /// Mosaic: LastPatientName is title-cased like "Collins Patrick".
    /// Match by comparing last name (first token) case-insensitively.
    /// </summary>
    private void CheckPatientMismatch()
    {
        var mosaicName = _mosaicReader.LastPatientName;
        if (string.IsNullOrWhiteSpace(mosaicName))
        {
            if (_patientMismatchActive)
            {
                _patientMismatchActive = false;
                BatchUI(() => _mainForm.SetPatientMismatchState(false, null, null));
            }
            return;
        }

        var ivDicomName = NativeWindows.GetTopmostInteleViewerPatientName();
        if (ivDicomName == null)
        {
            // No InteleViewer window open — clear mismatch
            if (_patientMismatchActive)
            {
                _patientMismatchActive = false;
                BatchUI(() => _mainForm.SetPatientMismatchState(false, null, null));
            }
            return;
        }

        // Parse InteleViewer DICOM name: "LASTNAME^FIRSTNAME^MID" → extract last name
        // Last name may be multi-word (e.g. "SOTO PENA^ALEXANDER^MARTIN")
        var ivParts = ivDicomName.Split('^');
        var ivLastName = ivParts[0].Trim();

        // Mosaic name is space-delimited: "Soto Pena Alexander Martin"
        // Check if Mosaic name starts with the IV last name (handles multi-word last names)
        bool isMatch = mosaicName.StartsWith(ivLastName, StringComparison.OrdinalIgnoreCase);

        if (!isMatch && !_patientMismatchActive)
        {
            _patientMismatchActive = true;
            // Format display names
            var ivDisplay = ivDicomName.Replace('^', ' ').Trim();
            BatchUI(() => _mainForm.SetPatientMismatchState(true, mosaicName, ivDisplay));
            Logger.Trace($"Patient mismatch detected: Mosaic='{mosaicName}', InteleViewer='{ivDicomName}'");
        }
        else if (isMatch && _patientMismatchActive)
        {
            _patientMismatchActive = false;
            BatchUI(() => _mainForm.SetPatientMismatchState(false, null, null));
            Logger.Trace("Patient mismatch resolved");
        }
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

    public void ManualCorrectTemplate()
    {
        var description = _mosaicReader.LastDescription;
        if (string.IsNullOrWhiteSpace(description))
        {
            InvokeUI(() => _mainForm.ShowStatusToast("No study description available"));
            return;
        }

        // Run on action thread via TriggerAction
        TriggerAction("__ManualCorrectTemplate__", "ContextMenu");
    }

    public string GetPriorDebugInfo()
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("=== Get Prior Debug ===");
        sb.AppendLine($"--- Raw (from InteleViewer) ---");
        sb.AppendLine(_lastPriorRawText ?? "(no prior captured)");
        sb.AppendLine($"--- Formatted ---");
        sb.AppendLine(_lastPriorFormattedText ?? "(none)");
        return sb.ToString().TrimEnd();
    }

    public string GetCriticalFindingsDebugInfo()
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("=== Critical Findings Debug ===");
        sb.AppendLine($"--- Raw (from Clario) ---");
        sb.AppendLine(_lastCriticalRawNote ?? "(no note captured)");
        sb.AppendLine($"--- Formatted ---");
        sb.AppendLine(_lastCriticalFormattedText ?? "(none)");
        return sb.ToString().TrimEnd();
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

        var success = CreateClarioCriticalNoteWithFallback();
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

    /// <summary>
    /// Create critical communication note in Clario — tries CDP first, falls back to FlaUI.
    /// </summary>
    private bool CreateClarioCriticalNoteWithFallback()
    {
        // [Clario CDP] Try CDP first
        if (_config.ClarioCdpEnabled && _cdpService?.IsClarioConnected == true)
        {
            var success = _cdpService.ClarioCreateCriticalNote();
            if (success)
            {
                Logger.Trace("Critical note created [via CDP]");
                return true;
            }
            Logger.Trace("Clario CDP critical note creation failed, falling back to FlaUI");
        }

        // FlaUI fallback
        return _automationService.CreateCriticalCommunicationNote();
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
        StopMosaicScrapeTimer();
        _hidService?.Dispose();
        _keyboardService?.Dispose();
        _automationService.Dispose();
        _pipeService?.Dispose();
        _sttService?.Dispose();  // [CustomSTT]
        _cdpService?.Dispose();  // [CDP]
        _llmService?.Dispose();  // [LLM]
        _actionEvent.Dispose();
    }
}
