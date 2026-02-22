using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Main floating widget window - the small "Mosaic Tools" bar.
/// Matches Python's MosaicToolsApp UI.
/// </summary>
public class MainForm : Form
{
    private readonly Configuration _config;
    private readonly ActionController _controller;
    
    // UI Elements
    private readonly Label _titleLabel;
    private readonly Label _dragHandle;
    private readonly Panel _innerFrame;
    private readonly Panel _rvuPanel;
    private readonly Label _rvuValueLabel;
    private readonly Label _rvuSeparatorLabel;
    private readonly Label _rvuSuffixLabel;

    // Critical studies indicator
    private readonly Panel _criticalPanel;
    private readonly Label _criticalIconLabel;
    private readonly Label _criticalCountLabel;

    // RVU Counter
    private readonly RvuCounterService _rvuCounterService;
    private System.Windows.Forms.Timer? _rvuTimer;

    // RVU multi-metric support
    private readonly List<Label> _rvuMetricLabels = new();
    private static readonly Font _metricFontBold = new("Segoe UI", 9, FontStyle.Bold);
    private static readonly Font _metricFontRegular = new("Segoe UI", 9, FontStyle.Regular);
    private System.Windows.Forms.Timer? _carouselTimer;
    private int _carouselIndex;
    private RvuPopupForm? _rvuPopup;
    private RvuPopupForm? _rvuDrawer; // persistent drawer for vertical stack

    // Pace car alternation
    private System.Windows.Forms.Timer? _paceCarTimer;
    private bool _showingPaceCar;

    // Connectivity Monitor
    private readonly ConnectivityService _connectivityService;
    private readonly Panel _connectivityPanel;
    private readonly Label[] _connectivityDots = new Label[4];
    private readonly ToolTip _connectivityTooltip;
    
    // [CustomSTT] Transcription window
    private TranscriptionForm? _transcriptionWindow;

    // Child Windows
    private FloatingToolbarForm? _toolbarWindow;
    private IndicatorForm? _indicatorWindow;
    private ClinicalHistoryForm? _clinicalHistoryWindow;
    private ImpressionForm? _impressionWindow;
    
    // System Tray (headless mode)
    private NotifyIcon? _trayIcon;
    
    // Drag state
    private Point _dragStart;
    private bool _dragging;

    // Settings dialog state (suppress topmost reassertion while open)
    private volatile bool _settingsOpen;
    
    // Toast management
    private readonly List<Form> _activeToasts = new();

    // Update service
    private readonly UpdateService _updateService = new();

    // RecoMD service
    private readonly RecoMdService _recoMdService = new();

    public MainForm(Configuration config)
    {
        _config = config;
        _controller = new ActionController(config, this, _recoMdService);
        _rvuCounterService = new RvuCounterService(config);
        _connectivityService = new ConnectivityService(config);
        _connectivityTooltip = new ToolTip { InitialDelay = 200, ReshowDelay = 100 };

        // Form properties (borderless, topmost, small)
        // Width adjusts based on RVU counter and connectivity monitor settings
        int baseWidth = 145; // Base for "Mosaic Tools" title + drag handle
        int rvuWidth = GetRvuPanelWidth();
        int connectivityWidth = GetConnectivityPanelWidth();
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        DoubleBuffered = true;
        Text = "MosaicToolsMainWindow";  // Hidden title for AHK targeting
        BackColor = Color.FromArgb(51, 51, 51); // #333333
        Size = new Size(baseWidth + rvuWidth + connectivityWidth, 40);
        StartPosition = FormStartPosition.Manual;
        Location = ScreenHelper.EnsureOnScreen(_config.WindowX, _config.WindowY, Size.Width, Size.Height);

        // Set application icon (embedded in exe)
        try
        {
            Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        }
        catch { /* Use default if extraction fails */ }

        // Inner frame (black)
        _innerFrame = new Panel
        {
            BackColor = Color.Black,
            Dock = DockStyle.Fill,
            Padding = new Padding(1)
        };
        Controls.Add(_innerFrame);
        
        // Drag handle
        _dragHandle = new Label
        {
            Text = "⋮",
            Font = new Font("Segoe UI", 14),
            ForeColor = Color.FromArgb(85, 85, 85),
            BackColor = Color.Black,
            AutoSize = false,
            Width = 20,
            Dock = DockStyle.Left,
            TextAlign = ContentAlignment.MiddleCenter,
            Cursor = Cursors.SizeAll
        };
        _dragHandle.MouseDown += OnDragStart;
        _dragHandle.MouseMove += OnDragMove;
        _dragHandle.MouseUp += OnDragEnd;

        // Title label (clickable -> settings)
        _titleLabel = new Label
        {
            Text = "Mosaic Tools",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(204, 204, 204),
            BackColor = Color.Black,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Cursor = Cursors.Hand
        };
        _titleLabel.MouseEnter += (_, _) => _titleLabel.ForeColor = Color.White;
        _titleLabel.MouseLeave += (_, _) => _titleLabel.ForeColor = Color.FromArgb(204, 204, 204);
        _titleLabel.Click += (_, _) => OpenSettings();

        // RVU panel (shown when RVU counter is enabled)
        _rvuPanel = new Panel
        {
            BackColor = Color.Black,
            Width = rvuWidth,
            Dock = DockStyle.Right,
            Visible = _config.RvuMetrics != RvuMetric.None || _config.PaceCarEnabled
        };
        typeof(Panel).GetProperty("DoubleBuffered",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .SetValue(_rvuPanel, true);
        _rvuPanel.MouseEnter += OnRvuPanelMouseEnter;
        _rvuPanel.MouseLeave += OnRvuPanelMouseLeave;

        // RVU value label (Carolina blue)
        _rvuValueLabel = new Label
        {
            Text = "",
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = Color.FromArgb(75, 156, 211), // Carolina blue
            BackColor = Color.Black,
            AutoSize = true
        };
        _rvuPanel.Controls.Add(_rvuValueLabel);

        // RVU separator label (white, only visible in Both mode)
        _rvuSeparatorLabel = new Label
        {
            Text = " |  ",
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(204, 204, 204), // White like title
            BackColor = Color.Black,
            AutoSize = true,
            Visible = false
        };
        _rvuPanel.Controls.Add(_rvuSeparatorLabel);

        // RVU suffix label
        _rvuSuffixLabel = new Label
        {
            Text = "RVU",
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(75, 156, 211), // Carolina blue
            BackColor = Color.Black,
            AutoSize = true
        };
        _rvuPanel.Controls.Add(_rvuSuffixLabel);

        // Connectivity monitor panel - 2x2 grid of small dots
        _connectivityPanel = new Panel
        {
            BackColor = Color.Black,
            Width = connectivityWidth,
            Dock = DockStyle.Right,
            Visible = _config.ConnectivityMonitorEnabled,
            Cursor = Cursors.Hand
        };
        _connectivityPanel.Click += OnConnectivityPanelClick;

        // Create 4 dots in 2x2 grid layout: [Mirth, Mosaic] / [Clario, IV]
        var dotSize = 4;
        var dotSpacing = 2;
        var gridWidth = 2 * dotSize + dotSpacing;
        var gridHeight = 2 * dotSize + dotSpacing;
        var startX = connectivityWidth - gridWidth - 5; // Align to right with 5px padding
        var startY = (40 - gridHeight) / 2; // Center vertically

        string[] serverNames = { "Mirth", "Mosaic", "Clario", "InteleViewer" };
        for (int i = 0; i < 4; i++)
        {
            int col = i % 2;
            int row = i / 2;
            _connectivityDots[i] = new Label
            {
                AutoSize = false,
                Size = new Size(dotSize, dotSize),
                Location = new Point(startX + col * (dotSize + dotSpacing), startY + row * (dotSize + dotSpacing)),
                BackColor = Color.FromArgb(74, 74, 74), // Dark gray (unknown)
                Tag = serverNames[i]
            };
            // Make it round
            using var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, dotSize, dotSize);
            _connectivityDots[i].Region = new Region(path);
            _connectivityDots[i].Click += OnConnectivityPanelClick;
            _connectivityPanel.Controls.Add(_connectivityDots[i]);
        }

        // Critical studies indicator panel (shown when critical notes placed)
        _criticalPanel = new Panel
        {
            BackColor = Color.Black,
            Width = 36,
            Dock = DockStyle.Left,
            Visible = false, // Hidden until critical studies exist
            Cursor = Cursors.Hand
        };
        _criticalPanel.Click += OnCriticalPanelClick;

        // Warning icon - bright red triangle with count
        _criticalIconLabel = new Label
        {
            Text = "\u26A0",  // ⚠ warning sign
            Font = new Font("Segoe UI", 12, FontStyle.Regular),
            ForeColor = Color.FromArgb(255, 60, 60), // Bright red
            BackColor = Color.Black,
            AutoSize = false,
            Size = new Size(24, 38),
            TextAlign = ContentAlignment.MiddleCenter,
            Location = new Point(0, 0)
        };
        _criticalIconLabel.Click += OnCriticalPanelClick;
        _criticalPanel.Controls.Add(_criticalIconLabel);

        // Count label (right of icon)
        _criticalCountLabel = new Label
        {
            Text = "0",
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            ForeColor = Color.FromArgb(255, 120, 120), // Light red text
            BackColor = Color.Black,
            AutoSize = false,
            Size = new Size(14, 38),
            TextAlign = ContentAlignment.MiddleLeft,
            Location = new Point(22, 0)
        };
        _criticalCountLabel.Click += OnCriticalPanelClick;
        _criticalPanel.Controls.Add(_criticalCountLabel);

        // Tooltip for critical panel
        var toolTip = new ToolTip();
        toolTip.SetToolTip(_criticalPanel, "Critical notes placed this session (click to view)");
        toolTip.SetToolTip(_criticalIconLabel, "Critical notes placed this session (click to view)");
        toolTip.SetToolTip(_criticalCountLabel, "Critical notes placed this session (click to view)");

        // Add controls: Fill first (laid out last), then Left/Right (laid out first)
        // WinForms docks in REVERSE order of Controls collection
        _innerFrame.Controls.Add(_titleLabel);   // index 0, laid out last (fills remaining)
        _innerFrame.Controls.Add(_criticalPanel); // index 1, laid out fourth (takes left after drag handle)
        _innerFrame.Controls.Add(_dragHandle);   // index 2, laid out third (takes left)
        _innerFrame.Controls.Add(_rvuPanel);     // index 3, laid out second from right
        _innerFrame.Controls.Add(_connectivityPanel); // index 4, laid out first (takes rightmost)

        // Context menu (for normal mode right-click)
        var contextMenu = new ContextMenuStrip();
        contextMenu.Items.Add("Settings", null, (_, _) => OpenSettings());
        contextMenu.Items.Add("Show Log", null, (_, _) => ShowLogFile());
        contextMenu.Items.Add("Reload", null, (_, _) => BeginInvoke(() => ReloadApp()));
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add("Exit", null, (_, _) => BeginInvoke(() => ExitApp()));
        ContextMenuStrip = contextMenu;

        // Always create system tray icon
        var trayMenu = new ContextMenuStrip();
        trayMenu.Items.Add("Settings", null, (_, _) => OpenSettings());
        trayMenu.Items.Add("Show Log", null, (_, _) => ShowLogFile());
        trayMenu.Items.Add(new ToolStripSeparator());
        trayMenu.Items.Add("Exit", null, (_, _) => BeginInvoke(() => ExitApp()));

        _trayIcon = new NotifyIcon
        {
            Text = App.IsHeadless ? "Mosaic Tools (Headless)" : "Mosaic Tools",
            Icon = Icon ?? SystemIcons.Application,
            ContextMenuStrip = trayMenu,
            Visible = true
        };
        _trayIcon.DoubleClick += (_, _) => OpenSettings();

        // Defer initialization until form is shown (handle exists)
        Shown += OnFormShown;
    }
    
    private void OnFormShown(object? sender, EventArgs e)
    {
        // Initialize child windows
        if (_config.FloatingToolbarEnabled)
        {
            ToggleFloatingToolbar(true);
        }
        
        // Indicator: only show if NOT headless AND setting is enabled
        // Don't show on startup if "hide when no study" is enabled - scrape will show it when study opens
        if (!App.IsHeadless && _config.IndicatorEnabled && !_config.HideIndicatorWhenNoStudy)
        {
            ToggleIndicator(true);
        }

        // Clinical History window: show if enabled and scrape mosaic is on
        // Don't show on startup if "hide when no study" is enabled - scrape will show it when study opens
        // Also don't show in alerts-only mode - window only appears when alerts trigger
        if (_config.ShowClinicalHistory &&
            _config.AlwaysShowClinicalHistory && !_config.HideClinicalHistoryWhenNoStudy)
        {
            ToggleClinicalHistory(true);
        }

        // Start controller (HID, hotkeys, etc.)
        _controller.Start();

        // Subscribe to critical studies changes
        _controller.CriticalStudiesChanged += UpdateCriticalIndicator;

        // Log window info for AHK debugging
        Logger.Trace($"MainForm HWND: 0x{Handle:X} Title: '{Text}' Class: WindowsForms");

        // Clean up old version from previous update
        UpdateService.CleanupOldVersion();

        // Check if version changed (for What's New popup)
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var versionStr = version != null ? $" v{version.Major}.{version.Minor}.{version.Build}" : "";
        var currentVersionStr = version != null ? $"{version.Major}.{version.Minor}.{version.Build}" : "1.0.0";
        string modeStr = App.IsHeadless ? " [Headless]" : "";

        var lastSeenVersion = _config.LastSeenVersion;
        var versionChanged = lastSeenVersion != currentVersionStr;

        if (versionChanged)
        {
            // Show What's New popup
            var whatsNew = new WhatsNewForm(lastSeenVersion, currentVersionStr);
            whatsNew.Show();

            // Update last seen version
            _config.LastSeenVersion = currentVersionStr;
            _config.Save();
        }
        else
        {
            ShowStatusToast($"Mosaic Tools{versionStr} Started ({_config.DoctorName}){modeStr}", 2000);
        }

        // Check for updates (async, non-blocking) - skip in headless mode
        if (_config.AutoUpdateEnabled && !App.IsHeadless)
        {
            _ = CheckForUpdatesAsync();
        }

        // Start RVU counter timer if enabled (skip in headless mode - no UI to display it)
        bool hasRvuContent = _config.RvuMetrics != RvuMetric.None || _config.PaceCarEnabled;
        if (hasRvuContent && !App.IsHeadless)
        {
            UpdateRvuDisplay(); // Initial update
            _rvuTimer = new System.Windows.Forms.Timer { Interval = 5000 }; // Update every 5 seconds
            _rvuTimer.Tick += (_, _) => UpdateRvuDisplay();
            _rvuTimer.Start();

            // Subscribe to pipe shift info updates for immediate RVU display refresh
            _controller.PipeService.ShiftInfoUpdated += () =>
            {
                if (InvokeRequired)
                    BeginInvoke(UpdateRvuDisplay);
                else
                    UpdateRvuDisplay();
            };

            // Start pace car alternation timer if enabled
            if (_config.PaceCarEnabled)
            {
                _paceCarTimer = new System.Windows.Forms.Timer
                {
                    Interval = _config.PaceCarAlternateSeconds * 1000
                };
                _paceCarTimer.Tick += (_, _) =>
                {
                    _showingPaceCar = !_showingPaceCar;
                    UpdateRvuDisplay();
                };
                _paceCarTimer.Start();
            }
        }

        // Subscribe to distraction alerts from RVUCounter — play escalating beeps
        // Higher pitch (1200Hz+) and longer than start/stop beeps (800Hz) to sound distinct
        _controller.PipeService.DistractionAlertReceived += alert =>
        {
            var volume = _config.DistractionAlertVolume;
            if (alert.AlertLevel <= 1)
            {
                // Level 1: two rapid beeps
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    AudioService.PlayBeep(1200, 400, volume);
                    Thread.Sleep(100);
                    AudioService.PlayBeep(1200, 400, volume);
                });
            }
            else if (alert.AlertLevel == 2)
            {
                // Level 2: three beeps, rising pitch
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    AudioService.PlayBeep(1200, 400, volume);
                    Thread.Sleep(100);
                    AudioService.PlayBeep(1400, 400, volume);
                    Thread.Sleep(100);
                    AudioService.PlayBeep(1600, 400, volume);
                });
            }
            else
            {
                // Level 3+: four rapid beeps, high pitch, louder
                var louder = Math.Min(volume * 1.5, 0.15);
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    AudioService.PlayBeep(1400, 350, louder);
                    Thread.Sleep(80);
                    AudioService.PlayBeep(1600, 350, louder);
                    Thread.Sleep(80);
                    AudioService.PlayBeep(1600, 350, louder);
                    Thread.Sleep(80);
                    AudioService.PlayBeep(1800, 350, louder);
                });
            }
        };

        // Start connectivity monitor if enabled (skip in headless mode - no UI to display it)
        if (_config.ConnectivityMonitorEnabled && !App.IsHeadless)
        {
            _connectivityService.StatusChanged += UpdateConnectivityDots;
            _connectivityService.Start();
        }
    }

    /// <summary>
    /// Calculate the RVU panel width based on selected metrics and layout.
    /// </summary>
    private int GetRvuPanelWidth()
    {
        if (_config.RvuMetrics == RvuMetric.None && !_config.PaceCarEnabled)
            return 0;

        var metrics = _config.RvuMetrics;
        int count = CountSetFlags(metrics);

        if (count == 0 && !_config.PaceCarEnabled)
            return 75; // Fallback

        int baseWidth;
        if (count == 0)
            baseWidth = 0;
        else if (count <= 2)
        {
            // Horizontal layout: ~80px per metric + right padding
            baseWidth = count * 80 + 8;
        }
        else
        {
            // 3+ metrics: depends on layout
            baseWidth = _config.RvuOverflowLayout switch
            {
                RvuOverflowLayout.Horizontal => count * 80 + 8,
                RvuOverflowLayout.VerticalStack => 0, // All metrics in drawer
                RvuOverflowLayout.HoverPopup => 80,    // First metric inline, rest on hover
                RvuOverflowLayout.Carousel => 85,      // Single metric slot
                _ => count * 80
            };
        }

        // Pace car: "Now: 85.3 | Best: 78.1 at 3:30 AM | +7.2 ahead" needs ~320px
        if (_config.PaceCarEnabled)
            baseWidth = Math.Max(baseWidth, 320);

        return baseWidth;
    }

    private static int CountSetFlags(RvuMetric m)
    {
        int count = 0;
        if (m.HasFlag(RvuMetric.Total)) count++;
        if (m.HasFlag(RvuMetric.PerHour)) count++;
        if (m.HasFlag(RvuMetric.CurrentHour)) count++;
        if (m.HasFlag(RvuMetric.PriorHour)) count++;
        if (m.HasFlag(RvuMetric.EstimatedTotal)) count++;
        if (m.HasFlag(RvuMetric.RvuPerStudy)) count++;
        if (m.HasFlag(RvuMetric.AvgPerHour)) count++;
        return count;
    }

    /// <summary>
    /// Calculate the connectivity panel width. Fixed 16px for 2x2 dot grid.
    /// </summary>
    private int GetConnectivityPanelWidth()
    {
        // 20px for small dots when enabled (includes right padding)
        return _config.ConnectivityMonitorEnabled ? 20 : 0;
    }

    private void UpdateRvuDisplay()
    {
        if (_config.RvuMetrics == RvuMetric.None && !_config.PaceCarEnabled)
        {
            _rvuPanel.Visible = false;
            return;
        }

        // Prefer pipe shift info from RVUCounter over SQLite fallback
        ShiftInfo? shiftInfo = null;
        ShiftInfoMessage? pipeShift = _controller.PipeService.LatestShiftInfo;
        if (_controller.PipeService.IsConnected && pipeShift != null && pipeShift.IsShiftActive)
        {
            shiftInfo = new ShiftInfo
            {
                TotalRvu = pipeShift.TotalRvu,
                RecordCount = pipeShift.RecordCount,
                ShiftStart = pipeShift.ShiftStart ?? "",
                ShiftId = 0
            };
        }
        else
        {
            shiftInfo = _rvuCounterService.GetCurrentShiftInfo();
        }

        // Hide the 3 legacy labels - we use _rvuMetricLabels now
        _rvuValueLabel.Visible = false;
        _rvuSeparatorLabel.Visible = false;
        _rvuSuffixLabel.Visible = false;

        // Clear old metric labels
        foreach (var lbl in _rvuMetricLabels)
            lbl.Dispose();
        _rvuMetricLabels.Clear();

        // Check if pace car should be shown this cycle
        bool hasPaceData = pipeShift?.PaceDiff != null;
        if (_showingPaceCar && _config.PaceCarEnabled && hasPaceData)
        {
            LayoutPaceCar(pipeShift!);
            _rvuPanel.Visible = true;
            return;
        }

        // Build metric entries
        var allMetrics = BuildMetricEntries(shiftInfo, pipeShift);
        var metrics = _config.RvuMetrics;
        int count = CountSetFlags(metrics);

        if (count == 0 || allMetrics.Count == 0)
        {
            // Show "--" fallback (or hide if only pace car with no data yet)
            if (_config.PaceCarEnabled)
            {
                _rvuPanel.Visible = false;
                return;
            }
            var fallback = CreateMetricLabel("--", Color.FromArgb(100, 100, 100));
            _rvuPanel.Controls.Add(fallback);
            _rvuMetricLabels.Add(fallback);
            int cy = (_rvuPanel.Height - fallback.PreferredHeight) / 2;
            int cx = Math.Max(2, (_rvuPanel.Width - fallback.PreferredWidth) / 2);
            fallback.Location = new Point(cx, cy);
            _rvuPanel.Visible = true;
            return;
        }

        // Hide drawer/popup if not needed for current layout
        bool useDrawer = count >= 3 && _config.RvuOverflowLayout == RvuOverflowLayout.VerticalStack;
        if (!useDrawer) HideRvuDrawer();

        // Layout depends on count and mode
        if (count <= 2 || _config.RvuOverflowLayout == RvuOverflowLayout.Horizontal)
        {
            LayoutHorizontal(allMetrics);
        }
        else
        {
            switch (_config.RvuOverflowLayout)
            {
                case RvuOverflowLayout.VerticalStack:
                    LayoutVerticalStack(allMetrics);
                    break;
                case RvuOverflowLayout.HoverPopup:
                    LayoutHoverPopup(allMetrics);
                    break;
                case RvuOverflowLayout.Carousel:
                    LayoutCarousel(allMetrics);
                    break;
                default:
                    LayoutHorizontal(allMetrics);
                    break;
            }
        }

        // Hide inline panel when vertical stack puts everything in the drawer
        bool isVerticalStack = count >= 3 && _config.RvuOverflowLayout == RvuOverflowLayout.VerticalStack;
        _rvuPanel.Visible = !isVerticalStack;
    }

    /// <summary>
    /// Build the list of (shortText, longLabel, value, color) for each enabled metric.
    /// </summary>
    private List<(string Short, string Label, string Value, Color Color)> BuildMetricEntries(
        ShiftInfo? shiftInfo, ShiftInfoMessage? pipeShift)
    {
        var carolinaBlue = Color.FromArgb(75, 156, 211);
        var lightRed = Color.FromArgb(255, 120, 120);
        var dimGray = Color.FromArgb(120, 120, 120);
        var metrics = _config.RvuMetrics;
        var result = new List<(string Short, string Label, string Value, Color Color)>();

        // Calculate RVU/hour from shift start time
        double rvuPerHour = 0;
        if (shiftInfo != null && DateTime.TryParse(shiftInfo.ShiftStart, out var shiftStart))
        {
            var hoursElapsed = (DateTime.Now - shiftStart).TotalHours;
            if (hoursElapsed > 0.01)
                rvuPerHour = shiftInfo.TotalRvu / hoursElapsed;
        }

        bool belowGoal = _config.RvuGoalEnabled && rvuPerHour < _config.RvuGoalPerHour;
        var rateColor = belowGoal ? lightRed : carolinaBlue;

        if (metrics.HasFlag(RvuMetric.Total))
        {
            var val = shiftInfo != null ? $"{shiftInfo.TotalRvu:F1}" : "--";
            result.Add(($"RVU: {val}", "RVU:", val, carolinaBlue));
        }

        if (metrics.HasFlag(RvuMetric.PerHour))
        {
            var val = shiftInfo != null ? $"{rvuPerHour:F1}" : "--";
            result.Add(($"{val}/h", "Rate:", $"{val}/h", rateColor));
        }

        if (metrics.HasFlag(RvuMetric.PriorHour))
        {
            var raw = pipeShift?.PriorHourRvu;
            var val = raw.HasValue ? $"{raw.Value:F1}" : "--";
            result.Add(($"{val} prev", "Prev hr:", val, carolinaBlue));
        }

        if (metrics.HasFlag(RvuMetric.CurrentHour))
        {
            var raw = pipeShift?.CurrentHourRvu;
            var val = raw.HasValue ? $"~{raw.Value:F1}" : "--";
            result.Add(($"{val} this", "This hr:", val, carolinaBlue));
        }

        if (metrics.HasFlag(RvuMetric.EstimatedTotal))
        {
            var raw = pipeShift?.EstimatedTotalRvu;
            var val = raw.HasValue ? $"~{raw.Value:F0}" : "--";
            result.Add(($"{val} total", "Total:", val, carolinaBlue));
        }

        if (metrics.HasFlag(RvuMetric.RvuPerStudy))
        {
            var raw = pipeShift?.RvuPerStudy;
            var val = raw.HasValue ? $"{raw.Value:F2}/st" : "--";
            result.Add((val, "RVU/st:", val, carolinaBlue));
        }

        if (metrics.HasFlag(RvuMetric.AvgPerHour))
        {
            var raw = pipeShift?.AvgPerHour;
            var val = raw.HasValue ? $"{raw.Value:F1}/h avg" : "--";
            result.Add((val, "Avg/h:", val, carolinaBlue));
        }

        return result;
    }

    private Label CreateMetricLabel(string text, Color foreColor, FontStyle style = FontStyle.Bold)
    {
        return new Label
        {
            Text = text,
            Font = style == FontStyle.Bold ? _metricFontBold : _metricFontRegular,
            ForeColor = foreColor,
            BackColor = Color.Black,
            AutoSize = true
        };
    }

    /// <summary>
    /// Horizontal layout: all metrics in a single row separated by " | ".
    /// </summary>
    private void LayoutHorizontal(List<(string Short, string Label, string Value, Color Color)> metrics)
    {
        var parts = new List<(string Text, Color Color)>();
        for (int i = 0; i < metrics.Count; i++)
        {
            if (i > 0)
                parts.Add((" | ", Color.FromArgb(80, 80, 80)));
            parts.Add((metrics[i].Short, metrics[i].Color));
        }

        // Create labels and measure
        int totalWidth = 0;
        var labels = new List<Label>();
        foreach (var (text, color) in parts)
        {
            var lbl = CreateMetricLabel(text, color, text == " | " ? FontStyle.Regular : FontStyle.Bold);
            _rvuPanel.Controls.Add(lbl);
            _rvuMetricLabels.Add(lbl);
            labels.Add(lbl);
            totalWidth += lbl.PreferredWidth;
        }

        // Center horizontally
        int startX = Math.Max(2, (_rvuPanel.Width - totalWidth) / 2);
        int centerY = (_rvuPanel.Height - (labels.Count > 0 ? labels[0].PreferredHeight : 16)) / 2;
        int x = startX;
        foreach (var lbl in labels)
        {
            lbl.Location = new Point(x, centerY);
            x += lbl.PreferredWidth;
        }
    }

    /// <summary>
    /// Vertical stack: ALL metrics go into a persistent drawer below the bar center.
    /// </summary>
    private void LayoutVerticalStack(List<(string Short, string Label, string Value, Color Color)> metrics)
    {
        if (metrics.Count == 0)
        {
            HideRvuDrawer();
            return;
        }

        // All metrics go into the drawer — nothing inline
        ShowRvuDrawer(metrics);
    }

    private void ShowRvuDrawer(List<(string Short, string Label, string Value, Color Color)> metrics)
    {
        if (_rvuDrawer == null || _rvuDrawer.IsDisposed)
            _rvuDrawer = new RvuPopupForm { Persistent = true };

        var popupData = metrics.Select(m => (m.Label, m.Value, m.Color)).ToList();
        _rvuDrawer.SetMetrics(popupData);

        // Center the drawer under the main form
        var formScreenPos = this.PointToScreen(new Point(0, this.Height - 1));
        int formCenterX = formScreenPos.X + this.Width / 2;
        int drawerX = formCenterX - _rvuDrawer.Width / 2;
        _rvuDrawer.Location = new Point(drawerX, formScreenPos.Y);

        if (!_rvuDrawer.Visible)
            _rvuDrawer.Show();
    }

    private void HideRvuDrawer()
    {
        if (_rvuDrawer != null && !_rvuDrawer.IsDisposed)
            _rvuDrawer.Hide();
    }

    /// <summary>
    /// Hover popup: show first metric in bar, rest in popup on hover.
    /// </summary>
    private void LayoutHoverPopup(List<(string Short, string Label, string Value, Color Color)> metrics)
    {
        if (metrics.Count == 0) return;

        // Show first metric inline
        var first = metrics[0];
        var lbl = CreateMetricLabel(first.Short, first.Color);
        lbl.MouseEnter += OnRvuPanelMouseEnter;
        lbl.MouseLeave += OnRvuPanelMouseLeave;
        _rvuPanel.Controls.Add(lbl);
        _rvuMetricLabels.Add(lbl);

        int cx = Math.Max(2, (_rvuPanel.Width - lbl.PreferredWidth) / 2);
        int cy = (_rvuPanel.Height - lbl.PreferredHeight) / 2;
        lbl.Location = new Point(cx, cy);

        // Store secondary metrics for popup display
        _rvuPopupMetrics = metrics.Count > 1 ? metrics.GetRange(1, metrics.Count - 1) : null;
    }

    private List<(string Short, string Label, string Value, Color Color)>? _rvuPopupMetrics;

    /// <summary>
    /// Carousel: cycle through metrics one at a time with label prefix.
    /// </summary>
    private void LayoutCarousel(List<(string Short, string Label, string Value, Color Color)> metrics)
    {
        if (metrics.Count == 0) return;

        // Ensure carousel index is valid
        if (_carouselIndex >= metrics.Count) _carouselIndex = 0;

        var current = metrics[_carouselIndex];
        var text = $"{current.Label} {current.Value}";
        var lbl = CreateMetricLabel(text, current.Color, FontStyle.Bold);
        _rvuPanel.Controls.Add(lbl);
        _rvuMetricLabels.Add(lbl);

        int cx = Math.Max(2, (_rvuPanel.Width - lbl.PreferredWidth) / 2);
        int cy = (_rvuPanel.Height - lbl.PreferredHeight) / 2;
        lbl.Location = new Point(cx, cy);

        // Ensure carousel timer is running
        if (_carouselTimer == null)
        {
            _carouselTimer = new System.Windows.Forms.Timer { Interval = 4000 };
            _carouselTimer.Tick += (_, _) =>
            {
                _carouselIndex++;
                UpdateRvuDisplay();
            };
            _carouselTimer.Start();
        }
    }

    /// <summary>
    /// Pace car inline layout: "85.3 now | 78.1 best | +7.2 ahead"
    /// Three segments: current RVU, comparison RVU with short label, diff with ahead/behind.
    /// </summary>
    private void LayoutPaceCar(ShiftInfoMessage pipeShift)
    {
        var diff = pipeShift.PaceDiff ?? 0;
        var diffSign = diff >= 0 ? "+" : "";
        var aheadBehind = diff >= 0 ? "ahead" : "behind";
        var diffColor = diff >= 0
            ? Color.FromArgb(100, 200, 100) // Green ahead
            : Color.FromArgb(255, 120, 120); // Red behind
        var carolinaBlue = Color.FromArgb(75, 156, 211);
        var sepColor = Color.FromArgb(80, 80, 80);

        // Shorten paceComparisonMode to a short label: "best", "week", "prior", etc.
        var modeLabel = ShortenPaceMode(pipeShift.PaceComparisonMode);

        var nowVal = pipeShift.PaceCurrentRvu?.ToString("F1") ?? "--";
        var targetVal = pipeShift.PaceTargetRvu?.ToString("F1") ?? "--";
        var diffVal = $"{diffSign}{Math.Abs(diff):F1}";

        // Build the prior/target segment with time: "Best: 78.1 at 3:30 AM"
        var timeText = pipeShift.PaceTimeText;
        var priorSegment = !string.IsNullOrEmpty(timeText)
            ? $"{modeLabel}: {targetVal} at {timeText}"
            : $"{modeLabel}: {targetVal}";

        // Build segments: "Now: 85.3 | Best: 78.1 at 3:30 AM | +7.2 ahead"
        var parts = new List<(string Text, Color Color, FontStyle Style)>
        {
            ($"Now: {nowVal}", carolinaBlue, FontStyle.Bold),
            (" | ", sepColor, FontStyle.Regular),
            (priorSegment, carolinaBlue, FontStyle.Regular),
            (" | ", sepColor, FontStyle.Regular),
            ($"{diffVal} {aheadBehind}", diffColor, FontStyle.Bold)
        };

        // Create labels and measure total width
        int totalWidth = 0;
        var labels = new List<Label>();
        foreach (var (text, color, style) in parts)
        {
            var lbl = CreateMetricLabel(text, color, style);
            _rvuPanel.Controls.Add(lbl);
            _rvuMetricLabels.Add(lbl);
            labels.Add(lbl);
            totalWidth += lbl.PreferredWidth;
        }

        // Tooltip with full description
        var ttText = $"Now: {nowVal} RVU\n" +
                     $"{pipeShift.PaceDescription ?? modeLabel}: {targetVal} RVU" +
                     (pipeShift.PaceTimeText != null ? $" at {pipeShift.PaceTimeText}" : "") +
                     $"\n{diffVal} {aheadBehind}";

        // Center horizontally
        int startX = Math.Max(2, (_rvuPanel.Width - totalWidth) / 2);
        int centerY = (_rvuPanel.Height - (labels.Count > 0 ? labels[0].PreferredHeight : 16)) / 2;
        int x = startX;
        foreach (var lbl in labels)
        {
            lbl.Location = new Point(x, centerY);
            x += lbl.PreferredWidth;
            // Apply tooltip to each label segment
            var tt = new ToolTip { InitialDelay = 200 };
            tt.SetToolTip(lbl, ttText);
        }
    }

    /// <summary>
    /// Shorten pace comparison mode to a compact label for inline display.
    /// </summary>
    private static string ShortenPaceMode(string? mode)
    {
        if (string.IsNullOrEmpty(mode)) return "Prior";
        return mode.ToLowerInvariant() switch
        {
            "best_week" => "Best",
            "best_this_week" => "Best",
            "prior_week" => "Prior",
            "last_week" => "Last",
            "average" => "Avg",
            "avg" => "Avg",
            _ => mode.Length > 6 ? mode[..6] : mode
        };
    }

    private void OnRvuPanelMouseEnter(object? sender, EventArgs e)
    {
        // Only show popup for HoverPopup mode with 3+ metrics
        if (_config.RvuOverflowLayout != RvuOverflowLayout.HoverPopup) return;
        if (CountSetFlags(_config.RvuMetrics) < 3) return;
        if (_rvuPopupMetrics == null || _rvuPopupMetrics.Count == 0) return;

        if (_rvuPopup == null || _rvuPopup.IsDisposed)
            _rvuPopup = new RvuPopupForm();

        var popupData = _rvuPopupMetrics.Select(m => (m.Label, m.Value, m.Color)).ToList();
        _rvuPopup.SetMetrics(popupData);

        // Position below the RVU panel
        var screenPos = _rvuPanel.PointToScreen(new Point(0, _rvuPanel.Height));
        _rvuPopup.Location = new Point(screenPos.X, screenPos.Y);
        _rvuPopup.Show();
    }

    private void OnRvuPanelMouseLeave(object? sender, EventArgs e)
    {
        if (_rvuPopup == null || _rvuPopup.IsDisposed || !_rvuPopup.Visible) return;

        // Small delay to allow mouse to move to popup
        var timer = new System.Windows.Forms.Timer { Interval = 200 };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            if (_rvuPopup != null && !_rvuPopup.IsDisposed && _rvuPopup.Visible)
            {
                var cursorPos = Cursor.Position;
                var panelRect = _rvuPanel.RectangleToScreen(_rvuPanel.ClientRectangle);
                var popupRect = _rvuPopup.Bounds;
                if (!panelRect.Contains(cursorPos) && !popupRect.Contains(cursorPos))
                {
                    _rvuPopup.Hide();
                }
            }
        };
        timer.Start();
    }


    #region Connectivity Monitor

    // Muted colors for status indicators (non-distracting)
    private static readonly Color ConnectivityColorGood = Color.FromArgb(90, 138, 90);      // Muted green
    private static readonly Color ConnectivityColorSlow = Color.FromArgb(154, 138, 80);     // Muted amber
    private static readonly Color ConnectivityColorDegraded = Color.FromArgb(154, 106, 74); // Muted orange
    private static readonly Color ConnectivityColorOffline = Color.FromArgb(138, 74, 74);   // Muted red
    private static readonly Color ConnectivityColorUnknown = Color.FromArgb(74, 74, 74);    // Dark gray

    private void UpdateConnectivityDots()
    {
        if (InvokeRequired)
        {
            Invoke(UpdateConnectivityDots);
            return;
        }

        // Update each dot's color based on server status
        for (int i = 0; i < _connectivityDots.Length; i++)
        {
            var dot = _connectivityDots[i];
            if (dot == null) continue;

            var serverName = dot.Tag as string;
            if (string.IsNullOrEmpty(serverName)) continue;

            var status = _connectivityService.GetStatus(serverName);
            var serverConfig = _config.ConnectivityServers.FirstOrDefault(s => s.Name == serverName);

            // Hide dot if server is disabled or has no host configured
            if (serverConfig == null || !serverConfig.Enabled || string.IsNullOrWhiteSpace(serverConfig.Host))
            {
                dot.Visible = false;
                continue;
            }

            dot.Visible = true;
            dot.BackColor = GetConnectivityStateColor(status?.State ?? ConnectivityState.Unknown);
        }

        // Update tooltip
        UpdateConnectivityTooltip();

        // Update panel visibility
        var hasVisibleDots = _connectivityDots.Any(d => d?.Visible == true);
        _connectivityPanel.Visible = _config.ConnectivityMonitorEnabled && hasVisibleDots;
    }

    private static Color GetConnectivityStateColor(ConnectivityState state)
    {
        return state switch
        {
            ConnectivityState.Good => ConnectivityColorGood,
            ConnectivityState.Slow => ConnectivityColorSlow,
            ConnectivityState.Degraded => ConnectivityColorDegraded,
            ConnectivityState.Offline => ConnectivityColorOffline,
            _ => ConnectivityColorUnknown
        };
    }

    private void UpdateConnectivityTooltip()
    {
        var tooltipLines = new System.Collections.Generic.List<string>();

        foreach (var dot in _connectivityDots)
        {
            if (dot == null || !dot.Visible) continue;

            var serverName = dot.Tag as string;
            if (string.IsNullOrEmpty(serverName)) continue;

            // Use abbreviated names for compact tooltip
            var shortName = serverName switch
            {
                "Mirth" => "Mi",
                "Mosaic" => "Mo",
                "Clario" => "Cl",
                "InteleViewer" => "IV",
                _ => serverName.Length > 2 ? serverName[..2] : serverName
            };

            var status = _connectivityService.GetStatus(serverName);
            if (status == null)
            {
                tooltipLines.Add($"{shortName}: ...");
            }
            else
            {
                var statusText = status.State switch
                {
                    ConnectivityState.Good => $"{status.CurrentLatencyMs:F0}ms",
                    ConnectivityState.Slow => $"Slow {status.CurrentLatencyMs:F0}ms",
                    ConnectivityState.Degraded => $"Slow {status.CurrentLatencyMs:F0}ms",
                    ConnectivityState.Offline => "Offline",
                    _ => "..."
                };
                tooltipLines.Add($"{shortName}: {statusText}");
            }
        }

        var tooltipText = string.Join("\n", tooltipLines);
        _connectivityTooltip.SetToolTip(_connectivityPanel, tooltipText);
        foreach (var dot in _connectivityDots)
        {
            if (dot != null)
                _connectivityTooltip.SetToolTip(dot, tooltipText);
        }
    }

    private ConnectivityDetailsForm? _connectivityDetailsPopup;

    private void OnConnectivityPanelClick(object? sender, EventArgs e)
    {
        // Toggle popup
        if (_connectivityDetailsPopup != null && !_connectivityDetailsPopup.IsDisposed && _connectivityDetailsPopup.Visible)
        {
            _connectivityDetailsPopup.Close();
            _connectivityDetailsPopup = null;
            return;
        }

        // Create and position popup below the panel
        var popupLocation = PointToScreen(new Point(_connectivityPanel.Left, Height));
        _connectivityDetailsPopup = new ConnectivityDetailsForm(_connectivityService, popupLocation);
        _connectivityDetailsPopup.Show();
    }

    /// <summary>
    /// Refresh RVU panel size and form width/height based on current metrics and layout.
    /// Called from SettingsForm after saving.
    /// </summary>
    public void RefreshRvuLayout()
    {
        // Stop carousel timer if layout changed away from carousel
        if (_carouselTimer != null && _config.RvuOverflowLayout != RvuOverflowLayout.Carousel)
        {
            _carouselTimer.Stop();
            _carouselTimer.Dispose();
            _carouselTimer = null;
            _carouselIndex = 0;
        }

        // Hide drawer if layout changed away from vertical stack
        if (_config.RvuOverflowLayout != RvuOverflowLayout.VerticalStack)
            HideRvuDrawer();

        // Stop/restart pace car timer
        if (_paceCarTimer != null)
        {
            _paceCarTimer.Stop();
            _paceCarTimer.Dispose();
            _paceCarTimer = null;
            _showingPaceCar = false;
        }
        if (_config.PaceCarEnabled && !App.IsHeadless)
        {
            _paceCarTimer = new System.Windows.Forms.Timer
            {
                Interval = _config.PaceCarAlternateSeconds * 1000
            };
            _paceCarTimer.Tick += (_, _) =>
            {
                _showingPaceCar = !_showingPaceCar;
                UpdateRvuDisplay();
            };
            _paceCarTimer.Start();
        }

        int newWidth = GetRvuPanelWidth();
        int oldWidth = _rvuPanel.Width;
        if (newWidth != oldWidth)
        {
            _rvuPanel.Width = newWidth;
            Width += (newWidth - oldWidth);
        }

        // Bar always stays 40px tall
        if (Height != 40)
            Height = 40;

        bool hasContent = _config.RvuMetrics != RvuMetric.None || _config.PaceCarEnabled;
        _rvuPanel.Visible = hasContent;

        // Restart or stop timer
        if (hasContent && !App.IsHeadless)
        {
            if (_rvuTimer == null)
            {
                _rvuTimer = new System.Windows.Forms.Timer { Interval = 5000 };
                _rvuTimer.Tick += (_, _) => UpdateRvuDisplay();
                _rvuTimer.Start();
            }
            UpdateRvuDisplay();
        }
        else if (_rvuTimer != null)
        {
            _rvuTimer.Stop();
            _rvuTimer.Dispose();
            _rvuTimer = null;
        }
    }

    /// <summary>
    /// Refresh connectivity service with current configuration.
    /// Called from SettingsForm after saving.
    /// </summary>
    public void RefreshConnectivityService()
    {
        const int panelWidth = 20;

        if (_config.ConnectivityMonitorEnabled && !App.IsHeadless)
        {
            // Expand form and panel if needed
            if (_connectivityPanel.Width == 0)
            {
                _connectivityPanel.Width = panelWidth;
                Width += panelWidth;
            }
            _connectivityPanel.Visible = true;
            _connectivityService.StatusChanged -= UpdateConnectivityDots;
            _connectivityService.StatusChanged += UpdateConnectivityDots;
            _connectivityService.Restart();
        }
        else
        {
            // Shrink form if panel was visible
            if (_connectivityPanel.Width > 0 && _connectivityPanel.Visible)
            {
                Width -= _connectivityPanel.Width;
                _connectivityPanel.Width = 0;
            }
            _connectivityService.Stop();
            _connectivityPanel.Visible = false;
        }
    }

    // [CustomSTT] Clean up transcription window after settings change
    public void RefreshSttPanel()
    {
        if (!_config.CustomSttEnabled)
        {
            // Close transcription window if open
            if (_transcriptionWindow != null && !_transcriptionWindow.IsDisposed)
            {
                _transcriptionWindow.Close();
                _transcriptionWindow = null;
            }
        }
    }

    #endregion

    private async System.Threading.Tasks.Task CheckForUpdatesAsync()
    {
        try
        {
            var updateAvailable = await _updateService.CheckForUpdateAsync();
            if (updateAvailable)
            {
                ShowStatusToast($"Downloading update v{_updateService.LatestVersion}...", 3000);
                var success = await _updateService.DownloadUpdateAsync();
                if (success)
                {
                    // Auto-restart - user will see "Updated!" toast on next start
                    UpdateService.RestartApp();
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"Update check error: {ex.Message}");
        }
    }

    /// <summary>
    /// Trigger update check manually (from Settings).
    /// </summary>
    public async System.Threading.Tasks.Task CheckForUpdatesManualAsync()
    {
        // If update already downloaded and ready, restart now
        if (_updateService.UpdateReady)
        {
            UpdateService.RestartApp();
            return;
        }

        ShowStatusToast("Checking for updates...", 2000);

        var updateAvailable = await _updateService.CheckForUpdateAsync();
        if (updateAvailable)
        {
            ShowStatusToast($"Downloading update v{_updateService.LatestVersion}...", 3000);
            var success = await _updateService.DownloadUpdateAsync();
            if (success)
            {
                // Auto-restart - user will see "Updated!" toast on next start
                UpdateService.RestartApp();
            }
            else
            {
                ShowStatusToast("Download failed. Try again later.", 5000);
            }
        }
        else
        {
            var currentVersion = UpdateService.GetCurrentVersion();
            ShowStatusToast($"You're up to date (v{currentVersion.Major}.{currentVersion.Minor}.{currentVersion.Build})", 3000);
        }
    }
    
    #region Drag Logic
    
    private void OnDragStart(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            _dragging = true;
            _dragStart = e.Location;
        }
    }
    
    private void OnDragMove(object? sender, MouseEventArgs e)
    {
        if (_dragging)
        {
            Location = new Point(
                Location.X + e.X - _dragStart.X,
                Location.Y + e.Y - _dragStart.Y
            );
        }
    }
    
    private void OnDragEnd(object? sender, MouseEventArgs e)
    {
        _dragging = false;
        SaveWindowPosition();
    }
    
    private void SaveWindowPosition()
    {
        _config.WindowX = Location.X;
        _config.WindowY = Location.Y;
        _config.Save();
    }
    
    #endregion
    
    #region Toast Notifications
    
    public void ShowStatusToast(string message, int durationMs = 5000)
    {
        if (InvokeRequired)
        {
            Invoke(() => ShowStatusToast(message, durationMs));
            return;
        }
        
        Logger.Trace($"Toast: {message}");
        
        var toast = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            ShowInTaskbar = false,
            TopMost = true,
            BackColor = Color.FromArgb(51, 51, 51),
            Opacity = 0.9,
            StartPosition = FormStartPosition.Manual
        };
        
        var label = new Label
        {
            Text = message,
            ForeColor = Color.White,
            BackColor = Color.FromArgb(51, 51, 51),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            AutoSize = true,
            Padding = new Padding(15, 8, 15, 8)
        };
        toast.Controls.Add(label);
        
        // Size toast to content
        toast.ClientSize = new Size(label.PreferredWidth, label.PreferredHeight);
        
        // Add to stack and position
        _activeToasts.Add(toast);
        RepositionToasts();
        
        toast.Show();
        
        // Auto-destroy
        var timer = new System.Windows.Forms.Timer { Interval = durationMs };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            _activeToasts.Remove(toast);
            if (!toast.IsDisposed)
                toast.Close();
            RepositionToasts();
        };
        timer.Start();
    }

    /// <summary>
    /// Show a toast with a "Restart Now" button for updates.
    /// Persists until user clicks a button (no auto-dismiss).
    /// </summary>
    public void ShowUpdateToast(string message)
    {
        if (InvokeRequired)
        {
            Invoke(() => ShowUpdateToast(message));
            return;
        }

        Logger.Trace($"Update Toast: {message}");

        var toast = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            ShowInTaskbar = false,
            TopMost = true,
            BackColor = Color.FromArgb(51, 51, 51),
            Opacity = 0.95,
            StartPosition = FormStartPosition.Manual
        };

        var panel = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            BackColor = Color.FromArgb(51, 51, 51),
            Padding = new Padding(10, 8, 10, 8)
        };

        var label = new Label
        {
            Text = message,
            ForeColor = Color.LightGreen,
            BackColor = Color.FromArgb(51, 51, 51),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            AutoSize = true,
            Margin = new Padding(5, 5, 10, 5)
        };
        panel.Controls.Add(label);

        var restartBtn = new Button
        {
            Text = "Restart Now",
            ForeColor = Color.White,
            BackColor = Color.FromArgb(0, 120, 215),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            AutoSize = true,
            Margin = new Padding(5),
            Cursor = Cursors.Hand
        };
        restartBtn.FlatAppearance.BorderSize = 0;
        restartBtn.Click += (_, _) =>
        {
            _activeToasts.Remove(toast);
            toast.Close();
            UpdateService.RestartApp();
        };
        panel.Controls.Add(restartBtn);

        var dismissBtn = new Button
        {
            Text = "Later",
            ForeColor = Color.Gray,
            BackColor = Color.FromArgb(70, 70, 70),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9),
            AutoSize = true,
            Margin = new Padding(5),
            Cursor = Cursors.Hand
        };
        dismissBtn.FlatAppearance.BorderSize = 0;
        dismissBtn.Click += (_, _) =>
        {
            _activeToasts.Remove(toast);
            toast.Close();
            RepositionToasts();
        };
        panel.Controls.Add(dismissBtn);

        toast.Controls.Add(panel);
        toast.ClientSize = panel.PreferredSize;

        _activeToasts.Add(toast);
        RepositionToasts();

        toast.Show();
        // No auto-dismiss - user must click "Restart Now" or "Later"
    }

    private void RepositionToasts()
    {
        // Stack from bottom-right upwards
        var screen = Screen.FromControl(this);
        var workingArea = screen.WorkingArea;
        var screenW = workingArea.Width;
        var screenH = workingArea.Height;
        int currentY = screenH - 100;
        
        foreach (var toast in _activeToasts.AsEnumerable().Reverse())
        {
            if (toast.IsDisposed) continue;
            
            int w = toast.Width;
            int h = toast.Height;
            int x = screenW - w - 20;
            int y = currentY - h;
            
            toast.Location = new Point(x, y);
            currentY = y - 10;
        }
    }
    
    #endregion
    
    #region Child Windows
    
    public void ToggleFloatingToolbar(bool show)
    {
        if (show)
        {
            if (_toolbarWindow == null || _toolbarWindow.IsDisposed)
            {
                _toolbarWindow = new FloatingToolbarForm(_config, _controller);
                _toolbarWindow.Show();
            }
        }
        else
        {
            if (_toolbarWindow != null && !_toolbarWindow.IsDisposed)
            {
                _toolbarWindow.Close();
            }
            _toolbarWindow = null;
        }
    }
    
    public void RefreshFloatingToolbar(FloatingButtonsConfig newConfig)
    {
        if (_toolbarWindow != null && !_toolbarWindow.IsDisposed)
        {
            _toolbarWindow.Refresh(newConfig);
        }
    }
    
    public void ToggleIndicator(bool show)
    {
        if (show)
        {
            if (_indicatorWindow == null || _indicatorWindow.IsDisposed)
            {
                _indicatorWindow = new IndicatorForm(_config);
                _indicatorWindow.Show();
            }
        }
        else
        {
            if (_indicatorWindow != null && !_indicatorWindow.IsDisposed)
            {
                _indicatorWindow.Close();
            }
            _indicatorWindow = null;
        }
    }
    
    /// <summary>
    /// Re-creates the indicator window if it was hidden but should be visible
    /// (e.g., user is dictating but scrape timer failed to detect study open).
    /// </summary>
    public void EnsureIndicatorVisible()
    {
        if (_indicatorWindow == null || _indicatorWindow.IsDisposed)
        {
            _indicatorWindow = new IndicatorForm(_config);
            _indicatorWindow.Show();
        }
    }

    public void UpdateIndicatorState(bool isRecording)
    {
        _indicatorWindow?.SetState(isRecording);
    }

    public void ToggleClinicalHistory(bool show)
    {
        if (show)
        {
            if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            {
                _clinicalHistoryWindow = new ClinicalHistoryForm(_config);
                // Wire up stroke note click callback (for stroke-specific click)
                _clinicalHistoryWindow.SetStrokeNoteClickCallback(() =>
                {
                    var accession = _controller.GetCurrentAccession();
                    return _controller.CreateStrokeCriticalNote(accession);
                });
                // Wire up critical note click callback (for Ctrl+Click and context menu)
                _clinicalHistoryWindow.SetCriticalNoteClickCallback(() =>
                    _controller.TriggerAction(Actions.CreateCriticalNote, "CtrlClick"));
                // Wire up session-wide clinical history fix tracking (prevents duplicate fixes on study reopen)
                _clinicalHistoryWindow.SetClinicalHistoryFixCallbacks(
                    _controller.HasClinicalHistoryFixedForAccession,
                    _controller.MarkClinicalHistoryFixedForAccession);
                // Wire up auto-fix completion callback (for Ignore Inpatient Drafted feature)
                _clinicalHistoryWindow.SetAutoFixCompleteCallback(
                    _controller.MarkAutoFixCompleteForCurrentAccession);
                // Wire up addendum check callback (block paste when addendum is open)
                _clinicalHistoryWindow.SetAddendumCheckCallback(
                    _controller.IsAddendumOpen);
                _clinicalHistoryWindow.Show();
            }
        }
        else
        {
            if (_clinicalHistoryWindow != null && !_clinicalHistoryWindow.IsDisposed)
            {
                _clinicalHistoryWindow.Close();
            }
            _clinicalHistoryWindow = null;
        }
    }

    /// <summary>
    /// Returns true if a study is currently open.
    /// </summary>
    public bool IsStudyOpen => _controller.IsStudyOpen;

    /// <summary>
    /// Updates indicator visibility based on current settings and study state.
    /// Call this when HideIndicatorWhenNoStudy setting changes.
    /// </summary>
    public void UpdateIndicatorVisibility()
    {
        if (!_config.IndicatorEnabled)
        {
            ToggleIndicator(false);
            return;
        }

        // If "hide when no study" is enabled and no study is open, hide it
        if (_config.HideIndicatorWhenNoStudy && !IsStudyOpen)
        {
            ToggleIndicator(false);
        }
        else
        {
            ToggleIndicator(true);
        }
    }

    /// <summary>
    /// Updates clinical history visibility based on current settings and study state.
    /// Call this when HideClinicalHistoryWhenNoStudy or AlwaysShowClinicalHistory setting changes.
    /// </summary>
    public void UpdateClinicalHistoryVisibility()
    {
        if (!_config.ShowClinicalHistory)
        {
            ToggleClinicalHistory(false);
            return;
        }

        // In alerts-only mode, don't automatically show the window
        // The ActionController will show it when an alert triggers
        if (!_config.AlwaysShowClinicalHistory)
        {
            // Don't hide if already showing an alert
            if (_clinicalHistoryWindow != null && !_clinicalHistoryWindow.IsDisposed &&
                _clinicalHistoryWindow.Visible && _clinicalHistoryWindow.IsShowingAlert)
            {
                return;
            }
            ToggleClinicalHistory(false);
            return;
        }

        // In always-show mode: If "hide when no study" is enabled and no study is open, hide it
        if (_config.HideClinicalHistoryWhenNoStudy && !IsStudyOpen)
        {
            ToggleClinicalHistory(false);
        }
        else
        {
            ToggleClinicalHistory(true);
        }
    }

    public void UpdateClinicalHistory(string? rawClarioText, string? accession = null, int? patientAge = null, string? patientGender = null)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        // Use the version that returns both pre-cleaned and cleaned for auto-fix detection
        var (preCleaned, cleaned) = ClinicalHistoryForm.ExtractClinicalHistoryWithFixInfo(rawClarioText);
        _clinicalHistoryWindow.SetClinicalHistoryWithAutoFix(preCleaned, cleaned, accession, patientAge, patientGender);
    }

    public void UpdateClinicalHistoryDraftedState(bool isDrafted)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        _clinicalHistoryWindow.SetDraftedState(isDrafted);
    }

    public void UpdateClinicalHistoryTemplateMismatch(bool isMismatch, string? description, string? templateName)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        _clinicalHistoryWindow.SetTemplateMismatchState(isMismatch, description, templateName);
    }

    public void OnClinicalHistoryStudyChanged(bool isNewStudy = true)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        _clinicalHistoryWindow.OnStudyChanged(isNewStudy);
    }

    public void UpdateClinicalHistoryTextColor(string? finalReportText)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        _clinicalHistoryWindow.UpdateTextColorFromFinalReport(finalReportText);
    }

    public void UpdateGenderCheck(string? reportText, string? patientGender)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        var mismatches = ClinicalHistoryForm.CheckGenderMismatch(reportText, patientGender);
        _clinicalHistoryWindow.SetGenderWarning(mismatches.Count > 0, patientGender, mismatches);
    }

    public void SetStrokeState(bool isStroke)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        _clinicalHistoryWindow.SetStrokeState(isStroke);
    }

    public void SetNoteCreatedState(bool created)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        _clinicalHistoryWindow.SetNoteCreated(created);
    }

    public void ShowAlertOnly(AlertType type, string details)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        _clinicalHistoryWindow.ShowAlertOnly(type, details);
    }

    public void ClearAlert()
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        _clinicalHistoryWindow.ClearAlert();
    }

    public void SetAidocAppend(List<FindingVerification>? findings)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        _clinicalHistoryWindow.SetAidocAppend(findings);
    }

    public void ShowImpressionWindow()
    {
        if (_impressionWindow == null || _impressionWindow.IsDisposed)
        {
            _impressionWindow = new ImpressionForm(_config);
        }
        _impressionWindow.SetImpression(null); // Show waiting message
        if (!_impressionWindow.Visible)
        {
            _impressionWindow.Show();
        }
    }

    public void ShowImpressionWindowIfNotVisible()
    {
        // Only create/show if not already visible - avoids flashing for auto-shown impressions
        if (_impressionWindow != null && !_impressionWindow.IsDisposed && _impressionWindow.Visible)
        {
            return; // Already visible, don't reset
        }

        if (_impressionWindow == null || _impressionWindow.IsDisposed)
        {
            _impressionWindow = new ImpressionForm(_config);
        }
        if (!_impressionWindow.Visible)
        {
            _impressionWindow.Show();
        }
    }

    public void HideImpressionWindow()
    {
        if (_impressionWindow != null && !_impressionWindow.IsDisposed)
        {
            _impressionWindow.Close();
        }
        _impressionWindow = null;
    }

    public void UpdateImpression(string? impression)
    {
        if (_impressionWindow == null || _impressionWindow.IsDisposed)
            return;

        _impressionWindow.SetImpression(impression);
    }

    public void EnsureWindowsOnTop()
    {
        if (_settingsOpen) return; // Don't fight with settings dialog for Z-order

        if (InvokeRequired)
        {
            BeginInvoke(EnsureWindowsOnTop);
            return;
        }

        NativeWindows.ForceTopMost(this.Handle);
        _toolbarWindow?.EnsureOnTop();
        _indicatorWindow?.EnsureOnTop();
        _clinicalHistoryWindow?.EnsureOnTop();
        if (_impressionWindow != null && !_impressionWindow.IsDisposed)
            _impressionWindow.EnsureOnTop();
        // [CustomSTT]
        if (_transcriptionWindow != null && !_transcriptionWindow.IsDisposed)
            _transcriptionWindow.EnsureOnTop();
    }

    #endregion

    #region Custom STT  // [CustomSTT]

    private void ToggleTranscriptionWindow()
    {
        if (_transcriptionWindow != null && !_transcriptionWindow.IsDisposed && _transcriptionWindow.Visible)
        {
            _transcriptionWindow.Hide();
        }
        else
        {
            ShowTranscriptionForm();
        }
    }

    public void ShowTranscriptionForm()
    {
        if (InvokeRequired) { BeginInvoke(ShowTranscriptionForm); return; }

        if (_transcriptionWindow == null || _transcriptionWindow.IsDisposed)
        {
            _transcriptionWindow = new TranscriptionForm(_config);
        }
        if (!_transcriptionWindow.Visible)
        {
            _transcriptionWindow.Show();
        }
    }

    public void OnSttTranscriptionReceived(SttResult result)
    {
        if (InvokeRequired) { BeginInvoke(() => OnSttTranscriptionReceived(result)); return; }

        if (_transcriptionWindow == null || _transcriptionWindow.IsDisposed) return;
        _transcriptionWindow.AppendResult(result);
    }

    public void UpdateSttRecordingState(bool recording)
    {
        if (InvokeRequired) { BeginInvoke(() => UpdateSttRecordingState(recording)); return; }
        _transcriptionWindow?.SetRecordingState(recording);
    }

    public void UpdateSttStatus(string status)
    {
        if (InvokeRequired) { BeginInvoke(() => UpdateSttStatus(status)); return; }
        _transcriptionWindow?.UpdateStatus(status);
    }

    public string GetTranscriptionText()
    {
        if (_transcriptionWindow == null || _transcriptionWindow.IsDisposed) return "";
        return _transcriptionWindow.GetTranscriptText();
    }

    public void ClearTranscriptionForm()
    {
        if (InvokeRequired) { BeginInvoke(ClearTranscriptionForm); return; }
        _transcriptionWindow?.ClearTranscript();
    }

    public void HideTranscriptionForm()
    {
        if (InvokeRequired) { BeginInvoke(HideTranscriptionForm); return; }
        _transcriptionWindow?.Hide();
    }

    #endregion

    #region Critical Studies Indicator

    private CriticalStudiesPopup? _criticalStudiesPopup;

    private void UpdateCriticalIndicator()
    {
        if (InvokeRequired)
        {
            Invoke(UpdateCriticalIndicator);
            return;
        }

        var count = _controller.CriticalStudies.Count;
        bool shouldShow = _config.TrackCriticalStudies && count > 0;
        bool wasVisible = _criticalPanel.Visible;
        _criticalPanel.Visible = shouldShow;
        _criticalCountLabel.Text = count.ToString();

        // Resize form to accommodate critical panel
        if (shouldShow && !wasVisible)
            Width += _criticalPanel.Width;
        else if (!shouldShow && wasVisible)
            Width -= _criticalPanel.Width;

        Logger.Trace($"UpdateCriticalIndicator: count={count}, visible={_criticalPanel.Visible}");
    }

    private void OnCriticalPanelClick(object? sender, EventArgs e)
    {
        // Toggle popup
        if (_criticalStudiesPopup != null && !_criticalStudiesPopup.IsDisposed && _criticalStudiesPopup.Visible)
        {
            _criticalStudiesPopup.Close();
            _criticalStudiesPopup = null;
            return;
        }

        // Create and position popup below the indicator (offset +2 to avoid obscuring top entry)
        var popupLocation = PointToScreen(new Point(_criticalPanel.Left, Height + 7));
        _criticalStudiesPopup = new CriticalStudiesPopup(_controller.CriticalStudies, popupLocation, _controller.Automation, entry => _controller.RemoveCriticalStudy(entry));
        _criticalStudiesPopup.Show();
    }

    #endregion

    #region Settings
    
    private void OpenSettings()
    {
        try
        {
            Logger.Trace("OpenSettings: Creating SettingsFormNew");
            _settingsOpen = true;
            using var settingsForm = new SettingsFormNew(_config, _controller, this);
            Logger.Trace("OpenSettings: Showing dialog");
            settingsForm.ShowDialog(this);
            Logger.Trace("OpenSettings: Dialog closed");
        }
        catch (Exception ex)
        {
            Logger.Trace($"OpenSettings: EXCEPTION - {ex.GetType().Name}: {ex.Message}");
            Logger.Trace($"OpenSettings: StackTrace - {ex.StackTrace}");
            MessageBox.Show($"Error opening settings: {ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _settingsOpen = false;
        }
    }
    
    #endregion
    
    #region App Control
    
    private void ReloadApp()
    {
        _controller.Stop();

        // Relaunch - preserve command line arguments (especially -headless)
        var args = Environment.GetCommandLineArgs();
        var argsToPass = args.Length > 1 ? string.Join(" ", args.Skip(1)) : "";

        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = Application.ExecutablePath,
            Arguments = argsToPass,
            UseShellExecute = true
        };

        // Launch new instance on background thread with delay to avoid mutex conflict
        // The delay gives this instance time to fully exit before new one checks mutex
        System.Threading.Tasks.Task.Run(async () =>
        {
            await System.Threading.Tasks.Task.Delay(800);
            System.Diagnostics.Process.Start(startInfo);
        });

        Application.Exit();
    }
    
    private void ShowLogFile()
    {
        var logPath = Logger.LogFilePath;
        if (File.Exists(logPath))
            System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{logPath}\"");
        else
            System.Diagnostics.Process.Start("explorer.exe", Path.GetDirectoryName(logPath)!);
    }

    private void ExitApp()
    {
        _controller.Stop();
        Application.Exit();
    }
    
    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        Logger.Trace($"App shutting down normally (CloseReason: {e.CloseReason})");

        // Detach context menus before disposal to prevent ObjectDisposedException
        // in WinForms message pump (ModalMenuFilter.ProcessActivationChange)
        ContextMenuStrip = null;
        if (_trayIcon != null)
            _trayIcon.ContextMenuStrip = null;

        _controller.Dispose();
        _toolbarWindow?.Close();
        _indicatorWindow?.Close();
        _clinicalHistoryWindow?.Close();
        _impressionWindow?.Close();

        // Cleanup RVU timer and carousel
        if (_rvuTimer != null)
        {
            _rvuTimer.Stop();
            _rvuTimer.Dispose();
            _rvuTimer = null;
        }
        if (_carouselTimer != null)
        {
            _carouselTimer.Stop();
            _carouselTimer.Dispose();
            _carouselTimer = null;
        }
        if (_paceCarTimer != null)
        {
            _paceCarTimer.Stop();
            _paceCarTimer.Dispose();
            _paceCarTimer = null;
        }
        if (_rvuPopup != null && !_rvuPopup.IsDisposed)
        {
            _rvuPopup.Close();
            _rvuPopup = null;
        }
        if (_rvuDrawer != null && !_rvuDrawer.IsDisposed)
        {
            _rvuDrawer.Close();
            _rvuDrawer = null;
        }

        // Cleanup RecoMD service
        _recoMdService.Dispose();

        // Cleanup connectivity service
        _connectivityService.Dispose();

        // Cleanup tray icon
        if (_trayIcon != null)
        {
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _trayIcon = null;
        }

        Logger.Trace("App shutdown complete");
        base.OnFormClosed(e);
    }
    
    public void ShowDebugResults(string raw, string formatted)
    {
        if (InvokeRequired)
        {
            Invoke(() => ShowDebugResults(raw, formatted));
            return;
        }

        var debugForm = new Form
        {
            Text = "Debug Scrape Results",
            Size = new Size(600, 400),
            StartPosition = FormStartPosition.CenterScreen,
            TopMost = true,
            BackColor = Color.FromArgb(30,30,30),
            ForeColor = Color.White
        };
        
        var textBox = new TextBox
        {
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill,
            Text = $"=== RAW ===\r\n{raw}\r\n\r\n=== FORMATTED ===\r\n{formatted}",
            Font = new Font("Consolas", 10),
            BackColor = Color.FromArgb(40,40,40),
            ForeColor = Color.White,
            ReadOnly = true
        };
        
        debugForm.Controls.Add(textBox);
        debugForm.Show(); // Use Show instead of ShowDialog to avoid blocking thread
    }
    
    #endregion
    
    #region Windows Message Handling (AHK Integration)

    protected override void WndProc(ref Message m)
    {
        // Legacy compat: translate old 0x04xx messages to new 0x80xx range (remove eventually)
        // Log custom messages in our range
        if (m.Msg >= 0x8001 && m.Msg <= 0x8011)
        {
            Logger.Trace($"WndProc received message 0x{m.Msg:X4} from HWND {m.WParam}");
        }

        switch (m.Msg)
        {
            case NativeWindows.WM_TRIGGER_SCRAPE:
                Logger.Trace("WndProc: Triggering CriticalFindings");
                BeginInvoke(() => _controller.TriggerAction(Actions.CriticalFindings));
                break;
            case NativeWindows.WM_TRIGGER_BEEP:
                Logger.Trace("WndProc: Triggering SystemBeep");
                BeginInvoke(() => _controller.TriggerAction(Actions.SystemBeep));
                break;
            case NativeWindows.WM_TRIGGER_SHOW_REPORT:
                Logger.Trace("WndProc: Triggering ShowReport");
                BeginInvoke(() => _controller.TriggerAction(Actions.ShowReport));
                break;
            case NativeWindows.WM_TRIGGER_CAPTURE_SERIES:
                Logger.Trace("WndProc: Triggering CaptureSeries");
                BeginInvoke(() => _controller.TriggerAction(Actions.CaptureSeries));
                break;
            case NativeWindows.WM_TRIGGER_GET_PRIOR:
                Logger.Trace("WndProc: Triggering GetPrior");
                BeginInvoke(() => _controller.TriggerAction(Actions.GetPrior));
                break;
            case NativeWindows.WM_TRIGGER_TOGGLE_RECORD:
                Logger.Trace("WndProc: Triggering ToggleRecord");
                BeginInvoke(() => _controller.TriggerAction(Actions.ToggleRecord));
                break;
            case NativeWindows.WM_TRIGGER_PROCESS_REPORT:
                Logger.Trace("WndProc: Triggering ProcessReport");
                BeginInvoke(() => _controller.TriggerAction(Actions.ProcessReport));
                break;
            case NativeWindows.WM_TRIGGER_SIGN_REPORT:
                Logger.Trace("WndProc: Triggering SignReport");
                BeginInvoke(() => _controller.TriggerAction(Actions.SignReport));
                break;
            case NativeWindows.WM_TRIGGER_OPEN_SETTINGS:
                Logger.Trace("WndProc: Triggering OpenSettings");
                BeginInvoke(() => OpenSettings());
                break;
            case NativeWindows.WM_TRIGGER_CREATE_IMPRESSION:
                Logger.Trace("WndProc: Triggering CreateImpression");
                BeginInvoke(() => _controller.TriggerAction(Actions.CreateImpression));
                break;
            case NativeWindows.WM_TRIGGER_DISCARD_STUDY:
                Logger.Trace("WndProc: Triggering DiscardStudy");
                BeginInvoke(() => _controller.TriggerAction(Actions.DiscardStudy));
                break;
            case NativeWindows.WM_TRIGGER_CHECK_UPDATES:
                Logger.Trace("WndProc: Triggering CheckForUpdates");
                BeginInvoke(() => { _ = CheckForUpdatesManualAsync(); });
                break;
            case NativeWindows.WM_TRIGGER_SHOW_PICK_LISTS:
                Logger.Trace("WndProc: Triggering ShowPickLists");
                BeginInvoke(() => _controller.TriggerAction(Actions.ShowPickLists));
                break;
            case NativeWindows.WM_TRIGGER_CREATE_CRITICAL_NOTE:
                Logger.Trace("WndProc: Triggering CreateCriticalNote");
                BeginInvoke(() => _controller.TriggerAction(Actions.CreateCriticalNote));
                break;
            case NativeWindows.WM_TRIGGER_RADAI_IMPRESSION:
                Logger.Trace("WndProc: Triggering RadAI Impression");
                BeginInvoke(() => _controller.TriggerAction(Actions.RadAiImpression));
                break;
            case NativeWindows.WM_TRIGGER_RECOMD:
                Logger.Trace("WndProc: Triggering RecoMD");
                BeginInvoke(() => _controller.TriggerAction(Actions.TriggerRecoMd));
                break;
            case NativeWindows.WM_TRIGGER_PASTE_RECOMD:
                Logger.Trace("WndProc: Triggering Paste RecoMD");
                BeginInvoke(() => _controller.TriggerAction(Actions.PasteRecoMd));
                break;
        }

        base.WndProc(ref m);
    }
    
    #endregion
}
