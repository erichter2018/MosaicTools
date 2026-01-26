using System;
using System.Drawing;
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
    
    // Toast management
    private readonly List<Form> _activeToasts = new();

    // Update service
    private readonly UpdateService _updateService = new();

    public MainForm(Configuration config)
    {
        _config = config;
        _controller = new ActionController(config, this);
        _rvuCounterService = new RvuCounterService(config);

        // Form properties (borderless, topmost, small)
        // Width adjusts based on RVU counter settings
        int baseWidth = 145; // Base for "Mosaic Tools" title + drag handle
        int rvuWidth = GetRvuPanelWidth();

        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        Text = "MosaicToolsMainWindow";  // Hidden title for AHK targeting
        BackColor = Color.FromArgb(51, 51, 51); // #333333
        Size = new Size(baseWidth + rvuWidth, 40);
        StartPosition = FormStartPosition.Manual;
        Location = new Point(_config.WindowX, _config.WindowY);

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

        // RVU panel with two labels (shown when RVU counter is enabled)
        _rvuPanel = new Panel
        {
            BackColor = Color.Black,
            Width = rvuWidth,
            Dock = DockStyle.Right,
            Visible = _config.RvuCounterEnabled
        };

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

        // Critical studies indicator panel (shown when critical notes placed)
        _criticalPanel = new Panel
        {
            BackColor = Color.Black,
            Width = 30,
            Dock = DockStyle.Left,
            Visible = false, // Hidden until critical studies exist
            Cursor = Cursors.Hand
        };
        _criticalPanel.Click += OnCriticalPanelClick;

        // Warning icon
        _criticalIconLabel = new Label
        {
            Text = "!",
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            ForeColor = Color.FromArgb(255, 100, 100), // Red
            BackColor = Color.Black,
            AutoSize = false,
            Size = new Size(20, 38),
            TextAlign = ContentAlignment.MiddleCenter,
            Location = new Point(0, 0)
        };
        _criticalIconLabel.Click += OnCriticalPanelClick;
        _criticalPanel.Controls.Add(_criticalIconLabel);

        // Count badge
        _criticalCountLabel = new Label
        {
            Text = "0",
            Font = new Font("Segoe UI", 7, FontStyle.Bold),
            ForeColor = Color.White,
            BackColor = Color.FromArgb(180, 50, 50), // Dark red background
            AutoSize = false,
            Size = new Size(14, 14),
            TextAlign = ContentAlignment.MiddleCenter,
            Location = new Point(14, 2) // Top-right of the panel
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
        _innerFrame.Controls.Add(_criticalPanel); // index 1, laid out third (takes left after drag handle)
        _innerFrame.Controls.Add(_dragHandle);   // index 2, laid out second (takes left)
        _innerFrame.Controls.Add(_rvuPanel);     // index 3, laid out first (takes right)

        // Context menu (for normal mode right-click)
        var contextMenu = new ContextMenuStrip();
        contextMenu.Items.Add("Settings", null, (_, _) => OpenSettings());
        contextMenu.Items.Add("Reload", null, (_, _) => ReloadApp());
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add("Exit", null, (_, _) => ExitApp());
        ContextMenuStrip = contextMenu;

        // Always create system tray icon
        var trayMenu = new ContextMenuStrip();
        trayMenu.Items.Add("Settings", null, (_, _) => OpenSettings());
        trayMenu.Items.Add(new ToolStripSeparator());
        trayMenu.Items.Add("Exit", null, (_, _) => ExitApp());

        _trayIcon = new NotifyIcon
        {
            Text = App.IsHeadless ? "Mosaic Tools (Headless)" : "Mosaic Tools",
            Icon = Icon ?? SystemIcons.Application,
            ContextMenuStrip = trayMenu,
            Visible = true
        };
        _trayIcon.DoubleClick += (_, _) => OpenSettings();

        // Headless mode: hide the widget bar
        if (App.IsHeadless)
        {
            Opacity = 0;
            ShowInTaskbar = false;
        }
        
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
        if (_config.ShowClinicalHistory && _config.ScrapeMosaicEnabled &&
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
        if (_config.RvuCounterEnabled && !App.IsHeadless)
        {
            UpdateRvuDisplay(); // Initial update
            _rvuTimer = new System.Windows.Forms.Timer { Interval = 5000 }; // Update every 5 seconds
            _rvuTimer.Tick += (_, _) => UpdateRvuDisplay();
            _rvuTimer.Start();
        }
    }

    /// <summary>
    /// Calculate the RVU panel width based on display mode.
    /// Max values to fit: Total="999.9 RVU", PerHour="99.9/h", Both="999 | 99.9/h"
    /// </summary>
    private int GetRvuPanelWidth()
    {
        if (!_config.RvuCounterEnabled)
            return 0;

        return _config.RvuDisplayMode switch
        {
            RvuDisplayMode.Total => 75,    // "999.9 RVU"
            RvuDisplayMode.PerHour => 60,  // "99.9/h"
            RvuDisplayMode.Both => 105,    // "999 | 99.9/h"
            _ => 75
        };
    }

    private void UpdateRvuDisplay()
    {
        Logger.Trace("UpdateRvuDisplay called");
        if (!_config.RvuCounterEnabled)
        {
            _rvuPanel.Visible = false;
            Logger.Trace("UpdateRvuDisplay: RVU counter not enabled");
            return;
        }

        var shiftInfo = _rvuCounterService.GetCurrentShiftInfo();
        Logger.Trace($"UpdateRvuDisplay: Got shift info = {(shiftInfo != null ? $"total={shiftInfo.TotalRvu:F1}" : "null")}");

        if (shiftInfo != null)
        {
            // Calculate RVU/hour
            double rvuPerHour = 0;
            if (DateTime.TryParse(shiftInfo.ShiftStart, out var shiftStart))
            {
                var hoursElapsed = (DateTime.Now - shiftStart).TotalHours;
                if (hoursElapsed > 0.01) // Avoid division by very small numbers
                {
                    rvuPerHour = shiftInfo.TotalRvu / hoursElapsed;
                }
            }

            // Determine color based on goal
            var carolinaBlue = Color.FromArgb(75, 156, 211);
            var lightRed = Color.FromArgb(255, 120, 120);

            // Check if below goal (only when goal is enabled and we have RVU/h to compare)
            bool belowGoal = _config.RvuGoalEnabled && rvuPerHour < _config.RvuGoalPerHour;
            var rateColor = belowGoal ? lightRed : carolinaBlue;

            // Build display based on mode
            switch (_config.RvuDisplayMode)
            {
                case RvuDisplayMode.Total:
                    _rvuValueLabel.Text = $"{shiftInfo.TotalRvu:F1}";
                    _rvuValueLabel.ForeColor = carolinaBlue;
                    _rvuSeparatorLabel.Visible = false;
                    _rvuSuffixLabel.Text = " RVU";
                    _rvuSuffixLabel.ForeColor = Color.FromArgb(120, 120, 120);
                    break;
                case RvuDisplayMode.PerHour:
                    _rvuValueLabel.Text = $"{rvuPerHour:F1}";
                    _rvuValueLabel.ForeColor = rateColor;
                    _rvuSeparatorLabel.Visible = false;
                    _rvuSuffixLabel.Text = "/h";
                    _rvuSuffixLabel.ForeColor = Color.FromArgb(120, 120, 120);
                    break;
                case RvuDisplayMode.Both:
                    _rvuValueLabel.Text = $"{shiftInfo.TotalRvu:F0}";
                    _rvuValueLabel.ForeColor = carolinaBlue;
                    _rvuSeparatorLabel.Visible = true;
                    _rvuSuffixLabel.Text = $"{rvuPerHour:F1}/h";
                    _rvuSuffixLabel.ForeColor = rateColor;
                    break;
                default:
                    _rvuValueLabel.Text = $"{shiftInfo.TotalRvu:F1}";
                    _rvuValueLabel.ForeColor = carolinaBlue;
                    _rvuSeparatorLabel.Visible = false;
                    _rvuSuffixLabel.Text = " RVU";
                    _rvuSuffixLabel.ForeColor = Color.FromArgb(120, 120, 120);
                    break;
            }

            _rvuValueLabel.Visible = true;
            _rvuSuffixLabel.Visible = true;

            // Center labels within the panel
            CenterRvuLabels();
        }
        else
        {
            // Show "--" in grey when no shift
            _rvuValueLabel.Text = "--";
            _rvuValueLabel.ForeColor = Color.FromArgb(100, 100, 100);
            _rvuValueLabel.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            _rvuValueLabel.Visible = true;
            _rvuSeparatorLabel.Visible = false;
            _rvuSuffixLabel.Visible = false;

            // Center the label
            CenterRvuLabels();
        }
        _rvuPanel.Visible = true;
    }

    private void CenterRvuLabels()
    {
        // Calculate total width of visible labels
        int totalWidth = _rvuValueLabel.PreferredWidth;
        if (_rvuSeparatorLabel.Visible)
        {
            totalWidth += _rvuSeparatorLabel.PreferredWidth;
        }
        if (_rvuSuffixLabel.Visible)
        {
            totalWidth += _rvuSuffixLabel.PreferredWidth;
        }

        // Center horizontally in panel
        int startX = Math.Max(2, (_rvuPanel.Width - totalWidth) / 2);
        int centerY = (_rvuPanel.Height - _rvuValueLabel.PreferredHeight) / 2;

        _rvuValueLabel.Location = new Point(startX, centerY);
        int nextX = _rvuValueLabel.Right;

        if (_rvuSeparatorLabel.Visible)
        {
            _rvuSeparatorLabel.Location = new Point(nextX, centerY);
            nextX = _rvuSeparatorLabel.Right;
        }

        if (_rvuSuffixLabel.Visible)
        {
            _rvuSuffixLabel.Location = new Point(nextX, centerY);
        }
    }

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
        var screenW = Screen.PrimaryScreen!.WorkingArea.Width;
        var screenH = Screen.PrimaryScreen.WorkingArea.Height;
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
        if (!_config.ShowClinicalHistory || !_config.ScrapeMosaicEnabled)
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

    public void UpdateClinicalHistory(string? rawClarioText, string? accession = null)
    {
        if (_clinicalHistoryWindow == null || _clinicalHistoryWindow.IsDisposed)
            return;

        // Use the version that returns both pre-cleaned and cleaned for auto-fix detection
        var (preCleaned, cleaned) = ClinicalHistoryForm.ExtractClinicalHistoryWithFixInfo(rawClarioText);
        _clinicalHistoryWindow.SetClinicalHistoryWithAutoFix(preCleaned, cleaned, accession);
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
        if (InvokeRequired)
        {
            Invoke(EnsureWindowsOnTop);
            return;
        }

        NativeWindows.ForceTopMost(this.Handle);
        _toolbarWindow?.EnsureOnTop();
        _indicatorWindow?.EnsureOnTop();
        _clinicalHistoryWindow?.EnsureOnTop();
        _impressionWindow?.EnsureOnTop();
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
        _criticalPanel.Visible = count > 0;
        _criticalCountLabel.Text = count.ToString();

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

        // Create and position popup below the indicator
        var popupLocation = PointToScreen(new Point(_criticalPanel.Left, Height));
        _criticalStudiesPopup = new CriticalStudiesPopup(_controller.CriticalStudies, popupLocation);
        _criticalStudiesPopup.Show();
    }

    #endregion

    #region Settings
    
    private void OpenSettings()
    {
        using var settingsForm = new SettingsForm(_config, _controller, this);
        settingsForm.ShowDialog(this);
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
    
    private void ExitApp()
    {
        _controller.Stop();
        Application.Exit();
    }
    
    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        Logger.Trace($"App shutting down normally (CloseReason: {e.CloseReason})");

        _controller.Dispose();
        _toolbarWindow?.Close();
        _indicatorWindow?.Close();
        _clinicalHistoryWindow?.Close();
        _impressionWindow?.Close();

        // Cleanup RVU timer
        if (_rvuTimer != null)
        {
            _rvuTimer.Stop();
            _rvuTimer.Dispose();
            _rvuTimer = null;
        }

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
        // Log custom messages in our range (0x0400-0x040F)
        if (m.Msg >= 0x0400 && m.Msg <= 0x040F)
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
                BeginInvoke(async () => await CheckForUpdatesManualAsync());
                break;
            case NativeWindows.WM_TRIGGER_SHOW_PICK_LISTS:
                Logger.Trace("WndProc: Triggering ShowPickLists");
                BeginInvoke(() => _controller.TriggerAction(Actions.ShowPickLists));
                break;
            case NativeWindows.WM_TRIGGER_CREATE_CRITICAL_NOTE:
                Logger.Trace("WndProc: Triggering CreateCriticalNote");
                BeginInvoke(() => _controller.TriggerAction(Actions.CreateCriticalNote));
                break;
        }

        base.WndProc(ref m);
    }
    
    #endregion
}
