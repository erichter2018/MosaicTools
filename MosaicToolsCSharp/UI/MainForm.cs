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
    
    // Child Windows
    private FloatingToolbarForm? _toolbarWindow;
    private IndicatorForm? _indicatorWindow;
    
    // System Tray (headless mode)
    private NotifyIcon? _trayIcon;
    
    // Drag state
    private Point _dragStart;
    private bool _dragging;
    
    // Toast management
    private readonly List<Form> _activeToasts = new();
    
    public MainForm(Configuration config)
    {
        _config = config;
        _controller = new ActionController(config, this);
        
        // Form properties (borderless, topmost, small)
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(51, 51, 51); // #333333
        Size = new Size(160, 40);
        StartPosition = FormStartPosition.Manual;
        Location = new Point(_config.WindowX, _config.WindowY);
        
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
        _innerFrame.Controls.Add(_dragHandle);
        
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
        _innerFrame.Controls.Add(_titleLabel);
        
        // Context menu (for normal mode right-click)
        var contextMenu = new ContextMenuStrip();
        contextMenu.Items.Add("Reload", null, (_, _) => ReloadApp());
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add("Exit", null, (_, _) => ExitApp());
        ContextMenuStrip = contextMenu;
        
        // Headless mode: hide the widget bar, show tray icon
        if (App.IsHeadless)
        {
            Opacity = 0;
            ShowInTaskbar = false;
            
            // Create system tray icon
            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Settings", null, (_, _) => OpenSettings());
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("Exit", null, (_, _) => ExitApp());
            
            _trayIcon = new NotifyIcon
            {
                Text = "Mosaic Tools (Headless)",
                Icon = SystemIcons.Application,
                ContextMenuStrip = trayMenu,
                Visible = true
            };
            _trayIcon.DoubleClick += (_, _) => OpenSettings();
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
        if (!App.IsHeadless && _config.IndicatorEnabled)
        {
            ToggleIndicator(true);
        }
        
        // Start controller (HID, hotkeys, etc.)
        _controller.Start();
        
        // Startup toast
        string modeStr = App.IsHeadless ? " [Headless]" : "";
        ShowStatusToast($"Mosaic Tools Started ({_config.DoctorName}){modeStr}", 2000);
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
        durationMs = 5000; // Force 5 seconds always!
        
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
        Close();
        
        // Relaunch
        System.Diagnostics.Process.Start(Application.ExecutablePath);
        Application.Exit();
    }
    
    private void ExitApp()
    {
        _controller.Stop();
        Application.Exit();
    }
    
    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _controller.Dispose();
        _toolbarWindow?.Close();
        _indicatorWindow?.Close();
        
        // Cleanup tray icon
        if (_trayIcon != null)
        {
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _trayIcon = null;
        }
        
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
        switch (m.Msg)
        {
            case NativeWindows.WM_TRIGGER_SCRAPE:
                BeginInvoke(() => _controller.TriggerAction(Actions.CriticalFindings));
                break;
            case NativeWindows.WM_TRIGGER_DEBUG:
                BeginInvoke(() => _controller.TriggerAction(Actions.DebugScrape));
                break;
            case NativeWindows.WM_TRIGGER_BEEP:
                BeginInvoke(() => _controller.TriggerAction(Actions.SystemBeep));
                break;
            case NativeWindows.WM_TRIGGER_SHOW_REPORT:
                BeginInvoke(() => _controller.TriggerAction(Actions.ShowReport));
                break;
            case NativeWindows.WM_TRIGGER_CAPTURE_SERIES:
                BeginInvoke(() => _controller.TriggerAction(Actions.CaptureSeries));
                break;
            case NativeWindows.WM_TRIGGER_GET_PRIOR:
                BeginInvoke(() => _controller.TriggerAction(Actions.GetPrior));
                break;
            case NativeWindows.WM_TRIGGER_TOGGLE_RECORD:
                BeginInvoke(() => _controller.TriggerAction(Actions.ToggleRecord));
                break;
            case NativeWindows.WM_TRIGGER_PROCESS_REPORT:
                BeginInvoke(() => _controller.TriggerAction(Actions.ProcessReport));
                break;
            case NativeWindows.WM_TRIGGER_SIGN_REPORT:
                BeginInvoke(() => _controller.TriggerAction(Actions.SignReport));
                break;
            case NativeWindows.WM_TRIGGER_OPEN_SETTINGS:
                BeginInvoke(() => OpenSettings());
                break;
        }
        
        base.WndProc(ref m);
    }
    
    #endregion
}
