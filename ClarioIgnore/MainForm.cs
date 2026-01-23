using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ClarioIgnore;

public class MainForm : Form
{
    private NotifyIcon _trayIcon = null!;
    private ContextMenuStrip _contextMenu = null!;
    private System.Windows.Forms.Timer _pollTimer = null!;
    private ClarioService _clarioService = null!;

    private int _skipCount = 0;

    public MainForm()
    {
        InitializeComponent();
        InitializeTrayIcon();
        InitializeTimer();

        _clarioService = new ClarioService();

        // Start minimized to tray
        WindowState = FormWindowState.Minimized;
        ShowInTaskbar = false;
        Visible = false;

        Logger.Log("ClarioIgnore started");
    }

    private void InitializeComponent()
    {
        Text = "ClarioIgnore";
        Size = new Size(300, 100);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
    }

    private void InitializeTrayIcon()
    {
        _contextMenu = new ContextMenuStrip();

        _contextMenu.Items.Add(new ToolStripMenuItem("Configure Rules...", null, ConfigureMenuItem_Click));
        _contextMenu.Items.Add(new ToolStripMenuItem("Pause", null, PauseMenuItem_Click));

        _contextMenu.Items.Add(new ToolStripSeparator());

        _contextMenu.Items.Add(new ToolStripMenuItem("Open Log", null, OpenLogMenuItem_Click));
        _contextMenu.Items.Add(new ToolStripMenuItem("Run Diagnostic...", null, DiagnosticMenuItem_Click));

        _contextMenu.Items.Add(new ToolStripSeparator());

        _contextMenu.Items.Add(new ToolStripMenuItem("Exit", null, ExitMenuItem_Click));

        _trayIcon = new NotifyIcon
        {
            Icon = CreateTrayIcon(),
            Text = "ClarioIgnore - Monitoring",
            ContextMenuStrip = _contextMenu,
            Visible = true
        };

        _trayIcon.DoubleClick += (s, e) => ConfigureMenuItem_Click(s, e);

        UpdateTrayIconState();
    }

    private Icon CreateTrayIcon()
    {
        // Create a simple icon programmatically
        var bitmap = new Bitmap(16, 16);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.Transparent);

            // Draw a curved arrow representing "skip"
            using (var pen = new Pen(Color.DodgerBlue, 2))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                // Arc for curved part
                g.DrawArc(pen, 2, 2, 10, 10, 180, 180);
                // Arrow head
                g.DrawLine(pen, 12, 7, 14, 4);
                g.DrawLine(pen, 12, 7, 14, 10);
            }
        }

        return Icon.FromHandle(bitmap.GetHicon());
    }

    private void UpdateTrayIconState()
    {
        _trayIcon.Text = $"ClarioIgnore (skipped: {_skipCount})";
    }

    private void InitializeTimer()
    {
        _pollTimer = new System.Windows.Forms.Timer
        {
            Interval = 10000 // 10 seconds
        };
        _pollTimer.Tick += PollTimer_Tick;
        _pollTimer.Start();
        Logger.Log("ClarioIgnore started (auto-skip every 10s)");
    }

    private void PollTimer_Tick(object? sender, EventArgs e)
    {
        try
        {
            var matches = _clarioService.GetMatchingItems();

            // Filter to only visible items that need skipping
            var toSkip = new List<(WorklistItem, SkipRule)>();
            foreach (var (item, rule) in matches)
            {
                if (_clarioService.IsRowVisible(item))
                {
                    toSkip.Add((item, rule));
                }
            }

            if (toSkip.Count == 0)
                return;

            bool anyClicked = false;

            // Save active window (harmless - just reads current window handle)
            _clarioService.SaveActiveWindow();

            foreach (var (item, rule) in toSkip)
            {
                if (_clarioService.ClickSkipButton(item))
                {
                    anyClicked = true;
                    _skipCount++;
                    Logger.Log($"Auto-skipped: {item.Procedure}");
                    System.Threading.Thread.Sleep(150);
                }
            }

            // Only restore focus if we actually clicked something
            if (anyClicked)
            {
                _clarioService.RestoreActiveWindow();
            }

            UpdateTrayIconState();
        }
        catch (Exception ex)
        {
            Logger.Log($"Poll error: {ex.Message}");
        }
    }

    private void ShowToast(string message, int durationMs = 2000)
    {
        // Create a small toast form
        var toast = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            TopMost = true,
            ShowInTaskbar = false,
            StartPosition = FormStartPosition.Manual,
            Size = new Size(300, 40),
            Opacity = 0.9
        };

        var label = new Label
        {
            Text = message,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 9)
        };
        toast.Controls.Add(label);

        // Position at bottom-right of screen
        var screen = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
        toast.Location = new Point(screen.Right - toast.Width - 10, screen.Bottom - toast.Height - 10);

        toast.Show();

        // Auto-close after duration
        var closeTimer = new System.Windows.Forms.Timer { Interval = durationMs };
        closeTimer.Tick += (s, e) => { closeTimer.Stop(); toast.Close(); };
        closeTimer.Start();
    }

    private void ConfigureMenuItem_Click(object? sender, EventArgs e)
    {
        using var form = new SettingsForm();
        form.ShowDialog();
    }

    private bool _isPaused = false;

    private void PauseMenuItem_Click(object? sender, EventArgs e)
    {
        _isPaused = !_isPaused;

        if (_isPaused)
        {
            _pollTimer.Stop();
            ShowToast("Paused");
            _trayIcon.Text = "ClarioIgnore (PAUSED)";
        }
        else
        {
            _pollTimer.Start();
            ShowToast("Resumed");
            UpdateTrayIconState();
        }

        // Update menu text
        if (sender is ToolStripMenuItem menuItem)
        {
            menuItem.Text = _isPaused ? "Resume" : "Pause";
        }
    }

    private void OpenLogMenuItem_Click(object? sender, EventArgs e)
    {
        var logPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ClarioIgnore", "clarioignore_log.txt");

        if (System.IO.File.Exists(logPath))
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = logPath,
                UseShellExecute = true
            });
        }
        else
        {
            MessageBox.Show("Log file not found yet.", "ClarioIgnore", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void DiagnosticMenuItem_Click(object? sender, EventArgs e)
    {
        try
        {
            var result = _clarioService.RunDiagnostic();

            // Show in a scrollable dialog
            using var form = new Form
            {
                Text = "ClarioIgnore Diagnostic",
                Size = new Size(700, 500),
                StartPosition = FormStartPosition.CenterScreen
            };

            var textBox = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9),
                Text = result,
                WordWrap = false
            };

            form.Controls.Add(textBox);
            form.ShowDialog();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Diagnostic failed: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ExitMenuItem_Click(object? sender, EventArgs e)
    {
        Logger.Log("ClarioIgnore exiting");
        _trayIcon.Visible = false;
        Application.Exit();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        // Minimize to tray instead of closing
        if (e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            WindowState = FormWindowState.Minimized;
            ShowInTaskbar = false;
            Visible = false;
        }
        else
        {
            _trayIcon?.Dispose();
            _pollTimer?.Dispose();
            _clarioService?.Dispose();
            base.OnFormClosing(e);
        }
    }

    protected override void SetVisibleCore(bool value)
    {
        // Start hidden
        if (!IsHandleCreated)
        {
            CreateHandle();
            value = false;
        }
        base.SetVisibleCore(value);
    }
}
