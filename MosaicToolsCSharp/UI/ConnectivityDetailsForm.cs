using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Popup form showing detailed connectivity status for all monitored servers.
/// </summary>
public class ConnectivityDetailsForm : Form
{
    private readonly ConnectivityService _connectivityService;
    private readonly Panel _contentPanel;
    private bool _allowDeactivateClose = false;

    // Muted colors for status indicators
    private static readonly Color ColorGood = Color.FromArgb(90, 138, 90);      // Muted green
    private static readonly Color ColorSlow = Color.FromArgb(154, 138, 80);     // Muted amber
    private static readonly Color ColorDegraded = Color.FromArgb(154, 106, 74); // Muted orange
    private static readonly Color ColorOffline = Color.FromArgb(138, 74, 74);   // Muted red
    private static readonly Color ColorUnknown = Color.FromArgb(74, 74, 74);    // Dark gray

    public ConnectivityDetailsForm(ConnectivityService connectivityService, Point location)
    {
        _connectivityService = connectivityService;

        // Form setup
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(35, 35, 35);
        StartPosition = FormStartPosition.Manual;
        Location = location;

        // Calculate size - ensure all 4 servers fit
        var enabledCount = _connectivityService.GetEnabledStatuses().Count();
        var minHeight = 150;
        var itemHeight = 72;
        var headerHeight = 40;
        var footerHeight = 50;
        var calculatedHeight = headerHeight + Math.Max(1, enabledCount) * itemHeight + footerHeight + 20;
        var height = Math.Max(minHeight, calculatedHeight); // No max cap - must fit all servers
        Size = new Size(340, height);

        // Border panel
        var borderPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(1),
            BackColor = Color.FromArgb(80, 80, 80)
        };
        Controls.Add(borderPanel);

        var innerPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(35, 35, 35)
        };
        borderPanel.Controls.Add(innerPanel);

        // Header
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = headerHeight,
            BackColor = Color.FromArgb(45, 45, 45)
        };

        var headerLabel = new Label
        {
            Text = "Network Status",
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Color.FromArgb(200, 200, 200),
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(12, 0, 0, 0)
        };
        headerPanel.Controls.Add(headerLabel);

        // Close button
        var closeBtn = new Label
        {
            Text = "X",
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = Color.Gray,
            AutoSize = false,
            Size = new Size(28, headerHeight),
            Dock = DockStyle.Right,
            TextAlign = ContentAlignment.MiddleCenter,
            Cursor = Cursors.Hand
        };
        closeBtn.Click += (s, e) => Close();
        closeBtn.MouseEnter += (s, e) => closeBtn.ForeColor = Color.White;
        closeBtn.MouseLeave += (s, e) => closeBtn.ForeColor = Color.Gray;
        headerPanel.Controls.Add(closeBtn);

        // Content panel (scrollable)
        _contentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = Color.FromArgb(40, 40, 40)
        };

        // Footer with "Check Now" button and last check time
        var footerPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = footerHeight,
            BackColor = Color.FromArgb(45, 45, 45)
        };

        var lastCheckLabel = new Label
        {
            Text = GetLastCheckText(),
            Font = new Font("Segoe UI", 8),
            ForeColor = Color.FromArgb(120, 120, 120),
            AutoSize = false,
            Location = new Point(10, 12),
            Size = new Size(150, 16),
            TextAlign = ContentAlignment.MiddleLeft
        };
        footerPanel.Controls.Add(lastCheckLabel);

        var checkNowBtn = new Button
        {
            Text = "Check Now",
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            ForeColor = Color.White,
            BackColor = Color.FromArgb(60, 90, 60),
            FlatStyle = FlatStyle.Flat,
            Size = new Size(85, 26),
            Location = new Point(Size.Width - 105, 7),
            Cursor = Cursors.Hand
        };
        checkNowBtn.FlatAppearance.BorderSize = 0;
        checkNowBtn.Click += async (s, e) =>
        {
            checkNowBtn.Enabled = false;
            checkNowBtn.Text = "Checking...";
            try
            {
                await _connectivityService.CheckNowAsync();
                RefreshContent();
                lastCheckLabel.Text = GetLastCheckText();
            }
            finally
            {
                if (!checkNowBtn.IsDisposed)
                {
                    checkNowBtn.Text = "Check Now";
                    checkNowBtn.Enabled = true;
                }
            }
        };
        footerPanel.Controls.Add(checkNowBtn);

        // Add controls in correct order for WinForms docking:
        // Fill first, then Top/Bottom (laid out in reverse order)
        innerPanel.Controls.Add(_contentPanel);
        innerPanel.Controls.Add(headerPanel);
        innerPanel.Controls.Add(footerPanel);

        // Populate content
        RefreshContent();

        // Handle click outside
        Deactivate += (s, e) =>
        {
            if (_allowDeactivateClose)
                Close();
        };

        // Enable close on deactivate after short delay
        Shown += (s, e) =>
        {
            var timer = new System.Windows.Forms.Timer { Interval = 200 };
            timer.Tick += (ts, te) =>
            {
                timer.Stop();
                timer.Dispose();
                _allowDeactivateClose = true;
            };
            timer.Start();
        };
    }

    private void RefreshContent()
    {
        _contentPanel.Controls.Clear();

        var statuses = _connectivityService.GetEnabledStatuses().ToList();

        if (statuses.Count == 0)
        {
            var emptyLabel = new Label
            {
                Text = "No servers configured.\nGo to Settings > Behavior to add servers.",
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                ForeColor = Color.FromArgb(120, 120, 120),
                AutoSize = false,
                Size = new Size(280, 50),
                Location = new Point(15, 15),
                TextAlign = ContentAlignment.TopLeft
            };
            _contentPanel.Controls.Add(emptyLabel);
            return;
        }

        int y = 5;
        int itemHeight = 70;

        foreach (var status in statuses)
        {
            var itemPanel = new Panel
            {
                Location = new Point(5, y),
                Size = new Size(_contentPanel.Width - 25, itemHeight - 5),
                BackColor = Color.FromArgb(40, 40, 40)
            };

            // Status dot
            var dotPanel = new Panel
            {
                Location = new Point(8, 12),
                Size = new Size(12, 12),
                BackColor = GetStateColor(status.State)
            };
            // Make it round
            using var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, 12, 12);
            dotPanel.Region = new Region(path);
            itemPanel.Controls.Add(dotPanel);

            // Server name
            var nameLabel = new Label
            {
                Text = status.Name,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(28, 8)
            };
            itemPanel.Controls.Add(nameLabel);

            // Current status (right side of first line)
            var statusText = GetStatusText(status);
            var statusLabel = new Label
            {
                Text = statusText,
                Font = new Font("Segoe UI", 9),
                ForeColor = GetStateColor(status.State),
                AutoSize = true
            };
            statusLabel.Location = new Point(itemPanel.Width - statusLabel.PreferredWidth - 10, 10);
            itemPanel.Controls.Add(statusLabel);

            // Second line: min/max/avg
            if (status.SuccessCount > 0)
            {
                var statsLabel = new Label
                {
                    Text = $"min {status.MinLatencyMs:F0}ms / max {status.MaxLatencyMs:F0}ms / avg {status.AvgLatencyMs:F0}ms",
                    Font = new Font("Segoe UI", 8),
                    ForeColor = Color.FromArgb(140, 140, 140),
                    AutoSize = true,
                    Location = new Point(28, 30)
                };
                itemPanel.Controls.Add(statsLabel);

                // Third line: packet loss
                var lossText = status.PacketLossPercent > 0
                    ? $"{status.PacketLossPercent:F0}% packet loss"
                    : "0% packet loss";
                var lossLabel = new Label
                {
                    Text = lossText,
                    Font = new Font("Segoe UI", 8),
                    ForeColor = status.PacketLossPercent > 0
                        ? Color.FromArgb(200, 120, 120)
                        : Color.FromArgb(120, 120, 120),
                    AutoSize = true,
                    Location = new Point(28, 46)
                };
                itemPanel.Controls.Add(lossLabel);
            }
            else if (status.State == ConnectivityState.Offline)
            {
                // Show error message
                var errorLabel = new Label
                {
                    Text = status.ErrorMessage ?? "Connection failed",
                    Font = new Font("Segoe UI", 8),
                    ForeColor = Color.FromArgb(180, 100, 100),
                    AutoSize = true,
                    Location = new Point(28, 30)
                };
                itemPanel.Controls.Add(errorLabel);

                // Last success
                if (status.LastSuccess.HasValue)
                {
                    var ago = DateTime.Now - status.LastSuccess.Value;
                    var agoText = ago.TotalMinutes < 1 ? "just now" :
                                  ago.TotalMinutes < 60 ? $"{(int)ago.TotalMinutes} min ago" :
                                  ago.TotalHours < 24 ? $"{(int)ago.TotalHours} hr ago" :
                                  $"{(int)ago.TotalDays} day(s) ago";
                    var lastOkLabel = new Label
                    {
                        Text = $"Last OK: {agoText}",
                        Font = new Font("Segoe UI", 8),
                        ForeColor = Color.FromArgb(120, 120, 120),
                        AutoSize = true,
                        Location = new Point(28, 46)
                    };
                    itemPanel.Controls.Add(lastOkLabel);
                }
            }
            else
            {
                var checkingLabel = new Label
                {
                    Text = "Checking...",
                    Font = new Font("Segoe UI", 8, FontStyle.Italic),
                    ForeColor = Color.FromArgb(120, 120, 120),
                    AutoSize = true,
                    Location = new Point(28, 30)
                };
                itemPanel.Controls.Add(checkingLabel);
            }

            // Separator line
            var separator = new Panel
            {
                Location = new Point(0, itemHeight - 6),
                Size = new Size(itemPanel.Width, 1),
                BackColor = Color.FromArgb(60, 60, 60)
            };
            itemPanel.Controls.Add(separator);

            _contentPanel.Controls.Add(itemPanel);
            y += itemHeight;
        }
    }

    private string GetStatusText(ServerStatus status)
    {
        return status.State switch
        {
            ConnectivityState.Good => $"{status.CurrentLatencyMs:F0}ms",
            ConnectivityState.Slow => $"Slow ({status.CurrentLatencyMs:F0}ms)",
            ConnectivityState.Degraded => $"Very slow ({status.CurrentLatencyMs:F0}ms)",
            ConnectivityState.Offline => "Offline",
            _ => "Checking..."
        };
    }

    private string GetLastCheckText()
    {
        var statuses = _connectivityService.GetEnabledStatuses().ToList();
        if (statuses.Count == 0)
            return "";

        var lastCheck = statuses.Max(s => s.LastCheck);
        if (lastCheck == DateTime.MinValue)
            return "Not checked yet";

        var ago = DateTime.Now - lastCheck;
        if (ago.TotalSeconds < 5)
            return "Last check: just now";
        if (ago.TotalSeconds < 60)
            return $"Last check: {(int)ago.TotalSeconds}s ago";
        return $"Last check: {(int)ago.TotalMinutes}m ago";
    }

    private static Color GetStateColor(ConnectivityState state)
    {
        return state switch
        {
            ConnectivityState.Good => ColorGood,
            ConnectivityState.Slow => ColorSlow,
            ConnectivityState.Degraded => ColorDegraded,
            ConnectivityState.Offline => ColorOffline,
            _ => ColorUnknown
        };
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape)
        {
            Close();
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }
}
