using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MosaicTools.Services;
using MosaicTools.UI.Settings;

namespace MosaicTools.UI;

/// <summary>
/// Modern settings form with sidebar navigation, search, and smooth scrolling.
/// Windows 11 Settings-inspired design.
/// </summary>
public class SettingsFormNew : Form, IMessageFilter
{
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
    private static extern int SetWindowTheme(IntPtr hWnd, string? pszSubAppName, string? pszSubIdList);

    private const int WM_MOUSEWHEEL = 0x020A;

    private readonly Configuration _config;
    private readonly ActionController _controller;
    private readonly MainForm _mainForm;

    private TextBox _searchBox = null!;
    private Panel _navPanel = null!;
    private Panel _contentPanel = null!;
    private ToolTip _toolTip = null!;

    private readonly List<SettingsSection> _sections = new();
    private List<int> _visibleSectionIndices = new(); // Indices of visible sections for nav
    private int _hoveredNavRow = -1;  // Row in nav (0-based, visible only)
    private int _selectedNavRow = 0;  // Row in nav (0-based, visible only)
    private Label _searchClearBtn = null!;

    // Cached fonts for nav panel
    private readonly Font _navFont = new("Segoe UI", 9.5f);
    private readonly Font _navVersionFont = new("Segoe UI", 8f);

    // Smooth scroll state
    private System.Windows.Forms.Timer? _scrollTimer;
    private int _scrollTargetY;
    private int _scrollStartY;
    private int _scrollStartTime;
    private const int ScrollDurationMs = 200;

    public SettingsFormNew(Configuration config, ActionController controller, MainForm mainForm)
    {
        _config = config;
        _controller = controller;
        _mainForm = mainForm;

        // Enable double buffering to prevent flicker/blank panel issues
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
        DoubleBuffered = true;

        InitializeUI();
        LoadAllSettings();
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        try
        {
            int value = 1;
            DwmSetWindowAttribute(Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
        }
        catch { }
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        Application.AddMessageFilter(this);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        Application.RemoveMessageFilter(this);
        base.OnFormClosed(e);
    }

    /// <summary>
    /// Intercept mouse wheel messages before they reach child controls.
    /// Only allow controls to handle wheel if they have focus.
    /// </summary>
    public bool PreFilterMessage(ref Message m)
    {
        // Only process when form is active and visible
        if (!Visible || !ContainsFocus) return false;

        if (m.Msg == WM_MOUSEWHEEL)
        {
            // Get the control under the mouse
            var mousePos = Cursor.Position;
            var control = FindControlAtPoint(this, mousePos);

            // Check if control or any parent is a wheel-responding control
            var wheelControl = FindWheelRespondingParent(control);

            // If it's a wheel-responding control without focus, intercept and scroll panel instead
            if (wheelControl != null && !wheelControl.ContainsFocus)
            {
                // Extract wheel delta from message (high word of wParam)
                int delta = (short)((m.WParam.ToInt64() >> 16) & 0xFFFF);

                // Scroll the content panel (smaller steps for smoother scrolling)
                int scrollAmount = delta > 0 ? -15 : 15;
                int currentY = -_contentPanel.AutoScrollPosition.Y;
                int newY = Math.Max(0, currentY + scrollAmount);
                _contentPanel.AutoScrollPosition = new Point(0, newY);

                // Update nav
                BeginInvoke(() => UpdateSelectedNavFromScroll());

                return true; // Message handled, don't pass to control
            }
        }
        return false; // Let message through
    }

    /// <summary>
    /// Find if the control or any of its parents is a wheel-responding control.
    /// </summary>
    private Control? FindWheelRespondingParent(Control? control)
    {
        while (control != null)
        {
            if (control is NumericUpDown or TrackBar or ComboBox or RichTextBox)
                return control;
            // Multiline textboxes respond to wheel for scrolling
            if (control is TextBox tb && tb.Multiline)
                return control;
            control = control.Parent;
        }
        return null;
    }

    private Control? FindControlAtPoint(Control parent, Point screenPoint)
    {
        var clientPoint = parent.PointToClient(screenPoint);
        var child = parent.GetChildAtPoint(clientPoint, GetChildAtPointSkip.Invisible);

        if (child == null)
            return parent.ClientRectangle.Contains(clientPoint) ? parent : null;

        return FindControlAtPoint(child, screenPoint);
    }

    private void InitializeUI()
    {
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var versionStr = version != null ? $" v{version.Major}.{version.Minor}.{version.Build}" : "";
        Text = $"Mosaic Tools Settings{versionStr}";
        Size = new Size(700, 650);
        StartPosition = FormStartPosition.Manual;
        Location = new Point(_config.SettingsX, _config.SettingsY);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.FromArgb(30, 30, 30);
        ForeColor = Color.White;

        // Tooltip system
        _toolTip = new ToolTip
        {
            AutoPopDelay = 15000,
            InitialDelay = 300,
            ReshowDelay = 100,
            ShowAlways = true
        };

        // Top bar with search
        var topBar = new Panel
        {
            Dock = DockStyle.Top,
            Height = 50,
            BackColor = Color.FromArgb(35, 35, 38)
        };
        Controls.Add(topBar);

        _searchBox = new TextBox
        {
            Location = new Point(15, 12),
            Size = new Size(200, 26),
            BackColor = Color.FromArgb(50, 50, 53),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Segoe UI", 10),
            Text = "Search settings..." // Set BEFORE wiring event
        };
        _searchBox.ForeColor = Color.Gray;
        // Placeholder text
        _searchBox.GotFocus += (s, e) => { if (_searchBox.Text == "Search settings...") { _searchBox.Text = ""; _searchBox.ForeColor = Color.White; } };
        _searchBox.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(_searchBox.Text)) { _searchBox.Text = "Search settings..."; _searchBox.ForeColor = Color.Gray; UpdateSearchClearVisibility(); } };
        // TextChanged wired later after all controls created
        topBar.Controls.Add(_searchBox);

        // X button to clear search
        _searchClearBtn = new Label
        {
            Text = "✕",
            Location = new Point(220, 14),
            Size = new Size(20, 20),
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 10),
            Cursor = Cursors.Hand,
            Visible = false
        };
        _searchClearBtn.Click += (s, e) =>
        {
            _searchBox.Text = "";
            _searchBox.Focus();
        };
        _searchClearBtn.MouseEnter += (s, e) => _searchClearBtn.ForeColor = Color.White;
        _searchClearBtn.MouseLeave += (s, e) => _searchClearBtn.ForeColor = Color.Gray;
        topBar.Controls.Add(_searchClearBtn);

        // Save/Cancel buttons in top bar
        var cancelBtn = new Button
        {
            Text = "Cancel",
            Size = new Size(75, 28),
            Location = new Point(Width - 100, 11),
            BackColor = Color.FromArgb(80, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9)
        };
        cancelBtn.FlatAppearance.BorderColor = Color.FromArgb(100, 60, 60);
        cancelBtn.Click += (s, e) => Close();
        topBar.Controls.Add(cancelBtn);

        var saveBtn = new Button
        {
            Text = "Save",
            Size = new Size(75, 28),
            Location = new Point(Width - 185, 11),
            BackColor = Color.FromArgb(50, 80, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9)
        };
        saveBtn.FlatAppearance.BorderColor = Color.FromArgb(60, 100, 60);
        saveBtn.Click += (s, e) => SaveAndClose();
        topBar.Controls.Add(saveBtn);

        // Navigation sidebar (double-buffered to prevent flicker)
        _navPanel = new DoubleBufferedPanel
        {
            Location = new Point(0, 50),
            Size = new Size(180, Height - 50),
            BackColor = Color.FromArgb(28, 28, 30)
        };
        _navPanel.Paint += OnNavPanelPaint;
        _navPanel.MouseMove += OnNavPanelMouseMove;
        _navPanel.MouseLeave += (s, e) => { _hoveredNavRow = -1; _navPanel.Invalidate(); };
        _navPanel.MouseClick += OnNavPanelClick;
        Controls.Add(_navPanel);

        // Content panel (scrollable, double-buffered)
        _contentPanel = new DoubleBufferedPanel
        {
            Location = new Point(180, 50),
            Size = new Size(Width - 180, Height - 50),
            AutoScroll = true,
            BackColor = Color.FromArgb(30, 30, 30)
        };
        // Smoother scrolling with smaller steps
        _contentPanel.VerticalScroll.SmallChange = 15;
        // Intercept mouse wheel to prevent controls from changing values
        _contentPanel.MouseWheel += OnContentPanelMouseWheel;
        // Track scrolling to update nav selection
        _contentPanel.Scroll += OnContentPanelScroll;
        Controls.Add(_contentPanel);

        // Apply dark scrollbar theme
        ApplyDarkScrollbars(_contentPanel);

        // Create sections
        CreateSections();

        // Layout sections in content panel
        LayoutSections();

        // Wire up search AFTER all controls are created
        _searchBox.TextChanged += OnSearchTextChanged;

        // Draw version at bottom of nav
        _navPanel.Invalidate();
    }

    private void CreateSections()
    {
        _sections.Clear();

        // Profile
        _sections.Add(new ProfileSection(_toolTip));

        // Desktop (non-headless only)
        if (!App.IsHeadless)
        {
            _sections.Add(new DesktopSection(_toolTip, _mainForm));
        }

        // Keys & Buttons
        _sections.Add(new KeysButtonsSection(_toolTip, _config, _controller, App.IsHeadless));

        // Text & Templates
        _sections.Add(new TextTemplatesSection(_toolTip, _config));

        // Alerts
        _sections.Add(new AlertsSection(_toolTip, _mainForm, _config));

        // Report Display
        _sections.Add(new ReportDisplaySection(_toolTip));

        // Behavior
        _sections.Add(new BehaviorSection(_toolTip, App.IsHeadless));

        // RVU & Metrics
        _sections.Add(new RvuMetricsSection(_toolTip));

        // Experimental
        _sections.Add(new ExperimentalSection(_toolTip, _mainForm, App.IsHeadless));

        // Reference
        _sections.Add(new ReferenceSection(_toolTip));

        // Wire up tooltip toggle from ProfileSection
        var profileSection = _sections.OfType<ProfileSection>().FirstOrDefault();
        if (profileSection != null)
        {
            profileSection.ShowTooltipsCheck.CheckedChanged += (s, e) =>
            {
                _toolTip.Active = profileSection.ShowTooltipsCheck.Checked;
            };
        }
    }

    // Store Y positions of each visible section for reliable navigation
    private List<int> _sectionYPositions = new();

    private void LayoutSections()
    {
        _contentPanel.SuspendLayout();
        _contentPanel.Controls.Clear();

        // Build list of visible section indices and their Y positions
        _visibleSectionIndices.Clear();
        _sectionYPositions.Clear();

        int y = 10;
        int sectionWidth = _contentPanel.Width - 40;

        for (int i = 0; i < _sections.Count; i++)
        {
            var section = _sections[i];
            if (!section.Visible) continue; // Skip hidden sections

            _visibleSectionIndices.Add(i);
            _sectionYPositions.Add(y); // Store the Y position
            section.Location = new Point(10, y);
            section.Width = sectionWidth;
            section.UpdateHeight();
            _contentPanel.Controls.Add(section);
            y += section.Height + 15;
        }

        // Add bottom padding so last section can scroll fully into view
        var bottomSpacer = new Panel
        {
            Location = new Point(10, y),
            Size = new Size(sectionWidth, 80),
            BackColor = Color.Transparent
        };
        _contentPanel.Controls.Add(bottomSpacer);

        // Reset nav selection to first visible
        _selectedNavRow = 0;
        _hoveredNavRow = -1;

        _contentPanel.ResumeLayout();
    }

    private void LoadAllSettings()
    {
        foreach (var section in _sections)
        {
            section.LoadSettings(_config);
        }

        // Set tooltip active state
        _toolTip.Active = _config.ShowTooltips;
    }

    private void SaveAndClose()
    {
        foreach (var section in _sections)
        {
            section.SaveSettings(_config);
        }

        // Save window position
        _config.SettingsX = Location.X;
        _config.SettingsY = Location.Y;

        _config.Save();

        // Refresh services (re-registers hotkeys, etc.)
        _controller.RefreshServices();

        // Apply UI changes
        _mainForm.ToggleFloatingToolbar(_config.FloatingToolbarEnabled);
        _mainForm.UpdateIndicatorVisibility();
        _mainForm.UpdateClinicalHistoryVisibility();
        _mainForm.RefreshFloatingToolbar(_config.FloatingButtons);
        _mainForm.RefreshRvuLayout();
        _mainForm.RefreshConnectivityService();

        Close();
    }

    #region Navigation

    private const int NavItemHeight = 32;
    private const int NavStartY = 15;

    private void OnNavPanelPaint(object? sender, PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(_navPanel.BackColor);

        int y = NavStartY;

        // Draw each visible section
        for (int row = 0; row < _visibleSectionIndices.Count; row++)
        {
            int sectionIdx = _visibleSectionIndices[row];
            var section = _sections[sectionIdx];
            var rect = new Rectangle(0, y, _navPanel.Width, NavItemHeight);

            // Background for selected/hovered
            if (row == _selectedNavRow)
            {
                using var selBrush = new SolidBrush(Color.FromArgb(50, 50, 55));
                g.FillRectangle(selBrush, rect);
                // Selection indicator bar
                using var barBrush = new SolidBrush(Color.FromArgb(0, 120, 215));
                g.FillRectangle(barBrush, 0, y, 3, NavItemHeight);
            }
            else if (row == _hoveredNavRow)
            {
                using var hoverBrush = new SolidBrush(Color.FromArgb(40, 40, 45));
                g.FillRectangle(hoverBrush, rect);
            }

            // Text
            g.DrawString(section.Title, _navFont, Brushes.White, 15, y + 7);

            y += NavItemHeight;
        }

        // Version at bottom
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        if (version != null)
        {
            var versionStr = $"v{version.Major}.{version.Minor}.{version.Build}";
            g.DrawString(versionStr, _navVersionFont, Brushes.Gray, 15, _navPanel.Height - 30);
        }
    }

    /// <summary>
    /// Get the nav row at a given y position.
    /// </summary>
    private int GetNavRowAtY(int mouseY)
    {
        if (mouseY < NavStartY) return -1;
        int row = (mouseY - NavStartY) / NavItemHeight;
        if (row >= 0 && row < _visibleSectionIndices.Count)
            return row;
        return -1;
    }

    private void OnNavPanelMouseMove(object? sender, MouseEventArgs e)
    {
        int row = GetNavRowAtY(e.Y);

        if (row != _hoveredNavRow)
        {
            _hoveredNavRow = row;
            _navPanel.Invalidate();
        }
    }

    private void OnNavPanelClick(object? sender, MouseEventArgs e)
    {
        int row = GetNavRowAtY(e.Y);

        if (row >= 0 && row < _sectionYPositions.Count)
        {
            _selectedNavRow = row;
            _navPanel.Invalidate();

            // Scroll to stored Y position
            _contentPanel.AutoScrollPosition = new Point(0, _sectionYPositions[row]);
        }
    }

    #endregion

    #region Search

    private void UpdateSearchClearVisibility()
    {
        var hasRealText = !string.IsNullOrWhiteSpace(_searchBox.Text) && _searchBox.Text != "Search settings...";
        _searchClearBtn.Visible = hasRealText;
    }

    private void OnSearchTextChanged(object? sender, EventArgs e)
    {
        var query = _searchBox.Text;
        if (query == "Search settings...") query = "";

        UpdateSearchClearVisibility();

        foreach (var section in _sections)
        {
            section.Visible = section.MatchesSearch(query);
        }

        // Re-layout visible sections
        LayoutSections();
        _navPanel.Invalidate();

        // If search cleared, scroll to top
        if (string.IsNullOrWhiteSpace(query))
        {
            _contentPanel.AutoScrollPosition = new Point(0, 0);
            _selectedNavRow = 0;
            _navPanel.Invalidate();
        }
    }

    #endregion

    #region Smooth Scrolling

    private void ScrollToSection(SettingsSection section)
    {
        int targetY = section.Top;
        int currentY = -_contentPanel.AutoScrollPosition.Y;

        // If already close, just snap
        if (Math.Abs(targetY - currentY) < 20)
        {
            _contentPanel.AutoScrollPosition = new Point(0, targetY);
            return;
        }

        // Start smooth scroll animation
        _scrollStartY = currentY;
        _scrollTargetY = targetY;
        _scrollStartTime = Environment.TickCount;

        if (_scrollTimer == null)
        {
            _scrollTimer = new System.Windows.Forms.Timer { Interval = 16 }; // ~60fps
            _scrollTimer.Tick += OnScrollTimerTick;
        }
        _scrollTimer.Start();
    }

    private void OnScrollTimerTick(object? sender, EventArgs e)
    {
        int elapsed = Environment.TickCount - _scrollStartTime;
        float progress = Math.Min(1f, elapsed / (float)ScrollDurationMs);

        // Ease-out quad
        float eased = 1f - (1f - progress) * (1f - progress);

        int newY = (int)(_scrollStartY + (_scrollTargetY - _scrollStartY) * eased);
        _contentPanel.AutoScrollPosition = new Point(0, newY);

        UpdateSelectedNavFromScroll();

        if (progress >= 1f)
        {
            _scrollTimer?.Stop();
        }
    }

    /// <summary>
    /// Intercept mouse wheel on content panel to prevent child controls
    /// (NumericUpDown, TrackBar) from changing values when scrolling.
    /// Only let controls respond to wheel if they have focus.
    /// </summary>
    private void OnContentPanelMouseWheel(object? sender, MouseEventArgs e)
    {
        // Check if mouse is over a control that would intercept wheel
        var pos = _contentPanel.PointToClient(Cursor.Position);
        var childAtPoint = _contentPanel.GetChildAtPoint(pos, GetChildAtPointSkip.Invisible);

        // Recursively find deepest control
        while (childAtPoint != null)
        {
            var localPos = childAtPoint.PointToClient(Cursor.Position);
            var deeper = childAtPoint.GetChildAtPoint(localPos, GetChildAtPointSkip.Invisible);
            if (deeper == null) break;
            childAtPoint = deeper;
        }

        // If it's a control that responds to wheel, only let it handle if it has focus
        // Otherwise, scroll the panel instead
        bool isWheelControl = childAtPoint is NumericUpDown || childAtPoint is TrackBar || childAtPoint is ComboBox || childAtPoint is RichTextBox;
        bool controlHasFocus = childAtPoint != null && childAtPoint.Focused;

        if (isWheelControl && !controlHasFocus)
        {
            // Manually scroll the panel (smaller steps for smoother scrolling)
            int scrollAmount = e.Delta > 0 ? -15 : 15;
            int currentY = -_contentPanel.AutoScrollPosition.Y;
            int newY = Math.Max(0, currentY + scrollAmount);
            _contentPanel.AutoScrollPosition = new Point(0, newY);

            // Mark handled to prevent control from receiving the event
            if (e is HandledMouseEventArgs handled)
            {
                handled.Handled = true;
            }
        }

        // Update nav selection after ANY mouse wheel (delayed to let scroll settle)
        BeginInvoke(() => UpdateSelectedNavFromScroll());
    }

    #endregion

    #region Helpers

    private void ApplyDarkScrollbars(Control parent)
    {
        foreach (Control c in parent.Controls)
        {
            try
            {
                if (c.IsHandleCreated)
                    SetWindowTheme(c.Handle, "DarkMode_Explorer", null);
                else
                    c.HandleCreated += (_, _) =>
                    {
                        try { SetWindowTheme(c.Handle, "DarkMode_Explorer", null); }
                        catch { }
                    };
            }
            catch { }

            if (c.Controls.Count > 0)
                ApplyDarkScrollbars(c);
        }
    }

    private void OnContentPanelScroll(object? sender, ScrollEventArgs e)
    {
        UpdateSelectedNavFromScroll();
    }

    private void UpdateSelectedNavFromScroll()
    {
        if (_sectionYPositions.Count == 0) return;

        // Get how far we've scrolled down
        int scrollPos = -_contentPanel.AutoScrollPosition.Y;

        // Find which section we're viewing
        // Change to next section when it's within 120px of viewport top
        int activeRow = 0;
        for (int row = 0; row < _sectionYPositions.Count; row++)
        {
            if (scrollPos >= _sectionYPositions[row] - 120)
            {
                activeRow = row;
            }
        }

        if (_selectedNavRow != activeRow)
        {
            _selectedNavRow = activeRow;
            _navPanel.Invalidate();
        }
    }

    #endregion

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _scrollTimer?.Dispose();
            _toolTip?.Dispose();
            _navFont?.Dispose();
            _navVersionFont?.Dispose();
        }
        base.Dispose(disposing);
    }
}

/// <summary>
/// Panel with double buffering enabled to prevent flicker.
/// </summary>
internal class DoubleBufferedPanel : Panel
{
    public DoubleBufferedPanel()
    {
        DoubleBuffered = true;
        SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
    }
}
