using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Floating toolbar with customizable InteleViewer control buttons.
/// Matches Python's FloatingButtonsWindow.
/// </summary>
public class FloatingToolbarForm : Form
{
    private readonly Configuration _config;
    private readonly ActionController _controller;
    private readonly Panel _buttonFrame;
    
    // Drag state
    private Point _dragStart;
    private bool _dragging;
    
    public FloatingToolbarForm(Configuration config, ActionController controller)
    {
        _config = config;
        _controller = controller;
        
        // Form properties
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.Black;
        StartPosition = FormStartPosition.Manual;
        Location = new Point(_config.FloatingToolbarX, _config.FloatingToolbarY);
        
        // Calculate size based on config
        Size = CalculateFormSize();
        
        // Main frame
        var mainFrame = new Panel
        {
            BackColor = Color.Black,
            Dock = DockStyle.Fill
        };
        Controls.Add(mainFrame);
        
        // Drag bar
        var dragBar = new Panel
        {
            BackColor = Color.Black,
            Height = 15,
            Dock = DockStyle.Top
        };
        var dragIndicator = new Label
        {
            Text = "⋯",
            Font = new Font("Segoe UI", 10),
            ForeColor = Color.FromArgb(102, 102, 102),
            BackColor = Color.Black,
            Dock = DockStyle.Right,
            AutoSize = true
        };
        dragBar.Controls.Add(dragIndicator);
        
        dragBar.MouseDown += OnDragStart;
        dragBar.MouseMove += OnDragMove;
        dragBar.MouseUp += OnDragEnd;
        dragIndicator.MouseDown += OnDragStart;
        dragIndicator.MouseMove += OnDragMove;
        dragIndicator.MouseUp += OnDragEnd;
        
        // Button frame - add BEFORE dragBar so Fill works correctly
        _buttonFrame = new Panel
        {
            BackColor = Color.Black,
            Dock = DockStyle.Fill
        };
        mainFrame.Controls.Add(_buttonFrame);
        mainFrame.Controls.Add(dragBar);  // Add dragBar LAST for Top dock priority
        
        // Context menu
        var menu = new ContextMenuStrip();
        menu.Items.Add("Close", null, (_, _) => Close());
        ContextMenuStrip = menu;
        
        // Render buttons
        RenderButtons();
    }
    
    // Shared sizing constants
    private const int BtnSize = 50;
    private const int WideBtnHeight = 34;  // Shorter height for wide buttons (20% taller than before)
    private const int BtnPadding = 3;
    private const int DragBarHeight = 15;
    private const int FormPadding = 4;
    
    private static int CalculateContentHeight(System.Collections.Generic.List<FloatingButtonDef> buttons, int columns)
    {
        int col = 0;
        
        int totalHeight = 0;
        
        foreach (var btn in buttons)
        {
            if (btn.Type == "wide")
            {
                if (col > 0)
                {
                    totalHeight += BtnSize + BtnPadding;
                    col = 0;
                }
                totalHeight += WideBtnHeight + BtnPadding;
            }
            else
            {
                col++;
                if (col >= columns)
                {
                    totalHeight += BtnSize + BtnPadding;
                    col = 0;
                }
            }
        }
        if (col > 0) totalHeight += BtnSize + BtnPadding;
        
        return totalHeight;
    }
    
    private Size CalculateFormSize()
    {
        var buttons = _config.FloatingButtons.Buttons;
        int columns = _config.FloatingButtons.Columns;
        int contentHeight = CalculateContentHeight(buttons, columns);
        
        int width = columns * (BtnSize + BtnPadding) + BtnPadding + FormPadding;
        int height = contentHeight + BtnPadding + DragBarHeight + FormPadding;
        
        return new Size(width, height);
    }
    
    private void RenderButtons()
    {
        _buttonFrame.Controls.Clear();
        
        var buttons = _config.FloatingButtons.Buttons;
        int columns = _config.FloatingButtons.Columns;
        
        var offWhite = Color.FromArgb(204, 204, 204);
        
        int yPos = BtnPadding;  // Track Y position instead of row count
        int col = 0;
        
        foreach (var btnCfg in buttons)
        {
            bool isWide = btnCfg.Type == "wide";
            int btnHeight = isWide ? WideBtnHeight : BtnSize;
            
            var borderPanel = new Panel
            {
                BackColor = offWhite,
                Size = isWide 
                    ? new Size(columns * (BtnSize + BtnPadding) - BtnPadding, btnHeight)
                    : new Size(BtnSize, btnHeight)
            };
            
            if (isWide)
            {
                if (col > 0)
                {
                    yPos += BtnSize + BtnPadding;
                    col = 0;
                }
                borderPanel.Location = new Point(BtnPadding, yPos);
                yPos += btnHeight + BtnPadding;
                col = 0;
            }
            else
            {
                borderPanel.Location = new Point(col * (BtnSize + BtnPadding) + BtnPadding, yPos);
                col++;
                if (col >= columns)
                {
                    col = 0;
                    yPos += BtnSize + BtnPadding;
                }
            }
            
            var btn = new Button
            {
                Text = !string.IsNullOrEmpty(btnCfg.Icon) ? btnCfg.Icon : btnCfg.Label,
                Font = isWide 
                    ? new Font("Segoe UI", 9, FontStyle.Bold)
                    : new Font("Segoe UI Symbol", 14, FontStyle.Bold),
                ForeColor = offWhite,
                BackColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Dock = DockStyle.Fill,
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter,
                Padding = new Padding(0)
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(51, 51, 51);
            
            var keystroke = btnCfg.Keystroke;
            btn.Click += (_, _) => SendKeysToInteleViewer(keystroke);
            
            borderPanel.Padding = new Padding(1);
            borderPanel.Controls.Add(btn);
            _buttonFrame.Controls.Add(borderPanel);
        }
    }
    
    private void SendKeysToInteleViewer(string keystroke)
    {
        if (string.IsNullOrEmpty(keystroke)) return;
        
        try
        {
            // Find InteleViewer window
            var hwnd = NativeWindows.FindWindowByTitle(new[] { "inteleviewer" });
            if (hwnd == IntPtr.Zero)
            {
                Logger.Trace("InteleViewer not found");
                return;
            }
            
            // Activate InteleViewer and send keys
            NativeWindows.ActivateWindow(hwnd);
            Thread.Sleep(50);
            
            // Use SendKeys for simpler keystrokes
            var sendKeysFormat = ConvertToSendKeysFormat(keystroke);
            System.Windows.Forms.SendKeys.SendWait(sendKeysFormat);
            
            Logger.Trace($"Sent to InteleViewer: {keystroke} -> {sendKeysFormat}");
        }
        catch (Exception ex)
        {
            Logger.Trace($"Key send error: {ex.Message}");
        }
    }
    
    private static string ConvertToSendKeysFormat(string keystroke)
    {
        // Convert "ctrl+v" format to SendKeys format "^v"
        var parts = keystroke.ToLower().Split('+');
        var result = new System.Text.StringBuilder();
        
        foreach (var part in parts)
        {
            switch (part.Trim())
            {
                case "ctrl": result.Append("^"); break;
                case "alt": result.Append("%"); break;
                case "shift": result.Append("+"); break;
                default:
                    // Handle special keys
                    if (part.StartsWith("f") && int.TryParse(part.Substring(1), out int fNum))
                        result.Append($"{{F{fNum}}}");
                    else if (part == "enter") result.Append("{ENTER}");
                    else if (part == "tab") result.Append("{TAB}");
                    else if (part == "esc") result.Append("{ESC}");
                    else if (part == "space") result.Append(" ");
                    else result.Append(part);
                    break;
            }
        }
        
        return result.ToString();
    }

    public void EnsureOnTop()
    {
        if (InvokeRequired)
        {
            Invoke(EnsureOnTop);
            return;
        }
        NativeWindows.ForceTopMost(this.Handle);
    }
    
    
    public void Refresh(FloatingButtonsConfig newConfig)
    {
        _config.FloatingButtons = newConfig;
        Size = CalculateFormSize();  // Resize form to fit new button layout
        RenderButtons();
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
        _config.FloatingToolbarX = Location.X;
        _config.FloatingToolbarY = Location.Y;
        _config.Save();
    }
    
    #endregion
}
