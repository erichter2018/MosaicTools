using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Report popup window.
/// Matches Python's ReportPopupWindow.
/// </summary>
public class ReportPopupForm : Form
{
    private readonly Configuration _config;
    private readonly RichTextBox _richTextBox;
    
    // Drag state
    private Point _dragStart;
    private bool _dragging;
    
    // Resize prevention
    private bool _isResizing;
    
    public ReportPopupForm(Configuration config, string reportText)
    {
        _config = config;
        
        // Form properties
        FormBorderStyle = FormBorderStyle.None; 
        Text = "";
        ShowInTaskbar = false;
        TopMost = true;
        DoubleBuffered = true;
        BackColor = Color.FromArgb(30, 30, 30);
        StartPosition = FormStartPosition.Manual;
        
        AutoSize = false;
        AutoScroll = false;
        
        int width = _config.ReportPopupWidth < 600 ? 600 : _config.ReportPopupWidth;
        int padding = 20;

        _richTextBox = new RichTextBox
        {
            Width = width - (padding * 2), // Initial width to force wrapping calc
            Location = new Point(padding, padding),
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.White,
            Font = new Font("Consolas", 11),
            Cursor = Cursors.Hand,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.None, 
            DetectUrls = false,
            ShortcutsEnabled = false
        };
        
        // Hook event BEFORE setting text so it fires on assignment
        _richTextBox.ContentsResized += RichTextBox_ContentsResized;
        
        // 1. Set Base Color (Slightly less bright)
        _richTextBox.ForeColor = Color.Gainsboro; // Off-white/Light Gray

        // 2. Set Text
        _richTextBox.Text = reportText.Replace("\n", Environment.NewLine);
        
        // 3. Set Text (formatting moved to OnLoad)
        _richTextBox.Text = reportText.Replace("\n", Environment.NewLine);
        
        Controls.Add(_richTextBox);
        
        // Initial Layout (fallback)
        this.ClientSize = new Size(width, 200); // Temporary size

        // Initial Location
        Location = new Point(_config.ReportPopupX, _config.ReportPopupY);
        
        // Config Save
        LocationChanged += (_, _) =>
        {
            _config.ReportPopupX = Location.X;
            _config.ReportPopupY = Location.Y;
        };
        
        FormClosed += (_, _) => _config.Save();
        
        // Actions
        SetupInteractions(this);
        SetupInteractions(_richTextBox);
        
        this.Load += (s,e) => 
        {
            this.Activate();
            ActiveControl = null; // Hide caret
            
            // Format AFTER load to ensure layout context
            FormatKeywords(_richTextBox, new[] { "IMPRESSION:", "FINDINGS:" });
            
            // Force Resize Calculation manually
            PerformResize();
        };
    }
    
    private void FormatKeywords(RichTextBox rtb, string[] keywords)
    {
        // Define Highlight Style
        // +2 points largest
        float baseSize = rtb.Font.Size;
        Font highlightFont = new Font(rtb.Font.FontFamily, baseSize + 2, FontStyle.Bold);
        Color highlightColor = Color.White;

        foreach (var word in keywords)
        {
            int startIndex = 0;
            while (startIndex < rtb.TextLength)
            {
                // Find next instance (Case Sensitive + Whole Word)
                int foundIndex = rtb.Find(word, startIndex, RichTextBoxFinds.WholeWord | RichTextBoxFinds.MatchCase);
                if (foundIndex == -1) break;

                // Select and Format
                rtb.Select(foundIndex, word.Length);
                rtb.SelectionColor = highlightColor;
                rtb.SelectionFont = highlightFont;

                // Move past this instance
                startIndex = foundIndex + word.Length;
            }
        }
        
        // Reset selection
        rtb.Select(0, 0);
    }
    
    private void RichTextBox_ContentsResized(object? sender, ContentsResizedEventArgs e)
    {
        if (_isResizing) return;
        
        // CRASH FIX: Check if handle is created.
        if (!this.IsHandleCreated) return; // Ignore event if not ready

        // Use BeginInvoke to decouple from the layout cycle
        this.BeginInvoke(new Action(() =>
        {
            if (_isResizing) return;
            _isResizing = true;
            try 
            {
                PerformResize(); 
            }
            finally
            {
                _isResizing = false;
            }
        }));
    }

    private void PerformResize()
    {
        // AUTHORITY: Ask the control for the position of the last character
        int contentHeight = 0;
        int textLength = _richTextBox.TextLength;
        
        if (textLength > 0)
        {
            // Get pixel position of last char
            Point pt = _richTextBox.GetPositionFromCharIndex(textLength - 1);
            // Height = Y pos + Line Height + Padding
            // Use Font.Height for the last line height approx (even if bolded, it's safer)
            // Getting specific line height is harder, but this is usually close enough.
            // We use the Base Font height + a buffer.
            contentHeight = pt.Y + _richTextBox.Font.Height + 10; 
        }
        else
        {
            contentHeight = _richTextBox.Font.Height;
        }

        int padding = 20;
        int requiredInnerHeight = contentHeight; 
        int requiredTotalHeight = requiredInnerHeight + (padding * 2);
        
        // Safety check for empty/small results
        if (requiredTotalHeight < 100) requiredTotalHeight = 100;

        // Get Max Height
        Rectangle workArea = Screen.FromControl(this).WorkingArea;
        int maxHeight = workArea.Height - 50;
        
        int finalFormHeight;
        bool needsScroll;

        if (requiredTotalHeight > maxHeight)
        {
            finalFormHeight = maxHeight;
            needsScroll = true;
        }
        else
        {
            finalFormHeight = requiredTotalHeight;
            needsScroll = false;
        }
        
        // Apply Scrollbars FIRST
        if (needsScroll && _richTextBox.ScrollBars != RichTextBoxScrollBars.Vertical)
        {
            _richTextBox.ScrollBars = RichTextBoxScrollBars.Vertical;
        }
        else if (!needsScroll && _richTextBox.ScrollBars != RichTextBoxScrollBars.None)
        {
            _richTextBox.ScrollBars = RichTextBoxScrollBars.None;
        }
        
        // Apply Form Size
        if (this.ClientSize.Height != finalFormHeight)
        {
            this.ClientSize = new Size(this.ClientSize.Width, finalFormHeight);
        }
        
        // Apply RTB Size
        int rtbHeight = finalFormHeight - (padding * 2);
        if (_richTextBox.Height != rtbHeight)
        {
            _richTextBox.Height = rtbHeight;
        }
    }

    private void SetupInteractions(Control control)
    {
        control.Click += (s, e) => Close();
        
        control.MouseDown += (s, e) =>
        {
            if (e.Button == MouseButtons.Left)
            {
                 // Capture offset relative to Form
                 _dragStart = new Point(e.X, e.Y);
                 
                 // If dragging via RTB, defer to OS Drag Logic
                 if (control != this) {
                     NativeMethods.ReleaseCapture();
                     NativeMethods.SendMessage(this.Handle, NativeMethods.WM_NCLBUTTONDOWN, NativeMethods.HT_CAPTION, 0);
                     _dragging = false; 
                 }
                 else 
                 {
                     _dragging = true;
                 }
            }
            if (e.Button == MouseButtons.Right) Close();
        };

        control.MouseMove += (s, e) =>
        {
            if (_dragging)
            {
                Point currentScreenPos = Cursor.Position;
                Location = new Point(
                    currentScreenPos.X - _dragStart.X,
                    currentScreenPos.Y - _dragStart.Y
                );
            }
        };
        
        control.MouseUp += (s, e) => _dragging = false;
    }

    private static class NativeMethods
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
    }
}
