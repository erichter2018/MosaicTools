using System;
using System.Drawing;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Floating indicator window showing microphone recording state.
/// Matches Python's FloatingIndicatorWindow.
/// </summary>
public class IndicatorForm : Form
{
    private readonly Configuration _config;
    private readonly Label _iconLabel;
    private readonly Panel _frame;
    
    private readonly Color _bgOff = Color.FromArgb(68, 68, 68);  // #444444
    private readonly Color _bgOn = Color.FromArgb(204, 0, 0);    // #CC0000
    
    // Drag state
    private Point _dragStart;
    private bool _dragging;
    
    public IndicatorForm(Configuration config)
    {
        _config = config;
        
        // Form properties
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        Size = new Size(50, 50);
        StartPosition = FormStartPosition.Manual;
        Location = ScreenHelper.EnsureOnScreen(_config.IndicatorX, _config.IndicatorY);
        
        // Frame
        _frame = new Panel
        {
            BackColor = _bgOff,
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.FixedSingle
        };
        Controls.Add(_frame);
        
        // Icon
        _iconLabel = new Label
        {
            Text = "🎙",
            Font = new Font("Segoe UI Symbol", 20),
            ForeColor = Color.White,
            BackColor = _bgOff,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter
        };
        _frame.Controls.Add(_iconLabel);
        
        // Drag events
        _iconLabel.MouseDown += OnDragStart;
        _iconLabel.MouseMove += OnDragMove;
        _iconLabel.MouseUp += OnDragEnd;
        _frame.MouseDown += OnDragStart;
        _frame.MouseMove += OnDragMove;
        _frame.MouseUp += OnDragEnd;
    }
    
    public void SetState(bool isRecording)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetState(isRecording));
            return;
        }
        
        var color = isRecording ? _bgOn : _bgOff;
        _frame.BackColor = color;
        _iconLabel.BackColor = color;
    }

    public void EnsureOnTop()
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            BeginInvoke(EnsureOnTop);
            return;
        }
        if (!IsDisposed && IsHandleCreated)
            NativeWindows.ForceTopMost(this.Handle);
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
        _config.IndicatorX = Location.X;
        _config.IndicatorY = Location.Y;
        _config.Save();
    }
    
    #endregion
}
