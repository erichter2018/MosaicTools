using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Draggable ruler overlays for InteleViewer — captures both vertical and horizontal rulers.
/// Drag to measure, release to snap back.
/// Periodically samples pixels — if yellow count changes, recaptures.
/// Uses WS_EX_LAYERED for per-pixel alpha.
/// </summary>
public class RulerOverlayForm : Form
{
    private readonly OcrService _ocrService;

    private Bitmap? _rulerBitmap;
    private Point _homeLocation;

    // Cached for thread-safe access from timer
    private volatile int _homeX, _homeY, _overlayW, _overlayH;

    // Last known viewport bounds for change detection
    private Rectangle _lastViewport;

    // Drag state
    private Point _dragStart;
    private bool _dragging;

    // Horizontal ruler (child overlay)
    private HorizontalRulerForm? _hRuler;

    [DllImport("user32.dll")]
    private static extern int ShowCursor(bool bShow);

    // Recheck timer
    private System.Threading.Timer? _checkTimer;
    private const int CheckIntervalMs = 1500;
    private int _checkRunning;
    private volatile int _baselineYellowCount;

    private static RulerOverlayForm? _instance;

    /// <summary>
    /// Called to hide/show other overlays (floating toolbar) during ruler capture.
    /// Set by MainForm. Parameter: true = hide, false = show.
    /// </summary>
    public static Action<bool>? HideOtherOverlays { get; set; }

    /// <summary>
    /// Cached toolbar bounds, updated periodically by MainForm/ActionController.
    /// Used to skip hiding the toolbar when it doesn't overlap the capture area.
    /// </summary>
    public static Rectangle? CachedToolbarBounds { get; set; }


    public static void Show(OcrService ocrService)
    {
        CloseIfOpen();
        _instance = new RulerOverlayForm(ocrService);
        _instance.CaptureAndShow();
    }

    public static void CloseIfOpen()
    {
        if (_instance != null && !_instance.IsDisposed)
        {
            _instance.Close();
            _instance = null;
        }
    }

    public static bool IsOpen => _instance != null && !_instance.IsDisposed;

    private RulerOverlayForm(OcrService ocrService)
    {
        _ocrService = ocrService;

        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;
        Size = new Size(40, 100);
        Cursor = Cursors.SizeAll;

        MouseDown += OnDragStart;
        MouseMove += OnDragMove;
        MouseUp += OnDragEnd;
    }

    private void CaptureAndShow()
    {
        var timer = new System.Windows.Forms.Timer { Interval = 500 };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            DoInitialCapture();
        };
        timer.Start();
    }

    private bool ToolbarOverlapsCapture(Rectangle viewport)
    {
        var tb = CachedToolbarBounds;
        if (!tb.HasValue) return false;

        // Vertical strip: left edge, 8% width, full height
        int vStripW = Math.Max(60, (int)(viewport.Width * 0.08));
        var vStrip = new Rectangle(viewport.Left, viewport.Top, vStripW, viewport.Height);

        // Horizontal strip: bottom edge, full width, 8% height
        int hStripH = Math.Max(60, (int)(viewport.Height * 0.08));
        var hStrip = new Rectangle(viewport.Left, viewport.Bottom - hStripH, viewport.Width, hStripH);

        return tb.Value.IntersectsWith(vStrip) || tb.Value.IntersectsWith(hStrip);
    }

    private void DoInitialCapture()
    {
        var viewport = _ocrService.FindYellowTarget();
        if (!viewport.HasValue) return;

        bool needHide = ToolbarOverlapsCapture(viewport.Value);
        if (needHide)
        {
            HideOtherOverlays?.Invoke(true);
            Thread.Sleep(50);
            viewport = _ocrService.FindYellowTarget();
            if (!viewport.HasValue) { HideOtherOverlays?.Invoke(false); return; }
        }

        _lastViewport = viewport.Value;

        // Capture both rulers
        var vResult = _ocrService.CaptureRulerStrip(viewport.Value);
        var hResult = _ocrService.CaptureHorizontalRulerStrip(viewport.Value);
        if (needHide) HideOtherOverlays?.Invoke(false);

        if (vResult != null)
        {
            var (bitmap, location) = vResult.Value;
            try
            {
                var debugPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MosaicTools", "ruler_debug.png");
                bitmap.Save(debugPath, ImageFormat.Png);
            }
            catch { }
            ApplyBitmap(bitmap, location);
        }

        if (hResult != null)
        {
            var (hBitmap, hLocation) = hResult.Value;
            ShowHorizontalRuler(hBitmap, hLocation);
        }

        // Delay baseline capture so floating toolbar has time to render back on screen
        var baselineTimer = new System.Windows.Forms.Timer { Interval = 300 };
        baselineTimer.Tick += (_, _) =>
        {
            baselineTimer.Stop();
            baselineTimer.Dispose();
            _baselineYellowCount = CountYellowInArea(_homeX, _homeY, _overlayW, _overlayH);
            Logger.Trace($"RulerOverlay: baseline yellow count = {_baselineYellowCount}");
            _checkTimer = new System.Threading.Timer(OnCheckTimer, null, CheckIntervalMs, CheckIntervalMs);
        };
        baselineTimer.Start();
    }

    private void ShowHorizontalRuler(Bitmap bitmap, Point location)
    {
        CloseHorizontalRuler();
        _hRuler = new HorizontalRulerForm(bitmap, location);
        _hRuler.Show();
    }

    private void CloseHorizontalRuler()
    {
        if (_hRuler != null && !_hRuler.IsDisposed)
        {
            _hRuler.Close();
            _hRuler.Dispose();
        }
        _hRuler = null;
    }

    private void ApplyBitmap(Bitmap bitmap, Point location)
    {
        LayeredWindowHelper.PremultiplyBitmapAlpha(bitmap);

        var old = _rulerBitmap;
        _rulerBitmap = bitmap;
        old?.Dispose();

        _homeLocation = location;
        _homeX = location.X;
        _homeY = location.Y;
        _overlayW = bitmap.Width;
        _overlayH = bitmap.Height;

        Size = new Size(bitmap.Width, bitmap.Height);
        if (!_dragging) Location = _homeLocation;

        if (!Visible) Show();
        LayeredWindowHelper.SetBitmap(this, _rulerBitmap);
    }

    private void OnCheckTimer(object? state)
    {
        if (Interlocked.CompareExchange(ref _checkRunning, 1, 0) != 0) return;
        if (_dragging || IsDisposed)
        {
            Interlocked.Exchange(ref _checkRunning, 0);
            return;
        }

        // Don't recapture while horizontal ruler is being dragged
        if (_hRuler != null && !_hRuler.IsDisposed && _hRuler.IsDragging)
        {
            Interlocked.Exchange(ref _checkRunning, 0);
            return;
        }

        try
        {
            int w = _overlayW, h = _overlayH;
            int x0 = _homeX, y0 = _homeY;
            if (w == 0 || h == 0) return;

            // Check 1: Has the active viewport changed? (user clicked a different viewport)
            var currentViewport = _ocrService.FindYellowTarget();
            bool viewportChanged = currentViewport.HasValue && currentViewport.Value != _lastViewport;

            if (!viewportChanged)
            {
                // Check 2: Has yellow appeared in the overlay area? (scroll/zoom changed ruler)
                int currentYellow = CountYellowInArea(x0, y0, w, h);
                int baseline = _baselineYellowCount;

                if (Math.Abs(currentYellow - baseline) < 5)
                    return;

                Logger.Trace($"RulerOverlay: yellow changed ({baseline} → {currentYellow}), recapturing");
            }
            else
            {
                Logger.Trace($"RulerOverlay: viewport changed, recapturing");
            }

            // Hide rulers (always needed — blue overlay covers yellow we're recapturing)
            SafeInvoke(() =>
            {
                if (!IsDisposed) Visible = false;
                if (_hRuler != null && !_hRuler.IsDisposed) _hRuler.Visible = false;
            });

            // Check if toolbar overlaps capture area using cached bounds
            var preViewport = _ocrService.FindYellowTarget();
            bool needHideToolbar = preViewport.HasValue && ToolbarOverlapsCapture(preViewport.Value);
            if (needHideToolbar)
                SafeInvoke(() => HideOtherOverlays?.Invoke(true));

            Thread.Sleep(50);

            // Re-find viewport (need clean capture without overlays)
            var viewport = _ocrService.FindYellowTarget();
            if (!viewport.HasValue)
            {
                SafeInvoke(() =>
                {
                    if (!IsDisposed) Visible = true;
                    if (_hRuler != null && !_hRuler.IsDisposed) _hRuler.Visible = true;
                    if (needHideToolbar) HideOtherOverlays?.Invoke(false);
                });
                return;
            }

            _lastViewport = viewport.Value;
            var vResult = _ocrService.CaptureRulerStrip(viewport.Value);
            var hResult = _ocrService.CaptureHorizontalRulerStrip(viewport.Value);

            SafeInvoke(() =>
            {
                if (IsDisposed) return;

                if (vResult != null)
                {
                    var (newBitmap, newLocation) = vResult.Value;
                    ApplyBitmap(newBitmap, newLocation);
                }
                else
                {
                    Visible = true;
                    if (_rulerBitmap != null) LayeredWindowHelper.SetBitmap(this, _rulerBitmap);
                }

                if (hResult != null)
                {
                    var (hBitmap, hLocation) = hResult.Value;
                    ShowHorizontalRuler(hBitmap, hLocation);
                }
                else if (_hRuler != null && !_hRuler.IsDisposed)
                {
                    _hRuler.Visible = true;
                }

                if (needHideToolbar) HideOtherOverlays?.Invoke(false);

                // Delay baseline so toolbar renders
                var bTimer = new System.Windows.Forms.Timer { Interval = 300 };
                bTimer.Tick += (_, _) =>
                {
                    bTimer.Stop();
                    bTimer.Dispose();
                    _baselineYellowCount = CountYellowInArea(_homeX, _homeY, _overlayW, _overlayH);
                    Logger.Trace($"RulerOverlay: new baseline = {_baselineYellowCount}");
                };
                bTimer.Start();
            });
        }
        catch (Exception ex)
        {
            Logger.Trace($"RulerOverlay check error: {ex.Message}");
        }
        finally
        {
            Interlocked.Exchange(ref _checkRunning, 0);
        }
    }

    private void SafeInvoke(Action action)
    {
        try
        {
            if (!IsDisposed && IsHandleCreated)
                Invoke(action);
        }
        catch { }
    }

    protected override void OnPaint(PaintEventArgs e) { }

    private void OnDragStart(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            _dragging = true;
            _dragStart = e.Location;
            ShowCursor(false);
        }
    }

    private void OnDragMove(object? sender, MouseEventArgs e)
    {
        if (!_dragging) return;
        Location = new Point(
            Location.X + e.X - _dragStart.X,
            Location.Y + e.Y - _dragStart.Y
        );
        if (_rulerBitmap != null && IsHandleCreated && !IsDisposed)
            LayeredWindowHelper.SetBitmap(this, _rulerBitmap);
    }

    private void OnDragEnd(object? sender, MouseEventArgs e)
    {
        _dragging = false;
        ShowCursor(true);
        Location = _homeLocation;
        if (_rulerBitmap != null && IsHandleCreated && !IsDisposed)
            LayeredWindowHelper.SetBitmap(this, _rulerBitmap);
    }

    private static int CountYellowInArea(int x0, int y0, int w, int h)
    {
        int count = 0;
        try
        {
            using var bmp = new Bitmap(w, h);
            using (var g = Graphics.FromImage(bmp))
                g.CopyFromScreen(x0, y0, 0, 0, new Size(w, h));

            var data = bmp.LockBits(new Rectangle(0, 0, w, h),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                unsafe
                {
                    byte* scan0 = (byte*)data.Scan0;
                    int stride = data.Stride;
                    for (int y = 0; y < h; y++)
                        for (int x = 0; x < w; x++)
                        {
                            byte* px = scan0 + y * stride + x * 4;
                            if (px[2] > 180 && px[1] > 140 && px[0] < 120)
                                count++;
                        }
                }
            }
            finally { bmp.UnlockBits(data); }
        }
        catch { }
        return count;
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        CloseHorizontalRuler();
        _checkTimer?.Dispose();
        _checkTimer = null;
        _rulerBitmap?.Dispose();
        _rulerBitmap = null;
        base.OnFormClosed(e);
        if (_instance == this) _instance = null;
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= 0x80;    // WS_EX_TOOLWINDOW — hide from Alt+Tab
            cp.ExStyle |= 0x80000; // WS_EX_LAYERED — per-pixel alpha
            return cp;
        }
    }

    /// <summary>
    /// Draggable horizontal ruler overlay — managed by the vertical RulerOverlayForm.
    /// </summary>
    private class HorizontalRulerForm : Form
    {
        private Bitmap? _bitmap;
        private Point _homeLocation;
        private Point _dragStart;
        private bool _dragging;
        public bool IsDragging => _dragging;

        [DllImport("user32.dll")]
        private static extern int ShowCursor(bool bShow);

        public HorizontalRulerForm(Bitmap bitmap, Point location)
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            Cursor = Cursors.SizeAll;

            LayeredWindowHelper.PremultiplyBitmapAlpha(bitmap);
            _bitmap = bitmap;
            _homeLocation = location;
            Size = new Size(bitmap.Width, bitmap.Height);
            Location = location;

            MouseDown += OnDragStart;
            MouseMove += OnDragMove;
            MouseUp += OnDragEnd;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (_bitmap != null)
                LayeredWindowHelper.SetBitmap(this, _bitmap);
        }

        protected override void OnPaint(PaintEventArgs e) { }

        private void OnDragStart(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _dragging = true;
                _dragStart = e.Location;
                ShowCursor(false);
            }
        }

        private void OnDragMove(object? sender, MouseEventArgs e)
        {
            if (!_dragging) return;
            Location = new Point(
                Location.X + e.X - _dragStart.X,
                Location.Y + e.Y - _dragStart.Y
            );
            if (_bitmap != null && IsHandleCreated && !IsDisposed)
                LayeredWindowHelper.SetBitmap(this, _bitmap);
        }

        private void OnDragEnd(object? sender, MouseEventArgs e)
        {
            _dragging = false;
            ShowCursor(true);
            Location = _homeLocation;
            if (_bitmap != null && IsHandleCreated && !IsDisposed)
                LayeredWindowHelper.SetBitmap(this, _bitmap);
        }

        public void UpdateBitmap(Bitmap bitmap, Point location)
        {
            LayeredWindowHelper.PremultiplyBitmapAlpha(bitmap);
            var old = _bitmap;
            _bitmap = bitmap;
            old?.Dispose();
            _homeLocation = location;
            Size = new Size(bitmap.Width, bitmap.Height);
            if (!_dragging) Location = location;
            if (IsHandleCreated && !IsDisposed)
                LayeredWindowHelper.SetBitmap(this, _bitmap);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _bitmap?.Dispose();
            _bitmap = null;
            base.OnFormClosed(e);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x80;    // WS_EX_TOOLWINDOW
                cp.ExStyle |= 0x80000; // WS_EX_LAYERED
                return cp;
            }
        }
    }
}
