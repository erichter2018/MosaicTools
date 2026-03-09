using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using OpenCvSharp;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace MosaicTools.Services;

/// <summary>
/// OCR service using Windows native Windows.Media.Ocr.
/// No Tesseract dependency - uses Windows 10/11 built-in OCR.
/// </summary>
public class OcrService
{
    private readonly OcrEngine? _engine;
    private readonly MobileSamService _mobileSam = new();
    private string? _lastVirtualScreenKey;
    private string? _lastYellowResult;
    
    public OcrService()
    {
        try
        {
            // Try to create OCR engine with user's language
            _engine = OcrEngine.TryCreateFromUserProfileLanguages();
            
            if (_engine == null)
            {
                // Fallback to English
                _engine = OcrEngine.TryCreateFromLanguage(new Windows.Globalization.Language("en-US"));
            }
            
            if (_engine != null)
            {
                Logger.Trace($"Windows OCR initialized: {_engine.RecognizerLanguage.DisplayName}");
            }
            else
            {
                Logger.Trace("Windows OCR engine not available");
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"OCR initialization failed: {ex.Message}");
        }
    }
    
    public bool IsAvailable => _engine != null;

    /// <summary>Cached px/cm from last ruler capture. Set by CaptureRulerStrip/CaptureHorizontalRulerStrip.</summary>
    public static double? LastVerticalPxPerCm { get; private set; }
    public static double? LastHorizontalPxPerCm { get; private set; }
    
    /// <summary>
    /// Run OCR on a bitmap and return recognized text.
    /// </summary>
    public async Task<string?> RecognizeAsync(Bitmap bitmap)
    {
        if (_engine == null) return null;
        
        try
        {
            // Convert Bitmap to SoftwareBitmap
            using var ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Bmp);
            ms.Position = 0;
            
            var decoder = await BitmapDecoder.CreateAsync(ms.AsRandomAccessStream());
            using var softwareBitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
            
            // Run OCR
            var result = await _engine.RecognizeAsync(softwareBitmap);
            
            return result.Text;
        }
        catch (Exception ex)
        {
            Logger.Trace($"OCR recognition failed: {ex.Message}");
            return null;
        }
    }
    
    /// <summary>
    /// Capture screen region and run OCR.
    /// Handles multi-monitor setups with negative coordinates.
    /// </summary>
    /// <summary>
    /// Capture screen region and run OCR.
    /// Uses full virtual screen capture + crop to reliably handle multi-monitor/negative coords.
    /// </summary>
    public async Task<string?> CaptureAndRecognizeAsync(Rectangle region)
    {
        try
        {
            Logger.Trace($"Capture region: X={region.X}, Y={region.Y}, W={region.Width}, H={region.Height}");
            
            // Get virtual screen bounds (same as FindYellowTarget which is known to work)
            var vScreenLeft = GetSystemMetrics(76); // SM_XVIRTUALSCREEN
            var vScreenTop = GetSystemMetrics(77);  // SM_YVIRTUALSCREEN
            var vScreenWidth = GetSystemMetrics(78); // SM_CXVIRTUALSCREEN
            var vScreenHeight = GetSystemMetrics(79); // SM_CYVIRTUALSCREEN
            
            using var fullScreen = new Bitmap(vScreenWidth, vScreenHeight);
            using var g = Graphics.FromImage(fullScreen);
            
            // Capture the entire virtual screen
            g.CopyFromScreen(vScreenLeft, vScreenTop, 0, 0, new Size(vScreenWidth, vScreenHeight));
            
            // Calculate crop coordinates relative to the full virtual screen bitmap
            // Bitmap (0,0) corresponds to (vScreenLeft, vScreenTop)
            int cropX = region.X - vScreenLeft;
            int cropY = region.Y - vScreenTop;
            
            // Ensure crop is within bounds
            if (cropX < 0 || cropY < 0 || cropX + region.Width > vScreenWidth || cropY + region.Height > vScreenHeight)
            {
                Logger.Trace($"Crop region out of bounds: ({cropX},{cropY}) size {region.Width}x{region.Height} in {vScreenWidth}x{vScreenHeight}");
                return null;
            }

            // Crop the specific region
            Rectangle cropRect = new Rectangle(cropX, cropY, region.Width, region.Height);
            Logger.Trace($"Cropping bitmap: Rect={cropRect.X},{cropRect.Y} {cropRect.Width}x{cropRect.Height} from full screen");
            
            using var rawCrop = fullScreen.Clone(cropRect, fullScreen.PixelFormat);
            
            // Scale up 2x for better OCR recognition of small text
            using var croppedBitmap = new Bitmap(rawCrop.Width * 2, rawCrop.Height * 2);
            using (var gScale = Graphics.FromImage(croppedBitmap))
            {
                gScale.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                gScale.DrawImage(rawCrop, 0, 0, croppedBitmap.Width, croppedBitmap.Height);
            }

            var result = await RecognizeAsync(croppedBitmap);
            
            string logText = result?.Replace("\r", "|").Replace("\n", "|") ?? "null";
            if (logText.Length > 200) logText = logText.Substring(0, 200) + "...";
            Logger.Trace($"Raw OCR text: {logText}");
            
            return result;
        }
        catch (Exception ex)
        {
            Logger.Trace($"Screen capture failed: {ex.Message}");
            return null;
        }
    }
    
    /// <summary>
    /// Find yellow target box on screen (for InteleViewer).
    /// Uses clustering to identify the largest dense yellow region, ignoring noise.
    /// Returns the full bounding box (Rectangle) in virtual screen coordinates.
    /// </summary>
    public Rectangle? FindYellowTarget()
    {
        try
        {
            // Multi-monitor: get virtual screen bounds
            var vScreenLeft = GetSystemMetrics(76); // SM_XVIRTUALSCREEN
            var vScreenTop = GetSystemMetrics(77);  // SM_YVIRTUALSCREEN
            var vScreenWidth = GetSystemMetrics(78); // SM_CXVIRTUALSCREEN
            var vScreenHeight = GetSystemMetrics(79); // SM_CYVIRTUALSCREEN
            
            // Log virtual screen bounds only once
            var vsKey = $"{vScreenLeft},{vScreenTop},{vScreenWidth},{vScreenHeight}";
            if (vsKey != _lastVirtualScreenKey)
            {
                _lastVirtualScreenKey = vsKey;
                Logger.Trace($"Virtual screen: Left={vScreenLeft}, Top={vScreenTop}, W={vScreenWidth}, H={vScreenHeight}");
            }
            
            using var screenshot = new Bitmap(vScreenWidth, vScreenHeight);
            using var g = Graphics.FromImage(screenshot);
            g.CopyFromScreen(vScreenLeft, vScreenTop, 0, 0, new Size(vScreenWidth, vScreenHeight));
            
            // List of detected yellow points (relative to bitmap)
            var yellowPoints = new System.Collections.Generic.List<Point>();
            
            // 1. Scan for yellow pixels using LockBits for performance
            var bmpData = screenshot.LockBits(new Rectangle(0, 0, vScreenWidth, vScreenHeight),
                System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            try
            {
                unsafe
                {
                    byte* scan0 = (byte*)bmpData.Scan0;
                    int stride = bmpData.Stride;
                    for (int y = 0; y < vScreenHeight; y += 4) // Step 4 for speed (large box = many pixels)
                    {
                        for (int x = 0; x < vScreenWidth; x += 4)
                        {
                            byte* pixel = scan0 + y * stride + x * 4;
                            byte pb = pixel[0]; // Blue
                            byte pg = pixel[1]; // Green
                            byte pr = pixel[2]; // Red
                            // Yellow detection: R > 180, G > 160, B < 120
                            if (pr > 180 && pg > 160 && pb < 120)
                            {
                                yellowPoints.Add(new Point(x, y));
                            }
                        }
                    }
                }
            }
            finally
            {
                screenshot.UnlockBits(bmpData);
            }
            
            if (yellowPoints.Count < 50)
            {
                if (_lastYellowResult != "none")
                {
                    _lastYellowResult = "none";
                    Logger.Trace($"Yellow scan: only {yellowPoints.Count} points, no viewport");
                }
                return null;
            }

            // 2. Cluster points
            // Simple approach: Iterate points, if point is within MergeDistance of existing cluster, add to it.
            // Since we scan linear order, most points will belong to the "current" or "previous" cluster.
            
            var clusters = new System.Collections.Generic.List<Rectangle>();
            const int MergeDist = 15; // Refined from 100 which was too aggressive
            
            // We use a simplified grid merging to avoid O(N^2)
            // But for <5000 points (11k / 4 step ~ 2500), N^2 is 6M ops, might be slow.
            // Better: Just maintain a list of Rectangles (bounds of clusters). 
            // If a point fits in or near a rect, expand it. If it fits multiple, merge them.
            
            foreach (var p in yellowPoints)
            {
                int bestClusterIdx = -1;
                
                for (int i = 0; i < clusters.Count; i++)
                {
                    var r = clusters[i];
                    // Check if point is close to this rectangle (inflate rect by MergeDist to check)
                    if (p.X >= r.Left - MergeDist && p.X <= r.Right + MergeDist &&
                        p.Y >= r.Top - MergeDist && p.Y <= r.Bottom + MergeDist)
                    {
                        if (bestClusterIdx == -1)
                        {
                            bestClusterIdx = i;
                        }
                        else
                        {
                            // Point belongs to two clusters -> Merge current cluster into bestCluster, remove current
                            var best = clusters[bestClusterIdx];
                            var current = clusters[i];
                            
                            int minX = Math.Min(best.Left, current.Left);
                            int minY = Math.Min(best.Top, current.Top);
                            int maxX = Math.Max(best.Right, current.Right);
                            int maxY = Math.Max(best.Bottom, current.Bottom);
                            
                            clusters[bestClusterIdx] = new Rectangle(minX, minY, maxX - minX, maxY - minY);
                            clusters.RemoveAt(i);
                            i--; // Adjust index
                        }
                    }
                }
                
                if (bestClusterIdx != -1)
                {
                    // Expand the found cluster
                    var r = clusters[bestClusterIdx];
                    int minX = Math.Min(r.Left, p.X);
                    int minY = Math.Min(r.Top, p.Y);
                    int maxX = Math.Max(r.Right, p.X);
                    int maxY = Math.Max(r.Bottom, p.Y);
                    clusters[bestClusterIdx] = new Rectangle(minX, minY, maxX - minX, maxY - minY);
                }
                else
                {
                    // Start new cluster 
                    // Initial size 1x1
                    clusters.Add(new Rectangle(p.X, p.Y, 1, 1));
                }
            }
            
            // Find best cluster
            Rectangle? bestRect = null;
            long maxArea = 0;

            foreach (var r in clusters)
            {
                // Filter noise
                if (r.Width < 50 || r.Height < 50) continue;

                long area = (long)r.Width * r.Height;

                if (area > maxArea)
                {
                    maxArea = area;
                    bestRect = r;
                }
            }

            if (bestRect.HasValue)
            {
                var r = bestRect.Value;
                // Convert back to screen coords
                var resultRect = new Rectangle(vScreenLeft + r.X, vScreenTop + r.Y, r.Width, r.Height);
                var key = $"{r.Width}x{r.Height}@{resultRect.X},{resultRect.Y}";
                if (key != _lastYellowResult)
                {
                    _lastYellowResult = key;
                    Logger.Trace($"Yellow viewport: {r.Width}x{r.Height} at ({resultRect.X},{resultRect.Y})");
                }
                return resultRect;
            }

            if (_lastYellowResult != "none")
            {
                _lastYellowResult = "none";
                Logger.Trace("No valid yellow clusters found");
            }
            return null;
        }
        catch (Exception ex)
        {
            Logger.Trace($"Yellow box find error: {ex.Message}");
            return null;
        }
    }
    
    /// <summary>
    /// Capture the left-edge ruler strip from the active IV viewport.
    /// Returns the bitmap and its screen location, or null if no viewport found.
    /// </summary>
    public (Bitmap bitmap, Point location)? CaptureRulerStrip(Rectangle viewportBounds)
    {
        try
        {
            // Capture from viewport left edge, full height.
            // We'll use pixel analysis to separate tick marks from borders and labels.
            int stripWidth = Math.Max(60, (int)(viewportBounds.Width * 0.08));
            int stripX = viewportBounds.Left;
            int stripY = viewportBounds.Top;
            int stripHeight = viewportBounds.Height;

            Logger.Trace($"CaptureRulerStrip: viewport={viewportBounds.Width}x{viewportBounds.Height}, strip={stripWidth}x{stripHeight}");

            var vScreenLeft = GetSystemMetrics(76);
            var vScreenTop = GetSystemMetrics(77);
            var vScreenWidth = GetSystemMetrics(78);
            var vScreenHeight = GetSystemMetrics(79);

            using var fullScreen = new Bitmap(vScreenWidth, vScreenHeight);
            using var g = Graphics.FromImage(fullScreen);
            g.CopyFromScreen(vScreenLeft, vScreenTop, 0, 0, new Size(vScreenWidth, vScreenHeight));

            int bmpX = stripX - vScreenLeft;
            int bmpY = stripY - vScreenTop;

            if (bmpX < 0 || bmpY < 0 || bmpX + stripWidth > vScreenWidth || bmpY + stripHeight > vScreenHeight)
                return null;

            var cropRect = new Rectangle(bmpX, bmpY, stripWidth, stripHeight);
            var strip = fullScreen.Clone(cropRect, PixelFormat.Format32bppArgb);

            // Save raw strip before processing
            try
            {
                var rawPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MosaicTools", "ruler_raw.png");
                strip.Save(rawPath, ImageFormat.Png);
            }
            catch { }

            var stripData = strip.LockBits(new Rectangle(0, 0, strip.Width, strip.Height),
                ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
            try
            {
                unsafe
                {
                    byte* scan0 = (byte*)stripData.Scan0;
                    int stride = stripData.Stride;
                    int w = strip.Width, h = strip.Height;

                    // Pass 1: Count yellow pixels per column and per row
                    int[] yellowPerCol = new int[w];
                    int[] yellowPerRow = new int[h];
                    for (int y = 0; y < h; y++)
                    {
                        for (int x = 0; x < w; x++)
                        {
                            byte* px = scan0 + y * stride + x * 4;
                            if (px[2] > 180 && px[1] > 80 && px[0] < 120)
                            {
                                yellowPerCol[x]++;
                                yellowPerRow[y]++;
                            }
                        }
                    }

                    // Vertical border columns: >80% of rows have yellow (full-height viewport border).
                    // The ruler's vertical line only spans the tick area (~60-70% density).
                    bool[] isBorderCol = new bool[w];
                    int borderColThreshold = (int)(h * 0.80);
                    for (int x = 0; x < w; x++)
                        isBorderCol[x] = yellowPerCol[x] > borderColThreshold;

                    // Horizontal border rows: yellow spans >60% of strip width
                    bool[] isBorderRow = new bool[h];
                    int borderRowThreshold = (int)(w * 0.6);
                    for (int y = 0; y < h; y++)
                        isBorderRow[y] = yellowPerRow[y] > borderRowThreshold;

                    int midY = h / 2;

                    // Debug: log border detection stats
                    int borderColCount = 0, borderRowCount = 0;
                    for (int x = 0; x < w; x++) if (isBorderCol[x]) borderColCount++;
                    for (int y = 0; y < h; y++) if (isBorderRow[y]) borderRowCount++;
                    // Log top 5 column densities and max row yellow count
                    var colDensities = new System.Text.StringBuilder();
                    for (int x = 0; x < Math.Min(10, w); x++)
                        colDensities.Append($"c{x}={yellowPerCol[x]} ");
                    int maxRowYellow = 0;
                    for (int y = 0; y < h; y++) if (yellowPerRow[y] > maxRowYellow) maxRowYellow = yellowPerRow[y];
                    Logger.Trace($"CaptureRulerStrip: borderCols={borderColCount}(thr={borderColThreshold}), borderRows={borderRowCount}(thr={borderRowThreshold}), maxRowYellow={maxRowYellow}, {colDensities}");

                    // Debug: sample pixel RGB at midpoint of strip to diagnose color issues
                    {
                        var rgbSample = new System.Text.StringBuilder("RGB samples col5: ");
                        int sampleCol = Math.Min(5, w - 1);
                        for (int si = 0; si < 5 && si < h; si++)
                        {
                            int sampleRow = h / 6 * (si + 1);
                            if (sampleRow >= h) sampleRow = h - 1;
                            byte* sp = scan0 + sampleRow * stride + sampleCol * 4;
                            rgbSample.Append($"y{sampleRow}=({sp[2]},{sp[1]},{sp[0]}) ");
                        }
                        Logger.Trace($"CaptureRulerStrip: {rgbSample}");
                    }

                    // Pass 2: Find the vertical range of tick marks (not the vertical line).
                    // The vertical ruler line is in the leftmost ~3 columns; ticks extend beyond.
                    // We use columns 4+ to determine where ticks exist vertically.
                    int tickMinRow = h, tickMaxRow = 0;
                    for (int y = 0; y < h; y++)
                    {
                        if (isBorderRow[y]) continue;
                        for (int x = 4; x < w; x++)
                        {
                            if (isBorderCol[x]) continue; // Skip ruler vertical line
                            byte* px = scan0 + y * stride + x * 4;
                            if (px[2] > 180 && px[1] > 80 && px[0] < 120)
                            {
                                if (y < tickMinRow) tickMinRow = y;
                                if (y > tickMaxRow) tickMaxRow = y;
                                break;
                            }
                        }
                    }

                    // Find ruler line start: first non-border column with significant yellow.
                    // Everything to the left (viewport border, R/L label) gets excluded.
                    int rulerStartCol = 0;
                    int significantThreshold = (int)(h * 0.3);
                    for (int x = 0; x < w; x++)
                    {
                        if (!isBorderCol[x] && yellowPerCol[x] > significantThreshold)
                        {
                            rulerStartCol = x;
                            break;
                        }
                    }
                    Logger.Trace($"CaptureRulerStrip: rulerStartCol={rulerStartCol}");

                    // Build tick mask: yellow pixels within tick range, at or right of ruler line
                    bool[] isTick = new bool[w * h];
                    for (int y = 0; y < h; y++)
                    {
                        bool inTickRange = (y >= tickMinRow && y <= tickMaxRow);
                        for (int x = 0; x < w; x++)
                        {
                            byte* px = scan0 + y * stride + x * 4;
                            bool isYellow = (px[2] > 180 && px[1] > 80 && px[0] < 120);
                            isTick[y * w + x] = isYellow && x >= rulerStartCol && !isBorderRow[y] && inTickRange;
                        }
                    }

                    // Calibration: count tick pixels per row to find tick spacing
                    {
                        int[] tickWidthPerRow = new int[h];
                        for (int y = tickMinRow; y <= tickMaxRow; y++)
                            for (int x = 0; x < w; x++)
                                if (isTick[y * w + x]) tickWidthPerRow[y]++;
                        LastVerticalPxPerCm = AnalyzeTickBands(tickWidthPerRow, h, "V");
                    }

                    // Dilate: expand tick mask by 1px in all directions for bolder lines
                    bool[] dilated = new bool[w * h];
                    for (int y = 0; y < h; y++)
                    {
                        for (int x = 0; x < w; x++)
                        {
                            if (isTick[y * w + x])
                            {
                                for (int dy = -1; dy <= 1; dy++)
                                    for (int dx = -1; dx <= 1; dx++)
                                    {
                                        int ny = y + dy, nx = x + dx;
                                        if (ny >= 0 && ny < h && nx >= 0 && nx < w)
                                            dilated[ny * w + nx] = true;
                                    }
                            }
                        }
                    }

                    // Apply colors using dilated mask
                    int minCol = w, maxCol = 0, minRow = h, maxRow = 0;
                    int tickCount = 0;

                    for (int y = 0; y < h; y++)
                    {
                        for (int x = 0; x < w; x++)
                        {
                            byte* px = scan0 + y * stride + x * 4;
                            if (dilated[y * w + x])
                            {
                                px[0] = 255; px[1] = 200; px[2] = 0; px[3] = 255; // Electric blue
                                tickCount++;
                                if (x < minCol) minCol = x;
                                if (x > maxCol) maxCol = x;
                                if (y < minRow) minRow = y;
                                if (y > maxRow) maxRow = y;
                            }
                            else
                            {
                                px[0] = 0; px[1] = 0; px[2] = 0; px[3] = 3;
                            }
                        }
                    }

                    Logger.Trace($"CaptureRulerStrip: {w}x{h}, tickRange={tickMinRow}-{tickMaxRow}, ticks={tickCount}");

                    // Auto-crop to tick extent + padding
                    if (minRow <= maxRow && minCol <= maxCol)
                    {
                        strip.UnlockBits(stripData);
                        int pad = 5;
                        int cL = Math.Max(0, minCol - pad);
                        int cT = Math.Max(0, minRow - pad);
                        int cR = Math.Min(w - 1, maxCol + pad);
                        int cB = Math.Min(h - 1, maxRow + pad);

                        var cropped = strip.Clone(
                            new Rectangle(cL, cT, cR - cL + 1, cB - cT + 1),
                            PixelFormat.Format32bppArgb);
                        strip.Dispose();
                        return (cropped, new Point(stripX + cL, stripY + cT));
                    }
                }
            }
            finally
            {
                try { strip.UnlockBits(stripData); } catch { }
            }

            return (strip, new Point(stripX, stripY));
        }
        catch (Exception ex)
        {
            Logger.Trace($"CaptureRulerStrip error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Captures and processes the horizontal ruler strip from the bottom edge of the viewport.
    /// The horizontal ruler line runs near the bottom; ticks extend upward from it.
    /// Below the ruler line: labels (P/A), text info, viewport border — all excluded.
    /// </summary>
    public (Bitmap bitmap, Point location)? CaptureHorizontalRulerStrip(Rectangle viewportBounds)
    {
        try
        {
            int stripHeight = Math.Max(60, (int)(viewportBounds.Height * 0.08));
            int stripX = viewportBounds.Left;
            int stripY = viewportBounds.Bottom - stripHeight;
            int stripWidth = viewportBounds.Width;

            Logger.Trace($"CaptureHorizontalRulerStrip: viewport={viewportBounds.Width}x{viewportBounds.Height}, strip={stripWidth}x{stripHeight}");

            var vScreenLeft = GetSystemMetrics(76);
            var vScreenTop = GetSystemMetrics(77);
            var vScreenWidth = GetSystemMetrics(78);
            var vScreenHeight = GetSystemMetrics(79);

            using var fullScreen = new Bitmap(vScreenWidth, vScreenHeight);
            using var g = Graphics.FromImage(fullScreen);
            g.CopyFromScreen(vScreenLeft, vScreenTop, 0, 0, new Size(vScreenWidth, vScreenHeight));

            int bmpX = stripX - vScreenLeft;
            int bmpY = stripY - vScreenTop;

            if (bmpX < 0 || bmpY < 0 || bmpX + stripWidth > vScreenWidth || bmpY + stripHeight > vScreenHeight)
                return null;

            var cropRect = new Rectangle(bmpX, bmpY, stripWidth, stripHeight);
            var strip = fullScreen.Clone(cropRect, PixelFormat.Format32bppArgb);

            try
            {
                var rawPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MosaicTools", "hruler_raw.png");
                strip.Save(rawPath, ImageFormat.Png);
            }
            catch { }

            var stripData = strip.LockBits(new Rectangle(0, 0, strip.Width, strip.Height),
                ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
            try
            {
                unsafe
                {
                    byte* scan0 = (byte*)stripData.Scan0;
                    int stride = stripData.Stride;
                    int w = strip.Width, h = strip.Height;

                    // Pass 1: Count yellow pixels per column and per row
                    int[] yellowPerCol = new int[w];
                    int[] yellowPerRow = new int[h];
                    for (int y = 0; y < h; y++)
                    {
                        for (int x = 0; x < w; x++)
                        {
                            byte* px = scan0 + y * stride + x * 4;
                            if (px[2] > 180 && px[1] > 80 && px[0] < 120)
                            {
                                yellowPerCol[x]++;
                                yellowPerRow[y]++;
                            }
                        }
                    }

                    // Border rows: >80% of columns have yellow (full-width viewport border)
                    bool[] isBorderRow = new bool[h];
                    int borderRowThreshold = (int)(w * 0.80);
                    for (int y = 0; y < h; y++)
                        isBorderRow[y] = yellowPerRow[y] > borderRowThreshold;

                    // Border columns: yellow spans >60% of strip height (vertical viewport borders)
                    bool[] isBorderCol = new bool[w];
                    int borderColThreshold = (int)(h * 0.6);
                    for (int x = 0; x < w; x++)
                        isBorderCol[x] = yellowPerCol[x] > borderColThreshold;

                    int borderRowCount = 0, borderColCount = 0;
                    for (int y = 0; y < h; y++) if (isBorderRow[y]) borderRowCount++;
                    for (int x = 0; x < w; x++) if (isBorderCol[x]) borderColCount++;

                    // Find the ruler end row: search bottom-up for the last non-border row
                    // with significant yellow density. This is the horizontal ruler line itself.
                    // Everything BELOW it (P label, text, viewport border) gets excluded.
                    int rulerEndRow = h - 1;
                    int significantThreshold = (int)(w * 0.3);
                    for (int y = h - 1; y >= 0; y--)
                    {
                        if (!isBorderRow[y] && yellowPerRow[y] > significantThreshold)
                        {
                            rulerEndRow = y;
                            break;
                        }
                    }

                    // Find tick range: scan non-border rows ABOVE the ruler line for yellow in non-border columns
                    int tickMinCol = w, tickMaxCol = 0;
                    int tickMinRow = h, tickMaxRow = 0;
                    for (int y = 0; y <= rulerEndRow; y++)
                    {
                        if (isBorderRow[y]) continue;
                        for (int x = 0; x < w; x++)
                        {
                            if (isBorderCol[x]) continue;
                            byte* px = scan0 + y * stride + x * 4;
                            if (px[2] > 180 && px[1] > 80 && px[0] < 120)
                            {
                                if (x < tickMinCol) tickMinCol = x;
                                if (x > tickMaxCol) tickMaxCol = x;
                                if (y < tickMinRow) tickMinRow = y;
                                if (y > tickMaxRow) tickMaxRow = y;
                            }
                        }
                    }

                    // Log row densities near the ruler for debugging
                    var rowDensities = new System.Text.StringBuilder();
                    for (int y = Math.Max(0, rulerEndRow - 5); y <= Math.Min(h - 1, rulerEndRow + 5); y++)
                        rowDensities.Append($"r{y}={yellowPerRow[y]}{(isBorderRow[y] ? "B" : "")} ");
                    Logger.Trace($"CaptureHorizontalRulerStrip: borderRows={borderRowCount}(thr={borderRowThreshold}), borderCols={borderColCount}(thr={borderColThreshold}), rulerEndRow={rulerEndRow}, tickRows={tickMinRow}-{tickMaxRow}, tickCols={tickMinCol}-{tickMaxCol}, {rowDensities}");

                    // Build tick mask: yellow at or above rulerEndRow, in tick column range, not border
                    bool[] isTick = new bool[w * h];
                    for (int y = 0; y <= rulerEndRow; y++)
                    {
                        if (isBorderRow[y]) continue;
                        for (int x = 0; x < w; x++)
                        {
                            if (isBorderCol[x]) continue;
                            bool inTickRange = (x >= tickMinCol && x <= tickMaxCol);
                            byte* px = scan0 + y * stride + x * 4;
                            bool isYellow = (px[2] > 180 && px[1] > 80 && px[0] < 120);
                            isTick[y * w + x] = isYellow && inTickRange;
                        }
                    }

                    // Calibration: count tick pixels per column to find tick spacing
                    {
                        int[] tickHeightPerCol = new int[w];
                        for (int x = tickMinCol; x <= tickMaxCol; x++)
                            for (int y = 0; y <= rulerEndRow; y++)
                                if (isTick[y * w + x]) tickHeightPerCol[x]++;
                        LastHorizontalPxPerCm = AnalyzeTickBands(tickHeightPerCol, w, "H");
                    }

                    // Dilate by 1px for bolder lines
                    bool[] dilated = new bool[w * h];
                    for (int y = 0; y < h; y++)
                    {
                        for (int x = 0; x < w; x++)
                        {
                            if (isTick[y * w + x])
                            {
                                for (int dy = -1; dy <= 1; dy++)
                                    for (int dx = -1; dx <= 1; dx++)
                                    {
                                        int ny = y + dy, nx = x + dx;
                                        if (ny >= 0 && ny < h && nx >= 0 && nx < w)
                                            dilated[ny * w + nx] = true;
                                    }
                            }
                        }
                    }

                    // Apply colors
                    int minCol = w, maxCol = 0, minRow = h, maxRow = 0;
                    int tickCount = 0;

                    for (int y = 0; y < h; y++)
                    {
                        for (int x = 0; x < w; x++)
                        {
                            byte* px = scan0 + y * stride + x * 4;
                            if (dilated[y * w + x])
                            {
                                px[0] = 255; px[1] = 200; px[2] = 0; px[3] = 255; // Electric blue
                                tickCount++;
                                if (x < minCol) minCol = x;
                                if (x > maxCol) maxCol = x;
                                if (y < minRow) minRow = y;
                                if (y > maxRow) maxRow = y;
                            }
                            else
                            {
                                px[0] = 0; px[1] = 0; px[2] = 0; px[3] = 3;
                            }
                        }
                    }

                    Logger.Trace($"CaptureHorizontalRulerStrip: {w}x{h}, ticks={tickCount}");

                    // Auto-crop to tick extent + padding
                    if (minRow <= maxRow && minCol <= maxCol)
                    {
                        strip.UnlockBits(stripData);
                        int pad = 5;
                        int cL = Math.Max(0, minCol - pad);
                        int cT = Math.Max(0, minRow - pad);
                        int cR = Math.Min(w - 1, maxCol + pad);
                        int cB = Math.Min(h - 1, maxRow + pad);

                        var cropped = strip.Clone(
                            new Rectangle(cL, cT, cR - cL + 1, cB - cT + 1),
                            PixelFormat.Format32bppArgb);
                        strip.Dispose();

                        try
                        {
                            var debugPath = System.IO.Path.Combine(
                                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                "MosaicTools", "hruler_debug.png");
                            cropped.Save(debugPath, ImageFormat.Png);
                        }
                        catch { }

                        return (cropped, new Point(stripX + cL, stripY + cT));
                    }
                }
            }
            finally
            {
                try { strip.UnlockBits(stripData); } catch { }
            }

            return (strip, new Point(stripX, stripY));
        }
        catch (Exception ex)
        {
            Logger.Trace($"CaptureHorizontalRulerStrip error: {ex.Message}");
            return null;
        }
    }

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    // ────────────────────────────────────────────────────────────────────
    //  Auto-Measurement: calibration + segmentation + bidimensional sizing
    // ────────────────────────────────────────────────────────────────────

    /// <summary>Result of an automatic object measurement.</summary>
    public record MeasurementResult(
        double MajorAxisCm, double MinorAxisCm,
        Point ScreenCenter,
        Point MajorStart, Point MajorEnd,
        Point MinorStart, Point MinorEnd);

    /// <summary>
    /// Returns cached px/cm from last ruler capture. Call CaptureRulerStrip /
    /// CaptureHorizontalRulerStrip first (they compute this automatically).
    /// If no rulers have been captured yet, triggers a fresh capture.
    /// </summary>
    public (double vertical, double horizontal)? CalculatePixelsPerCm(Rectangle viewportBounds)
    {
        // If not cached yet, trigger a capture to populate them
        if (LastVerticalPxPerCm == null && LastHorizontalPxPerCm == null)
        {
            var vr = CaptureRulerStrip(viewportBounds);
            vr?.bitmap.Dispose();
            var hr = CaptureHorizontalRulerStrip(viewportBounds);
            hr?.bitmap.Dispose();
        }

        if (LastVerticalPxPerCm == null && LastHorizontalPxPerCm == null) return null;

        double v = LastVerticalPxPerCm ?? LastHorizontalPxPerCm!.Value;
        double h = LastHorizontalPxPerCm ?? LastVerticalPxPerCm!.Value;

        // Try OCR-based refinement on the horizontal ruler area
        TryOcrRulerCalibration(viewportBounds, ref h, ref v);

        return (v, h);
    }

    /// <summary>
    /// OCR the horizontal ruler strip to find numeric labels and refine px/cm.
    /// Looks for two numbers whose positions match major tick positions.
    /// </summary>
    private void TryOcrRulerCalibration(Rectangle viewportBounds, ref double hPxPerCm, ref double vPxPerCm)
    {
        try
        {
            // Capture the horizontal ruler area (wider than normal to include labels)
            int stripHeight = Math.Max(80, (int)(viewportBounds.Height * 0.10));
            int stripX = viewportBounds.Left;
            int stripY = viewportBounds.Bottom - stripHeight;
            int stripWidth = viewportBounds.Width;

            var vScreenLeft = GetSystemMetrics(76);
            var vScreenTop = GetSystemMetrics(77);
            var vScreenWidth = GetSystemMetrics(78);
            var vScreenHeight = GetSystemMetrics(79);

            using var fullScreen = new Bitmap(vScreenWidth, vScreenHeight);
            using (var g = Graphics.FromImage(fullScreen))
                g.CopyFromScreen(vScreenLeft, vScreenTop, 0, 0, new Size(vScreenWidth, vScreenHeight));

            int bmpX = stripX - vScreenLeft;
            int bmpY = stripY - vScreenTop;
            if (bmpX < 0 || bmpY < 0 || bmpX + stripWidth > vScreenWidth || bmpY + stripHeight > vScreenHeight)
                return;

            using var strip = fullScreen.Clone(
                new Rectangle(bmpX, bmpY, stripWidth, stripHeight), PixelFormat.Format32bppArgb);

            // Scale up 3x for better OCR
            using var scaled = new Bitmap(strip.Width * 3, strip.Height * 3);
            using (var g2 = Graphics.FromImage(scaled))
            {
                g2.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g2.DrawImage(strip, 0, 0, scaled.Width, scaled.Height);
            }

            // Run OCR
            var task = RecognizeWithPositionsAsync(scaled);
            task.Wait();
            var words = task.Result;
            if (words == null || words.Count == 0) return;

            // Find numeric words and their X positions (in original strip coords)
            var numbers = new List<(double value, double xCenter)>();
            foreach (var (text, bounds) in words)
            {
                if (double.TryParse(text, out double val) && val >= 0 && val <= 100)
                {
                    // Convert scaled coords back to original strip coords
                    double xCenter = (bounds.X + bounds.Width / 2) / 3.0;
                    numbers.Add((val, xCenter));
                }
            }

            if (numbers.Count < 2)
            {
                Logger.Trace($"OcrRulerCalibration: found {numbers.Count} numbers (need 2+), " +
                    $"raw words: [{string.Join(", ", words.Select(w => w.text))}]");
                return;
            }

            // Sort by X position and find two adjacent numbers with consistent spacing
            numbers.Sort((a, b) => a.xCenter.CompareTo(b.xCenter));
            for (int i = 1; i < numbers.Count; i++)
            {
                double valueDiff = Math.Abs(numbers[i].value - numbers[i - 1].value);
                double pixelDiff = Math.Abs(numbers[i].xCenter - numbers[i - 1].xCenter);
                if (valueDiff > 0 && pixelDiff > 20)
                {
                    double pxPerCm = pixelDiff / valueDiff;
                    Logger.Trace($"OcrRulerCalibration: '{numbers[i-1].value}' at x={numbers[i-1].xCenter:F0} → " +
                        $"'{numbers[i].value}' at x={numbers[i].xCenter:F0} → " +
                        $"{pixelDiff:F0}px / {valueDiff}cm = {pxPerCm:F1} px/cm " +
                        $"(was H={hPxPerCm:F1}, V={vPxPerCm:F1})");
                    hPxPerCm = pxPerCm;
                    vPxPerCm = pxPerCm; // assume isotropic unless vertical OCR says otherwise
                    return;
                }
            }

            Logger.Trace($"OcrRulerCalibration: no valid number pairs from {numbers.Count} numbers");
        }
        catch (Exception ex)
        {
            Logger.Trace($"OcrRulerCalibration error: {ex.Message}");
        }
    }

    /// <summary>
    /// Run OCR and return words with their bounding rectangles.
    /// </summary>
    public async Task<List<(string text, Windows.Foundation.Rect bounds)>?> RecognizeWithPositionsAsync(Bitmap bitmap)
    {
        if (_engine == null) return null;

        try
        {
            using var ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Bmp);
            ms.Position = 0;

            var decoder = await BitmapDecoder.CreateAsync(ms.AsRandomAccessStream());
            using var softwareBitmap = await decoder.GetSoftwareBitmapAsync(
                BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);

            var result = await _engine.RecognizeAsync(softwareBitmap);

            var words = new List<(string text, Windows.Foundation.Rect bounds)>();
            foreach (var line in result.Lines)
                foreach (var word in line.Words)
                    words.Add((word.Text, word.BoundingRect));

            return words;
        }
        catch (Exception ex)
        {
            Logger.Trace($"RecognizeWithPositionsAsync error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Common logic: given per-row (or per-column) tick pixel counts from the raw
    /// tick mask, find tick positions and calculate px/cm.
    /// Uses the spacing between ALL consecutive ticks (minor tick interval),
    /// assumes that interval = 1cm.
    /// </summary>
    private static double? AnalyzeTickBands(int[] tickPixelCount, int count, string label)
    {
        // Find baseline (spine-only rows have ~1 pixel from the ruler line)
        var nonZero = tickPixelCount.Where(c => c > 0).OrderBy(c => c).ToArray();
        if (nonZero.Length < 10)
        {
            Logger.Trace($"TickSpacing({label}): only {nonZero.Length} non-zero positions");
            return null;
        }

        int baseline = nonZero[nonZero.Length / 4];
        int tickThreshold = baseline + Math.Max(2, baseline);

        // Group consecutive positions above threshold into tick bands
        var bands = new List<(int center, int maxWidth)>();
        int bandStart = -1, bandMax = 0;
        for (int i = 0; i <= count; i++)
        {
            bool isTick = i < count && tickPixelCount[i] > tickThreshold;
            if (isTick)
            {
                if (bandStart < 0) bandStart = i;
                bandMax = Math.Max(bandMax, tickPixelCount[i]);
            }
            else if (bandStart >= 0)
            {
                bands.Add(((bandStart + i - 1) / 2, bandMax));
                bandStart = -1;
                bandMax = 0;
            }
        }

        if (bands.Count < 3)
        {
            Logger.Trace($"TickSpacing({label}): baseline={baseline}, thr={tickThreshold}, " +
                $"max={nonZero.Last()}, only {bands.Count} bands (need 3+)");
            return null;
        }

        // Calculate ALL consecutive spacings (minor tick interval)
        var allSpacings = new List<double>();
        for (int i = 1; i < bands.Count; i++)
            allSpacings.Add(bands[i].center - bands[i - 1].center);
        allSpacings.Sort();
        double minorSpacing = allSpacings[allSpacings.Count / 2]; // median

        // Identify major ticks by width and count subs between first two majors
        int longest = bands.Max(b => b.maxWidth);
        int majorThr = (int)(longest * 0.6);
        var majors = bands.Where(b => b.maxWidth >= majorThr).OrderBy(b => b.center).ToList();
        int subsBetween = majors.Count >= 2
            ? bands.Count(b => b.maxWidth < majorThr
                && b.center > majors[0].center && b.center < majors[1].center)
            : -1;

        // px/cm = minor tick spacing, assuming each minor tick = 1cm
        double pxPerCm = minorSpacing;

        Logger.Trace($"TickSpacing({label}): {bands.Count} bands, {majors.Count} major, " +
            $"{subsBetween} subs, minorSpacing={minorSpacing:F1}px, majorSpacing=" +
            $"{(majors.Count >= 2 ? (majors[1].center - majors[0].center).ToString("F1") : "?")}px " +
            $"→ {pxPerCm:F1} px/cm");

        return pxPerCm;
    }

    /// <summary>
    /// Measure an object at a screen point using MobileSAM (ONNX) for segmentation,
    /// then OpenCV FitEllipse for axis measurements.
    /// </summary>
    public MeasurementResult? MeasureObjectAtPoint(
        Point screenPoint, Rectangle viewport, double pxPerCmH, double pxPerCmV)
    {
        try
        {
            // Use a small crop so the lesion dominates the frame — SAM segments
            // "the object" so the lesion needs to be the largest thing visible.
            int captureRadius = 80;
            int captureX = Math.Max(viewport.Left, screenPoint.X - captureRadius);
            int captureY = Math.Max(viewport.Top, screenPoint.Y - captureRadius);
            int captureW = Math.Min(captureRadius * 2, viewport.Right - captureX);
            int captureH = Math.Min(captureRadius * 2, viewport.Bottom - captureY);

            int seedX = screenPoint.X - captureX;
            int seedY = screenPoint.Y - captureY;
            if (seedX < 0 || seedY < 0 || seedX >= captureW || seedY >= captureH) return null;

            // Capture screen region
            var vScreenLeft = GetSystemMetrics(76);
            var vScreenTop = GetSystemMetrics(77);
            var vScreenWidth = GetSystemMetrics(78);
            var vScreenHeight = GetSystemMetrics(79);

            using var fullScreen = new Bitmap(vScreenWidth, vScreenHeight);
            using (var g = Graphics.FromImage(fullScreen))
                g.CopyFromScreen(vScreenLeft, vScreenTop, 0, 0, new Size(vScreenWidth, vScreenHeight));

            int bmpX = captureX - vScreenLeft;
            int bmpY = captureY - vScreenTop;
            if (bmpX < 0 || bmpY < 0 || bmpX + captureW > vScreenWidth || bmpY + captureH > vScreenHeight)
                return null;

            using var region = fullScreen.Clone(
                new Rectangle(bmpX, bmpY, captureW, captureH), PixelFormat.Format32bppArgb);

            // === Region Growing + Edge-Snap Refinement ===
            // Region growing was "almost there" — slightly over-measures.
            // Fix: grow the region, then trim boundary pixels that sit on strong edges.
            var sw = System.Diagnostics.Stopwatch.StartNew();

            // Convert region to grayscale
            var rDataRg = region.LockBits(new Rectangle(0, 0, captureW, captureH),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            using var regionMat = Mat.FromPixelData(captureH, captureW, MatType.CV_8UC4, rDataRg.Scan0, rDataRg.Stride);
            using var grayMat = new Mat();
            Cv2.CvtColor(regionMat, grayMat, ColorConversionCodes.BGRA2GRAY);
            region.UnlockBits(rDataRg);

            // 1. Bilateral filter — edge-preserving smoothing
            using var filtered = new Mat();
            Cv2.BilateralFilter(grayMat, filtered, 9, 75, 75);

            // 2. Gradient magnitude for edge stopping during grow
            using var gradX = new Mat();
            using var gradY = new Mat();
            using var gradMag = new Mat();
            Cv2.Sobel(filtered, gradX, MatType.CV_64F, 1, 0, 3);
            Cv2.Sobel(filtered, gradY, MatType.CV_64F, 0, 1, 3);
            Cv2.Magnitude(gradX, gradY, gradMag);
            using var gradNorm = new Mat();
            Cv2.Normalize(gradMag, gradNorm, 0, 255, NormTypes.MinMax);
            gradNorm.ConvertTo(gradNorm, MatType.CV_8UC1);

            // 3. Sample seed neighborhood for adaptive threshold
            int nhRadius = 3;
            int nhX1 = Math.Max(0, seedX - nhRadius), nhY1 = Math.Max(0, seedY - nhRadius);
            int nhX2 = Math.Min(captureW - 1, seedX + nhRadius), nhY2 = Math.Min(captureH - 1, seedY + nhRadius);
            using var seedRegion = new Mat(filtered, new OpenCvSharp.Rect(nhX1, nhY1, nhX2 - nhX1 + 1, nhY2 - nhY1 + 1));
            Cv2.MeanStdDev(seedRegion, out Scalar meanVal, out Scalar stddevVal);

            double seedMean = meanVal[0];
            double seedStddev = Math.Max(stddevVal[0], 5.0);
            double multiplier = 2.5;
            double lo = seedMean - multiplier * seedStddev;
            double hi = seedMean + multiplier * seedStddev;
            double gradThreshold = 40;

            Logger.Trace($"AutoMeasure: grow seed=({seedX},{seedY}), mean={seedMean:F1}, " +
                $"std={seedStddev:F1}, range=[{lo:F0},{hi:F0}]");

            // 4. Region growing with 8-connectivity
            var visited = new bool[captureH, captureW];
            var member = new bool[captureH, captureW];
            var queue = new Queue<(int x, int y)>();
            queue.Enqueue((seedX, seedY));
            visited[seedY, seedX] = true;
            member[seedY, seedX] = true;
            int totalPixels = 1;
            int checkpointCount = 0;
            int[] dx = { -1, 1, 0, 0, -1, -1, 1, 1 };
            int[] dy = { 0, 0, -1, 1, -1, 1, -1, 1 };

            while (queue.Count > 0)
            {
                var (px, py) = queue.Dequeue();
                for (int d = 0; d < 8; d++)
                {
                    int nx = px + dx[d], ny = py + dy[d];
                    if (nx < 0 || ny < 0 || nx >= captureW || ny >= captureH) continue;
                    if (visited[ny, nx]) continue;
                    visited[ny, nx] = true;

                    byte val = filtered.At<byte>(ny, nx);
                    byte grad = gradNorm.At<byte>(ny, nx);

                    if (val >= lo && val <= hi && grad < gradThreshold)
                    {
                        member[ny, nx] = true;
                        queue.Enqueue((nx, ny));
                        totalPixels++;
                    }
                }

                if (totalPixels - checkpointCount >= 500)
                {
                    if (checkpointCount > 200)
                    {
                        double growthRate = (double)totalPixels / checkpointCount;
                        if (growthRate > 3.0) break;
                    }
                    checkpointCount = totalPixels;
                    if (totalPixels > captureW * captureH * 0.6) break;
                }
            }

            Logger.Trace($"AutoMeasure: grew to {totalPixels}px in {sw.ElapsedMilliseconds}ms");

            if (totalPixels < 50)
            {
                Logger.Trace("AutoMeasure: region too small");
                return null;
            }

            // 5. Convert to mask
            using var growMask = new Mat(captureH, captureW, MatType.CV_8UC1, Scalar.All(0));
            for (int y = 0; y < captureH; y++)
                for (int x = 0; x < captureW; x++)
                    if (member[y, x])
                        growMask.Set(y, x, (byte)255);

            // 6. Edge-snap refinement: erode, then trim pixels near strong edges
            //    This tightens the boundary to the actual lesion edge.
            using var edges = new Mat();
            Cv2.Canny(filtered, edges, 30, 90);
            using var edgeDilK = Cv2.GetStructuringElement(MorphShapes.Ellipse, new OpenCvSharp.Size(3, 3));
            Cv2.Dilate(edges, edges, edgeDilK);

            // Subtract edge pixels from the grown region (snaps boundary inward)
            using var roiMask = new Mat();
            Cv2.Subtract(growMask, edges, roiMask);

            // 7. Keep only the connected component containing the seed
            using var ccLabels = new Mat();
            int nLabels = Cv2.ConnectedComponents(roiMask, ccLabels);
            int seedLabel = ccLabels.At<int>(seedY, seedX);

            if (seedLabel > 0)
            {
                using var seedMask = new Mat(captureH, captureW, MatType.CV_8UC1, Scalar.All(0));
                int finalCount = 0;
                for (int y = 0; y < captureH; y++)
                    for (int x = 0; x < captureW; x++)
                        if (ccLabels.At<int>(y, x) == seedLabel)
                        {
                            seedMask.Set(y, x, (byte)255);
                            finalCount++;
                        }
                seedMask.CopyTo(roiMask);
                Logger.Trace($"AutoMeasure: edge-snap refined to {finalCount}px");
            }

            // 8. Morphological cleanup: close holes from edge subtraction, smooth boundary
            using var closeKernel = Cv2.GetStructuringElement(MorphShapes.Ellipse, new OpenCvSharp.Size(5, 5));
            Cv2.MorphologyEx(roiMask, roiMask, MorphTypes.Close, closeKernel, iterations: 2);
            using var openKernel = Cv2.GetStructuringElement(MorphShapes.Ellipse, new OpenCvSharp.Size(3, 3));
            Cv2.MorphologyEx(roiMask, roiMask, MorphTypes.Open, openKernel);

            int segCount = Cv2.CountNonZero(roiMask);
            if (segCount < 50)
            {
                Logger.Trace("AutoMeasure: final mask too small after edge-snap");
                return null;
            }

            // === FindContours → pick largest → FitEllipse ===
            Cv2.FindContours(roiMask, out OpenCvSharp.Point[][] finalContours, out _,
                RetrievalModes.External, ContourApproximationModes.ApproxSimple);

            if (finalContours.Length == 0)
            {
                Logger.Trace("AutoMeasure: no contours from mask");
                return null;
            }

            // Pick the largest contour
            var fitContour = finalContours.OrderByDescending(c => Cv2.ContourArea(c)).First();
            if (fitContour.Length < 5)
            {
                Logger.Trace($"AutoMeasure: contour too small ({fitContour.Length} pts)");
                return null;
            }

            // === FitEllipse for precise axes ===
            var ellipse = Cv2.FitEllipse(fitContour);
            float majorPx = Math.Max(ellipse.Size.Width, ellipse.Size.Height);
            float minorPx = Math.Min(ellipse.Size.Width, ellipse.Size.Height);

            // Angle of major axis
            double angleDeg = ellipse.Angle;
            if (ellipse.Size.Height > ellipse.Size.Width)
                angleDeg += 90;
            double angleRad = angleDeg * Math.PI / 180.0;

            // Convert px → cm using direction-weighted scale
            double cosM = Math.Cos(angleRad), sinM = Math.Sin(angleRad);
            double majorPxPerCm = Math.Sqrt(cosM * cosM * pxPerCmH * pxPerCmH + sinM * sinM * pxPerCmV * pxPerCmV);
            double minorPxPerCm = Math.Sqrt(sinM * sinM * pxPerCmH * pxPerCmH + cosM * cosM * pxPerCmV * pxPerCmV);

            double majorCm = majorPx / majorPxPerCm;
            double minorCm = minorPx / minorPxPerCm;

            Logger.Trace($"AutoMeasure: contour={fitContour.Length}pts, " +
                $"ellipse={majorPx:F1}x{minorPx:F1}px@{angleDeg:F0}°, " +
                $"center=({ellipse.Center.X:F0},{ellipse.Center.Y:F0}), " +
                $"{majorCm:F1}x{minorCm:F1}cm");

            // Calculate endpoint coords from ellipse
            double cx = ellipse.Center.X, cy = ellipse.Center.Y;
            double majCos = Math.Cos(angleRad), majSin = Math.Sin(angleRad);
            double perpRad = angleRad + Math.PI / 2;
            double perCos = Math.Cos(perpRad), perSin = Math.Sin(perpRad);

            int majX1 = (int)(cx - majCos * majorPx / 2);
            int majY1 = (int)(cy - majSin * majorPx / 2);
            int majX2 = (int)(cx + majCos * majorPx / 2);
            int majY2 = (int)(cy + majSin * majorPx / 2);

            int minX1 = (int)(cx - perCos * minorPx / 2);
            int minY1 = (int)(cy - perSin * minorPx / 2);
            int minX2 = (int)(cx + perCos * minorPx / 2);
            int minY2 = (int)(cy + perSin * minorPx / 2);

            // Save debug image
            try
            {
                var rData = region.LockBits(new Rectangle(0, 0, captureW, captureH),
                    ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                using var src = Mat.FromPixelData(captureH, captureW, MatType.CV_8UC4, rData.Scan0, rData.Stride);
                using var dbgGray = new Mat();
                Cv2.CvtColor(src, dbgGray, ColorConversionCodes.BGRA2GRAY);
                region.UnlockBits(rData);

                using var debugMat = new Mat();
                Cv2.CvtColor(dbgGray, debugMat, ColorConversionCodes.GRAY2BGR);
                for (int y = 0; y < captureH; y++)
                    for (int x = 0; x < captureW; x++)
                        if (roiMask.At<byte>(y, x) > 0)
                            debugMat.Set(y, x, new Vec3b(0, 180, 0));
                Cv2.DrawContours(debugMat, new[] { fitContour }, 0, new Scalar(0, 255, 0), 1);
                Cv2.Ellipse(debugMat, ellipse, new Scalar(0, 255, 255), 1);
                Cv2.Line(debugMat, new OpenCvSharp.Point(majX1, majY1), new OpenCvSharp.Point(majX2, majY2), new Scalar(200, 255, 0), 2);
                Cv2.Line(debugMat, new OpenCvSharp.Point(minX1, minY1), new OpenCvSharp.Point(minX2, minY2), new Scalar(0, 200, 255), 2);
                Cv2.Circle(debugMat, new OpenCvSharp.Point(seedX, seedY), 4, new Scalar(0, 0, 255), -1);
                var debugPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MosaicTools", "measure_debug.png");
                debugMat.SaveImage(debugPath);
            }
            catch { }

            // Convert local → screen coords
            return new MeasurementResult(majorCm, minorCm, screenPoint,
                new Point(captureX + majX1, captureY + majY1),
                new Point(captureX + majX2, captureY + majY2),
                new Point(captureX + minX1, captureY + minY1),
                new Point(captureX + minX2, captureY + minY2));
        }
        catch (Exception ex)
        {
            Logger.Trace($"MeasureObjectAtPoint error: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Extract series/image numbers from OCR text.
    /// Focus on the LAST Se: and Im: entries (the first Se: is usually a date).
    /// </summary>
    public static string? ExtractSeriesImageNumbers(string? text, string? template = null)
    {
        if (string.IsNullOrEmpty(text)) return null;
        
        string logText = text.Replace("\r", " ").Replace("\n", " ");
        if (logText.Length > 100) logText = logText.Substring(0, 100) + "...";
        Logger.Trace($"Extracting from: [{logText}]");
        
        // Clean common OCR typos
        text = Regex.Replace(text, @"\b(0e|5e|8e|Be|S8|S6)\b", "Se", RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"\b(Irm|1rn|1m|Lm|1n|In|ln)\b", "Im", RegexOptions.IgnoreCase);
        
        string? finalSeries = null;
        string? finalImage = null;
        
        // Find ALL Se: patterns in the text  
        // Pattern: Se: followed by optional letters/symbols, then digits
        // Exclude if immediately followed by month names (dates)
        var seMatches = Regex.Matches(text, @"[Ss]e:?\s*([A-Z]*\s*#?)(\d+)", RegexOptions.IgnoreCase);
        Logger.Trace($"Found {seMatches.Count} 'Se:' candidates");

        foreach (Match m in seMatches)
        {
            var prefix = m.Groups[1].Value; // e.g., "CT #" or "DX #" or just ""
            var number = m.Groups[2].Value;
            
            // Get context after the match to check for date patterns
            var afterMatch = text.Substring(m.Index, Math.Min(50, text.Length - m.Index));
            
            // Skip if this is a date pattern (Se: Jan 16, 2026)
            if (Regex.IsMatch(afterMatch, @"^[Ss]e:?\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)", RegexOptions.IgnoreCase))
            {
                Logger.Trace($"Skipping date pattern match: '{m.Value}'");
                continue;
            }
            
            Logger.Trace($"Series candidate accepted: prefix='{prefix}', number='{number}'");
            finalSeries = number;
        }
        
        // Find ALL Im: patterns or X/Y patterns
        var imMatches = Regex.Matches(text, @"[Ii]m:?\s*(\d+)", RegexOptions.IgnoreCase);
        Logger.Trace($"Found {imMatches.Count} 'Im:' candidates");
        foreach (Match m in imMatches)
        {
            finalImage = m.Groups[1].Value;
            Logger.Trace($"Image candidate: {finalImage}");
        }
        
        // Fallback: look for X/Y pattern (e.g., "34/156")
        // Also handles cases where '/' is misread as '1', 'I', 'l', '|'
        if (finalImage == null)
        {
            // Pattern: Number + (Separator) + Number
            // Separators: /, \, |, 1, I, l
            var xyMatches = Regex.Matches(text, @"\b(\d+)\s*([/\\|1Il])\s*(\d+)\b");
            
            foreach (Match m in xyMatches)
            {
                if (int.TryParse(m.Groups[1].Value, out int current) && 
                    int.TryParse(m.Groups[3].Value, out int total))
                {
                    string separator = m.Groups[2].Value;
                    
                    // If separator is ambiguous (looks like a digit/letter), enforce strict rules
                    // to avoid false positives (e.g. dates, or just "Image 1 of 5")
                    bool isAmbiguous = separator == "1" || separator == "I" || separator == "l";

                    // Rule: Current image must be <= Total images (and Total > 0)
                    // When separator is ambiguous, also require image <= series to reduce OCR false positives
                    if (isAmbiguous && finalSeries != null && int.TryParse(finalSeries, out int seriesNum) && current > seriesNum)
                        continue;

                    if (total > 0 && current <= total)
                    {
                        finalImage = current.ToString();
                        Logger.Trace($"Image from fallback pattern '{m.Value}': {finalImage} (Sep='{separator}', {current}/{total})");
                        break; // Stop at first valid match
                    }
                }
            }
        }
        
        if (finalSeries != null || finalImage != null)
        {
            var t = template ?? "(series {series}, image {image})";
            var result = t
                .Replace("{series}", finalSeries ?? "?")
                .Replace("{image}", finalImage ?? "?");

            Logger.Trace($"Extraction result: {result}");
            return result;
        }
        
        Logger.Trace("No series/image found");
        return null;
    }
}
