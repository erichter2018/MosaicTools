using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace MosaicTools.Services;

/// <summary>
/// OCR service using Windows native Windows.Media.Ocr.
/// No Tesseract dependency - uses Windows 10/11 built-in OCR.
/// </summary>
public class OcrService
{
    private readonly OcrEngine? _engine;
    
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
            
            Logger.Trace($"Virtual screen: Left={vScreenLeft}, Top={vScreenTop}, W={vScreenWidth}, H={vScreenHeight}");
            
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
            
            Logger.Trace($"Yellow scan found {yellowPoints.Count} points (step 4)");
            
            if (yellowPoints.Count < 50)
            {
                Logger.Trace("Not enough yellow pixels found");
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
            
            Logger.Trace($"Found {clusters.Count} clusters");

            // Find best cluster
            Rectangle? bestRect = null;
            long maxArea = 0;

            foreach (var r in clusters)
            {
                Logger.Trace($"Cluster: {r.Width}x{r.Height} at {r.X},{r.Y}");

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
                Logger.Trace($"Selected Best Cluster: {r.Width}x{r.Height} at ({resultRect.X},{resultRect.Y})");
                return resultRect;
            }

            Logger.Trace("No valid clusters found");
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
                            if (px[2] > 180 && px[1] > 140 && px[0] < 120)
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
                            if (px[2] > 180 && px[1] > 140 && px[0] < 120)
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
                            bool isYellow = (px[2] > 180 && px[1] > 140 && px[0] < 120);
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
                            if (px[2] > 180 && px[1] > 140 && px[0] < 120)
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
                            if (px[2] > 180 && px[1] > 140 && px[0] < 120)
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
                            bool isYellow = (px[2] > 180 && px[1] > 140 && px[0] < 120);
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
    /// Measure an object at a screen point using region-growing segmentation.
    /// Captures a region around the point, segments by intensity similarity,
    /// then calculates bidimensional measurements using the calibrated ruler scale.
    /// </summary>
    public MeasurementResult? MeasureObjectAtPoint(
        Point screenPoint, Rectangle viewport, double pxPerCmH, double pxPerCmV)
    {
        try
        {
            int captureRadius = 200;
            int captureX = Math.Max(viewport.Left, screenPoint.X - captureRadius);
            int captureY = Math.Max(viewport.Top, screenPoint.Y - captureRadius);
            int captureW = Math.Min(captureRadius * 2, viewport.Right - captureX);
            int captureH = Math.Min(captureRadius * 2, viewport.Bottom - captureY);

            int seedX = screenPoint.X - captureX;
            int seedY = screenPoint.Y - captureY;
            if (seedX < 0 || seedY < 0 || seedX >= captureW || seedY >= captureH) return null;

            // Capture
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

            // Convert to grayscale
            byte[,] gray = new byte[captureH, captureW];
            var rData = region.LockBits(new Rectangle(0, 0, captureW, captureH),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                unsafe
                {
                    byte* scan0 = (byte*)rData.Scan0;
                    int stride = rData.Stride;
                    for (int y = 0; y < captureH; y++)
                        for (int x = 0; x < captureW; x++)
                        {
                            byte* px = scan0 + y * stride + x * 4;
                            gray[y, x] = (byte)((px[2] * 299 + px[1] * 587 + px[0] * 114) / 1000);
                        }
                }
            }
            finally { region.UnlockBits(rData); }

            // Adaptive region growing — increase tolerance until region is large enough
            byte seedIntensity = gray[seedY, seedX];
            bool[,]? segmented = null;
            int segCount = 0;
            int maxPixels = captureW * captureH / 4;

            for (int tol = 15; tol <= 80; tol += 10)
            {
                var (mask, count) = FloodFill(gray, captureW, captureH, seedX, seedY, seedIntensity, tol, maxPixels);
                if (count >= 16)
                {
                    segmented = mask;
                    segCount = count;
                    if (count >= 50 || count > maxPixels) break;
                }
            }

            if (segmented == null || segCount < 16)
            {
                Logger.Trace($"AutoMeasure: segmentation failed (count={segCount})");
                return null;
            }

            Logger.Trace($"AutoMeasure: segmented {segCount} pixels, seed=({seedX},{seedY}), " +
                $"seedIntensity={seedIntensity}, capture={captureW}x{captureH}");

            // Save debug image
            try
            {
                using var debugBmp = new Bitmap(captureW, captureH, PixelFormat.Format32bppArgb);
                for (int y = 0; y < captureH; y++)
                    for (int x = 0; x < captureW; x++)
                    {
                        if (segmented[y, x])
                            debugBmp.SetPixel(x, y, Color.FromArgb(255, 0, 255, 0));
                        else
                            debugBmp.SetPixel(x, y, Color.FromArgb(255, gray[y, x], gray[y, x], gray[y, x]));
                    }
                // Mark seed
                for (int d = -3; d <= 3; d++)
                {
                    if (seedX + d >= 0 && seedX + d < captureW) debugBmp.SetPixel(seedX + d, seedY, Color.Red);
                    if (seedY + d >= 0 && seedY + d < captureH) debugBmp.SetPixel(seedX, seedY + d, Color.Red);
                }
                var debugPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MosaicTools", "measure_debug.png");
                debugBmp.Save(debugPath, ImageFormat.Png);
            }
            catch { }

            // Collect edge pixels
            var edgePixels = new List<(int x, int y)>();
            for (int y = 0; y < captureH; y++)
                for (int x = 0; x < captureW; x++)
                {
                    if (!segmented[y, x]) continue;
                    bool isEdge = false;
                    for (int dy = -1; dy <= 1 && !isEdge; dy++)
                        for (int dx = -1; dx <= 1 && !isEdge; dx++)
                        {
                            int ny = y + dy, nx = x + dx;
                            if (ny < 0 || ny >= captureH || nx < 0 || nx >= captureW || !segmented[ny, nx])
                                isEdge = true;
                        }
                    if (isEdge) edgePixels.Add((x, y));
                }

            if (edgePixels.Count < 4) return null;

            // Rotating calipers: find angle with maximum spread (= major axis direction)
            double bestAngle = 0, maxSpread = 0;
            for (int deg = 0; deg < 180; deg += 5)
            {
                double rad = deg * Math.PI / 180.0;
                double cosA = Math.Cos(rad), sinA = Math.Sin(rad);
                double minP = double.MaxValue, maxP = double.MinValue;
                foreach (var (ex, ey) in edgePixels)
                {
                    double p = ex * cosA + ey * sinA;
                    if (p < minP) minP = p;
                    if (p > maxP) maxP = p;
                }
                double spread = maxP - minP;
                if (spread > maxSpread) { maxSpread = spread; bestAngle = deg; }
            }

            // Minor axis = perpendicular
            double perpRad = (bestAngle + 90) * Math.PI / 180.0;
            double perpCos = Math.Cos(perpRad), perpSin = Math.Sin(perpRad);
            double perpMin = double.MaxValue, perpMax = double.MinValue;
            foreach (var (ex, ey) in edgePixels)
            {
                double p = ex * perpCos + ey * perpSin;
                if (p < perpMin) perpMin = p;
                if (p > perpMax) perpMax = p;
            }
            double minorSpread = perpMax - perpMin;

            // Convert px → cm using direction-weighted px/cm (Pythagorean for anisotropic pixels)
            double majorRad = bestAngle * Math.PI / 180.0;
            double cosM = Math.Cos(majorRad), sinM = Math.Sin(majorRad);
            double majorPxPerCm = Math.Sqrt(cosM * cosM * pxPerCmH * pxPerCmH + sinM * sinM * pxPerCmV * pxPerCmV);
            double minorPxPerCm = Math.Sqrt(sinM * sinM * pxPerCmH * pxPerCmH + cosM * cosM * pxPerCmV * pxPerCmV);

            double majorCm = maxSpread / majorPxPerCm;
            double minorCm = minorSpread / minorPxPerCm;

            Logger.Trace($"AutoMeasure: edges={edgePixels.Count}, majorSpread={maxSpread:F1}px@{bestAngle}°, " +
                $"minorSpread={minorSpread:F1}px, pxPerCm={majorPxPerCm:F1}/{minorPxPerCm:F1}");

            // Find actual endpoints for measurement lines
            double majCos = Math.Cos(majorRad), majSin = Math.Sin(majorRad);
            (int x, int y) majMinPt = edgePixels[0], majMaxPt = edgePixels[0];
            double majMinProj = double.MaxValue, majMaxProj = double.MinValue;
            (int x, int y) perMinPt = edgePixels[0], perMaxPt = edgePixels[0];
            double perMinProj = double.MaxValue, perMaxProj = double.MinValue;

            foreach (var (ex, ey) in edgePixels)
            {
                double mp = ex * majCos + ey * majSin;
                if (mp < majMinProj) { majMinProj = mp; majMinPt = (ex, ey); }
                if (mp > majMaxProj) { majMaxProj = mp; majMaxPt = (ex, ey); }
                double pp = ex * perpCos + ey * perpSin;
                if (pp < perMinProj) { perMinProj = pp; perMinPt = (ex, ey); }
                if (pp > perMaxProj) { perMaxProj = pp; perMaxPt = (ex, ey); }
            }

            // Convert local → screen coords
            return new MeasurementResult(majorCm, minorCm, screenPoint,
                new Point(captureX + majMinPt.x, captureY + majMinPt.y),
                new Point(captureX + majMaxPt.x, captureY + majMaxPt.y),
                new Point(captureX + perMinPt.x, captureY + perMinPt.y),
                new Point(captureX + perMaxPt.x, captureY + perMaxPt.y));
        }
        catch (Exception ex)
        {
            Logger.Trace($"MeasureObjectAtPoint error: {ex.Message}");
            return null;
        }
    }

    private static (bool[,] mask, int count) FloodFill(
        byte[,] gray, int w, int h, int seedX, int seedY, byte seedIntensity, int tolerance, int maxPixels)
    {
        var mask = new bool[h, w];
        int count = 0;
        int lo = Math.Max(0, seedIntensity - tolerance);
        int hi = Math.Min(255, seedIntensity + tolerance);

        var queue = new Queue<(int x, int y)>();
        queue.Enqueue((seedX, seedY));
        mask[seedY, seedX] = true;

        while (queue.Count > 0 && count < maxPixels)
        {
            var (cx, cy) = queue.Dequeue();
            count++;

            for (int dy = -1; dy <= 1; dy++)
                for (int dx = -1; dx <= 1; dx++)
                {
                    if (dx == 0 && dy == 0) continue;
                    int nx = cx + dx, ny = cy + dy;
                    if (nx < 0 || nx >= w || ny < 0 || ny >= h) continue;
                    if (mask[ny, nx]) continue;
                    int gv = gray[ny, nx];
                    if (gv >= lo && gv <= hi)
                    {
                        mask[ny, nx] = true;
                        queue.Enqueue((nx, ny));
                    }
                }
        }

        return (mask, count);
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
