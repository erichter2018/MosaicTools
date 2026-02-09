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
    
    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);
    
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
