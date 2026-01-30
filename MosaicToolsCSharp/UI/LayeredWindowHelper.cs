using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MosaicTools.UI;

/// <summary>
/// Shared infrastructure for WS_EX_LAYERED window rendering.
/// Provides P/Invoke declarations and bitmap manipulation helpers
/// for transparent overlay forms (ReportPopup, ClinicalHistory, Impression).
/// </summary>
internal static class LayeredWindowHelper
{
    #region P/Invoke

    [StructLayout(LayoutKind.Sequential)]
    public struct W32Point { public int X, Y; }

    [StructLayout(LayoutKind.Sequential)]
    public struct W32Size { public int Width, Height; }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct BLENDFUNCTION
    {
        public byte BlendOp, BlendFlags, SourceConstantAlpha, AlphaFormat;
    }

    public const int WS_EX_LAYERED = 0x80000;
    public const int WM_NCLBUTTONDOWN = 0xA1;
    public const int HT_CAPTION = 0x2;

    private const byte AC_SRC_OVER = 0x00;
    private const byte AC_SRC_ALPHA = 0x01;
    private const uint ULW_ALPHA = 0x02;

    [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
    private static extern bool UpdateLayeredWindow(
        IntPtr hwnd, IntPtr hdcDst, ref W32Point pptDst, ref W32Size psize,
        IntPtr hdcSrc, ref W32Point pptSrc, uint crKey,
        ref BLENDFUNCTION pblend, uint dwFlags);

    [DllImport("user32.dll", ExactSpelling = true)]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll", ExactSpelling = true)]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("gdi32.dll", ExactSpelling = true)]
    private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll", ExactSpelling = true)]
    private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

    [DllImport("gdi32.dll", ExactSpelling = true)]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("gdi32.dll", ExactSpelling = true)]
    private static extern bool DeleteDC(IntPtr hdc);

    [DllImport("user32.dll")]
    public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

    [DllImport("user32.dll")]
    public static extern bool ReleaseCapture();

    #endregion

    /// <summary>
    /// Present a 32-bit ARGB bitmap as the layered window content.
    /// </summary>
    public static void SetBitmap(Form form, Bitmap bitmap)
    {
        if (!form.IsHandleCreated || form.IsDisposed) return;

        IntPtr screenDc = GetDC(IntPtr.Zero);
        IntPtr memDc = CreateCompatibleDC(screenDc);
        IntPtr hBitmap = bitmap.GetHbitmap(Color.FromArgb(0));
        IntPtr oldBitmap = SelectObject(memDc, hBitmap);

        try
        {
            var blend = new BLENDFUNCTION
            {
                BlendOp = AC_SRC_OVER,
                BlendFlags = 0,
                SourceConstantAlpha = 255,
                AlphaFormat = AC_SRC_ALPHA
            };

            var size = new W32Size { Width = bitmap.Width, Height = bitmap.Height };
            var source = new W32Point { X = 0, Y = 0 };
            var topPos = new W32Point { X = form.Left, Y = form.Top };

            UpdateLayeredWindow(form.Handle, screenDc, ref topPos, ref size,
                memDc, ref source, 0, ref blend, ULW_ALPHA);
        }
        finally
        {
            SelectObject(memDc, oldBitmap);
            DeleteObject(hBitmap);
            DeleteDC(memDc);
            ReleaseDC(IntPtr.Zero, screenDc);
        }
    }

    /// <summary>
    /// Stamp ClearType text pixels from textLayer onto dst bitmap.
    /// Pixels that differ from the reference background color (bgR, bgG, bgB)
    /// are text pixels - their RGB is copied and alpha set to 255 (fully opaque).
    /// Background pixels are left untouched (semi-transparent).
    /// </summary>
    public static void MergeTextLayer(Bitmap dst, Bitmap textLayer,
        byte bgR = 0, byte bgG = 0, byte bgB = 0, int threshold = 8)
    {
        var rect = new Rectangle(0, 0, dst.Width, dst.Height);
        var dstData = dst.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        var srcData = textLayer.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

        int byteCount = Math.Abs(dstData.Stride) * dstData.Height;
        byte[] dstPx = new byte[byteCount];
        byte[] srcPx = new byte[byteCount];
        Marshal.Copy(dstData.Scan0, dstPx, 0, byteCount);
        Marshal.Copy(srcData.Scan0, srcPx, 0, byteCount);

        for (int i = 0; i < byteCount; i += 4)
        {
            int diff = Math.Abs(srcPx[i] - bgB) +
                       Math.Abs(srcPx[i + 1] - bgG) +
                       Math.Abs(srcPx[i + 2] - bgR);

            if (diff > threshold)
            {
                dstPx[i]     = srcPx[i];     // B
                dstPx[i + 1] = srcPx[i + 1]; // G
                dstPx[i + 2] = srcPx[i + 2]; // R
                dstPx[i + 3] = 255;          // A
            }
        }

        Marshal.Copy(dstPx, 0, dstData.Scan0, byteCount);
        dst.UnlockBits(dstData);
        textLayer.UnlockBits(srcData);
    }

    /// <summary>
    /// Premultiply alpha for correct UpdateLayeredWindow rendering.
    /// GDI+ stores straight alpha, but UpdateLayeredWindow expects premultiplied.
    /// </summary>
    public static void PremultiplyBitmapAlpha(Bitmap bmp)
    {
        var rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
        var data = bmp.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
        int byteCount = Math.Abs(data.Stride) * data.Height;
        byte[] pixels = new byte[byteCount];
        Marshal.Copy(data.Scan0, pixels, 0, byteCount);

        for (int i = 0; i < byteCount; i += 4)
        {
            byte a = pixels[i + 3];
            if (a == 255) continue;
            if (a == 0)
            {
                pixels[i] = pixels[i + 1] = pixels[i + 2] = 0;
                continue;
            }
            pixels[i]     = (byte)(pixels[i]     * a / 255);
            pixels[i + 1] = (byte)(pixels[i + 1] * a / 255);
            pixels[i + 2] = (byte)(pixels[i + 2] * a / 255);
        }

        Marshal.Copy(pixels, 0, data.Scan0, byteCount);
        bmp.UnlockBits(data);
    }
}
