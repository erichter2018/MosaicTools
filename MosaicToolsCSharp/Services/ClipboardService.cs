using System;
using System.Windows.Forms;

namespace MosaicTools.Services;

/// <summary>
/// Clipboard operations with retry logic.
/// </summary>
public static class ClipboardService
{
    /// <summary>
    /// Get text from clipboard with retry.
    /// </summary>
    public static string? GetText(int retries = 5, int delayMs = 50)
    {
        for (int i = 0; i < retries; i++)
        {
            try
            {
                if (Clipboard.ContainsText())
                    return Clipboard.GetText();
                else
                    Thread.Sleep(50);
            }
            catch
            {
                Thread.Sleep(delayMs);
            }
        }
        return null;
    }
    
    /// <summary>
    /// Set text to clipboard with retry.
    /// </summary>
    public static bool SetText(string text, int retries = 5, int delayMs = 50)
    {
        for (int i = 0; i < retries; i++)
        {
            try
            {
                Clipboard.SetText(text);
                return true;
            }
            catch
            {
                Thread.Sleep(delayMs);
            }
        }
        return false;
    }
    
    /// <summary>
    /// Clear clipboard with retry.
    /// </summary>
    public static void Clear(int retries = 3)
    {
        for (int i = 0; i < retries; i++)
        {
            try
            {
                Clipboard.Clear();
                return;
            }
            catch
            {
                Thread.Sleep(50);
            }
        }
    }
}
