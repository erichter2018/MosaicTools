using System.Drawing;
using System.Windows.Forms;

namespace MosaicTools.UI;

/// <summary>
/// Ensures saved window positions are still on a visible monitor.
/// If the point is offscreen (e.g., monitor disconnected), clamps to the nearest screen.
/// </summary>
internal static class ScreenHelper
{
    /// <summary>
    /// Returns the point as-is if it's within any monitor's working area,
    /// otherwise clamps it to the nearest screen's working area.
    /// </summary>
    internal static Point EnsureOnScreen(int x, int y, int width = 0, int height = 0)
    {
        var point = new Point(x, y);

        // Check if the top-left corner is within any screen's working area
        foreach (var screen in Screen.AllScreens)
        {
            if (screen.WorkingArea.Contains(point))
                return point;
        }

        // Off-screen — clamp to the nearest screen's working area
        var nearest = Screen.FromPoint(point);
        var area = nearest.WorkingArea;
        // Math.Max ensures min<=max when form is wider/taller than the screen
        int clampedX = Math.Clamp(x, area.Left, Math.Max(area.Left, area.Right - Math.Max(width, 1)));
        int clampedY = Math.Clamp(y, area.Top, Math.Max(area.Top, area.Bottom - Math.Max(height, 1)));
        return new Point(clampedX, clampedY);
    }

    /// <summary>
    /// Overload for center-based positioning (ClinicalHistoryForm, ImpressionForm).
    /// Takes the saved center point and form size, returns clamped top-left.
    /// </summary>
    internal static Point EnsureOnScreenFromCenter(int centerX, int centerY, int width, int height)
    {
        int x = centerX - width / 2;
        int y = centerY - height / 2;
        return EnsureOnScreen(x, y, width, height);
    }
}
