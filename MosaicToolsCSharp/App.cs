namespace MosaicTools;

/// <summary>
/// Global application state and flags.
/// </summary>
public static class App
{
    /// <summary>
    /// True if the application was started with -headless flag.
    /// In headless mode: no widget bar, no hotkeys, no indicator.
    /// PowerMic, floating buttons, toasts, and Windows messages still work.
    /// </summary>
    public static bool IsHeadless { get; internal set; } = false;
}
