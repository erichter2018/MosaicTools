namespace MosaicTools.Services;

/// <summary>
/// Command interface for mutating Mosaic UI state (focus, clicks).
/// Today: implemented by AutomationService via FlaUI.
/// Tomorrow: can be backed by a Mosaic API.
/// </summary>
public interface IMosaicCommander
{
    bool FocusTranscriptBox();
    bool FocusFinalReportBox();
    bool ClickCreateImpression();
    bool SelectImpressionContent();
    bool ClickDiscardStudy();
}
