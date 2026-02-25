namespace MosaicTools.Services;

/// <summary>
/// Read-only interface for Mosaic report/study state.
/// Today: implemented by AutomationService via FlaUI scraping.
/// Tomorrow: can be backed by a Mosaic REST/WebSocket API.
/// </summary>
public interface IMosaicReader
{
    // State properties (populated by GetFinalReportFast)
    string? LastFinalReport { get; }
    string? LastAccession { get; }
    string? LastDescription { get; }
    string? LastTemplateName { get; }
    string? LastPatientGender { get; }
    int? LastPatientAge { get; }
    string? LastPatientName { get; }
    string? LastSiteCode { get; }
    string? LastMrn { get; }
    bool LastDraftedState { get; }
    bool IsAddendumDetected { get; }

    // Data retrieval
    string? GetFinalReportFast(bool checkDraftedStatus = false);
    string? GetReportTextContent();
    bool IsDiscardDialogVisible();
    bool IsDictationActiveUIA();
    void ClearLastReport();

    /// <summary>
    /// Invalidate the cached ProseMirror editor element, forcing the next scrape
    /// to do a fresh search. Call after Process Report when Mosaic rebuilds the editor.
    /// </summary>
    void InvalidateEditorCache();
}
