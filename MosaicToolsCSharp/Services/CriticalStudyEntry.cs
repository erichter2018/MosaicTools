namespace MosaicTools.Services;

/// <summary>
/// Represents a study where a critical note was placed via the Critical Findings action.
/// These are tracked per-session (volatile, not persisted).
/// </summary>
public class CriticalStudyEntry
{
    /// <summary>
    /// The study accession number.
    /// </summary>
    public string Accession { get; set; } = "";

    /// <summary>
    /// Patient name in "Lastname Firstname" format (title case).
    /// </summary>
    public string PatientName { get; set; } = "";

    /// <summary>
    /// Site code (e.g., "MLC", "UNM").
    /// </summary>
    public string SiteCode { get; set; } = "";

    /// <summary>
    /// Study description (e.g., "CT ABDOMEN PELVIS").
    /// </summary>
    public string Description { get; set; } = "";

    /// <summary>
    /// Medical Record Number (MRN) - required for XML file drop to open studies.
    /// Memory only, not persisted.
    /// </summary>
    public string Mrn { get; set; } = "";

    /// <summary>
    /// When the critical note was placed.
    /// </summary>
    public DateTime CriticalNoteTime { get; set; }
}
