using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace MosaicTools.Services;

/// <summary>
/// Clinical note formatting service.
/// Matches Python's format_note() function for Clario exam notes.
/// </summary>
public class NoteFormatter
{
    private readonly string _doctorName;
    private readonly string? _template;
    private readonly string? _targetTimezone;  // null = preserve original, otherwise convert

    public NoteFormatter(string doctorName, string? template = null, string? targetTimezone = null)
    {
        _doctorName = doctorName;
        _template = template;
        _targetTimezone = targetTimezone;
    }
    
    /// <summary>
    /// Format Clario exam note into standardized critical findings statement.
    /// </summary>
    public string FormatNote(string rawText)
    {
        try
        {
            // 1. Extract Name (Look for clinician patterns)
            var segments = new List<string>();

            // Prefer title-based matches (Dr., Nurse, NP, PA, etc.) — most reliable
            // \b prevents matching inside words (e.g. "RN" in "Eastern")
            var namePattern = @"(\b(?:Dr\.?|Nurse|NP|PA|RN|MD|DO)\s+.+?)(?=\s+(?:at|with|to|w/|@|said|confirmed|stated|reported|declined|who|confirm|are|is|and|&|connected|contacted|spoke|discussed|transferred|called|notified|informed|reached|paged)\b|\s*[-;:,]|\s+\b(?:Dr\.?|Nurse|NP|PA|RN|MD|DO)|\s+\d{2}/\d{2}/|$)";
            var nameMatches = Regex.Matches(rawText, namePattern, RegexOptions.IgnoreCase);
            foreach (Match m in nameMatches)
            {
                segments.Add(m.Groups[1].Value);
            }

            string? titleAndName = TryExtractName(segments);

            // Fallback: verb-based extraction if title-based found nothing usable
            // (e.g. all title matches were the current doctor)
            // Negative lookbehind prevents matching passive "to be connected" / "been connected"
            if (titleAndName == null)
            {
                segments.Clear();
                var verbPattern = @"(?<!\bto\s+be\s+)(?<!\bbeen\s+)(?:Transferred|Connected\s+with|Connected\s+to|Connected|Was\s+connected\s+with|Spoke\s+to|Spoke\s+with|Discussed\s+with)\s+(.+?)(?=\s+(?:at|@|said|confirmed|stated|reported|declined|who|are|is)|\s*[-;:,]|\s+\d{2}/\d{2}/|$)";
                var verbMatch = Regex.Match(rawText, verbPattern, RegexOptions.IgnoreCase);
                if (verbMatch.Success)
                {
                    var fullSegment = verbMatch.Groups[1].Value;
                    var parts = Regex.Split(fullSegment, @"\s+(?:to|with|w/|and|&)\s+", RegexOptions.IgnoreCase);
                    segments.AddRange(parts);
                }
                titleAndName = TryExtractName(segments);
            }

            titleAndName ??= "Dr. / Nurse [Name not found]";
            
            // 2. Extract End Timestamp
            var endMatch = Regex.Match(rawText, @"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2}\s*(?:AM|PM))", RegexOptions.IgnoreCase);
            string endDate = "N/A";
            string endTimeCtStr = "N/A";
            DateTime? dtEnd = null;
            
            Match? dialogTimeMatch = null;
            if (!endMatch.Success)
            {
                dialogTimeMatch = Regex.Match(rawText, @"(?:@|at|\.)\s+(\d{1,2}:\d{2}\s*(?:AM|PM))\s*([a-zA-Z\s]+)?", RegexOptions.IgnoreCase);
            }
            
            if (endMatch.Success)
            {
                endDate = endMatch.Groups[1].Value;
                endTimeCtStr = NormalizeTimeStr(endMatch.Groups[2].Value);
                
                if (DateTime.TryParseExact($"{endDate} {endTimeCtStr}", "MM/dd/yyyy h:mm tt", 
                    System.Globalization.CultureInfo.InvariantCulture, 
                    System.Globalization.DateTimeStyles.None, out var parsed))
                {
                    dtEnd = parsed;
                }
            }
            else if (dialogTimeMatch != null && dialogTimeMatch.Success)
            {
                // Note Dialog format - infer date from current time
                var textTimeStr = NormalizeTimeStr(dialogTimeMatch.Groups[1].Value);
                var rawTz = dialogTimeMatch.Groups[2].Success ? dialogTimeMatch.Groups[2].Value.Trim() : "";
                var sourceTimezoneStr = GetTimezoneDisplay(rawTz);

                if (DateTime.TryParseExact(textTimeStr, "h:mm tt",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var noteTime))
                {
                    // Use timezone-aware "now" for date inference
                    var targetZone = ResolveTimeZone(_targetTimezone ?? "Central Time")
                                     ?? TimeZoneInfo.Local;
                    var sourceZone = ResolveTimeZone(sourceTimezoneStr)
                                     ?? targetZone;
                    var nowInTarget = TimeZoneInfo.ConvertTime(DateTime.UtcNow, targetZone);
                    var today = nowInTarget.Date;
                    var yesterday = today.AddDays(-1);

                    // Convert the note time from source zone to target zone for date comparison
                    var unspecifiedNote = DateTime.SpecifyKind(
                        today.Add(noteTime.TimeOfDay), DateTimeKind.Unspecified);
                    var noteInTarget = TimeZoneInfo.ConvertTime(unspecifiedNote, sourceZone, targetZone);

                    if (noteInTarget > nowInTarget)
                        endDate = yesterday.ToString("MM/dd/yyyy");
                    else
                        endDate = today.ToString("MM/dd/yyyy");

                    endTimeCtStr = textTimeStr;

                    // Apply timezone conversion if target is set
                    var (finalTime, finalTzAbbr) = ConvertToTargetTimezone(noteTime, sourceTimezoneStr);

                    return $"Critical findings were discussed with and acknowledged by {titleAndName} at {finalTime:h:mm tt} {finalTzAbbr} on {endDate}.";
                }
            }
            
            // 3. Extract Text Time and Timezone
            string finalTimeDisplay = "N/A";
            string diffWarning = "";
            var textTimeMatch = Regex.Match(rawText, @"(?:at|@|\.|-)\s*(\d{1,2}:\d{2}\s*(?i:AM|PM))\s*([a-zA-Z\s]+)?", RegexOptions.IgnoreCase);

            if (textTimeMatch.Success && dtEnd.HasValue)
            {
                var textTimeStr = NormalizeTimeStr(textTimeMatch.Groups[1].Value);
                var rawTz = textTimeMatch.Groups[2].Success ? textTimeMatch.Groups[2].Value.Trim() : "";

                // Determine source timezone from note text (defaults to Central if not detected)
                var sourceTimezoneStr = GetTimezoneDisplay(rawTz);

                if (DateTime.TryParseExact($"{endDate} {textTimeStr}", "MM/dd/yyyy h:mm tt",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var dtText))
                {
                    // Convert discovered text time to Central using DST-aware TimeZoneInfo
                    var sourceZoneInfo = ResolveTimeZone(sourceTimezoneStr);
                    var centralZoneInfo = ResolveTimeZone("Central Time");
                    DateTime dtTextCt;
                    if (sourceZoneInfo != null && centralZoneInfo != null)
                    {
                        var unspecified = DateTime.SpecifyKind(dtText, DateTimeKind.Unspecified);
                        dtTextCt = TimeZoneInfo.ConvertTime(unspecified, sourceZoneInfo, centralZoneInfo);
                    }
                    else
                    {
                        dtTextCt = dtText; // Fallback: assume already Central
                    }

                    // Cross-reference with system "end" timestamp (which is already Central)
                    var diffSeconds = (dtTextCt - dtEnd.Value).TotalSeconds;
                    if (diffSeconds > 43200) dtTextCt = dtTextCt.AddDays(-1);
                    else if (diffSeconds < -43200) dtTextCt = dtTextCt.AddDays(1);

                    var timeDiffMin = Math.Abs((dtTextCt - dtEnd.Value).TotalMinutes);

                    // Apply timezone conversion based on settings
                    var (finalTime, finalTzAbbr) = ConvertToTargetTimezone(dtText, sourceTimezoneStr);
                    finalTimeDisplay = $"{finalTime:h:mm tt} {finalTzAbbr}";

                    if (timeDiffMin > 55)
                        diffWarning = $"NOTE: Entry logged {(int)timeDiffMin} minutes after reported communication.\n";
                }
                else
                {
                    // Parsing failed - use source time with appropriate abbreviation
                    var (_, tzAbbr) = ConvertToTargetTimezone(DateTime.Now, sourceTimezoneStr);
                    finalTimeDisplay = $"{textTimeStr} {tzAbbr}";
                }
            }
            else
            {
                // No text time found - use end timestamp with CST (assumed Central)
                if (!string.IsNullOrEmpty(endTimeCtStr) && endTimeCtStr != "N/A")
                {
                    if (DateTime.TryParseExact(endTimeCtStr, "h:mm tt",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out var endTime))
                    {
                        var (finalTime, finalTzAbbr) = ConvertToTargetTimezone(endTime, "Central Time");
                        finalTimeDisplay = $"{finalTime:h:mm tt} {finalTzAbbr}";
                    }
                    else
                    {
                        finalTimeDisplay = $"{endTimeCtStr} CT";
                    }
                }
            }
            
            string template = _template ?? "Critical findings were discussed with and acknowledged by {name} at {time} on {date}.";
            var result = template
                .Replace("{name}", titleAndName)
                .Replace("{time}", finalTimeDisplay)
                .Replace("{date}", endDate);
            return diffWarning + result;
        }
        catch (Exception ex)
        {
            Logger.Trace($"NoteFormatter error: {ex.Message}\nRaw: {rawText}");
            return $"[Error parsing critical findings note - check log for details]";
        }
    }
    
    private static string NormalizeTimeStr(string ts)
    {
        ts = ts.ToUpper().Trim();
        ts = Regex.Replace(ts, @"([0-9])(AM|PM)", "$1 $2");
        return ts;
    }
    
    /// <summary>
    /// Maps timezone abbreviations/names from note text to Windows TimeZoneInfo IDs.
    /// </summary>
    private static readonly Dictionary<string, string> TzNameToWindowsId = new(StringComparer.OrdinalIgnoreCase)
    {
        ["east"] = "Eastern Standard Time",
        ["eastern"] = "Eastern Standard Time",
        ["est"] = "Eastern Standard Time",
        ["edt"] = "Eastern Standard Time",
        ["central"] = "Central Standard Time",
        ["cst"] = "Central Standard Time",
        ["cdt"] = "Central Standard Time",
        ["mountain"] = "Mountain Standard Time",
        ["mst"] = "Mountain Standard Time",
        ["mdt"] = "Mountain Standard Time",
        ["pacific"] = "Pacific Standard Time",
        ["pst"] = "Pacific Standard Time",
        ["pdt"] = "Pacific Standard Time",
    };

    /// <summary>
    /// Maps friendly timezone display names to Windows TimeZoneInfo IDs.
    /// </summary>
    private static readonly Dictionary<string, string> DisplayNameToWindowsId = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Eastern Time"] = "Eastern Standard Time",
        ["Central Time"] = "Central Standard Time",
        ["Mountain Time"] = "Mountain Standard Time",
        ["Pacific Time"] = "Pacific Standard Time",
    };

    private static string GetTimezoneDisplay(string rawTz)
    {
        if (string.IsNullOrWhiteSpace(rawTz)) return "Central Time";

        var firstWord = rawTz.Split(' ')[0];
        if (TzNameToWindowsId.TryGetValue(firstWord, out var windowsId))
        {
            return windowsId switch
            {
                "Eastern Standard Time" => "Eastern Time",
                "Mountain Standard Time" => "Mountain Time",
                "Pacific Standard Time" => "Pacific Time",
                _ => "Central Time"
            };
        }

        return "Central Time";
    }

    private static string GetTimezoneAbbreviation(string timezone)
    {
        return timezone switch
        {
            "Eastern Time" => "ET",
            "Central Time" => "CT",
            "Mountain Time" => "MT",
            "Pacific Time" => "PT",
            _ => "CT"
        };
    }

    /// <summary>
    /// Resolves a friendly timezone name (e.g. "Eastern Time") or a target timezone setting
    /// to a TimeZoneInfo instance. Returns null if not recognized.
    /// </summary>
    private static TimeZoneInfo? ResolveTimeZone(string displayName)
    {
        if (string.IsNullOrEmpty(displayName)) return null;

        if (DisplayNameToWindowsId.TryGetValue(displayName, out var windowsId))
        {
            try { return TimeZoneInfo.FindSystemTimeZoneById(windowsId); }
            catch { return null; }
        }

        // Also try treating it directly as a Windows timezone ID
        try { return TimeZoneInfo.FindSystemTimeZoneById(displayName); }
        catch { return null; }
    }

    /// <summary>
    /// Convert time from source timezone to target timezone using DST-aware TimeZoneInfo.
    /// Returns the converted time and the timezone abbreviation to display.
    /// </summary>
    private (DateTime time, string tzAbbr) ConvertToTargetTimezone(DateTime sourceTime, string sourceTimezone)
    {
        // If no target timezone is set, keep the original
        if (string.IsNullOrEmpty(_targetTimezone))
        {
            return (sourceTime, GetTimezoneAbbreviation(sourceTimezone));
        }

        var sourceZone = ResolveTimeZone(sourceTimezone);
        var targetZone = ResolveTimeZone(_targetTimezone);

        if (sourceZone != null && targetZone != null)
        {
            // Treat sourceTime as unspecified so ConvertTime interprets it in the source zone
            var unspecified = DateTime.SpecifyKind(sourceTime, DateTimeKind.Unspecified);
            var convertedTime = TimeZoneInfo.ConvertTime(unspecified, sourceZone, targetZone);
            return (convertedTime, GetTimezoneAbbreviation(_targetTimezone));
        }

        // Fallback: can't resolve zones, return original with target abbreviation
        return (sourceTime, GetTimezoneAbbreviation(_targetTimezone));
    }

    /// <summary>
    /// Try to extract a usable name from a list of candidate segments.
    /// Returns null if no usable name found (all are current doctor or too short).
    /// </summary>
    private string? TryExtractName(List<string> segments)
    {
        foreach (var s in segments)
        {
            var cleaned = s.Trim();
            cleaned = Regex.Replace(cleaned, @"^(?:with|to|at|from|and|@|w/)\s+", "", RegexOptions.IgnoreCase);

            bool isCurrentDoctor = false;
            if (!string.IsNullOrWhiteSpace(_doctorName))
            {
                var nameParts = _doctorName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                isCurrentDoctor = nameParts.Any(part =>
                    part.Length > 2
                    && !TitleWords.Contains(part.TrimEnd('.', ','))
                    && cleaned.Contains(part, StringComparison.OrdinalIgnoreCase));
            }

            if (!isCurrentDoctor && cleaned.Length > 2)
            {
                cleaned = Regex.Split(cleaned, @"\s+(?:to|at|@|w/)\s+|[-;:,]", RegexOptions.IgnoreCase)[0];
                return ToTitleCase(cleaned);
            }
        }
        return null;
    }

    /// <summary>Titles to ignore when matching DoctorName parts against extracted names.</summary>
    private static readonly HashSet<string> TitleWords = new(StringComparer.OrdinalIgnoreCase)
        { "Dr", "Nurse", "NP", "PA", "RN", "MD", "DO", "LPN" };

    private static readonly HashSet<string> PreservedAcronyms = new(StringComparer.OrdinalIgnoreCase)
        { "NP", "MD", "DO", "PA", "RN", "LPN", "BSN", "MSN", "DNP", "PhD" };

    private static string ToTitleCase(string s)
    {
        // Title-case all words except known post-name credential acronyms
        var words = s.Split(' ');
        for (int i = 0; i < words.Length; i++)
        {
            var w = words[i].Trim();
            if (w.Length <= 1) { words[i] = w.ToUpper(); continue; }
            if (PreservedAcronyms.Contains(w.TrimEnd('.', ',')))
                continue; // Known credential acronym — leave it
            words[i] = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(w.ToLower());
        }
        return string.Join(" ", words);
    }
}
