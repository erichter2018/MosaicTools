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
            var namePattern = @"((?:Dr\.?|Nurse|NP|PA|RN|MD)\s+.+?)(?=\s+(?:at|with|to|w/|@|said|confirmed|stated|reported|declined|who|confirm|are|is|and|&|connected|contacted|spoke|discussed|transferred|called|notified|informed|reached|paged)\b|\s*[-;:,]|\s+(?:Dr\.?|Nurse|NP|PA|RN|MD)|\s+\d{2}/\d{2}/|$)";
            var nameMatches = Regex.Matches(rawText, namePattern, RegexOptions.IgnoreCase);
            foreach (Match m in nameMatches)
            {
                segments.Add(m.Groups[1].Value);
            }

            // Fallback: verb-based extraction (only if no title-based match found)
            // Negative lookbehind prevents matching passive "to be connected" / "been connected"
            if (segments.Count == 0)
            {
                var verbPattern = @"(?<!\bto\s+be\s+)(?<!\bbeen\s+)(?:Transferred|Connected\s+with|Connected\s+to|Connected|Was\s+connected\s+with|Spoke\s+to|Spoke\s+with|Discussed\s+with)\s+(.+?)(?=\s+(?:at|@|said|confirmed|stated|reported|declined|who|are|is)|\s*[-;:,]|\s+\d{2}/\d{2}/|$)";
                var verbMatch = Regex.Match(rawText, verbPattern, RegexOptions.IgnoreCase);
                if (verbMatch.Success)
                {
                    var fullSegment = verbMatch.Groups[1].Value;
                    var parts = Regex.Split(fullSegment, @"\s+(?:to|with|w/|and|&)\s+", RegexOptions.IgnoreCase);
                    segments.AddRange(parts);
                }
            }
            
            string titleAndName = "Dr. / Nurse [Name not found]";
            foreach (var s in segments)
            {
                var cleaned = s.Trim();
                cleaned = Regex.Replace(cleaned, @"^(?:with|to|at|from|and|@|w/)\s+", "", RegexOptions.IgnoreCase);

                // Skip if it's the current doctor or too short
                // Check both "Firstname Lastname" and "Lastname, Firstname" formats
                bool isCurrentDoctor = false;
                if (!string.IsNullOrWhiteSpace(_doctorName))
                {
                    var nameParts = _doctorName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    // If any part of the doctor's name appears in this segment, skip it
                    // Exclude common titles (Dr, Dr., Nurse, etc.) so they don't match every name
                    isCurrentDoctor = nameParts.Any(part =>
                        part.Length > 2
                        && !TitleWords.Contains(part.TrimEnd('.', ','))
                        && cleaned.Contains(part, StringComparison.OrdinalIgnoreCase));
                }

                if (!isCurrentDoctor && cleaned.Length > 2)
                {
                    cleaned = Regex.Split(cleaned, @"\s+(?:to|at|@|w/)\s+|[-;:,]", RegexOptions.IgnoreCase)[0];
                    titleAndName = ToTitleCase(cleaned);
                    break;
                }
            }
            
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

                var tzOffsets = new Dictionary<string, int>
                {
                    ["Eastern Time"] = 1, ["Central Time"] = 0, ["Mountain Time"] = -1, ["Pacific Time"] = -2
                };
                int sourceOffset = tzOffsets.GetValueOrDefault(sourceTimezoneStr, 0);

                var now = DateTime.Now;
                var today = now.Date;
                var yesterday = today.AddDays(-1);

                if (DateTime.TryParseExact(textTimeStr, "h:mm tt",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var noteTime))
                {
                    var noteHourCt = noteTime.Hour - sourceOffset;
                    if (noteHourCt < 0) noteHourCt += 24;
                    else if (noteHourCt >= 24) noteHourCt -= 24;

                    var noteDtToday = today.AddHours(noteHourCt).AddMinutes(noteTime.Minute);

                    if (noteDtToday > now)
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
                int sourceOffsetFromCentral = GetTzHoursFromCentral(rawTz);
                var sourceTimezoneStr = GetTimezoneDisplay(rawTz);

                if (DateTime.TryParseExact($"{endDate} {textTimeStr}", "MM/dd/yyyy h:mm tt",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var dtText))
                {
                    // Convert discovered text time to Central (CDT/CST)
                    // If Mountain (-1), we subtract -1 (add 1 hour) to get Central.
                    var dtTextCt = dtText.AddHours(-sourceOffsetFromCentral);

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
            return $"Error parsing: {ex.Message}\nRaw: {rawText}";
        }
    }
    
    private static string NormalizeTimeStr(string ts)
    {
        ts = ts.ToUpper().Trim();
        ts = Regex.Replace(ts, @"([0-9])(AM|PM)", "$1 $2");
        return ts;
    }
    
    private static int GetTzHoursFromCentral(string rawTz)
    {
        if (string.IsNullOrWhiteSpace(rawTz)) return 0;
        
        var firstWord = rawTz.Split(' ')[0].ToLowerInvariant();
        
        if (firstWord is "east" or "eastern" or "est" or "edt") return 1;
        if (firstWord is "central" or "cst" or "cdt") return 0;
        if (firstWord is "mountain" or "mst" or "mdt") return -1;
        if (firstWord is "pacific" or "pst" or "pdt") return -2;
        
        return 0;
    }
    
    private static string GetTimezoneDisplay(string rawTz)
    {
        int offset = GetTzHoursFromCentral(rawTz);
        return offset switch
        {
            1 => "Eastern Time",
            -1 => "Mountain Time",
            -2 => "Pacific Time",
            _ => "Central Time"
        };
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
    /// Convert time from source timezone to target timezone (or keep original if no target set).
    /// Returns the converted time and the timezone abbreviation to display.
    /// </summary>
    private (DateTime time, string tzAbbr) ConvertToTargetTimezone(DateTime sourceTime, string sourceTimezone)
    {
        // If no target timezone is set, keep the original
        if (string.IsNullOrEmpty(_targetTimezone))
        {
            return (sourceTime, GetTimezoneAbbreviation(sourceTimezone));
        }

        // Get offsets from Central for source and target
        int sourceOffsetFromCentral = sourceTimezone switch
        {
            "Eastern Time" => 1,
            "Central Time" => 0,
            "Mountain Time" => -1,
            "Pacific Time" => -2,
            _ => 0  // Default to Central if unknown
        };

        int targetOffsetFromCentral = _targetTimezone switch
        {
            "Eastern Time" => 1,
            "Central Time" => 0,
            "Mountain Time" => -1,
            "Pacific Time" => -2,
            _ => 0  // Default to Central if unknown
        };

        // Convert: source → Central → target
        // Offset difference: how many hours to add to convert source to target
        int hoursDiff = targetOffsetFromCentral - sourceOffsetFromCentral;
        var convertedTime = sourceTime.AddHours(hoursDiff);

        return (convertedTime, GetTimezoneAbbreviation(_targetTimezone));
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
            if (w.Length <= 1) continue;
            if (PreservedAcronyms.Contains(w.TrimEnd('.', ',')))
                continue; // Known credential acronym — leave it
            words[i] = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(w.ToLower());
        }
        return string.Join(" ", words);
    }
}
