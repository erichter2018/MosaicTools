using System;
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
    
    public NoteFormatter(string doctorName, string? template = null)
    {
        _doctorName = doctorName;
        _template = template;
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
            
            // Look for contact verbs
            var verbPattern = @"(?:Transferred|Connected\s+with|Connected\s+to|Connected|Was\s+connected\s+with|Spoke\s+to|Spoke\s+with|Discussed\s+with)\s+(.+?)(?=\s+(?:at|@|said|confirmed|stated|reported|declined|who|are|is)|\s*[-;:,]|\s+\d{2}/\d{2}/|$)";
            var verbMatch = Regex.Match(rawText, verbPattern, RegexOptions.IgnoreCase);
            if (verbMatch.Success)
            {
                var fullSegment = verbMatch.Groups[1].Value;
                var parts = Regex.Split(fullSegment, @"\s+(?:to|with|w/)\s+", RegexOptions.IgnoreCase);
                segments.AddRange(parts);
            }
            
            // Fallback: Individual title matches (Dr. or Nurse)
            var namePattern = @"((?:Dr\.|Nurse)\s+.+?)(?=\s+(?:at|with|to|w/|@|said|confirmed|stated|reported|declined|who|confirm|are|is)|\s*[-;:,]|\s+(?:Dr\.|Nurse)|\s+\d{2}/\d{2}/|$)";
            var nameMatches = Regex.Matches(rawText, namePattern, RegexOptions.IgnoreCase);
            foreach (Match m in nameMatches)
            {
                segments.Add(m.Groups[1].Value);
            }
            
            string titleAndName = "Dr. / Nurse [Name not found]";
            foreach (var s in segments)
            {
                var cleaned = s.Trim();
                cleaned = Regex.Replace(cleaned, @"^(?:with|to|at|from|and|@|w/)\s+", "", RegexOptions.IgnoreCase);
                
                // Skip if it's the current doctor or too short
                if (!cleaned.Contains(_doctorName, StringComparison.OrdinalIgnoreCase) && cleaned.Length > 2)
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
                var timezoneStr = GetTimezoneDisplay(rawTz);
                
                var tzOffsets = new Dictionary<string, int>
                {
                    ["Eastern Time"] = 1, ["Central Time"] = 0, ["Mountain Time"] = -1, ["Pacific Time"] = -2
                };
                int offset = tzOffsets.GetValueOrDefault(timezoneStr, 0);
                
                var now = DateTime.Now;
                var today = now.Date;
                var yesterday = today.AddDays(-1);
                
                if (DateTime.TryParseExact(textTimeStr, "h:mm tt", 
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var noteTime))
                {
                    var noteHourCt = noteTime.Hour - offset;
                    if (noteHourCt < 0) noteHourCt += 24;
                    else if (noteHourCt >= 24) noteHourCt -= 24;
                    
                    var noteDtToday = today.AddHours(noteHourCt).AddMinutes(noteTime.Minute);
                    
                    if (noteDtToday > now)
                        endDate = yesterday.ToString("MM/dd/yyyy");
                    else
                        endDate = today.ToString("MM/dd/yyyy");
                    
                    endTimeCtStr = textTimeStr;
                    
                    return $"Critical findings were discussed with and acknowledged by {titleAndName} at {textTimeStr} {timezoneStr} on {endDate}.";
                }
            }
            
            // 3. Extract Text Time and Timezone
            string finalTimeDisplay = "N/A";
            var textTimeMatch = Regex.Match(rawText, @"(?:at|@|\.)\s+(\d{1,2}:\d{2}\s*(?:AM|PM))\s*([a-zA-Z\s]+)?", RegexOptions.IgnoreCase);
            
            if (textTimeMatch.Success && dtEnd.HasValue)
            {
                var textTimeStr = NormalizeTimeStr(textTimeMatch.Groups[1].Value);
                var rawTz = textTimeMatch.Groups[2].Success ? textTimeMatch.Groups[2].Value.Trim() : "";
                var timezoneStr = GetTimezoneDisplay(rawTz);
                
                var tzOffsets = new Dictionary<string, int>
                {
                    ["Eastern Time"] = 1, ["Central Time"] = 0, ["Mountain Time"] = -1, ["Pacific Time"] = -2
                };
                int offset = tzOffsets.GetValueOrDefault(timezoneStr, 0);
                
                if (DateTime.TryParseExact($"{endDate} {textTimeStr}", "MM/dd/yyyy h:mm tt",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var dtText))
                {
                    var dtTextCt = dtText.AddHours(-offset);
                    var diffSeconds = (dtTextCt - dtEnd.Value).TotalSeconds;
                    
                    if (diffSeconds > 43200) dtTextCt = dtTextCt.AddDays(-1);
                    else if (diffSeconds < -43200) dtTextCt = dtTextCt.AddDays(1);
                    
                    var timeDiffMin = Math.Abs((dtTextCt - dtEnd.Value).TotalMinutes);
                    
                    if (timeDiffMin <= 10)
                        finalTimeDisplay = $"{textTimeStr} {timezoneStr}";
                    else
                        finalTimeDisplay = $"{textTimeStr} {timezoneStr} (? Diff: {(int)timeDiffMin}m)";
                }
                else
                {
                    finalTimeDisplay = textTimeStr;
                }
            }
            else
            {
                finalTimeDisplay = endTimeCtStr;
            }
            
            string template = _template ?? "Critical findings were discussed with and acknowledged by {name} at {time} on {date}.";
            return template
                .Replace("{name}", titleAndName)
                .Replace("{time}", finalTimeDisplay)
                .Replace("{date}", endDate);
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
    
    private static string GetTimezoneDisplay(string rawTz)
    {
        var firstWord = rawTz.Split(' ')[0].ToLowerInvariant();
        
        string[] validTokens = { "eastern", "central", "mountain", "pacific", "east", "est", "cst", "mst", "pst", "edt", "cdt", "mdt", "pdt" };
        
        if (!validTokens.Contains(firstWord))
            return "Central Time";
        
        if (firstWord is "east" or "eastern" or "est" or "edt") return "Eastern Time";
        if (firstWord is "central" or "cst" or "cdt") return "Central Time";
        if (firstWord is "mountain" or "mst" or "mdt") return "Mountain Time";
        if (firstWord is "pacific" or "pst" or "pdt") return "Pacific Time";
        
        return "Central Time";
    }
    
    private static string ToTitleCase(string s)
    {
        return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLower());
    }
}
