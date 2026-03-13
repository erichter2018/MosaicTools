// [CustomSTT] Shared text processing transforms for STT transcripts.
// Extracted from ActionController so ensemble merger and single-provider paths share the same logic.
using System.Text.RegularExpressions;

namespace MosaicTools.Services;

public static class SttTextProcessor
{
    /// <summary>
    /// Convert spoken punctuation words to symbols, except "colon" (medical term).
    /// Replaces Deepgram's dictation mode so we have full control over which words convert.
    /// </summary>
    public static string ApplySpokenPunctuation(string text, bool expandContractions = true)
    {
        var replacements = new (string pattern, string replacement)[]
        {
            (@"\bnew\s+paragraph\b",  "\n\n"),
            (@"\bnew\s+line\b",       "\n"),
            (@"\bexclamation\s+mark\b[!.]?", "!"),
            (@"\bquestion\s+mark\b[?.]?",   "?"),
            (@"\bperiod\b\.?",        "."),   // consume trailing dot (auto-punctuator may add one)
            (@"\bcomma\b,?",          ","),
            (@"\bsemicolon\b;?",      ";"),
            (@"\bhyphen\b-?",         "-"),
        };

        var result = text;
        foreach (var (pattern, replacement) in replacements)
        {
            result = Regex.Replace(result, pattern, replacement, RegexOptions.IgnoreCase);
        }

        if (expandContractions)
            result = ExpandContractions(result);

        // Clean up spaces before punctuation marks (e.g., "word . next" → "word. next")
        result = Regex.Replace(result, @"\s+([.,;!?])", "$1");

        // Clean up redundant punctuation (e.g., ",." → "." from "something, period.")
        result = Regex.Replace(result, @"[,;]\s*\.", ".");
        result = Regex.Replace(result, @"\.{2,}", ".");

        return result;
    }

    /// <summary>
    /// Expand contractions that STT providers favor over spoken full forms.
    /// </summary>
    public static string ExpandContractions(string text)
    {
        var contractions = new (string pattern, string replacement)[]
        {
            (@"\bthere's\b",   "there is"),
            (@"\bit's\b",      "it is"),
            (@"\bthat's\b",    "that is"),
            (@"\bdoesn't\b",   "does not"),
            (@"\bdon't\b",     "do not"),
            (@"\bdidn't\b",    "did not"),
            (@"\bwasn't\b",    "was not"),
            (@"\bisn't\b",     "is not"),
            (@"\baren't\b",    "are not"),
            (@"\bweren't\b",   "were not"),
            (@"\bwouldn't\b",  "would not"),
            (@"\bcouldn't\b",  "could not"),
            (@"\bshouldn't\b", "should not"),
            (@"\bhasn't\b",    "has not"),
            (@"\bhaven't\b",   "have not"),
            (@"\bhadn't\b",    "had not"),
            (@"\bcan't\b",     "cannot"),
            (@"\bwon't\b",     "will not"),
        };
        foreach (var (pattern, replacement) in contractions)
        {
            text = Regex.Replace(
                text, pattern,
                m => char.IsUpper(m.Value[0])
                    ? char.ToUpper(replacement[0]) + replacement[1..]
                    : replacement,
                RegexOptions.IgnoreCase);
        }
        return text;
    }

    // ── Radiology transcript cleanup ──────────────────────────────────────

    private static readonly Dictionary<string, string> NumberWords = new(StringComparer.OrdinalIgnoreCase)
    {
        ["zero"] = "0", ["one"] = "1", ["two"] = "2", ["three"] = "3",
        ["four"] = "4", ["five"] = "5", ["six"] = "6", ["seven"] = "7",
        ["eight"] = "8", ["nine"] = "9", ["ten"] = "10", ["eleven"] = "11",
        ["twelve"] = "12", ["thirteen"] = "13", ["fourteen"] = "14",
        ["fifteen"] = "15", ["sixteen"] = "16", ["seventeen"] = "17",
        ["eighteen"] = "18", ["nineteen"] = "19", ["twenty"] = "20",
        ["thirty"] = "30", ["forty"] = "40", ["fifty"] = "50",
        ["sixty"] = "60", ["seventy"] = "70", ["eighty"] = "80", ["ninety"] = "90",
    };

    private const string NumPat = @"(?<n>\d+|zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)";

    private static string NumReplace(Match m, string groupName = "n")
    {
        var val = m.Groups[groupName].Value;
        return NumberWords.TryGetValue(val, out var d) ? d : val;
    }

    private static readonly Dictionary<string, int> OrdinalWords = new(StringComparer.OrdinalIgnoreCase)
    {
        ["first"] = 1, ["second"] = 2, ["third"] = 3, ["fourth"] = 4, ["fifth"] = 5,
        ["sixth"] = 6, ["seventh"] = 7, ["eighth"] = 8, ["ninth"] = 9, ["tenth"] = 10,
        ["eleventh"] = 11, ["twelfth"] = 12, ["thirteenth"] = 13, ["fourteenth"] = 14,
        ["fifteenth"] = 15, ["sixteenth"] = 16, ["seventeenth"] = 17, ["eighteenth"] = 18,
        ["nineteenth"] = 19, ["twentieth"] = 20, ["thirtieth"] = 30,
    };

    private static int TryParseSpokenInt(string text)
    {
        text = text.Trim();
        if (string.IsNullOrEmpty(text)) return -1;

        var stripped = Regex.Replace(text, @"(?:st|nd|rd|th)$", "", RegexOptions.IgnoreCase);
        if (int.TryParse(stripped, out var n)) return n;

        if (NumberWords.TryGetValue(text, out var nw) && int.TryParse(nw, out var nwi)) return nwi;
        if (OrdinalWords.TryGetValue(text, out var ow)) return ow;

        var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 2)
        {
            if (string.Equals(parts[0], "oh", StringComparison.OrdinalIgnoreCase))
            {
                var r = TryParseSpokenInt(parts[1]);
                if (r >= 0 && r <= 9) return r;
            }

            int tens = parts[0].ToLower() switch
            {
                "twenty" => 20, "thirty" => 30, "forty" => 40, "fifty" => 50,
                "sixty" => 60, "seventy" => 70, "eighty" => 80, "ninety" => 90, _ => -1
            };
            if (tens > 0)
            {
                if (OrdinalWords.TryGetValue(parts[1], out var o) && o >= 1 && o <= 9) return tens + o;
                if (NumberWords.TryGetValue(parts[1], out var c) && int.TryParse(c, out var ci) && ci >= 1 && ci <= 9) return tens + ci;
            }
        }

        return -1;
    }

    private static int TryParseSpokenYear(string text)
    {
        text = text.Trim();
        if (string.IsNullOrEmpty(text)) return -1;
        if (int.TryParse(text, out var n) && n >= 1900 && n <= 2099) return n;

        var words = text.ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length < 1) return -1;

        if (words.Length >= 2 && words[0] == "two" && words[1] == "thousand")
        {
            if (words.Length == 2) return 2000;
            int start = 2;
            if (start < words.Length && words[start] == "and") start++;
            if (start >= words.Length) return 2000;
            var offset = TryParseSpokenInt(string.Join(" ", words[start..]));
            if (offset >= 1 && offset <= 99) return 2000 + offset;
            return -1;
        }

        if (words.Length >= 2 && words[0] == "twenty" && words[1] == "twenty")
        {
            if (words.Length == 2) return 2020;
            if (words.Length == 3)
            {
                var ones = TryParseSpokenInt(words[2]);
                if (ones >= 1 && ones <= 9) return 2020 + ones;
            }
            return -1;
        }

        if (words.Length == 2 && words[0] == "twenty")
        {
            var second = TryParseSpokenInt(words[1]);
            if (second >= 10 && second <= 19) return 2000 + second;
        }

        if (words.Length >= 2 && words[0] == "nineteen")
        {
            var offset = TryParseSpokenInt(string.Join(" ", words[1..]));
            if (offset >= 0 && offset <= 99) return 1900 + offset;
            return -1;
        }

        var bare = TryParseSpokenInt(text);
        if (bare >= 0 && bare <= 99) return 2000 + bare;

        return -1;
    }

    /// <summary>
    /// Convert spoken radiology shorthand into standard written forms.
    /// Spine levels, decimals, units, dimensions, dates.
    /// </summary>
    public static string ApplyRadiologyCleanup(string text)
    {
        const RegexOptions IC = RegexOptions.IgnoreCase;

        var result = text;

        // ── Spine levels ──
        result = Regex.Replace(result,
            @"\b(?<seg>[CTLS])\s*" + NumPat.Replace("<n>", "<n1>") +
            @"(?:\s*(?:-|through)?\s*(?<seg2>[CTLS])?\s*" + NumPat.Replace("<n>", "<n2>") + @")?\b",
            m =>
            {
                var seg = m.Groups["seg"].Value.ToUpper();
                var n1 = m.Groups["n1"].Value;
                if (NumberWords.TryGetValue(n1, out var d1)) n1 = d1;

                if (!m.Groups["n2"].Success)
                    return seg + n1;

                var n2 = m.Groups["n2"].Value;
                if (NumberWords.TryGetValue(n2, out var d2)) n2 = d2;

                var seg2 = m.Groups["seg2"].Success ? m.Groups["seg2"].Value.ToUpper() : "";
                return seg + n1 + "-" + seg2 + n2;
            }, IC);

        // Fix concatenated digit spine levels: C34 → C3-4, T12 stays T12
        result = Regex.Replace(result,
            @"\b([CTLS])(\d)(\d)\b",
            m =>
            {
                var seg = m.Groups[1].Value.ToUpper();
                var d1 = m.Groups[2].Value;
                var d2 = m.Groups[3].Value;
                var combined = int.Parse(d1 + d2);
                var max = seg[0] switch { 'C' => 7, 'T' => 12, 'L' => 5, 'S' => 5, _ => 0 };
                if (combined <= max) return seg + d1 + d2;
                return seg + d1 + "-" + d2;
            }, IC);

        // ── "point" between numbers → decimal ──
        result = Regex.Replace(result,
            NumPat.Replace("<n>", "<n1>") + @"\s+point\s+" + NumPat.Replace("<n>", "<n2>"),
            m =>
            {
                var n1 = m.Groups["n1"].Value;
                if (NumberWords.TryGetValue(n1, out var d1)) n1 = d1;
                var n2 = m.Groups["n2"].Value;
                if (NumberWords.TryGetValue(n2, out var d2)) n2 = d2;
                return n1 + "." + n2;
            }, IC);

        // ── Unit words → abbreviations (only after a number) ──
        var unitMap = new (string pattern, string abbrev)[]
        {
            ("centimeters?", "cm"), ("millimeters?", "mm"), ("meters?", "m"),
        };
        foreach (var (unitPat, abbrev) in unitMap)
        {
            result = Regex.Replace(result,
                NumPat + @"\s+" + unitPat + @"\b",
                m => NumReplace(m) + " " + abbrev, IC);
        }

        // ── Dimensions: "by" between numbers → "x" ──
        result = Regex.Replace(result,
            @"(\d+(?:\.\d+)?)\s+by\s+(\d+(?:\.\d+)?)", "$1 x $2", IC);

        result = Regex.Replace(result,
            @"(\d+(?:\.\d+)?)\s+by\s*$", "$1 x", IC);
        result = Regex.Replace(result,
            @"^\s*by\s+(\d+(?:\.\d+)?)", "x $1", IC);

        // ── Dates ──
        var onesW = @"one|two|three|four|five|six|seven|eight|nine";
        var teensW = @"ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen";
        var tensW = @"twenty|thirty";
        var ordOnesW = @"first|second|third|fourth|fifth|sixth|seventh|eighth|ninth";
        var ordTeensW = @"tenth|eleventh|twelfth|thirteenth|fourteenth|fifteenth|sixteenth|seventeenth|eighteenth|nineteenth";

        var dayPat =
            $@"(?:(?:{tensW})\s+(?:{ordOnesW})" +
            $@"|(?:{ordOnesW}|{ordTeensW}|twentieth|thirtieth)" +
            $@"|(?:{tensW})\s+(?:{onesW})" +
            $@"|(?:{onesW}|{teensW}|twenty|thirty)" +
            @"|\d{1,2}(?:st|nd|rd|th)?)";

        var year2020 = $@"twenty\s+twenty(?:\s+(?:{onesW}))?";
        var year2010 = $@"twenty\s+(?:{teensW})";
        var year19xx = $@"nineteen\s+(?:(?:ninety|eighty|seventy|sixty|fifty|forty|thirty|twenty|ten)(?:\s+(?:{onesW}))?|(?:{teensW})|(?:{onesW}))";
        var year2000 = $@"two\s+thousand(?:\s+(?:and\s+)?(?:twenty(?:\s+(?:{onesW}))?|(?:{teensW})|(?:{onesW})))?";
        var yearDigit = @"20[0-9]\d|19\d\d";
        var yearPat = $@"(?:{year2020}|{year2010}|{year19xx}|{year2000}|{yearDigit})";
        var yearOrBare = $@"(?:{yearPat}|(?:{tensW})\s+(?:{onesW})|(?:{teensW})|(?:{onesW})|\d{{1,2}})";

        var monthNamePat = @"(?:january|february|march|april|may|june|july|august|september|october|november|december)";
        var numMonthPat = $@"(?:oh\s+(?:{onesW})|{onesW}|ten|eleven|twelve|\d{{1,2}})";

        // 1. Month-name + day + year
        result = Regex.Replace(result,
            $@"\b({monthNamePat})\s+({dayPat})\s*,?\s+({yearPat})\b",
            m =>
            {
                var month = char.ToUpper(m.Groups[1].Value[0]) + m.Groups[1].Value[1..].ToLower();
                var day = TryParseSpokenInt(m.Groups[2].Value);
                var year = TryParseSpokenYear(m.Groups[3].Value);
                if (day >= 1 && day <= 31 && year >= 1900)
                    return $"{month} {day}, {year}";
                return m.Value;
            }, IC);

        // 2. Month-name + year (no day)
        result = Regex.Replace(result,
            $@"\b({monthNamePat})\s+({yearPat})\b",
            m =>
            {
                var month = char.ToUpper(m.Groups[1].Value[0]) + m.Groups[1].Value[1..].ToLower();
                var year = TryParseSpokenYear(m.Groups[2].Value);
                if (year >= 1900)
                    return $"{month} {year}";
                return m.Value;
            }, IC);

        // 3. Month-name + day (no year)
        result = Regex.Replace(result,
            $@"\b({monthNamePat})\s+({dayPat})\b",
            m =>
            {
                var month = char.ToUpper(m.Groups[1].Value[0]) + m.Groups[1].Value[1..].ToLower();
                var day = TryParseSpokenInt(m.Groups[2].Value);
                if (day >= 1 && day <= 31)
                    return $"{month} {day}";
                return m.Value;
            }, IC);

        // 4. Numeric date with 4-number year anchor
        result = Regex.Replace(result,
            $@"\b({numMonthPat})\s+({dayPat})\s+({yearPat})\b",
            m =>
            {
                var month = TryParseSpokenInt(m.Groups[1].Value);
                var day = TryParseSpokenInt(m.Groups[2].Value);
                var year = TryParseSpokenYear(m.Groups[3].Value);
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 1900)
                    return $"{month}/{day}/{year}";
                return m.Value;
            }, IC);

        // 5. Numeric date with 3-number year (assumes 20xx)
        result = Regex.Replace(result,
            $@"\b({numMonthPat})\s+({dayPat})\s+({yearOrBare})\b",
            m =>
            {
                var month = TryParseSpokenInt(m.Groups[1].Value);
                var day = TryParseSpokenInt(m.Groups[2].Value);
                var year = TryParseSpokenYear(m.Groups[3].Value);
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 2000 && year <= 2099)
                    return $"{month}/{day}/{year}";
                return m.Value;
            }, IC);

        return result;
    }

    /// <summary>
    /// Apply user-defined word replacements.
    /// Uses word-boundary matching with case preservation.
    /// </summary>
    public static string ApplyCustomReplacements(string text, List<SttReplacementEntry> replacements)
    {
        if (replacements.Count == 0) return text;

        foreach (var entry in replacements)
        {
            if (!entry.Enabled || string.IsNullOrWhiteSpace(entry.Find)) continue;
            var pattern = @"\b" + Regex.Escape(entry.Find) + @"\b";
            var replacement = entry.Replace ?? "";
            text = Regex.Replace(
                text, pattern,
                m => PreserveCaseReplace(m.Value, replacement),
                RegexOptions.IgnoreCase);
        }
        return text;
    }

    /// <summary>
    /// Insert a newline after each sentence-ending period.
    /// Skips periods in decimal numbers (e.g. "1.2", "3.5").
    /// </summary>
    public static string InsertNewlineAfterSentences(string text)
    {
        return Regex.Replace(text, @"(?<!\d)\.(\s+)", ".\n");
    }

    /// <summary>
    /// Replace text while preserving the case pattern of the original match.
    /// ALL CAPS → ALL CAPS, Title Case → Title Case, otherwise use replacement as-is.
    /// </summary>
    public static string PreserveCaseReplace(string original, string replacement)
    {
        if (replacement.Length == 0) return replacement;
        if (original.Length > 0 && original == original.ToUpperInvariant())
            return replacement.ToUpperInvariant();
        if (original.Length > 0 && char.IsUpper(original[0]))
            return char.ToUpper(replacement[0]) + replacement[1..];
        return replacement;
    }

    /// <summary>
    /// Strip auto-punctuation from provider output so it looks like manual dictation.
    /// Used when SttAutoPunctuateFinalReport is on but text is going to the transcript editor.
    /// Lowercase first char and remove trailing sentence-ending period.
    /// </summary>
    public static string StripAutoPunctuation(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        // Drop standalone punctuation tokens (Soniox can emit "." as a separate final result)
        var trimmed = text.Trim();
        if (trimmed.Length > 0 && trimmed.All(c => char.IsPunctuation(c)))
            return "";

        // Remove all sentence-ending punctuation (periods, commas, semicolons, etc.)
        // but preserve decimal points between digits (e.g., "2.5")
        text = Regex.Replace(text, @"(?<!\d)[.,;!?](?!\d)", "");

        // Lowercase first char (unless it's an acronym — check if second char is also upper)
        if (text.Length > 0 && char.IsUpper(text[0]) && !(text.Length > 1 && char.IsUpper(text[1])))
            text = char.ToLower(text[0]) + text[1..];

        return text.Trim();
    }

    /// <summary>
    /// Apply all enabled STT text transforms in the correct order.
    /// Convenience method for both single-provider paste and ensemble merger.
    /// </summary>
    public static string ProcessTranscript(string rawText, Configuration config)
    {
        // Always replace spoken punctuation words ("period" → ".", "comma" → ",", etc.)
        // Even with auto-punctuate, providers sometimes output these as literal words.
        var transcript = ApplySpokenPunctuation(rawText, config.SttExpandContractions).Trim();

        if (config.SttRadiologyCleanup)
            transcript = ApplyRadiologyCleanup(transcript);

        // In manual punctuation mode, lowercase first char (providers don't capitalize)
        // Skip when final-report-only punctuation is on — stripping happens per-editor at paste time
        if (!config.SttAutoPunctuate && !config.SttAutoPunctuateFinalReport
            && transcript.Length > 0 && !(transcript.Length > 1 && char.IsDigit(transcript[1])))
            transcript = char.ToLower(transcript[0]) + transcript[1..];

        transcript = ApplyCustomReplacements(transcript, config.SttCustomReplacements);

        if (config.SttNewlineAfterSentence)
            transcript = InsertNewlineAfterSentences(transcript);

        return transcript;
    }
}
