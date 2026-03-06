using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MosaicTools.Services;

/// <summary>
/// A single consistency mismatch found between report sections or study description.
/// </summary>
public record ConsistencyResult(
    string Key,           // "meas_calculus_4mm_vs_5mm" or "lat_kidney_left_vs_right"
    string DisplayName,   // "calculus: 4mm→5mm" or "kidney: left→right"
    string CheckType,     // "measurement" | "laterality" | "study_laterality"
    string[] SearchTerms  // For CDP text flashing
);

/// <summary>
/// Checks measurement consistency (FINDINGS vs IMPRESSION), laterality consistency
/// (FINDINGS vs IMPRESSION), and study description laterality against dictated report.
/// </summary>
public static class ConsistencyChecker
{
    // Measurement regex: captures number and unit (mm or cm)
    // Handles ranges like "3-4 mm" and dimensions like "3 x 4 cm"
    private static readonly Regex MeasurementRegex = new(
        @"(\d+(?:\.\d+)?)\s*(?:[-x×]\s*(\d+(?:\.\d+)?)\s*)*\s*(mm|cm)\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Laterality terms with word boundary enforcement
    // Words (left/right/LT/RT) are case-insensitive to handle ALL CAPS study descriptions.
    // Single-letter L/R only match uppercase and must NOT be followed by a digit (avoids L4, L5 spine levels).
    private static readonly Regex LeftRegex = new(
        @"\b(?:left|LT)\b|(?<![a-zA-Z])L(?![a-zA-Z0-9])",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex RightRegex = new(
        @"\b(?:right|RT)\b|(?<![a-zA-Z])R(?![a-zA-Z0-9])",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex BilateralRegex = new(
        @"\b(?:bilateral|bilaterally)\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Anatomy synonym map for measurement matching
    private static readonly Dictionary<string, string> AnatomySynonyms = new(StringComparer.OrdinalIgnoreCase)
    {
        { "renal", "kidney" }, { "kidneys", "kidney" },
        { "hepatic", "liver" },
        { "pulmonary", "lung" }, { "lungs", "lung" },
        { "calculus", "stone" }, { "calculi", "stone" }, { "stones", "stone" }, { "urolithiasis", "stone" },
        { "nodule", "nodule" }, { "nodules", "nodule" },
        { "mass", "mass" }, { "masses", "mass" },
        { "cyst", "cyst" }, { "cysts", "cyst" },
        { "lesion", "lesion" }, { "lesions", "lesion" },
        { "lymph", "lymph" }, { "node", "lymph" }, { "nodes", "lymph" },
        { "aorta", "aorta" }, { "aortic", "aorta" },
        { "aneurysm", "aneurysm" }, { "aneurysmal", "aneurysm" },
        { "gallstone", "gallstone" }, { "gallstones", "gallstone" }, { "cholelithiasis", "gallstone" },
        { "adrenal", "adrenal" }, { "adrenals", "adrenal" },
        { "thyroid", "thyroid" },
        { "pancreas", "pancreas" }, { "pancreatic", "pancreas" },
        { "spleen", "spleen" }, { "splenic", "spleen" },
        { "ovary", "ovary" }, { "ovarian", "ovary" }, { "ovaries", "ovary" },
    };

    // Laterality anatomy synonyms — maps variants to canonical form
    private static readonly Dictionary<string, string> LateralitySynonyms = new(StringComparer.OrdinalIgnoreCase)
    {
        { "renal", "kidney" }, { "kidneys", "kidney" },
        { "hepatic", "liver" },
        { "pulmonary", "lung" }, { "lungs", "lung" },
        { "adrenal", "adrenal" }, { "adrenals", "adrenal" },
        { "ovary", "ovary" }, { "ovarian", "ovary" }, { "ovaries", "ovary" },
        { "pleural", "pleural" },
        { "inguinal", "inguinal" },
        { "femoral", "femoral" },
        { "popliteal", "popliteal" },
        { "subclavian", "subclavian" },
        { "carotid", "carotid" },
        { "humeral", "humerus" },
        { "radial", "radius" },
        { "ulnar", "ulna" },
        { "tibial", "tibia" },
        { "fibular", "fibula" },
    };

    // Negation phrases (reused from FIM checker pattern)
    private static readonly string[] NegationPhrases =
    {
        "no findings to suggest", "no evidence of", "no convincing", "no definite",
        "no significant", "no signs of", "no findings of", "no acute",
        "negative for", "absence of", "rules out", "ruled out",
        "without", "not ", "no ",
    };

    // Body parts that are inherently unilateral (for study description check)
    private static readonly HashSet<string> UnilateralBodyParts = new(StringComparer.OrdinalIgnoreCase)
    {
        "hand", "wrist", "elbow", "shoulder", "hip", "knee", "ankle", "foot",
        "finger", "toe", "thumb",
        "humerus", "radius", "ulna", "femur", "tibia", "fibula", "clavicle",
        "forearm", "arm", "leg", "thigh", "shin",
        "scapula", "patella", "calcaneus", "talus",
    };

    // Keywords in study description that suppress laterality checks (bilateral/midline anatomy)
    private static readonly HashSet<string> BilateralSuppressors = new(StringComparer.OrdinalIgnoreCase)
    {
        "pelvis", "chest", "abdomen", "spine", "brain", "head", "neck",
        "bilateral", "both", "ap", "cervical", "thoracic", "lumbar", "sacral",
        "face", "facial", "sinus", "orbit", "skull", "sacrum", "coccyx",
    };

    // Anatomy words to extract as context around measurements
    private static readonly Regex AnatomyWordRegex = new(
        @"\b([a-z]{3,})\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static List<ConsistencyResult> Check(
        string? findings, string? impression, string? studyDescription)
    {
        var results = new List<ConsistencyResult>();

        if (!string.IsNullOrWhiteSpace(findings) && !string.IsNullOrWhiteSpace(impression))
        {
            CheckMeasurements(findings, impression, results);
            CheckLaterality(findings, impression, results);
        }

        if (!string.IsNullOrWhiteSpace(studyDescription))
        {
            var combinedReport = (findings ?? "") + "\n" + (impression ?? "");
            if (!string.IsNullOrWhiteSpace(combinedReport))
                CheckStudyLaterality(studyDescription, combinedReport, results);
        }

        return results;
    }

    #region Measurement Consistency

    private static void CheckMeasurements(string findings, string impression, List<ConsistencyResult> results)
    {
        var findingsMeasurements = ExtractMeasurements(findings);
        var impressionMeasurements = ExtractMeasurements(impression);

        // For each anatomy in impression that has a measurement, check against findings
        foreach (var (anatomy, impMeasures) in impressionMeasurements)
        {
            if (!findingsMeasurements.TryGetValue(anatomy, out var findMeasures))
                continue;

            // Compare: flag if measurements differ
            foreach (var impM in impMeasures)
            {
                foreach (var findM in findMeasures)
                {
                    // Both have explicit measurements — compare in mm
                    if (Math.Abs(impM.ValueMm - findM.ValueMm) > 0.5)
                    {
                        var key = $"meas_{anatomy}_{FormatMm(findM.ValueMm)}_vs_{FormatMm(impM.ValueMm)}";
                        if (results.Any(r => r.Key == key)) continue;

                        var display = $"{anatomy}: {FormatDisplay(findM)}→{FormatDisplay(impM)}";
                        var searchTerms = new[] { findM.OriginalText, impM.OriginalText };
                        results.Add(new ConsistencyResult(key, display, "measurement", searchTerms));
                        break; // One flag per anatomy is enough
                    }
                }
                if (results.Any(r => r.Key.StartsWith($"meas_{anatomy}_"))) break;
            }
        }
    }

    private record MeasurementInfo(double ValueMm, string OriginalText, string Unit);

    private static Dictionary<string, List<MeasurementInfo>> ExtractMeasurements(string text)
    {
        var result = new Dictionary<string, List<MeasurementInfo>>(StringComparer.OrdinalIgnoreCase);
        var textLower = text.ToLowerInvariant();
        var matches = MeasurementRegex.Matches(textLower);

        foreach (Match match in matches)
        {
            // Parse all numbers in the match (handles "3 x 4 cm" — we take the largest)
            var numbers = new List<double>();
            if (double.TryParse(match.Groups[1].Value, out var n1))
                numbers.Add(n1);
            if (match.Groups[2].Success && double.TryParse(match.Groups[2].Value, out var n2))
                numbers.Add(n2);

            if (numbers.Count == 0) continue;

            var maxValue = numbers.Max();
            var unit = match.Groups[3].Value.ToLowerInvariant();
            var valueMm = unit == "cm" ? maxValue * 10 : maxValue;
            var originalText = match.Value.Trim();

            // Extract anatomy context: look at 1-3 words before and after the measurement
            var anatomy = ExtractAnatomyContext(textLower, match.Index, match.Length);
            if (string.IsNullOrEmpty(anatomy)) continue;

            // Normalize anatomy
            anatomy = NormalizeAnatomy(anatomy);

            if (!result.ContainsKey(anatomy))
                result[anatomy] = new List<MeasurementInfo>();
            result[anatomy].Add(new MeasurementInfo(valueMm, originalText, unit));
        }

        return result;
    }

    private static string? ExtractAnatomyContext(string text, int matchStart, int matchLength)
    {
        // Look at words before the measurement (up to 40 chars back)
        int lookbackStart = Math.Max(0, matchStart - 40);
        var before = text.Substring(lookbackStart, matchStart - lookbackStart);

        // Look at words after the measurement (up to 40 chars forward)
        int afterStart = matchStart + matchLength;
        int lookforwardEnd = Math.Min(text.Length, afterStart + 40);
        var after = text.Substring(afterStart, lookforwardEnd - afterStart);

        // Get anatomy words from context (prefer before, then after)
        var beforeWords = AnatomyWordRegex.Matches(before)
            .Cast<Match>()
            .Select(m => m.Groups[1].Value)
            .Where(IsAnatomyWord)
            .ToList();

        var afterWords = AnatomyWordRegex.Matches(after)
            .Cast<Match>()
            .Select(m => m.Groups[1].Value)
            .Where(IsAnatomyWord)
            .Take(2)
            .ToList();

        // Take the last anatomy word before measurement, or first after
        if (beforeWords.Count > 0)
            return beforeWords.Last();
        if (afterWords.Count > 0)
            return afterWords.First();

        return null;
    }

    private static bool IsAnatomyWord(string word)
    {
        if (word.Length < 3) return false;
        // Skip common non-anatomy words
        if (_nonAnatomyWords.Contains(word)) return false;
        return true;
    }

    private static readonly HashSet<string> _nonAnatomyWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "the", "and", "with", "from", "into", "for", "was", "are", "has", "had",
        "been", "being", "this", "that", "which", "what", "where", "when", "how",
        "measures", "measuring", "approximately", "about", "around", "nearly",
        "largest", "smallest", "stable", "unchanged", "previously", "prior",
        "new", "increased", "decreased", "compared", "interval", "since",
        "again", "also", "noted", "seen", "identified", "demonstrated", "compatible",
        "consistent", "suggestive", "suspicious", "likely", "possibly", "probably",
        "mild", "moderate", "severe", "minimal", "small", "large", "tiny",
        "multiple", "several", "few", "numerous",
    };

    private static string NormalizeAnatomy(string anatomy)
    {
        if (AnatomySynonyms.TryGetValue(anatomy, out var canonical))
            return canonical;
        return anatomy.ToLowerInvariant();
    }

    private static string FormatMm(double mm)
    {
        if (mm >= 10 && mm % 10 == 0)
            return $"{mm / 10:0.#}cm";
        return $"{mm:0.#}mm";
    }

    private static string FormatDisplay(MeasurementInfo m)
    {
        return m.OriginalText;
    }

    #endregion

    #region Laterality Consistency (FINDINGS vs IMPRESSION)

    private static void CheckLaterality(string findings, string impression, List<ConsistencyResult> results)
    {
        var findingsLats = ExtractLateralizedAnatomy(findings);
        var impressionLats = ExtractLateralizedAnatomy(impression);

        foreach (var (anatomy, impSides) in impressionLats)
        {
            if (!findingsLats.TryGetValue(anatomy, out var findSides))
                continue;

            // bilateral is compatible with anything
            if (impSides.Contains("bilateral") || findSides.Contains("bilateral"))
                continue;

            // Check for side mismatch
            bool impLeft = impSides.Contains("left");
            bool impRight = impSides.Contains("right");
            bool findLeft = findSides.Contains("left");
            bool findRight = findSides.Contains("right");

            // Flag if impression says one side and findings says the other (without also saying both)
            if (impLeft && !impRight && findRight && !findLeft)
            {
                var key = $"lat_{anatomy}_right_vs_left";
                if (!results.Any(r => r.Key == key))
                {
                    var display = $"{anatomy}: right→left";
                    results.Add(new ConsistencyResult(key, display, "laterality",
                        new[] { "right", "left", anatomy }));
                }
            }
            else if (impRight && !impLeft && findLeft && !findRight)
            {
                var key = $"lat_{anatomy}_left_vs_right";
                if (!results.Any(r => r.Key == key))
                {
                    var display = $"{anatomy}: left→right";
                    results.Add(new ConsistencyResult(key, display, "laterality",
                        new[] { "left", "right", anatomy }));
                }
            }
        }
    }

    private static Dictionary<string, HashSet<string>> ExtractLateralizedAnatomy(string text)
    {
        var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var sentences = SplitSentences(text);

        foreach (var sentence in sentences)
        {
            // Skip negated sentences
            bool hasLeft = LeftRegex.IsMatch(sentence);
            bool hasRight = RightRegex.IsMatch(sentence);
            bool hasBilateral = BilateralRegex.IsMatch(sentence);

            if (!hasLeft && !hasRight && !hasBilateral)
                continue;

            // Extract anatomy words from this sentence
            var anatomyWords = AnatomyWordRegex.Matches(sentence.ToLowerInvariant())
                .Cast<Match>()
                .Select(m => m.Groups[1].Value)
                .Where(w => IsLateralityAnatomyWord(w))
                .Select(w => NormalizeLateralityAnatomy(w))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Build list of laterality positions in the sentence
            var latPositions = new List<(int pos, string side)>();
            foreach (Match m in LeftRegex.Matches(sentence))
                if (!IsNegated(sentence, m.Index))
                    latPositions.Add((m.Index, "left"));
            foreach (Match m in RightRegex.Matches(sentence))
                if (!IsNegated(sentence, m.Index))
                    latPositions.Add((m.Index, "right"));
            foreach (Match m in BilateralRegex.Matches(sentence))
                latPositions.Add((m.Index, "bilateral"));

            if (latPositions.Count == 0) continue;

            // For each anatomy word, associate with the NEAREST laterality (avoids cross-contamination
            // when a sentence mentions multiple anatomies with different sides)
            var sentenceLower = sentence.ToLowerInvariant();
            foreach (var anatomy in anatomyWords)
            {
                if (!result.ContainsKey(anatomy))
                    result[anatomy] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Find position of this anatomy word in the sentence
                var anatomyMatch = Regex.Match(sentenceLower, $@"\b{Regex.Escape(anatomy)}\b");
                if (!anatomyMatch.Success) continue;
                int anatomyPos = anatomyMatch.Index;

                // Associate with nearest laterality
                var nearest = latPositions.OrderBy(lp => Math.Abs(lp.pos - anatomyPos)).First();
                result[anatomy].Add(nearest.side);
            }
        }

        return result;
    }

    private static readonly HashSet<string> _lateralityAnatomyWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "kidney", "renal", "adrenal", "lung", "pulmonary", "ovary", "ovarian",
        "breast", "pleural", "inguinal", "femoral", "popliteal", "subclavian", "carotid",
        "parotid", "submandibular", "axillary", "iliac",
        "hip", "knee", "ankle", "foot", "hand", "wrist", "elbow", "shoulder",
        "humerus", "humeral", "radius", "radial", "ulna", "ulnar",
        "femur", "tibia", "tibial", "fibula", "fibular",
        "clavicle", "scapula", "patella", "calcaneus",
        "hydronephrosis", "hydroureter", "nephrolithiasis", "ureterolithiasis",
        "pneumothorax", "effusion", "atelectasis",
        "eye", "orbit", "orbital",
    };

    private static bool IsLateralityAnatomyWord(string word)
    {
        return _lateralityAnatomyWords.Contains(word);
    }

    private static string NormalizeLateralityAnatomy(string word)
    {
        if (LateralitySynonyms.TryGetValue(word, out var canonical))
            return canonical;
        return word.ToLowerInvariant();
    }

    #endregion

    #region Study Description Laterality

    private static void CheckStudyLaterality(string studyDescription, string reportText, List<ConsistencyResult> results)
    {
        var descLower = studyDescription.ToLowerInvariant();

        // Check for bilateral suppressors first — any match = skip entirely
        foreach (var suppressor in BilateralSuppressors)
        {
            if (Regex.IsMatch(descLower, $@"\b{Regex.Escape(suppressor)}\b"))
                return;
        }
        if (BilateralRegex.IsMatch(descLower))
            return;

        // Parse study description for side + body part
        bool descHasLeft = LeftRegex.IsMatch(studyDescription);
        bool descHasRight = RightRegex.IsMatch(studyDescription);

        if (!descHasLeft && !descHasRight)
            return; // No laterality in description

        // If description has BOTH left and right (e.g., "LEFT AND RIGHT KNEE"), treat as bilateral
        if (descHasLeft && descHasRight)
            return;

        // Find body parts in description
        var descBodyParts = new List<string>();
        foreach (var part in UnilateralBodyParts)
        {
            if (Regex.IsMatch(descLower, $@"\b{Regex.Escape(part)}\b"))
                descBodyParts.Add(part);
        }

        if (descBodyParts.Count == 0)
            return; // No recognizable unilateral body part

        // Determine study side
        string studySide = descHasLeft ? "left" : "right";
        string oppositeSide = descHasLeft ? "right" : "left";
        var oppositeRegex = descHasLeft ? RightRegex : LeftRegex;

        // Scan report for opposite-side mentions of same body parts
        var reportSentences = SplitSentences(reportText);

        foreach (var bodyPart in descBodyParts)
        {
            var bodyPartRegex = new Regex($@"\b{Regex.Escape(bodyPart)}\b", RegexOptions.IgnoreCase);

            foreach (var sentence in reportSentences)
            {
                if (!bodyPartRegex.IsMatch(sentence))
                    continue;

                var oppositeMatch = oppositeRegex.Match(sentence);
                if (!oppositeMatch.Success)
                    continue;

                // Skip negated context
                if (IsNegated(sentence, oppositeMatch.Index))
                    continue;

                var key = $"studylat_{bodyPart}_{studySide}_vs_{oppositeSide}";
                if (!results.Any(r => r.Key == key))
                {
                    var display = $"study {studySide} {bodyPart}, report says {oppositeSide}";
                    results.Add(new ConsistencyResult(key, display, "study_laterality",
                        new[] { oppositeSide, bodyPart }));
                }
                break; // One flag per body part
            }
        }
    }

    #endregion

    #region Shared Helpers

    private static bool IsNegated(string sentence, int matchIndex)
    {
        var prefix = sentence.Substring(0, matchIndex).ToLowerInvariant();

        foreach (var neg in NegationPhrases)
        {
            int negIdx = prefix.LastIndexOf(neg, StringComparison.Ordinal);
            if (negIdx >= 0)
            {
                var between = prefix.Substring(negIdx + neg.Length);
                if (!between.Contains(';') && !between.Contains(':')
                    && !between.Contains(" however") && !between.Contains(" but ")
                    && !between.Contains(" although"))
                    return true;
            }
        }

        return false;
    }

    private static List<string> SplitSentences(string text)
    {
        var parts = Regex.Split(text, @"(?<=[.!?])\s+|\r?\n|(?=\d+\.\s)");
        return parts
            .Select(p => p.Trim())
            .Where(p => p.Length > 0)
            .ToList();
    }

    #endregion
}
