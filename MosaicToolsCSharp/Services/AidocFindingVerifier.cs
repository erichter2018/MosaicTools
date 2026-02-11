using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MosaicTools.Services;

public record FindingVerification(string FindingType, bool IsAddressed);

public static class AidocFindingVerifier
{
    private static readonly Dictionary<string, string[]> FindingTerms = new(StringComparer.OrdinalIgnoreCase)
    {
        ["ICH"] = new[] { "hemorrhage", "hematoma", "ICH", "SAH", "SDH", "EDH", "IVH", "subarachnoid", "subdural", "epidural", "intraparenchymal", "intraventricular" },
        ["MLS"] = new[] { "midline shift", "midline deviation", "MLS", "mass effect", "subfalcine herniation" },
        ["LVO"] = new[] { "occlusion", "occluded", "thrombus", "thrombosis", "LVO" },
        ["M1-LVO"] = new[] { "occlusion", "occluded", "thrombus", "thrombosis", "LVO" },
        ["VO"] = new[] { "vertebral occlusion", "vertebral artery occlusion", "vertebral thrombosis" },
        ["BA"] = new[] { "aneurysm", "aneurysmal", "brain aneurysm", "cerebral aneurysm", "intracranial aneurysm" },
        ["PE"] = new[] { "pulmonary embolism", "pulmonary emboli", "pulmonary thrombus", "filling defect" },
        ["IPE"] = new[] { "pulmonary embolism", "pulmonary emboli", "pulmonary thrombus", "filling defect" },
        ["Ptx"] = new[] { "pneumothorax", "pneumothoraces" },
        ["iptx"] = new[] { "pneumothorax", "pneumothoraces" },
        ["PN"] = new[] { "pneumonia", "consolidation", "airspace disease", "airspace opacity", "infiltrate" },
        ["AD"] = new[] { "dissection", "intimal flap", "dissection flap" },
        ["M-Aorta"] = new[] { "aneurysm", "aneurysmal", "ectasia", "ectatic", "aortic dilation", "aortic dilatation" },
        ["M-AbdAo"] = new[] { "aneurysm", "aneurysmal", "AAA", "ectasia", "ectatic", "aortic dilation", "aortic dilatation" },
        ["CSpFx"] = new[] { "fracture", "fx" },
        ["VCFx"] = new[] { "compression fracture", "compression deformity", "vertebral fracture", "height loss" },
        ["FreeAir"] = new[] { "pneumoperitoneum", "free air", "free gas", "extraluminal air" },
        ["GIB"] = new[] { "extravasation", "GI bleed", "gastrointestinal hemorrhage", "active bleeding" },
        ["CAC"] = new[] { "coronary calcification", "coronary atherosclerosis", "CAC" },
        ["RibFx"] = new[] { "rib fracture", "fractured rib", "rib fx", "costal fracture" },
        ["MalETT"] = new[] { "endotracheal", "ETT", "ET tube" },
        ["RV/LV"] = new[] { "RV/LV", "right ventricular enlargement", "right ventricular strain", "RV dilation", "right heart strain" },
    };

    private static readonly string[] NegationPhrases = new[]
    {
        "no evidence of", "without evidence of", "negative for",
        "not identified", "not seen", "not visualized", "not detected",
        "not demonstrated", "absence of", "ruled out",
        "has resolved", "have resolved"
    };

    private static readonly HashSet<string> NegationWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "no", "not", "without", "absent", "negative", "denies"
    };

    public static List<FindingVerification> VerifyFindings(List<string> findingTypes, string reportText)
    {
        var results = new List<FindingVerification>();
        if (string.IsNullOrWhiteSpace(reportText))
        {
            foreach (var ft in findingTypes)
                results.Add(new FindingVerification(ft, false));
            return results;
        }

        // Clean pipe-delimited text (Mosaic sometimes uses pipes as line separators)
        var cleaned = reportText.Replace("|", "\n");

        // Extract only FINDINGS and IMPRESSION sections
        var relevantText = ExtractFindingsAndImpression(cleaned);
        if (string.IsNullOrWhiteSpace(relevantText))
            relevantText = cleaned; // Fallback to full text if no sections found

        foreach (var findingType in findingTypes)
        {
            if (FindingTerms.TryGetValue(findingType, out var terms))
            {
                bool addressed = false;
                foreach (var term in terms)
                {
                    if (IsTermPositivelyMentioned(relevantText, term))
                    {
                        addressed = true;
                        break;
                    }
                }
                results.Add(new FindingVerification(findingType, addressed));
                if (addressed)
                    Logger.Trace($"AidocVerifier: '{findingType}' addressed (term match in report)");
            }
            else
            {
                Logger.Trace($"AidocVerifier: Unknown finding type '{findingType}', marking as not addressed");
                results.Add(new FindingVerification(findingType, false));
            }
        }

        return results;
    }

    private static string ExtractFindingsAndImpression(string text)
    {
        // Look for FINDINGS and IMPRESSION section headers (common radiology report structure)
        // Headers are typically on their own line, possibly followed by colon
        var sections = new List<string>();

        var lines = text.Split('\n');
        bool inRelevantSection = false;
        var currentSection = new List<string>();

        foreach (var line in lines)
        {
            var trimmed = line.Trim().ToUpperInvariant();

            // Check if this line is a section header
            bool isSectionHeader = IsSectionHeader(trimmed);

            if (isSectionHeader)
            {
                // Save previous section if it was relevant
                if (inRelevantSection && currentSection.Count > 0)
                    sections.Add(string.Join("\n", currentSection));

                currentSection.Clear();

                // Check if new section is one we care about
                inRelevantSection = IsRelevantSectionHeader(trimmed);
            }
            else if (inRelevantSection)
            {
                currentSection.Add(line);
            }
        }

        // Don't forget the last section
        if (inRelevantSection && currentSection.Count > 0)
            sections.Add(string.Join("\n", currentSection));

        return string.Join("\n", sections);
    }

    private static bool IsSectionHeader(string trimmedUpper)
    {
        // Section headers: line starts with a known header keyword, optionally followed by colon
        var headers = new[] {
            "FINDINGS", "IMPRESSION", "CLINICAL HISTORY", "HISTORY",
            "CLINICAL INFORMATION", "INDICATION", "TECHNIQUE",
            "COMPARISON", "EXAM", "PROCEDURE", "CONCLUSION"
        };

        foreach (var h in headers)
        {
            if (trimmedUpper == h || trimmedUpper == h + ":" ||
                trimmedUpper.StartsWith(h + ":") || trimmedUpper.StartsWith(h + " "))
                return true;
        }
        return false;
    }

    private static bool IsRelevantSectionHeader(string trimmedUpper)
    {
        // We only want FINDINGS and IMPRESSION sections (where radiologist writes their assessment)
        var relevant = new[] { "FINDINGS", "IMPRESSION", "CONCLUSION" };
        foreach (var h in relevant)
        {
            if (trimmedUpper == h || trimmedUpper == h + ":" ||
                trimmedUpper.StartsWith(h + ":") || trimmedUpper.StartsWith(h + " "))
                return true;
        }
        return false;
    }

    private static bool IsTermPositivelyMentioned(string text, string term)
    {
        // Find all occurrences of the term (case-insensitive, word-boundary, optional plural 's')
        var pattern = @"\b" + Regex.Escape(term) + @"s?\b";
        // Handle terms with "/" which don't play nice with \b
        if (term.Contains('/'))
            pattern = Regex.Escape(term);

        var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);

        foreach (Match match in matches)
        {
            var (clause, clauseStart) = ExtractClause(text, match.Index, match.Length);
            var textBeforeTerm = clause.Substring(0, Math.Max(0, match.Index - clauseStart));

            if (!IsNegated(textBeforeTerm))
                return true; // Found a non-negated mention
        }

        return false; // All mentions negated, or term not found
    }

    private static (string Clause, int ClauseStart) ExtractClause(string text, int matchPos, int matchLen)
    {
        // Find clause boundaries: . \n ; : before and after the match
        int start = matchPos;
        for (int i = matchPos - 1; i >= 0; i--)
        {
            char c = text[i];
            if (c == '.' || c == '\n' || c == ';' || c == ':')
            {
                start = i + 1;
                break;
            }
            if (i == 0)
                start = 0;
        }

        int end = text.Length;
        for (int i = matchPos + matchLen; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '.' || c == '\n' || c == ';' || c == ':')
            {
                end = i;
                break;
            }
        }

        return (text.Substring(start, end - start), start);
    }

    private static bool IsNegated(string textBeforeTerm)
    {
        if (string.IsNullOrWhiteSpace(textBeforeTerm))
            return false;

        var lowerText = textBeforeTerm.ToLowerInvariant().Trim();

        // Check multi-word negation phrases anywhere before the term in the clause
        foreach (var phrase in NegationPhrases)
        {
            if (lowerText.Contains(phrase))
                return true;
        }

        // Check single negation words within 6-word window before term
        var words = lowerText.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        var lastWords = words.Length <= 6 ? words : words.Skip(words.Length - 6).ToArray();

        foreach (var word in lastWords)
        {
            // Strip punctuation for matching
            var clean = word.TrimEnd(',', '.', ';', ':');
            if (NegationWords.Contains(clean))
                return true;
        }

        return false;
    }
}
