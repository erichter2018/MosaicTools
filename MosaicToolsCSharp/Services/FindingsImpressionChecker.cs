using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MosaicTools.Services;

/// <summary>
/// A single mismatch between FINDINGS and IMPRESSION sections.
/// </summary>
public record MismatchResult(string Key, string DisplayName, string Direction, string[] SearchTerms);

/// <summary>
/// Detects contradictions between FINDINGS and IMPRESSION sections of a radiology report.
/// For example, "no bowel obstruction" in FINDINGS but "bowel obstruction" in IMPRESSION.
/// Uses both exact phrase matching and proximity matching (all key words in same sentence).
/// </summary>
public static class FindingsImpressionChecker
{
    /// <summary>
    /// Medical concept with synonym patterns and proximity word sets for matching.
    /// </summary>
    private class Concept
    {
        public string Key { get; init; } = "";
        public string DisplayName { get; init; } = "";
        public string[] Terms { get; init; } = Array.Empty<string>();
        public Regex[] Patterns { get; init; } = Array.Empty<Regex>();
        /// <summary>
        /// For each multi-word term, the set of significant words that must all appear
        /// in the same sentence for a proximity match. Single-word terms have no entry.
        /// </summary>
        public string[][] ProximityWordSets { get; init; } = Array.Empty<string[]>();
    }

    private static readonly Regex WordBoundary = new(@"\b(\w+)\b", RegexOptions.Compiled);

    /// <summary>
    /// Small words to skip when building proximity word sets.
    /// </summary>
    private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "of", "the", "a", "an", "in", "on", "for", "and", "or", "to", "with"
    };

    private static readonly Concept[] Concepts = BuildConcepts();

    /// <summary>
    /// Negation phrases ordered longest-first so greedy matching works correctly.
    /// </summary>
    private static readonly string[] NegationPhrases =
    {
        "no findings to suggest",
        "no evidence of",
        "no convincing",
        "no definite",
        "no significant",
        "no signs of",
        "no findings of",
        "no acute",
        "negative for",
        "absence of",
        "rules out",
        "ruled out",
        "without",
        "not ",
        "no ",
    };

    /// <summary>
    /// Check for mismatches between findings and impression sections.
    /// Returns list of concepts where one section negates what the other asserts.
    /// </summary>
    public static List<MismatchResult> Check(string? findings, string? impression)
    {
        var results = new List<MismatchResult>();
        if (string.IsNullOrWhiteSpace(findings) || string.IsNullOrWhiteSpace(impression))
            return results;

        var findingsLower = findings.ToLowerInvariant();
        var impressionLower = impression.ToLowerInvariant();

        // Pre-split into sentences for negation lookback
        var findingsSentences = SplitSentences(findingsLower);
        var impressionSentences = SplitSentences(impressionLower);

        foreach (var concept in Concepts)
        {
            var findingsMatches = ClassifyMatches(concept, findingsSentences);
            var impressionMatches = ClassifyMatches(concept, impressionSentences);

            // Skip if concept not present in both sections
            if (findingsMatches.Count == 0 || impressionMatches.Count == 0)
                continue;

            bool findingsAllNegated = findingsMatches.All(m => m.isNegated);
            bool findingsAnyPositive = findingsMatches.Any(m => !m.isNegated);
            bool impressionAllNegated = impressionMatches.All(m => m.isNegated);
            bool impressionAnyPositive = impressionMatches.Any(m => !m.isNegated);

            // Mismatch: findings negated + impression positive
            if (findingsAllNegated && impressionAnyPositive)
            {
                results.Add(new MismatchResult(concept.Key, concept.DisplayName, "findings_neg_impression_pos", concept.Terms));
            }
            // Mismatch: findings positive + impression negated
            else if (findingsAnyPositive && impressionAllNegated)
            {
                results.Add(new MismatchResult(concept.Key, concept.DisplayName, "findings_pos_impression_neg", concept.Terms));
            }
        }

        return results;
    }

    /// <summary>
    /// Find all matches of concept patterns in sentences and classify as negated or positive.
    /// Tries exact phrase patterns first, then proximity word matching for multi-word terms.
    /// </summary>
    private static List<(string sentence, bool isNegated)> ClassifyMatches(
        Concept concept, List<string> sentences)
    {
        var results = new List<(string, bool)>();

        foreach (var sentence in sentences)
        {
            // Try exact phrase patterns first
            bool matched = false;
            foreach (var pattern in concept.Patterns)
            {
                var match = pattern.Match(sentence);
                if (match.Success)
                {
                    bool negated = IsNegated(sentence, match.Index);
                    results.Add((sentence, negated));
                    matched = true;
                    break;
                }
            }

            if (matched) continue;

            // Try proximity matching: all key words present in the sentence (any order)
            foreach (var wordSet in concept.ProximityWordSets)
            {
                if (wordSet.Length < 2) continue; // single-word already covered by exact patterns

                int lastWordPos = -1;
                bool allFound = true;
                foreach (var word in wordSet)
                {
                    var wordMatch = Regex.Match(sentence, $@"\b{Regex.Escape(word)}\b");
                    if (!wordMatch.Success)
                    {
                        allFound = false;
                        break;
                    }
                    // Track the rightmost key word position for negation check
                    if (wordMatch.Index > lastWordPos)
                        lastWordPos = wordMatch.Index;
                }

                if (allFound && lastWordPos >= 0)
                {
                    bool negated = IsNegated(sentence, lastWordPos);
                    results.Add((sentence, negated));
                    break;
                }
            }
        }

        return results;
    }

    /// <summary>
    /// Check if a match at the given position in a sentence is preceded by a negation phrase.
    /// Looks back from the match position to the start of the sentence.
    /// </summary>
    private static bool IsNegated(string sentence, int matchIndex)
    {
        // Get text before the match
        var prefix = sentence.Substring(0, matchIndex);

        // Check each negation phrase (longest first)
        foreach (var neg in NegationPhrases)
        {
            // Look for negation phrase near the match (within the same clause)
            int negIdx = prefix.LastIndexOf(neg, StringComparison.Ordinal);
            if (negIdx >= 0)
            {
                // Ensure negation is in the same clause (no intervening clause breaks)
                var between = prefix.Substring(negIdx + neg.Length);
                if (!between.Contains(';') && !between.Contains(':')
                    && !ContainsClauseBreak(between))
                    return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Check for clause-breaking conjunctions that would isolate a negation from the match.
    /// </summary>
    private static bool ContainsClauseBreak(string text)
    {
        // "however", "but", "although" create clause breaks
        return text.Contains(" however") || text.Contains(" but ") || text.Contains(" although");
    }

    /// <summary>
    /// Split text into sentences at period, newline, or numbered-list boundaries.
    /// </summary>
    private static List<string> SplitSentences(string text)
    {
        // Split on sentence-ending punctuation, newlines, or numbered list items
        var parts = Regex.Split(text, @"(?<=[.!?])\s+|\r?\n|(?=\d+\.\s)");
        return parts
            .Select(p => p.Trim())
            .Where(p => p.Length > 0)
            .ToList();
    }

    /// <summary>
    /// Build a regex pattern for a term. Short abbreviations use word boundaries;
    /// longer terms use a looser match.
    /// </summary>
    private static Regex MakePattern(string term)
    {
        var escaped = Regex.Escape(term);
        // Short abbreviations (<=4 chars, has uppercase) get strict word boundaries
        if (term.Length <= 4 && term.Any(char.IsUpper))
            return new Regex($@"\b{escaped}\b", RegexOptions.Compiled);
        // Everything else: word boundaries, case-insensitive
        return new Regex($@"\b{escaped}\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    }

    /// <summary>
    /// Extract significant words from a multi-word term for proximity matching.
    /// Returns null for single-word terms (already handled by exact patterns).
    /// </summary>
    private static string[]? ExtractProximityWords(string term)
    {
        var words = WordBoundary.Matches(term.ToLowerInvariant())
            .Cast<Match>()
            .Select(m => m.Groups[1].Value)
            .Where(w => w.Length >= 2 && !StopWords.Contains(w))
            .ToArray();

        return words.Length >= 2 ? words : null;
    }

    private static Concept[] BuildConcepts()
    {
        var defs = new (string key, string display, string[] terms)[]
        {
            ("bowel_obstruction", "bowel obstruction", new[] { "bowel obstruction", "SBO", "small bowel obstruction", "large bowel obstruction" }),
            ("pneumothorax", "pneumothorax", new[] { "pneumothorax", "PTX" }),
            ("pulmonary_embolism", "pulmonary embolism", new[] { "pulmonary embolism", "pulmonary embolus", "PE" }),
            ("fracture", "fracture", new[] { "fracture", "fx" }),
            ("hemorrhage", "hemorrhage", new[] { "hemorrhage", "hematoma", "bleeding" }),
            ("aortic_dissection", "aortic dissection", new[] { "aortic dissection", "dissection flap" }),
            ("appendicitis", "appendicitis", new[] { "appendicitis" }),
            ("cholecystitis", "cholecystitis", new[] { "cholecystitis" }),
            ("pancreatitis", "pancreatitis", new[] { "pancreatitis" }),
            ("diverticulitis", "diverticulitis", new[] { "diverticulitis" }),
            ("colitis", "colitis", new[] { "colitis" }),
            ("pneumonia", "pneumonia", new[] { "pneumonia" }),
            ("consolidation", "consolidation", new[] { "consolidation" }),
            ("abscess", "abscess", new[] { "abscess" }),
            ("perforation", "perforation", new[] { "perforation", "perforated" }),
            ("pneumoperitoneum", "pneumoperitoneum", new[] { "pneumoperitoneum", "free air", "free intraperitoneal air" }),
            ("hydronephrosis", "hydronephrosis", new[] { "hydronephrosis" }),
            ("aneurysm", "aneurysm", new[] { "aneurysm", "aneurysmal" }),
            ("stenosis", "stenosis", new[] { "stenosis", "stenotic" }),
            ("thrombosis", "thrombosis", new[] { "thrombosis", "thrombus" }),
            ("dvt", "DVT", new[] { "DVT", "deep vein thrombosis", "deep venous thrombosis" }),
            ("lymphadenopathy", "lymphadenopathy", new[] { "lymphadenopathy" }),
            ("metastasis", "metastasis", new[] { "metastasis", "metastatic", "metastases" }),
            ("cardiomegaly", "cardiomegaly", new[] { "cardiomegaly" }),
            ("pulmonary_edema", "pulmonary edema", new[] { "pulmonary edema", "pulmonary oedema" }),
            ("pleural_effusion", "pleural effusion", new[] { "pleural effusion" }),
            ("pericardial_effusion", "pericardial effusion", new[] { "pericardial effusion" }),
            ("ascites", "ascites", new[] { "ascites" }),
            ("volvulus", "volvulus", new[] { "volvulus" }),
            ("intussusception", "intussusception", new[] { "intussusception" }),
            ("torsion", "torsion", new[] { "torsion" }),
            ("cord_compression", "cord compression", new[] { "cord compression", "spinal cord compression" }),
            ("midline_shift", "midline shift", new[] { "midline shift" }),
            ("herniation", "herniation", new[] { "herniation", "herniating" }),
            ("occlusion", "occlusion", new[] { "occlusion", "occluded" }),
            ("infarct", "infarct", new[] { "infarct", "infarction" }),
        };

        return defs.Select(d => new Concept
        {
            Key = d.key,
            DisplayName = d.display,
            Terms = d.terms,
            Patterns = d.terms.Select(MakePattern).ToArray(),
            ProximityWordSets = d.terms
                .Select(ExtractProximityWords)
                .Where(w => w != null)
                .Select(w => w!)
                .ToArray()
        }).ToArray();
    }
}
