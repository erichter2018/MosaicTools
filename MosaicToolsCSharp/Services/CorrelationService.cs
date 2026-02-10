using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace MosaicTools.Services;

/// <summary>
/// A single correlation between an impression item and its matched findings.
/// </summary>
public class CorrelationItem
{
    public int ImpressionIndex { get; set; }
    public string ImpressionText { get; set; } = "";
    public List<string> MatchedFindings { get; set; } = new();
    public int ColorIndex { get; set; }
}

/// <summary>
/// Result of a correlation analysis between FINDINGS and IMPRESSION sections.
/// </summary>
public class CorrelationResult
{
    public List<CorrelationItem> Items { get; set; } = new();
    public string Source { get; set; } = "heuristic";
}

/// <summary>
/// Heuristic-based correlation between FINDINGS and IMPRESSION sections of a radiology report.
/// Extracts canonical medical terms and matches impression items to their supporting findings.
/// </summary>
public static class CorrelationService
{
    /// <summary>
    /// Color palette for correlation highlighting (8 muted pastels).
    /// </summary>
    public static readonly Color[] Palette = new[]
    {
        Color.FromArgb(255, 50, 50),    // Red
        Color.FromArgb(255, 140, 0),    // Orange
        Color.FromArgb(255, 255, 0),    // Yellow
        Color.FromArgb(0, 200, 0),      // Green
        Color.FromArgb(0, 150, 255),    // Blue
        Color.FromArgb(75, 0, 180),     // Indigo
        Color.FromArgb(180, 0, 255),    // Violet
    };

    /// <summary>
    /// Blend a palette color at ~30% alpha with a dark background.
    /// </summary>
    public static Color BlendWithBackground(Color paletteColor, Color background)
    {
        const float alpha = 0.50f;
        return Color.FromArgb(
            (int)(background.R + (paletteColor.R - background.R) * alpha),
            (int)(background.G + (paletteColor.G - background.G) * alpha),
            (int)(background.B + (paletteColor.B - background.B) * alpha));
    }

    /// <summary>
    /// Synonym dictionary mapping adjective/alternate forms to canonical nouns.
    /// </summary>
    private static readonly Dictionary<string, string> Synonyms = new(StringComparer.OrdinalIgnoreCase)
    {
        // Anatomical
        { "hepatic", "liver" }, { "hepatomegaly", "liver" },
        { "renal", "kidney" }, { "nephro", "kidney" },
        { "pulmonary", "lung" }, { "pulmonic", "lung" },
        { "cardiac", "heart" }, { "myocardial", "heart" }, { "cardiomegaly", "heart" },
        { "cerebral", "brain" }, { "intracranial", "brain" }, { "cranial", "brain" },
        { "osseous", "bone" }, { "skeletal", "bone" }, { "bony", "bone" },
        { "splenic", "spleen" }, { "splenomegaly", "spleen" },
        { "pancreatic", "pancreas" },
        { "aortic", "aorta" },
        { "pleural", "pleura" },
        { "colonic", "colon" },
        { "biliary", "bile duct" }, { "choledochal", "bile duct" },
        { "thyroid", "thyroid" }, { "thyroidal", "thyroid" },
        { "adrenal", "adrenal gland" }, { "suprarenal", "adrenal gland" },
        { "prostatic", "prostate" },
        { "uterine", "uterus" }, { "endometrial", "uterus" },
        { "ovarian", "ovary" },
        { "vesical", "bladder" },
        { "gastric", "stomach" },
        { "duodenal", "duodenum" },
        { "ileal", "ileum" }, { "jejunal", "jejunum" },
        { "cecal", "cecum" }, { "appendiceal", "appendix" },
        { "rectal", "rectum" },
        { "esophageal", "esophagus" },
        { "tracheal", "trachea" },
        { "bronchial", "bronchus" },
        { "vertebral", "vertebra" }, { "spinal", "spine" },
        { "pericardial", "pericardium" },
        { "peritoneal", "peritoneum" },
        { "retroperitoneal", "retroperitoneum" },
        { "mediastinal", "mediastinum" },
        { "hilar", "hilum" },
        { "mesenteric", "mesentery" },
        { "inguinal", "groin" },
        { "femoral", "femur" },
        { "humeral", "humerus" },
        { "tibial", "tibia" },
        { "pelvic", "pelvis" },

        // Pathology
        { "nephrolithiasis", "kidney stone" }, { "urolithiasis", "kidney stone" },
        { "cholelithiasis", "gallstone" }, { "gallbladder stone", "gallstone" },
        { "calculus", "stone" }, { "calculi", "stone" }, { "calcified", "calcification" },
        { "thrombus", "clot" }, { "thrombosis", "clot" }, { "thrombotic", "clot" },
        { "embolus", "clot" }, { "embolism", "clot" }, { "embolic", "clot" }, { "filling defect", "clot" },
        { "hemorrhage", "bleeding" }, { "hemorrhagic", "bleeding" },
        { "hematoma", "bleeding" },
        { "edema", "swelling" }, { "edematous", "swelling" },
        { "stenosis", "narrowing" }, { "stenotic", "narrowing" },
        { "hernia", "herniation" }, { "herniated", "herniation" },
        { "fracture", "fracture" }, { "fractured", "fracture" },
        { "effusion", "effusion" },
        { "consolidation", "consolidation" }, { "consolidated", "consolidation" },
        { "atelectasis", "atelectasis" }, { "atelectatic", "atelectasis" },
        { "pneumothorax", "pneumothorax" },
        { "pneumonia", "pneumonia" },
        { "abscess", "abscess" },
        { "mass", "mass" }, { "lesion", "lesion" }, { "nodule", "nodule" }, { "nodular", "nodule" },
        { "tumor", "tumor" }, { "neoplasm", "tumor" }, { "neoplastic", "tumor" },
        { "malignant", "malignancy" }, { "malignancy", "malignancy" },
        { "metastasis", "metastatic disease" }, { "metastatic", "metastatic disease" }, { "metastases", "metastatic disease" },
        { "lymphadenopathy", "lymph node enlargement" }, { "lymph node", "lymph node" },
        { "aneurysm", "aneurysm" }, { "aneurysmal", "aneurysm" },
        { "dissection", "dissection" },
        { "obstruction", "obstruction" }, { "obstructing", "obstruction" }, { "obstructive", "obstruction" },
        { "dilatation", "dilation" }, { "dilated", "dilation" }, { "dilation", "dilation" },
        { "hydronephrosis", "hydronephrosis" }, { "hydroureter", "hydroureter" },
        { "ascites", "ascites" },
        { "swollen", "swelling" },

        // Radiology descriptive
        { "opacification", "opacity" }, { "opacified", "opacity" }, { "opaque", "opacity" },
        { "lucent", "lucency" }, { "radiolucent", "lucency" },
        { "sclerotic", "sclerosis" },
        { "osteolytic", "lytic" },

        // Degenerative / MSK
        { "degenerative", "degeneration" }, { "spondylosis", "degeneration" },
        { "arthritic", "arthritis" }, { "arthropathy", "arthritis" },
        { "scoliotic", "scoliosis" },
        { "kyphotic", "kyphosis" },
        { "lordotic", "lordosis" },

        // Vascular
        { "ectatic", "dilation" }, { "ectasia", "dilation" },
        { "tortuous", "tortuosity" },
        { "atherosclerotic", "atherosclerosis" }, { "atheromatous", "atherosclerosis" },

        // Tissue changes
        { "fibrotic", "fibrosis" },
        { "emphysematous", "emphysema" },
        { "necrotic", "necrosis" },
        { "ischemic", "ischemia" },
        { "infarcted", "infarction" }, { "infarct", "infarction" },

        // Structural
        { "perforated", "perforation" },
        { "displaced", "displacement" },
        { "compressed", "compression" },
        { "distended", "distension" },
        { "collapsed", "collapse" },

        // Neoplastic
        { "carcinoma", "malignancy" }, { "carcinomatous", "malignancy" },

        // Size/shape descriptors
        { "thickened", "thickening" },
        { "enlarged", "enlargement" },
        { "prominent", "prominence" },

        // Short but important pathology terms (< 5 chars, need synonym to be included)
        { "cystic", "cyst" }, { "cyst", "cyst" },

        // Missing anatomical adjective → noun
        { "abdominal", "abdomen" },
        { "sacral", "sacrum" },
    };

    /// <summary>
    /// Laterality terms that alone are too generic to establish a correlation match.
    /// These still contribute to scoring (to pick left vs right impression) but
    /// at least one non-laterality term must also match.
    /// </summary>
    private static readonly HashSet<string> LateralityTerms = new(StringComparer.OrdinalIgnoreCase)
    {
        "left", "right"
    };

    /// <summary>
    /// Anatomical terms for direct matching (these are already canonical).
    /// </summary>
    private static readonly HashSet<string> AnatomicalTerms = new(StringComparer.OrdinalIgnoreCase)
    {
        "liver", "kidney", "lung", "brain", "heart", "spleen", "pancreas", "aorta",
        "pleura", "colon", "bile duct", "thyroid", "adrenal", "prostate", "uterus",
        "ovary", "bladder", "stomach", "duodenum", "ileum", "jejunum", "cecum",
        "appendix", "rectum", "esophagus", "trachea", "bronchus", "vertebra", "spine",
        "pericardium", "peritoneum", "retroperitoneum", "mediastinum", "hilum",
        "mesentery", "femur", "humerus", "tibia", "pelvis", "gallbladder",
        "bowel", "small bowel", "large bowel", "chest wall", "diaphragm",
        "bone", "lymph node",
        // Laterality and short anatomical terms (< 5 chars, need explicit listing)
        "left", "right", "lobe", "rib", "disc", "disk",
        "atrium", "ventricle", "artery", "vein", "sacrum"
    };

    /// <summary>
    /// Correlate findings and impression sections of a radiology report.
    /// </summary>
    public static CorrelationResult Correlate(string reportText, int? seed = null)
    {
        var result = new CorrelationResult { Source = "heuristic" };

        Logger.Trace($"Correlate: input length={reportText.Length}");

        var contextTerms = ExtractExamContextTerms(reportText);
        var (findingsText, impressionText) = ExtractSections(reportText);
        Logger.Trace($"Correlate: findings={findingsText.Length} chars, impression={impressionText.Length} chars");
        if (string.IsNullOrWhiteSpace(findingsText) || string.IsNullOrWhiteSpace(impressionText))
        {
            Logger.Trace("Correlate: Missing FINDINGS or IMPRESSION section");
            return result;
        }

        var impressionItems = ParseImpressionItems(impressionText);
        var findingsSentences = ParseFindingsSentences(findingsText);

        Logger.Trace($"Correlate: {impressionItems.Count} impression items, {findingsSentences.Count} findings sentences");

        if (impressionItems.Count == 0 || findingsSentences.Count == 0)
            return result;

        // Group findings sentences by subsection (contiguous blocks)
        var findingsBlocks = GroupFindingsBySubsection(findingsText);
        Logger.Trace($"Correlate: {findingsBlocks.Count} findings subsection blocks");

        // Shuffle color order — seed from report text so same report gets same colors
        var colorOrder = Enumerable.Range(0, Palette.Length).ToList();
        var rng = new Random(seed ?? reportText.GetHashCode());
        for (int i = colorOrder.Count - 1; i > 0; i--)
        {
            int j = rng.Next(i + 1);
            (colorOrder[i], colorOrder[j]) = (colorOrder[j], colorOrder[i]);
        }

        int colorIndex = 0;
        foreach (var (index, text) in impressionItems)
        {
            // Skip negative/normal impression statements — these match too broadly
            if (NegativeSentencePattern.IsMatch(text))
            {
                Logger.Trace($"Correlate: Impression #{index} skipped (negative statement): {text.Substring(0, Math.Min(80, text.Length))}");
                continue;
            }

            var impressionTerms = ExtractTerms(text, contextTerms);
            Logger.Trace($"Correlate: Impression #{index} terms: [{string.Join(", ", impressionTerms)}] from: {text.Substring(0, Math.Min(80, text.Length))}");

            if (impressionTerms.Count == 0) continue;

            // Find the best matching subsection block (most shared terms)
            // This prevents a single common word from spreading across the whole report
            int bestScore = 0;
            int bestSubstantiveScore = 0;
            List<string>? bestBlock = null;

            foreach (var block in findingsBlocks)
            {
                // Pool all terms from the block's sentences
                var blockTerms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var sentence in block)
                {
                    foreach (var t in ExtractTerms(sentence, contextTerms))
                        blockTerms.Add(t);
                }

                var shared = impressionTerms.Intersect(blockTerms, StringComparer.OrdinalIgnoreCase).ToList();
                int substantiveCount = shared.Count(t => !LateralityTerms.Contains(t));
                if (shared.Count > bestScore)
                {
                    bestScore = shared.Count;
                    bestSubstantiveScore = substantiveCount;
                    bestBlock = block;
                    Logger.Trace($"Correlate:   Block match score={shared.Count} substantive={substantiveCount} terms=[{string.Join(", ", shared)}]");
                }
            }

            // Require at least 1 non-laterality shared term with the best block
            if (bestSubstantiveScore >= 1 && bestBlock != null)
            {
                // Only include sentences from the best block that actually share substantive terms
                var matchedSentences = new List<string>();
                foreach (var sentence in bestBlock)
                {
                    var sentenceTerms = ExtractTerms(sentence, contextTerms);
                    var sentenceShared = impressionTerms.Intersect(sentenceTerms, StringComparer.OrdinalIgnoreCase).ToList();
                    if (sentenceShared.Any(t => !LateralityTerms.Contains(t)))
                    {
                        matchedSentences.Add(sentence);
                    }
                }

                if (matchedSentences.Count > 0)
                {
                    result.Items.Add(new CorrelationItem
                    {
                        ImpressionIndex = index,
                        ImpressionText = text,
                        MatchedFindings = matchedSentences,
                        ColorIndex = colorOrder[colorIndex % colorOrder.Count]
                    });
                    colorIndex++;
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Regex pattern matching negative/normal/incidental finding sentences that don't need impression representation.
    /// Start-anchored patterns prevent filtering mixed sentences like
    /// "The kidney shows no hydronephrosis but a new 2 cm mass."
    /// Unanchored patterns match anywhere for unambiguous indicators.
    /// </summary>
    private static readonly Regex NegativeSentencePattern = new(
        @"^\s*(?:no\s|normal\b|negative\b|stable\b|there\s+(?:is|are)\s+no\s|not\s)" +
        @"|\bunremarkable\b|\bwithin\s+normal\b" +
        @"|\b(?:was|has\s+been)\s+performed\b" +               // surgical history
        @"|\bdemonstrates?\s+no\s" +                            // mid-sentence negative
        @"|\bis\s+normal\s+in\b" +                              // normal descriptive
        @"|\bnot\s+clearly\s+(?:seen|visualized|identified)\b" + // non-visualization
        @"|\b(?:is|are)\s+clear\b" +                            // "airways are clear"
        @"|\bno\s+evidence\b",                                  // "no evidence of acute..."
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Reversed correlation: starts from dictated findings and checks whether each
    /// is represented in the impression. Unmatched dictated findings get a unique
    /// "orphan" color that appears nowhere in the impression.
    /// </summary>
    public static CorrelationResult CorrelateReversed(string reportText, HashSet<string>? dictatedSentences = null, int? seed = null)
    {
        var result = new CorrelationResult { Source = "heuristic-reversed" };

        Logger.Trace($"CorrelateReversed: input length={reportText.Length}, dictatedSentences={dictatedSentences?.Count ?? -1}");

        var contextTerms = ExtractExamContextTerms(reportText);
        var (findingsText, impressionText) = ExtractSections(reportText);
        if (string.IsNullOrWhiteSpace(findingsText) || string.IsNullOrWhiteSpace(impressionText))
        {
            Logger.Trace("CorrelateReversed: Missing FINDINGS or IMPRESSION section");
            return result;
        }

        var findingsSentences = ParseFindingsSentences(findingsText);
        var impressionItems = ParseImpressionItems(impressionText);

        if (findingsSentences.Count == 0 || impressionItems.Count == 0)
            return result;

        // Filter findings to only dictated sentences (if baseline was available)
        List<string> filteredFindings;
        if (dictatedSentences != null && dictatedSentences.Count > 0)
        {
            filteredFindings = new List<string>();
            foreach (var sentence in findingsSentences)
            {
                var normalized = Regex.Replace(sentence.ToLowerInvariant().Trim(), @"\s+", " ");
                if (dictatedSentences.Contains(normalized))
                    filteredFindings.Add(sentence);
            }
            Logger.Trace($"CorrelateReversed: {filteredFindings.Count}/{findingsSentences.Count} findings are dictated");
        }
        else
        {
            filteredFindings = new List<string>(findingsSentences);
        }

        // Identify negative/normal statements (these can match impressions but won't become orphans)
        var negativeFindings = new HashSet<string>(StringComparer.Ordinal);
        foreach (var sentence in filteredFindings)
        {
            if (NegativeSentencePattern.IsMatch(sentence))
            {
                negativeFindings.Add(sentence);
                Logger.Trace($"CorrelateReversed: Negative finding (can match, won't orphan): {sentence.Substring(0, Math.Min(60, sentence.Length))}");
            }
        }
        int significantCount = filteredFindings.Count - negativeFindings.Count;
        Logger.Trace($"CorrelateReversed: {significantCount} significant + {negativeFindings.Count} negative findings");

        // Safety net: if no dictated sentences provided AND most findings are significant (suggesting
        // baseline wasn't captured), fall back to the old impression-first approach
        if (dictatedSentences == null && significantCount > findingsSentences.Count * 0.8)
        {
            Logger.Trace("CorrelateReversed: No baseline, most findings pass filter — falling back to Correlate()");
            return Correlate(reportText, seed);
        }

        if (significantCount == 0)
            return result;

        // Shuffle color order for consistent colors per report
        var colorOrder = Enumerable.Range(0, Palette.Length).ToList();
        var rng = new Random(seed ?? reportText.GetHashCode());
        for (int i = colorOrder.Count - 1; i > 0; i--)
        {
            int j = rng.Next(i + 1);
            (colorOrder[i], colorOrder[j]) = (colorOrder[j], colorOrder[i]);
        }

        // For each finding, find the best-matching impression item
        // All findings (including negative) participate in matching,
        // but only positive findings become orphans when unmatched
        var matchedGroups = new Dictionary<int, (string impressionText, List<string> findings)>();
        var orphanFindings = new List<string>();

        foreach (var finding in filteredFindings)
        {
            bool isNegative = negativeFindings.Contains(finding);
            var findingTerms = ExtractTerms(finding, contextTerms);
            Logger.Trace($"CorrelateReversed: Finding terms=[{string.Join(", ", findingTerms)}] neg={isNegative} for: {finding.Substring(0, Math.Min(60, finding.Length))}");
            if (findingTerms.Count == 0)
            {
                if (!isNegative) orphanFindings.Add(finding);
                continue;
            }

            int bestScore = 0;
            int bestSubstantiveScore = 0;
            int bestImpressionIdx = -1;
            string bestImpressionText = "";

            foreach (var (index, text) in impressionItems)
            {
                // Skip negative/normal impression statements — these match too broadly
                if (NegativeSentencePattern.IsMatch(text))
                    continue;

                var impressionTerms = ExtractTerms(text, contextTerms);
                var shared = findingTerms.Intersect(impressionTerms, StringComparer.OrdinalIgnoreCase).ToList();
                int substantiveCount = shared.Count(t => !LateralityTerms.Contains(t));
                Logger.Trace($"CorrelateReversed:   vs Impression #{index} terms=[{string.Join(", ", impressionTerms)}] shared=[{string.Join(", ", shared)}] score={shared.Count} substantive={substantiveCount}");
                if (shared.Count > bestScore)
                {
                    bestScore = shared.Count;
                    bestSubstantiveScore = substantiveCount;
                    bestImpressionIdx = index;
                    bestImpressionText = text;
                }
            }

            // Require at least one non-laterality shared term to count as a match
            if (bestSubstantiveScore >= 1 && bestImpressionIdx >= 0)
            {
                Logger.Trace($"CorrelateReversed: MATCHED finding to Impression #{bestImpressionIdx} (score={bestScore}, substantive={bestSubstantiveScore}, neg={isNegative})");
                // Negative findings participate in scoring (so the impression gets matched)
                // but are excluded from highlighted output to avoid highlighting boilerplate
                if (!isNegative)
                {
                    if (!matchedGroups.ContainsKey(bestImpressionIdx))
                        matchedGroups[bestImpressionIdx] = (bestImpressionText, new List<string>());
                    matchedGroups[bestImpressionIdx].findings.Add(finding);
                }
            }
            else if (!isNegative)
            {
                Logger.Trace($"CorrelateReversed: ORPHAN finding (bestScore={bestScore}, substantive={bestSubstantiveScore})");
                orphanFindings.Add(finding);
            }
            else
            {
                Logger.Trace($"CorrelateReversed: Negative finding unmatched, not orphaned");
            }
        }

        // Build result: matched groups get a color shared between impression item and findings
        int colorIndex = 0;
        foreach (var (groupIdx, (groupText, groupFindings)) in matchedGroups)
        {
            result.Items.Add(new CorrelationItem
            {
                ImpressionIndex = groupIdx,
                ImpressionText = groupText,
                MatchedFindings = groupFindings,
                ColorIndex = colorOrder[colorIndex % colorOrder.Count]
            });
            colorIndex++;
        }

        // Orphan findings: each gets its own color, no impression text
        foreach (var orphan in orphanFindings)
        {
            result.Items.Add(new CorrelationItem
            {
                ImpressionIndex = -1,
                ImpressionText = "",
                MatchedFindings = new List<string> { orphan },
                ColorIndex = colorOrder[colorIndex % colorOrder.Count]
            });
            colorIndex++;
        }

        Logger.Trace($"CorrelateReversed: {matchedGroups.Count} matched groups, {orphanFindings.Count} orphans");
        return result;
    }

    /// <summary>
    /// Extract context terms from the EXAM description line to exclude from correlation.
    /// Terms like "lumbar", "spine" in a lumbar spine study appear everywhere and don't indicate pathology.
    /// </summary>
    internal static HashSet<string> ExtractExamContextTerms(string reportText)
    {
        var contextTerms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Find EXAM: line (inline "EXAM: description" or "EXAM:" followed by description on next line)
        var lines = reportText.Replace("\r\n", "\n").Split('\n');
        string? examLine = null;
        bool sawExamHeader = false;

        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith("EXAM:", StringComparison.OrdinalIgnoreCase))
            {
                var rest = trimmed.Substring(5).Trim();
                if (!string.IsNullOrEmpty(rest))
                {
                    examLine = rest;
                    break;
                }
                // Standalone "EXAM:" — description is on the next line
                sawExamHeader = true;
                continue;
            }
            if (sawExamHeader && !string.IsNullOrWhiteSpace(trimmed))
            {
                // Don't grab another section header as the exam description
                if (Regex.IsMatch(trimmed, @"^[A-Z\s]+:\s*$"))
                    break;
                examLine = trimmed;
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(examLine)) return contextTerms;

        // Strip date/time suffix (e.g., "02/06/2026 06:59:00 PM")
        examLine = Regex.Replace(examLine, @"\d{1,2}/\d{1,2}/\d{4}.*$", "").Trim();

        // Extract significant words from exam description
        var words = Regex.Split(examLine.ToLowerInvariant(), @"[^a-z]+")
            .Where(w => w.Length >= 4) // Only meaningful words
            .ToArray();

        foreach (var word in words)
        {
            // Skip very common modifiers that aren't anatomical context
            if (word is "with" or "without" or "views" or "view" or "contrast" or "portable")
                continue;
            contextTerms.Add(word);
            // Also add synonym-mapped forms
            if (Synonyms.TryGetValue(word, out var canonical))
                contextTerms.Add(canonical);
            if (AnatomicalTerms.Contains(word))
                contextTerms.Add(word);
        }

        if (contextTerms.Count > 0)
            Logger.Trace($"Exam context terms (excluded from matching): [{string.Join(", ", contextTerms)}]");

        return contextTerms;
    }

    /// <summary>
    /// Extract FINDINGS and IMPRESSION sections from report text.
    /// </summary>
    internal static (string findings, string impression) ExtractSections(string text)
    {
        // Normalize
        text = text.Replace("\r\n", "\n").Replace("\r", "\n");

        // Find section boundaries using regex
        var findingsMatch = Regex.Match(text, @"^FINDINGS:\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
        var impressionMatch = Regex.Match(text, @"^IMPRESSION:\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);

        if (!findingsMatch.Success || !impressionMatch.Success)
        {
            // Try inline format: "FINDINGS: content" on same line
            findingsMatch = Regex.Match(text, @"FINDINGS:", RegexOptions.IgnoreCase);
            impressionMatch = Regex.Match(text, @"IMPRESSION:", RegexOptions.IgnoreCase);
        }

        if (!findingsMatch.Success || !impressionMatch.Success)
            return ("", "");

        // Ensure FINDINGS comes before IMPRESSION
        if (findingsMatch.Index >= impressionMatch.Index)
            return ("", "");

        string findings = text.Substring(
            findingsMatch.Index + findingsMatch.Length,
            impressionMatch.Index - findingsMatch.Index - findingsMatch.Length).Trim();

        string impression = text.Substring(
            impressionMatch.Index + impressionMatch.Length).Trim();

        // Trim impression at next major section if any
        var nextSection = Regex.Match(impression, @"^\s*[A-Z][A-Z\s]+:", RegexOptions.Multiline);
        if (nextSection.Success && nextSection.Index > 0)
        {
            impression = impression.Substring(0, nextSection.Index).Trim();
        }

        return (findings, impression);
    }

    /// <summary>
    /// Parse impression text into numbered items. If no numbered items, treat whole text as one item.
    /// </summary>
    private static List<(int index, string text)> ParseImpressionItems(string impressionText)
    {
        var items = new List<(int, string)>();

        // Match numbered items: "1. text" or "1) text"
        var matches = Regex.Matches(impressionText, @"(?:^|\n)\s*(\d+)[.)]\s*(.+?)(?=\n\s*\d+[.)]|\n\s*$|$)", RegexOptions.Singleline);

        if (matches.Count > 0)
        {
            foreach (Match m in matches)
            {
                int num = int.Parse(m.Groups[1].Value);
                string text = m.Groups[2].Value.Trim();
                if (!string.IsNullOrWhiteSpace(text))
                    items.Add((num, text));
            }
        }
        else
        {
            // No numbered items - treat entire impression as one item
            if (!string.IsNullOrWhiteSpace(impressionText))
                items.Add((1, impressionText.Trim()));
        }

        return items;
    }

    /// <summary>
    /// Parse findings into individual sentences, splitting on period + space or newlines.
    /// Strip subsection headers to get content only.
    /// </summary>
    private static List<string> ParseFindingsSentences(string findingsText)
    {
        var sentences = new List<string>();

        // Split on newlines first to separate subsections
        var lines = findingsText.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line)) continue;

            // Strip subsection header prefix (e.g., "LUNGS AND PLEURA: text" -> "text")
            string content = StripSubsectionHeader(line);
            if (string.IsNullOrWhiteSpace(content)) continue;

            // Don't include pure section headers
            if (IsAllCapsHeader(content)) continue;

            // Split content into sentences on ". " boundary
            var sentenceParts = Regex.Split(content, @"(?<=\.)\s+");
            foreach (var part in sentenceParts)
            {
                var trimmed = part.Trim();
                if (!string.IsNullOrWhiteSpace(trimmed) && trimmed.Length >= 3)
                    sentences.Add(trimmed);
            }
        }

        return sentences;
    }

    /// <summary>
    /// Group findings text into subsection blocks. Each block is a list of sentences
    /// under the same subsection header. Sentences without a subsection header form their own block.
    /// This ensures correlation matches contiguous related sentences, not scattered ones.
    /// </summary>
    private static List<List<string>> GroupFindingsBySubsection(string findingsText)
    {
        var blocks = new List<List<string>>();
        var currentBlock = new List<string>();

        var lines = findingsText.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line)) continue;

            // Check if this line starts a new subsection (ALL CAPS header with colon)
            bool isHeader = false;
            int colonIdx = line.IndexOf(':');
            if (colonIdx > 0 && colonIdx < line.Length - 1)
            {
                var prefix = line.Substring(0, colonIdx);
                isHeader = IsAllCapsHeader(prefix);
            }
            else if (IsAllCapsHeader(line))
            {
                isHeader = true;
            }

            if (isHeader && currentBlock.Count > 0)
            {
                // Save previous block, start new one
                blocks.Add(currentBlock);
                currentBlock = new List<string>();
            }

            // Strip header prefix and split into sentences
            string content = StripSubsectionHeader(line);
            if (string.IsNullOrWhiteSpace(content) || IsAllCapsHeader(content)) continue;

            var sentenceParts = Regex.Split(content, @"(?<=\.)\s+");
            foreach (var part in sentenceParts)
            {
                var trimmed = part.Trim();
                if (!string.IsNullOrWhiteSpace(trimmed) && trimmed.Length >= 3)
                    currentBlock.Add(trimmed);
            }
        }

        if (currentBlock.Count > 0)
            blocks.Add(currentBlock);

        // If no subsection headers were found (simple reports), treat all sentences as one block
        if (blocks.Count == 0 && findingsText.Length > 0)
        {
            var allSentences = ParseFindingsSentences(findingsText);
            if (allSentences.Count > 0)
                blocks.Add(allSentences);
        }

        return blocks;
    }

    private static string StripSubsectionHeader(string line)
    {
        int colonIdx = line.IndexOf(':');
        if (colonIdx <= 0 || colonIdx >= line.Length - 1) return line;

        var prefix = line.Substring(0, colonIdx);
        if (IsAllCapsHeader(prefix))
            return line.Substring(colonIdx + 1).Trim();

        return line;
    }

    private static bool IsAllCapsHeader(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var trimmed = text.Trim().TrimEnd(':');
        if (trimmed.Length < 2) return false;

        int upper = 0, letters = 0;
        foreach (char c in trimmed)
        {
            if (char.IsLetter(c))
            {
                letters++;
                if (char.IsUpper(c)) upper++;
            }
        }
        return letters > 0 && (double)upper / letters > 0.8;
    }

    /// <summary>
    /// Common English stopwords to exclude from content word matching.
    /// </summary>
    private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "about", "above", "after", "again", "along", "also", "appears", "are",
        "been", "before", "being", "below", "between", "both", "cannot",
        "change", "changes", "compatible", "compared", "could",
        "demonstrate", "demonstrates", "demonstrated", "does",
        "each", "either", "evaluate", "evaluation", "evidence",
        "findings", "following", "from", "given", "have", "having",
        "identified", "impression", "including", "interval",
        "large", "likely", "limited", "measure", "measures", "measuring",
        "mild", "mildly", "minimal", "moderate", "moderately",
        "neither", "normal", "noted", "number",
        "other", "otherwise", "overall",
        "partially", "patient", "please", "possible", "possibly",
        "prior", "probably",
        "redemonstrated", "related", "remain", "remains",
        "seen", "series", "severe", "severely", "several", "should",
        "since", "small", "stable", "status", "study", "suggest", "suggests",
        "there", "these", "those", "through", "total",
        "unchanged", "under", "unremarkable",
        "visualized", "well", "where", "which", "within", "without"
    };

    /// <summary>
    /// Simple depluralization: strips trailing 's' or 'es' to normalize plurals.
    /// E.g., "tissues" → "tissue", "masses" → "mass", "calculi" handled by synonyms.
    /// </summary>
    private static string Depluralize(string word)
    {
        if (word.Length <= 3) return word;
        if (word.EndsWith("ses") || word.EndsWith("zes") || word.EndsWith("xes") || word.EndsWith("ches") || word.EndsWith("shes"))
            return word.Substring(0, word.Length - 2);
        if (word.EndsWith("ies"))
            return word.Substring(0, word.Length - 3) + "y";
        if (word.EndsWith("s") && !word.EndsWith("ss") && !word.EndsWith("us") && !word.EndsWith("is"))
            return word.Substring(0, word.Length - 1);
        return word;
    }

    /// <summary>
    /// Extract canonical medical terms from a text fragment.
    /// Returns a set of normalized terms for matching.
    /// Includes: synonym-mapped terms, anatomical terms, and significant content words (5+ chars).
    /// </summary>
    internal static HashSet<string> ExtractTerms(string text) => ExtractTerms(text, null);

    internal static HashSet<string> ExtractTerms(string text, HashSet<string>? contextExclusions)
    {
        var terms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(text)) return terms;

        // Tokenize: split on non-letter characters
        var words = Regex.Split(text.ToLowerInvariant(), @"[^a-z]+")
            .Where(w => w.Length >= 3)
            .ToArray();

        foreach (var word in words)
        {
            // Check synonym dictionary (try both raw word and depluralized form)
            var singular = Depluralize(word);
            if (Synonyms.TryGetValue(word, out var canonical))
            {
                terms.Add(canonical);
            }
            else if (singular != word && Synonyms.TryGetValue(singular, out var canonical2))
            {
                terms.Add(canonical2);
            }

            // Check if word itself is an anatomical term
            if (AnatomicalTerms.Contains(word))
            {
                terms.Add(word);
            }
            else if (singular != word && AnatomicalTerms.Contains(singular))
            {
                terms.Add(singular);
            }

            // Include significant content words (5+ chars, not stopwords)
            // This catches medical terms not in our dictionaries
            // Use depluralized form so "tissues" and "tissue" match
            if (word.Length >= 5 && !StopWords.Contains(word))
            {
                terms.Add(singular);
            }
        }

        // Also check multi-word anatomical terms
        var lowerText = text.ToLowerInvariant();
        foreach (var term in AnatomicalTerms)
        {
            if (term.Contains(' ') && lowerText.Contains(term))
            {
                terms.Add(term);
            }
        }

        // Check multi-word synonyms
        foreach (var kvp in Synonyms)
        {
            if (kvp.Key.Contains(' ') && lowerText.Contains(kvp.Key.ToLowerInvariant()))
            {
                terms.Add(kvp.Value);
            }
        }

        // Remove exam context terms (e.g., "lumbar", "spine" in a lumbar spine study)
        if (contextExclusions != null && contextExclusions.Count > 0)
        {
            terms.ExceptWith(contextExclusions);
        }

        return terms;
    }
}
