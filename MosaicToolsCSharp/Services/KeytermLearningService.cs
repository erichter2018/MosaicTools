using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace MosaicTools.Services;

public class KeytermLearningData
{
    [JsonPropertyName("version")]
    public int Version { get; set; } = 1;

    [JsonPropertyName("entries")]
    public Dictionary<string, KeytermEntry> Entries { get; set; } = new();
}

public class KeytermEntry
{
    [JsonPropertyName("word")]
    public string Word { get; set; } = "";

    [JsonPropertyName("source")]
    public string Source { get; set; } = "survived";

    [JsonPropertyName("occurrences")]
    public int Occurrences { get; set; }

    [JsonPropertyName("total_confidence")]
    public double TotalConfidence { get; set; }

    [JsonPropertyName("first_seen")]
    public string FirstSeen { get; set; } = "";

    [JsonPropertyName("last_seen")]
    public string LastSeen { get; set; } = "";
}

public class LowConfidenceCapture
{
    public string Word { get; set; } = "";
    public string OriginalForm { get; set; } = "";
    public float Confidence { get; set; }
    public string[] ContextBefore { get; set; } = Array.Empty<string>();
    public string[] ContextAfter { get; set; } = Array.Empty<string>();
}

public class KeytermLearningService
{
    private static readonly string DataPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MosaicTools", "KeytermLearning.json");

    private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "the", "a", "an", "is", "are", "was", "were", "be", "been", "being",
        "have", "has", "had", "do", "does", "did", "will", "would", "shall",
        "should", "may", "might", "must", "can", "could",
        "of", "in", "to", "for", "with", "on", "at", "from", "by", "as",
        "into", "through", "during", "before", "after", "above", "below",
        "between", "out", "off", "over", "under", "again", "further", "then",
        "once", "here", "there", "when", "where", "why", "how", "all", "each",
        "every", "both", "few", "more", "most", "other", "some", "such",
        "no", "nor", "not", "only", "own", "same", "so", "than", "too",
        "very", "just", "about", "up", "down",
        "and", "but", "or", "if", "while", "because", "until", "although",
        "i", "me", "my", "we", "our", "you", "your", "he", "him", "his",
        "she", "her", "it", "its", "they", "them", "their",
        "this", "that", "these", "those", "what", "which", "who", "whom",
        "also", "still", "well", "back", "even", "new", "old", "right", "left"
    };

    private KeytermLearningData _data = new();
    private readonly List<LowConfidenceCapture> _sessionBuffer = new();
    private readonly object _lock = new();

    public void CollectLowConfidenceWords(SttWord[] words, float threshold)
    {
        if (words.Length == 0) return;

        lock (_lock)
        {
            for (int i = 0; i < words.Length; i++)
            {
                var w = words[i];
                if (w.Confidence >= threshold) continue;

                var clean = CleanWord(w.Text);
                if (!IsUsefulWord(clean)) continue;

                // Gather context: up to 2 high-confidence words before and after
                var before = new List<string>();
                for (int b = Math.Max(0, i - 2); b < i; b++)
                {
                    if (words[b].Confidence >= threshold)
                        before.Add(CleanWord(words[b].Text));
                }

                var after = new List<string>();
                for (int a = i + 1; a <= Math.Min(words.Length - 1, i + 2); a++)
                {
                    if (words[a].Confidence >= threshold)
                        after.Add(CleanWord(words[a].Text));
                }

                _sessionBuffer.Add(new LowConfidenceCapture
                {
                    Word = clean.ToLowerInvariant(),
                    OriginalForm = w.PunctuatedText ?? w.Text,
                    Confidence = w.Confidence,
                    ContextBefore = before.ToArray(),
                    ContextAfter = after.ToArray()
                });
            }
        }

        if (_sessionBuffer.Count > 0)
            Logger.Trace($"KeytermLearning: Collected {_sessionBuffer.Count} low-confidence words this session");
    }

    public void VerifyAgainstReport(string finalReportText)
    {
        List<LowConfidenceCapture> captures;
        lock (_lock)
        {
            if (_sessionBuffer.Count == 0) return;
            captures = new List<LowConfidenceCapture>(_sessionBuffer);
        }

        var now = DateTime.UtcNow.ToString("o");
        int survived = 0, corrected = 0, discarded = 0;

        // Normalize report for matching
        var reportNorm = NormalizeForSearch(finalReportText);

        foreach (var cap in captures)
        {
            // 1. Survival check: word exists in final report
            var wordPattern = @"\b" + Regex.Escape(cap.Word) + @"\b";
            if (Regex.IsMatch(reportNorm, wordPattern, RegexOptions.IgnoreCase))
            {
                RecordEntry(cap.Word, cap.OriginalForm, "survived", cap.Confidence, now);
                survived++;
                continue;
            }

            // 2. Context anchor correction
            var correction = FindCorrectionViaAnchors(cap, reportNorm);
            if (correction != null)
            {
                RecordEntry(correction.ToLowerInvariant(), correction, "corrected", cap.Confidence, now);
                corrected++;
            }
            else
            {
                discarded++;
            }
        }

        if (survived + corrected > 0)
        {
            Save();
            Logger.Trace($"KeytermLearning: Verified {captures.Count} words — {survived} survived, {corrected} corrected, {discarded} discarded ({_data.Entries.Count} total entries)");
        }
    }

    public void ClearSessionBuffer()
    {
        lock (_lock)
        {
            _sessionBuffer.Clear();
        }
    }

    public List<string> GetTopKeyterms(int maxCount)
    {
        if (maxCount <= 0) return new List<string>();

        lock (_lock)
        {
            return _data.Entries.Values
                .OrderByDescending(e => ComputeScore(e))
                .Take(maxCount)
                .Select(e => e.Word)
                .ToList();
        }
    }

    public IReadOnlyList<(string Word, double Score, double AvgConfidence, string Source, int Occurrences)> GetAllEntries()
    {
        lock (_lock)
        {
            return _data.Entries.Values
                .OrderByDescending(e => ComputeScore(e))
                .Select(e => (e.Word, ComputeScore(e),
                    e.Occurrences > 0 ? e.TotalConfidence / e.Occurrences : 0.0,
                    e.Source, e.Occurrences))
                .ToList();
        }
    }

    public int EntryCount
    {
        get { lock (_lock) { return _data.Entries.Count; } }
    }

    public void Load()
    {
        try
        {
            if (!File.Exists(DataPath))
            {
                _data = new KeytermLearningData();
                return;
            }

            var json = File.ReadAllText(DataPath);
            _data = JsonSerializer.Deserialize<KeytermLearningData>(json) ?? new KeytermLearningData();
            Logger.Trace($"KeytermLearning: Loaded {_data.Entries.Count} entries from {DataPath}");
        }
        catch (Exception ex)
        {
            Logger.Trace($"KeytermLearning: Load error: {ex.Message}");
            _data = new KeytermLearningData();
        }
    }

    public void Save()
    {
        try
        {
            var dir = Path.GetDirectoryName(DataPath);
            if (dir != null && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(_data, new JsonSerializerOptions
            {
                WriteIndented = true
            });
            File.WriteAllText(DataPath, json);
        }
        catch (Exception ex)
        {
            Logger.Trace($"KeytermLearning: Save error: {ex.Message}");
        }
    }

    private void RecordEntry(string key, string displayForm, string source, float confidence, string now)
    {
        var lowerKey = key.ToLowerInvariant();

        lock (_lock)
        {
            if (_data.Entries.TryGetValue(lowerKey, out var existing))
            {
                existing.Occurrences++;
                existing.TotalConfidence += confidence;
                existing.LastSeen = now;
                // Upgrade source if we found a correction (more valuable)
                if (source == "corrected") existing.Source = "corrected";
            }
            else
            {
                _data.Entries[lowerKey] = new KeytermEntry
                {
                    Word = displayForm.Trim(),
                    Source = source,
                    Occurrences = 1,
                    TotalConfidence = confidence,
                    FirstSeen = now,
                    LastSeen = now
                };
            }
        }
    }

    private static double ComputeScore(KeytermEntry entry)
    {
        var avgConfidence = entry.Occurrences > 0 ? entry.TotalConfidence / entry.Occurrences : 0.5;
        var gap = 1.0 - avgConfidence;
        var sourceWeight = entry.Source == "corrected" ? 2.0 : 1.0;
        return Math.Log(1 + entry.Occurrences) * (gap * gap) * sourceWeight;
    }

    private static string? FindCorrectionViaAnchors(LowConfidenceCapture cap, string reportNorm)
    {
        // Need at least one before-anchor and one after-anchor
        if (cap.ContextBefore.Length == 0 || cap.ContextAfter.Length == 0) return null;

        // Build pattern: beforeAnchors ... (captured word(s)) ... afterAnchors
        // Use the last before-anchor and first after-anchor for tightest match
        var beforeAnchor = Regex.Escape(cap.ContextBefore[^1]);
        var afterAnchor = Regex.Escape(cap.ContextAfter[0]);

        // Match 1-3 words between anchors (the correction)
        var pattern = @"\b" + beforeAnchor + @"\b\s+([\w'-]+(?:\s+[\w'-]+){0,2})\s+\b" + afterAnchor + @"\b";

        var match = Regex.Match(reportNorm, pattern, RegexOptions.IgnoreCase);
        if (!match.Success) return null;

        var extracted = match.Groups[1].Value.Trim();

        // Don't return if it's the same word we started with
        if (string.Equals(extracted, cap.Word, StringComparison.OrdinalIgnoreCase)) return null;

        // Don't return stop words or very short words
        if (!IsUsefulWord(extracted)) return null;

        return extracted;
    }

    private static string CleanWord(string word)
    {
        // Strip trailing/leading punctuation but keep hyphens and apostrophes internal to the word
        return Regex.Replace(word, @"^[^\w'-]+|[^\w'-]+$", "");
    }

    private static string NormalizeForSearch(string text)
    {
        // Collapse whitespace for easier matching
        return Regex.Replace(text, @"\s+", " ").Trim();
    }

    private static bool IsUsefulWord(string word)
    {
        if (string.IsNullOrWhiteSpace(word)) return false;
        if (word.Length <= 2) return false;
        if (StopWords.Contains(word)) return false;
        // Pure numbers
        if (double.TryParse(word, out _)) return false;
        return true;
    }
}
