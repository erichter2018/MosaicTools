// [CustomSTT] Ensemble merger: collects finals from all 3 providers and produces a merged result.
namespace MosaicTools.Services;

/// <summary>
/// Snapshot of ensemble performance stats for live display.
/// </summary>
public record EnsembleStats(
    int TotalWords, int LowConfidenceWords, int CorrectedWords,
    double AverageConfidence, int MergeCount,
    int S1Arrivals, int S2Arrivals,
    int S1Corrections, int S2Corrections, int ConsensusCorrections,
    int S1Confirms, int S2Confirms,
    string? LastCorrectionDetail);

/// <summary>
/// Collects final results from Deepgram (primary), AssemblyAI, and Speechmatics.
/// Uses word-level alignment + voting to correct low-confidence Deepgram words.
/// Post-processes each provider's transcript before comparison for apples-to-apples matching.
/// </summary>
public class SttEnsembleMerger
{
    private readonly Configuration _config;
    private readonly int _waitMs;
    private readonly double _confidenceThreshold;
    private readonly object _lock = new();

    // Current pending merge
    private SttResult? _primaryResult;
    private SttResult? _secondary1Result;
    private SttResult? _secondary2Result;
    private CancellationTokenSource? _timerCts;

    // Session stats
    private int _totalWords;
    private int _lowConfidenceWords;
    private int _correctedWords;
    private double _totalConfidence;   // Sum of all word confidences (for average)
    private int _mergeCount;           // Number of merge operations
    private int _s1Arrivals;           // AssemblyAI finals received
    private int _s2Arrivals;           // Speechmatics finals received

    // Per-provider contribution stats
    private int _s1Corrections;        // Times AAI word replaced Deepgram
    private int _s2Corrections;        // Times SM word replaced Deepgram
    private int _consensusCorrections; // Both secondaries agreed → overruled Deepgram
    private int _s1Confirms;           // AAI agreed with low-conf Deepgram word (validated it)
    private int _s2Confirms;           // SM agreed with low-conf Deepgram word (validated it)
    private string? _lastCorrectionDetail; // e.g. "pole → poll (AAI+SM)"

    public int TotalWords => _totalWords;
    public int LowConfidenceWords => _lowConfidenceWords;
    public int CorrectedWords => _correctedWords;
    public double AverageConfidence => _totalWords > 0 ? _totalConfidence / _totalWords : 0;
    public int MergeCount => _mergeCount;
    public int S1Arrivals => _s1Arrivals;
    public int S2Arrivals => _s2Arrivals;

    public event Action<SttResult>? MergedResultReady;
    /// <summary>
    /// Fires after every merge with a snapshot of current stats for live display.
    /// </summary>
    public event Action<EnsembleStats>? StatsUpdated;

    public SttEnsembleMerger(Configuration config, int waitMs = 500, double confidenceThreshold = 0.80)
    {
        _config = config;
        _waitMs = waitMs;
        _confidenceThreshold = confidenceThreshold;
    }

    public void SubmitResult(SttResult result)
    {
        lock (_lock)
        {
            switch (result.ProviderName)
            {
                case "deepgram":
                    _primaryResult = result;
                    break;
                case "assemblyai":
                    _secondary1Result = result;
                    Interlocked.Increment(ref _s1Arrivals);
                    break;
                case "speechmatics":
                    _secondary2Result = result;
                    Interlocked.Increment(ref _s2Arrivals);
                    break;
                default:
                    return;
            }

            // Primary hasn't arrived yet — wait
            if (_primaryResult == null) return;

            // Check if all words are high-confidence — emit immediately
            if (AllWordsHighConfidence(_primaryResult))
            {
                CancelTimer();
                EmitMerged();
                return;
            }

            // Check if all 3 have arrived
            if (_secondary1Result != null && _secondary2Result != null)
            {
                CancelTimer();
                EmitMerged();
                return;
            }

            // Start timer if not already running
            if (_timerCts == null)
            {
                var cts = new CancellationTokenSource();
                _timerCts = cts;
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await Task.Delay(_waitMs, cts.Token);
                        lock (_lock)
                        {
                            if (!cts.IsCancellationRequested)
                                EmitMerged();
                        }
                    }
                    catch (OperationCanceledException) { }
                });
            }
        }
    }

    public void ResetStats()
    {
        _totalWords = 0;
        _lowConfidenceWords = 0;
        _correctedWords = 0;
        _totalConfidence = 0;
        _mergeCount = 0;
        _s1Arrivals = 0;
        _s2Arrivals = 0;
        _s1Corrections = 0;
        _s2Corrections = 0;
        _consensusCorrections = 0;
        _s1Confirms = 0;
        _s2Confirms = 0;
        _lastCorrectionDetail = null;
    }

    private bool AllWordsHighConfidence(SttResult result)
    {
        if (result.Words == null || result.Words.Length == 0) return true;
        foreach (var w in result.Words)
            if (w.Confidence < _confidenceThreshold) return false;
        return true;
    }

    private void CancelTimer()
    {
        _timerCts?.Cancel();
        _timerCts?.Dispose();
        _timerCts = null;
    }

    /// <summary>
    /// Must be called under _lock.
    /// </summary>
    private void EmitMerged()
    {
        var primary = _primaryResult;
        if (primary == null) return;

        // Clear pending state
        var s1 = _secondary1Result;
        var s2 = _secondary2Result;
        _primaryResult = null;
        _secondary1Result = null;
        _secondary2Result = null;
        _timerCts = null;

        Interlocked.Increment(ref _mergeCount);

        // If no secondaries arrived, just post-process and emit primary
        if (s1 == null && s2 == null)
        {
            var processed = SttTextProcessor.ProcessTranscript(primary.Transcript, _config);
            var wc = primary.Words?.Length ?? 0;
            Interlocked.Add(ref _totalWords, wc);
            if (primary.Words != null)
                foreach (var w in primary.Words)
                    AddConfidence(w.Confidence);
            MergedResultReady?.Invoke(primary with
            {
                Transcript = processed,
                ProviderName = "ensemble"
            });
            FireStatsUpdated();
            return;
        }

        // Post-process all transcripts for normalized comparison
        var primaryProcessed = SttTextProcessor.ProcessTranscript(primary.Transcript, _config);
        var s1Processed = s1 != null ? SttTextProcessor.ProcessTranscript(s1.Transcript, _config) : null;
        var s2Processed = s2 != null ? SttTextProcessor.ProcessTranscript(s2.Transcript, _config) : null;

        // Build processed word arrays
        var primaryWords = TokenizeProcessed(primaryProcessed);
        var s1Words = s1Processed != null ? TokenizeProcessed(s1Processed) : Array.Empty<string>();
        var s2Words = s2Processed != null ? TokenizeProcessed(s2Processed) : Array.Empty<string>();

        // Merge: use primary timing structure, replace low-confidence words via voting
        var mergedWords = new List<string>();
        int corrections = 0;
        int lowConf = 0;
        var originalWords = primary.Words ?? Array.Empty<SttWord>();

        for (int i = 0; i < primaryWords.Length; i++)
        {
            var word = primaryWords[i];
            var origWord = i < originalWords.Length ? originalWords[i] : null;
            var confidence = origWord?.Confidence ?? 1f;

            Interlocked.Increment(ref _totalWords);
            AddConfidence(confidence);

            if (confidence >= _confidenceThreshold)
            {
                mergedWords.Add(word);
                continue;
            }

            lowConf++;
            Interlocked.Increment(ref _lowConfidenceWords);

            // Find overlapping words from secondaries by position ratio
            var posRatio = primaryWords.Length > 1 ? (double)i / (primaryWords.Length - 1) : 0.5;
            var s1Match = FindMatchByPosition(s1Words, posRatio);
            var s2Match = FindMatchByPosition(s2Words, posRatio);

            // Also try timestamp overlap if original word has timing
            if (origWord != null && origWord.StartTime > 0)
            {
                var s1TimedMatch = FindMatchByTimestamp(s1?.Words, origWord.StartTime, origWord.EndTime);
                var s2TimedMatch = FindMatchByTimestamp(s2?.Words, origWord.StartTime, origWord.EndTime);
                if (s1TimedMatch != null) s1Match = s1TimedMatch;
                if (s2TimedMatch != null) s2Match = s2TimedMatch;
            }

            // Vote: 2-of-3 consensus wins, all differ → highest confidence
            var winner = Vote(word, confidence, s1Match, s2Match);

            if (!string.Equals(winner, word, StringComparison.OrdinalIgnoreCase))
            {
                // A correction happened — track which provider(s) contributed
                corrections++;
                Interlocked.Increment(ref _correctedWords);

                bool fromS1 = s1Match != null && string.Equals(winner, s1Match, StringComparison.OrdinalIgnoreCase);
                bool fromS2 = s2Match != null && string.Equals(winner, s2Match, StringComparison.OrdinalIgnoreCase);

                if (fromS1 && fromS2)
                {
                    Interlocked.Increment(ref _consensusCorrections);
                    Interlocked.Increment(ref _s1Corrections);
                    Interlocked.Increment(ref _s2Corrections);
                }
                else if (fromS1)
                    Interlocked.Increment(ref _s1Corrections);
                else if (fromS2)
                    Interlocked.Increment(ref _s2Corrections);

                var source = fromS1 && fromS2 ? "AAI+SM" : fromS1 ? "AAI" : fromS2 ? "SM" : "?";
                _lastCorrectionDetail = $"\"{word}\" \u2192 \"{winner}\" ({source})";

                Logger.Trace($"Ensemble merge: \"{word}\" ({confidence:P0}) \u2192 \"{winner}\" [s1={s1Match ?? "?"}, s2={s2Match ?? "?"}] source={source}");
            }
            else
            {
                // Primary word kept — track which secondaries confirmed it
                if (s1Match != null && string.Equals(word, s1Match, StringComparison.OrdinalIgnoreCase))
                    Interlocked.Increment(ref _s1Confirms);
                if (s2Match != null && string.Equals(word, s2Match, StringComparison.OrdinalIgnoreCase))
                    Interlocked.Increment(ref _s2Confirms);
            }

            mergedWords.Add(winner);
        }

        var mergedTranscript = string.Join(" ", mergedWords);

        if (corrections > 0)
            Logger.Trace($"Ensemble merge: {corrections} corrections out of {lowConf} low-confidence words");

        MergedResultReady?.Invoke(primary with
        {
            Transcript = mergedTranscript,
            ProviderName = "ensemble"
        });
        FireStatsUpdated();
    }

    private static string[] TokenizeProcessed(string text)
    {
        return text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
    }

    private static string? FindMatchByPosition(string[] words, double posRatio)
    {
        if (words.Length == 0) return null;
        var idx = (int)Math.Round(posRatio * (words.Length - 1));
        idx = Math.Clamp(idx, 0, words.Length - 1);
        return words[idx];
    }

    private static string? FindMatchByTimestamp(SttWord[]? words, double startTime, double endTime)
    {
        if (words == null || words.Length == 0) return null;

        string? bestMatch = null;
        double bestOverlap = 0;

        foreach (var w in words)
        {
            if (w.EndTime <= startTime || w.StartTime >= endTime) continue;
            var overlapStart = Math.Max(w.StartTime, startTime);
            var overlapEnd = Math.Min(w.EndTime, endTime);
            var overlap = overlapEnd - overlapStart;
            var duration = endTime - startTime;
            if (duration <= 0) continue;
            var ratio = overlap / duration;
            if (ratio > 0.40 && ratio > bestOverlap)
            {
                bestOverlap = ratio;
                bestMatch = w.Text;
            }
        }

        return bestMatch;
    }

    private static string Vote(string primary, float primaryConf,
        string? secondary1, string? secondary2)
    {
        // 2-of-3 consensus
        if (secondary1 != null && string.Equals(primary, secondary1, StringComparison.OrdinalIgnoreCase))
            return primary;
        if (secondary2 != null && string.Equals(primary, secondary2, StringComparison.OrdinalIgnoreCase))
            return primary;
        if (secondary1 != null && secondary2 != null &&
            string.Equals(secondary1, secondary2, StringComparison.OrdinalIgnoreCase))
            return secondary1;

        // All differ → highest confidence wins. Without per-word confidence from secondaries,
        // prefer any secondary over a low-confidence primary.
        if (secondary1 != null && secondary2 != null)
        {
            // Pick whichever secondary seems more reasonable (prefer longer, less likely truncated)
            return secondary1.Length >= secondary2.Length ? secondary1 : secondary2;
        }
        if (secondary1 != null) return secondary1;
        if (secondary2 != null) return secondary2;

        return primary; // No secondaries available
    }

    private void AddConfidence(double conf)
    {
        // Not thread-critical — approximate is fine for display
        _totalConfidence += conf;
    }

    private void FireStatsUpdated()
    {
        StatsUpdated?.Invoke(new EnsembleStats(
            _totalWords, _lowConfidenceWords, _correctedWords,
            AverageConfidence, _mergeCount,
            _s1Arrivals, _s2Arrivals,
            _s1Corrections, _s2Corrections, _consensusCorrections,
            _s1Confirms, _s2Confirms,
            _lastCorrectionDetail));
    }
}
