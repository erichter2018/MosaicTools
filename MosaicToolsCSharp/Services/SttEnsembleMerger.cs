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
    string? LastCorrectionDetail,
    // All-time (persisted across sessions)
    int AlltimeWords, int AlltimeCorrected,
    double AlltimeAverageConfidence, int AlltimeMerges,
    // Correction validation (verified against signed reports)
    int SessionValidated, int SessionRejected,
    int AlltimeValidated, int AlltimeRejected);

/// <summary>
/// Records a single ensemble correction for later validation against the signed report.
/// </summary>
public record CorrectionRecord(string Original, string Replacement, string Source);

/// <summary>
/// Collects final results from Deepgram (primary), Soniox, and Speechmatics.
/// Uses word-level alignment + voting to correct low-confidence Deepgram words.
/// Post-processes each provider's transcript before comparison for apples-to-apples matching.
/// </summary>
public class SttEnsembleMerger
{
    private const int BonusWaitMs = 150; // Short wait even for high-confidence phrases

    private readonly Configuration _config;
    private readonly int _waitMs;
    private readonly double _confidenceThreshold;
    private readonly string _s1Name; // Secondary provider 1 name (e.g. "soniox")
    private readonly string _s2Name; // Secondary provider 2 name (e.g. "speechmatics")
    private readonly object _lock = new();

    // Current pending merge
    private SttResult? _primaryResult;
    private long _primaryArrivedTicks; // Wall-clock arrival time for staleness detection
    private readonly List<(SttResult Result, long ArrivedTicks)> _secondary1Buffer = new();
    private readonly List<(SttResult Result, long ArrivedTicks)> _secondary2Buffer = new();
    private CancellationTokenSource? _timerCts;

    // Session stats
    private int _totalWords;
    private int _lowConfidenceWords;
    private int _correctedWords;
    private double _totalConfidence;   // Sum of all word confidences (for average)
    private int _mergeCount;           // Number of merge operations
    private int _s1Arrivals;           // Soniox finals received
    private int _s2Arrivals;           // Speechmatics finals received

    // Per-provider contribution stats
    private int _s1Corrections;        // Times Soniox word replaced Deepgram
    private int _s2Corrections;        // Times SM word replaced Deepgram
    private int _consensusCorrections; // Both secondaries agreed → overruled Deepgram
    private int _s1Confirms;           // Soniox agreed with low-conf Deepgram word (validated it)
    private int _s2Confirms;           // SM agreed with low-conf Deepgram word (validated it)
    private string? _lastCorrectionDetail; // e.g. "pole → poll (SNX+SM)"

    // Per-study delta — only committed to all-time when report is signed
    private int _studyWords;
    private int _studyCorrected;
    private double _studyConfSum;
    private int _studyMerges;

    public int TotalWords => _totalWords;
    public int LowConfidenceWords => _lowConfidenceWords;
    public int CorrectedWords => _correctedWords;
    public double AverageConfidence => _totalWords > 0 ? _totalConfidence / _totalWords : 0;
    public int MergeCount => _mergeCount;
    public int S1Arrivals => _s1Arrivals;
    public int S2Arrivals => _s2Arrivals;
    public string S1Name => _s1Name;
    public string S2Name => _s2Name;

    private static string Short(string n) => n switch
    {
        "soniox" => "SNX", "speechmatics" => "SM", "assemblyai" => "AAI", "none" => "—",
        _ => n.ToUpperInvariant()[..Math.Min(3, n.Length)]
    };

    public event Action<SttResult>? MergedResultReady;
    /// <summary>
    /// Fires after every merge with a snapshot of current stats for live display.
    /// </summary>
    public event Action<EnsembleStats>? StatsUpdated;
    /// <summary>
    /// Fires after a merge that produced ≥1 correction, for later validation against signed report.
    /// </summary>
    public event Action<List<CorrectionRecord>>? CorrectionsEmitted;

    public SttEnsembleMerger(Configuration config, int waitMs = 500, double confidenceThreshold = 0.80,
        string s1Name = "soniox", string s2Name = "speechmatics")
    {
        _config = config;
        _waitMs = waitMs;
        _confidenceThreshold = confidenceThreshold;
        _s1Name = s1Name;
        _s2Name = s2Name;
    }

    public void SubmitResult(SttResult result)
    {
        lock (_lock)
        {
            var name = result.ProviderName;
            var now = Environment.TickCount64;
            if (name == "deepgram")
            {
                _primaryResult = result;
                _primaryArrivedTicks = now;
                // Purge secondary entries from previous utterance (arrived >2s before this primary)
                PurgeStaleEntries(_secondary1Buffer, now - 2000);
                PurgeStaleEntries(_secondary2Buffer, now - 2000);
            }
            else if (name == _s1Name)
            {
                _secondary1Buffer.Add((result, now));
                Interlocked.Increment(ref _s1Arrivals);
            }
            else if (name == _s2Name)
            {
                _secondary2Buffer.Add((result, now));
                Interlocked.Increment(ref _s2Arrivals);
            }
            else
            {
                return;
            }

            // Primary hasn't arrived yet — accumulate secondaries
            if (_primaryResult == null) return;

            // Check if all active secondaries have contributed — emit immediately
            // Only count entries that arrived within 1.5s of primary (avoid stale cross-utterance data)
            var s1Ready = _s1Name == "none" || HasFreshEntries(_secondary1Buffer);
            var s2Ready = _s2Name == "none" || HasFreshEntries(_secondary2Buffer);
            if (s1Ready && s2Ready)
            {
                Logger.Trace($"Ensemble: all 3 providers arrived ({Short(_s1Name)}={_secondary1Buffer.Count} finals, {Short(_s2Name)}={_secondary2Buffer.Count} finals) — merging immediately");
                CancelTimer();
                EmitMerged();
                return;
            }

            // Determine wait time: short bonus wait for high-confidence (secondaries
            // might still catch confident-but-wrong words like "anatomy" → "sternotomy"),
            // full wait for low-confidence words that need secondary correction
            var allHighConf = AllWordsHighConfidence(_primaryResult);
            var delay = allHighConf ? BonusWaitMs : _waitMs;

            // Start timer if not already running
            if (_timerCts == null)
            {
                Logger.Trace($"Ensemble: starting {delay}ms timer ({(allHighConf ? "bonus" : "full")} wait), {Short(_s1Name)}={_secondary1Buffer.Count} finals, {Short(_s2Name)}={_secondary2Buffer.Count} finals");
                var cts = new CancellationTokenSource();
                _timerCts = cts;
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await Task.Delay(delay, cts.Token);
                        lock (_lock)
                        {
                            if (!cts.IsCancellationRequested)
                            {
                                Logger.Trace($"Ensemble: timer expired ({delay}ms), {Short(_s1Name)}={_secondary1Buffer.Count} finals, {Short(_s2Name)}={_secondary2Buffer.Count} finals");
                                EmitMerged();
                            }
                        }
                    }
                    catch (OperationCanceledException) { }
                    finally { cts.Dispose(); }
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
        _timerCts = null; // Disposal handled by lambda's finally block
    }

    private static void PurgeStaleEntries(List<(SttResult Result, long ArrivedTicks)> buffer, long cutoffTicks)
    {
        var removed = buffer.RemoveAll(b => b.ArrivedTicks < cutoffTicks);
        if (removed > 0)
            Logger.Trace($"Ensemble: purged {removed} stale secondary buffer entries");
    }

    private bool HasFreshEntries(List<(SttResult Result, long ArrivedTicks)> buffer)
    {
        if (buffer.Count == 0) return false;
        // At least one entry arrived within 1.5s of primary arrival
        var cutoff = _primaryArrivedTicks - 1500;
        return buffer.Exists(b => b.ArrivedTicks >= cutoff);
    }

    /// <summary>
    /// Combine buffered secondary finals into a single SttResult with merged transcript and word arrays.
    /// Speechmatics sends word-by-word finals, so we may have 10+ fragments that need concatenation.
    /// </summary>
    private static SttResult? CombineBufferedResults(List<(SttResult Result, long ArrivedTicks)> buffer)
    {
        if (buffer.Count == 0) return null;
        if (buffer.Count == 1) return buffer[0].Result;

        // Concatenate transcripts and word arrays
        var allWords = new List<SttWord>();
        var transcriptParts = new List<string>();

        foreach (var (r, _) in buffer)
        {
            if (!string.IsNullOrEmpty(r.Transcript))
                transcriptParts.Add(r.Transcript);
            if (r.Words != null)
                allWords.AddRange(r.Words);
        }

        var combined = new SttResult(
            Transcript: string.Join(" ", transcriptParts),
            Words: allWords.ToArray(),
            Confidence: allWords.Count > 0 ? allWords.Average(w => w.Confidence) : 0f,
            IsFinal: true,
            SpeechFinal: true,
            Duration: allWords.Count > 0 ? allWords.Max(w => w.EndTime) - allWords.Min(w => w.StartTime) : 0,
            ProviderName: buffer[0].Result.ProviderName
        );

        Logger.Trace($"Ensemble: combined {buffer.Count} {buffer[0].Result.ProviderName} finals → {allWords.Count} words, \"{combined.Transcript}\"");
        return combined;
    }

    /// <summary>
    /// Must be called under _lock.
    /// </summary>
    private void EmitMerged()
    {
        var primary = _primaryResult;
        if (primary == null) return;

        // Combine buffered secondaries into single results and clear state
        var s1 = CombineBufferedResults(_secondary1Buffer);
        var s2 = CombineBufferedResults(_secondary2Buffer);
        _primaryResult = null;
        _secondary1Buffer.Clear();
        _secondary2Buffer.Clear();
        _timerCts = null;

        Interlocked.Increment(ref _mergeCount);
        _studyMerges++;

        var s1Status = s1 != null ? $"{s1.Words?.Length ?? 0} words" : "missing";
        var s2Status = s2 != null ? $"{s2.Words?.Length ?? 0} words" : "missing";
        Logger.Trace($"Ensemble merge #{_mergeCount}: DG={primary.Words?.Length ?? 0} words, {Short(_s1Name)}={s1Status}, {Short(_s2Name)}={s2Status}");

        // If no secondaries arrived, just post-process and emit primary
        if (s1 == null && s2 == null)
        {
            Logger.Trace("Ensemble merge: no secondaries arrived, emitting primary only");
            var processed = SttTextProcessor.ProcessTranscript(primary.Transcript, _config);
            var wc = primary.Words?.Length ?? 0;
            Interlocked.Add(ref _totalWords, wc);
            _studyWords += wc;
            if (primary.Words != null)
                foreach (var w in primary.Words)
                {
                    AddConfidence(w.Confidence);
                    _studyConfSum += w.Confidence;
                }
            MergedResultReady?.Invoke(primary with
            {
                Transcript = processed,
                ProviderName = "ensemble"
            });
            FireStatsUpdated();
            return;
        }

        // Post-process primary transcript for output
        var primaryProcessed = SttTextProcessor.ProcessTranscript(primary.Transcript, _config);

        // Build processed primary word array (determines output structure)
        var primaryWords = TokenizeProcessed(primaryProcessed);

        // Secondary word arrays for timestamp-based alignment (raw words with times)
        var s1WordArr = s1?.Words ?? Array.Empty<SttWord>();
        var s2WordArr = s2?.Words ?? Array.Empty<SttWord>();
        var originalWords = primary.Words ?? Array.Empty<SttWord>();

        // Align secondary words to primary using monotonic nearest-midpoint matching.
        // This handles the ~200ms timestamp offset between providers (SM/AAI vs DG)
        // that caused per-word overlap matching to consistently hit the adjacent word.
        var s1Aligned = AlignWordsByTimestamp(originalWords, s1WordArr);
        var s2Aligned = AlignWordsByTimestamp(originalWords, s2WordArr);

        // Merge: use primary timing structure, replace low-confidence words via voting
        var mergedWords = new List<string>();
        var correctionRecords = new List<CorrectionRecord>();
        int corrections = 0;
        int lowConf = 0;

        for (int i = 0; i < primaryWords.Length; i++)
        {
            var word = primaryWords[i];
            var origWord = i < originalWords.Length ? originalWords[i] : null;
            var confidence = origWord?.Confidence ?? 1f;

            Interlocked.Increment(ref _totalWords);
            _studyWords++;
            AddConfidence(confidence);
            _studyConfSum += confidence;

            // Get pre-aligned secondary matches (full SttWord for confidence access)
            var s1Word = i < s1Aligned.Length ? s1Aligned[i] : null;
            var s2Word = i < s2Aligned.Length ? s2Aligned[i] : null;
            var s1Match = StripTrailingPunctuation(s1Word?.Text);
            var s2Match = StripTrailingPunctuation(s2Word?.Text);

            bool isLowConf = confidence < _confidenceThreshold;
            bool hasSecondaries = s1Match != null || s2Match != null;

            // High-confidence with no secondaries available — skip matching
            if (!isLowConf && !hasSecondaries)
            {
                mergedWords.Add(word);
                continue;
            }

            if (isLowConf)
            {
                lowConf++;
                Interlocked.Increment(ref _lowConfidenceWords);
            }

            // Phrase-anchored context validation: check ±2 words around target to confirm
            // the alignment is correct before trusting the comparison. This prevents false
            // corrections from misaligned secondaries (e.g., off-by-one timestamp drift).
            var s1Anchors = s1Match != null ? CountContextAnchors(originalWords, s1Aligned, i) : 0;
            var s2Anchors = s2Match != null ? CountContextAnchors(originalWords, s2Aligned, i) : 0;
            var contextPositions = CountContextPositions(i, originalWords.Length);

            // A secondary is "trusted" if it has context anchors confirming alignment,
            // or if the phrase is a single word (no context positions to check at all)
            bool s1Trusted = s1Match != null && (s1Anchors >= 1 || contextPositions == 0);
            bool s2Trusted = s2Match != null && (s2Anchors >= 1 || contextPositions == 0);

            // Log word-level comparison with timestamps and anchor info
            var confTag = isLowConf ? "LOW" : "hi";
            var timeTag = origWord != null && origWord.StartTime > 0 ? $" {origWord.StartTime:F2}-{origWord.EndTime:F2}s" : "";
            Logger.Trace($"Ensemble word[{i}]: \"{word}\" ({confidence:F2} {confTag}{timeTag}) vs s1=\"{s1Match ?? "?"}\" s2=\"{s2Match ?? "?"}\" ctx[s1:{s1Anchors} s2:{s2Anchors}/{contextPositions}]");

            // For high-confidence words: only override if BOTH trusted secondaries agree
            // on a different word AND at least one has ≥2 context anchors (strong evidence).
            // For low-confidence: normal 2-of-3 vote using only trusted secondaries.
            string winner;
            if (isLowConf)
            {
                // First try normal 2-of-3 consensus vote with trusted secondaries
                winner = Vote(word, confidence,
                    s1Trusted ? s1Match : null,
                    s2Trusted ? s2Match : null);

                // If Vote kept the primary (no consensus), allow a single HIGH-CONFIDENCE
                // trusted secondary to override a LOW-CONFIDENCE primary (<0.70).
                // This handles the common case where only SM arrives in time (AAI is late)
                // and DG is clearly guessing (0.50) while SM is confident (0.95).
                if (string.Equals(winner, word, StringComparison.OrdinalIgnoreCase) && confidence < 0.70f)
                {
                    if (s1Trusted && s1Word != null && s1Word.Confidence >= 0.85f && s2Match == null &&
                        !string.Equals(word, s1Match, StringComparison.OrdinalIgnoreCase))
                    {
                        Logger.Trace($"Ensemble single-secondary override: s1 \"{s1Match}\" ({s1Word.Confidence:F2}) vs DG \"{word}\" ({confidence:F2})");
                        winner = s1Match!;
                    }
                    else if (s2Trusted && s2Word != null && s2Word.Confidence >= 0.85f && s1Match == null &&
                        !string.Equals(word, s2Match, StringComparison.OrdinalIgnoreCase))
                    {
                        Logger.Trace($"Ensemble single-secondary override: s2 \"{s2Match}\" ({s2Word.Confidence:F2}) vs DG \"{word}\" ({confidence:F2})");
                        winner = s2Match!;
                    }
                }
            }
            else
            {
                // High-confidence override requires strong phrase-level evidence
                bool strongContext = s1Anchors >= 2 || s2Anchors >= 2 || contextPositions == 0;
                if (s1Trusted && s2Trusted && strongContext &&
                    string.Equals(s1Match, s2Match, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(word, s1Match, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Trace($"Ensemble hi-conf override: \"{word}\" → \"{s1Match}\" (both agree, anchors: s1={s1Anchors} s2={s2Anchors})");
                    winner = s1Match;
                }
                else
                {
                    mergedWords.Add(word);
                    continue;
                }
            }

            // If the only difference is trailing punctuation, keep DG's version
            // (DG has smart_format which handles punctuation correctly; secondaries often omit it)
            if (string.Equals(StripTrailingPunctuation(winner), StripTrailingPunctuation(word), StringComparison.OrdinalIgnoreCase))
                winner = word;

            if (!string.Equals(winner, word, StringComparison.OrdinalIgnoreCase))
            {
                // A correction happened — track which provider(s) contributed
                corrections++;
                Interlocked.Increment(ref _correctedWords);
                _studyCorrected++;

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

                var source = fromS1 && fromS2 ? $"{Short(_s1Name)}+{Short(_s2Name)}" : fromS1 ? Short(_s1Name) : fromS2 ? Short(_s2Name) : "?";
                _lastCorrectionDetail = $"\"{word}\" \u2192 \"{winner}\" ({source})";
                correctionRecords.Add(new CorrectionRecord(word, winner, source));

                Logger.Trace($"Ensemble CORRECTION: \"{word}\" ({confidence:F2} {confTag}) → \"{winner}\" [s1={s1Match ?? "?"}, s2={s2Match ?? "?"}] anchors[s1:{s1Anchors} s2:{s2Anchors}] source={source}");
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
        if (correctionRecords.Count > 0)
            CorrectionsEmitted?.Invoke(correctionRecords);
        FireStatsUpdated();
    }

    private static string[] TokenizeProcessed(string text)
    {
        return text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
    }

    private static string? StripTrailingPunctuation(string? word)
    {
        if (string.IsNullOrEmpty(word)) return word;
        return word.TrimEnd('.', ',', ';', ':', '!', '?');
    }

    /// <summary>
    /// Count how many context words (±windowSize around target) match between primary and
    /// secondary aligned arrays. High anchor count = alignment is correct and the comparison
    /// at the target position is trustworthy.
    /// </summary>
    private static int CountContextAnchors(
        SttWord[] originalWords, SttWord?[] secondaryAligned, int targetIdx, int windowSize = 2)
    {
        int anchors = 0;
        for (int d = -windowSize; d <= windowSize; d++)
        {
            if (d == 0) continue; // Skip the target word itself
            int pi = targetIdx + d;
            if (pi < 0 || pi >= originalWords.Length || pi >= secondaryAligned.Length) continue;
            if (secondaryAligned[pi] == null) continue;

            var pWord = StripTrailingPunctuation(originalWords[pi].Text);
            var sWord = StripTrailingPunctuation(secondaryAligned[pi]?.Text);
            if (pWord != null && sWord != null &&
                string.Equals(pWord, sWord, StringComparison.OrdinalIgnoreCase))
                anchors++;
        }
        return anchors;
    }

    /// <summary>
    /// Count how many context positions exist around target (for determining if anchoring is possible).
    /// </summary>
    private static int CountContextPositions(int targetIdx, int arrayLen, int windowSize = 2)
    {
        int count = 0;
        for (int d = -windowSize; d <= windowSize; d++)
        {
            if (d == 0) continue;
            if (targetIdx + d >= 0 && targetIdx + d < arrayLen) count++;
        }
        return count;
    }

    /// <summary>
    /// Estimate the systematic timestamp offset between providers by finding words
    /// that match by text and computing the median midpoint difference.
    /// Returns offset to subtract from secondary midpoints to align with primary.
    /// </summary>
    private static double EstimateTimestampOffset(SttWord[] primary, SttWord[] secondary)
    {
        if (primary.Length == 0 || secondary.Length == 0) return 0;

        var offsets = new List<double>();
        int secIdx = 0;

        foreach (var pw in primary)
        {
            if (pw.StartTime <= 0) continue;
            var pMid = (pw.StartTime + pw.EndTime) / 2;
            var pText = pw.Text.TrimEnd('.', ',', ';', ':').ToLowerInvariant();

            for (int j = secIdx; j < secondary.Length; j++)
            {
                var sw = secondary[j];
                if (sw.EndTime <= 0) continue;
                var sText = sw.Text.TrimEnd('.', ',', ';', ':').ToLowerInvariant();

                if (string.Equals(pText, sText, StringComparison.OrdinalIgnoreCase))
                {
                    offsets.Add((sw.StartTime + sw.EndTime) / 2 - pMid);
                    secIdx = j + 1;
                    break;
                }

                // Don't search too far ahead
                if (sw.StartTime > pw.EndTime + 2.0) break;
            }
        }

        if (offsets.Count < 2) return 0; // Not enough data to calibrate

        // Median is robust to outliers (mismatched words with same text)
        offsets.Sort();
        var median = offsets[offsets.Count / 2];
        if (Math.Abs(median) > 0.05) // Only log if non-trivial offset
            Logger.Trace($"Ensemble: timestamp offset calibration = {median * 1000:F0}ms from {offsets.Count} matching words");
        return median;
    }

    /// <summary>
    /// Align secondary words to primary words using monotonic nearest-midpoint matching.
    /// Calibrates for systematic timestamp offset between providers (~200-400ms).
    /// Returns array of matched secondary SttWords (null where no match found).
    /// </summary>
    private static SttWord?[] AlignWordsByTimestamp(SttWord[] primaryWords, SttWord[] secondaryWords)
    {
        var result = new SttWord?[primaryWords.Length];
        if (secondaryWords.Length == 0) return result;

        // Calibrate: find the systematic timestamp offset so SM/AAI midpoints
        // align with DG midpoints (providers assign slightly different times to same audio)
        var offset = EstimateTimestampOffset(primaryWords, secondaryWords);

        int secStart = 0; // Monotonic: only advance forward in secondary array

        for (int i = 0; i < primaryWords.Length; i++)
        {
            var pw = primaryWords[i];
            if (pw.StartTime <= 0) continue;
            var pMid = (pw.StartTime + pw.EndTime) / 2;

            SttWord? bestWord = null;
            double bestDist = 0.5; // max 500ms between calibrated midpoints
            int bestIdx = -1;

            for (int j = secStart; j < secondaryWords.Length; j++)
            {
                var sw = secondaryWords[j];
                if (sw.EndTime <= 0) continue;
                var sMid = (sw.StartTime + sw.EndTime) / 2 - offset; // Apply calibration
                var dist = Math.Abs(sMid - pMid);

                if (dist < bestDist)
                {
                    bestDist = dist;
                    bestWord = sw;
                    bestIdx = j;
                }

                // Stop searching if secondary is too far ahead (use calibrated time)
                if (sw.StartTime - offset > pw.EndTime + 0.5) break;
            }

            if (bestWord != null)
            {
                result[i] = bestWord;
                secStart = bestIdx + 1; // Advance past matched word for monotonicity
            }
        }

        return result;
    }

    private static string Vote(string primary, float primaryConf,
        string? secondary1, string? secondary2)
    {
        // 2-of-3 consensus: any two providers agree → that word wins
        if (secondary1 != null && string.Equals(primary, secondary1, StringComparison.OrdinalIgnoreCase))
            return primary;
        if (secondary2 != null && string.Equals(primary, secondary2, StringComparison.OrdinalIgnoreCase))
            return primary;
        if (secondary1 != null && secondary2 != null &&
            string.Equals(secondary1, secondary2, StringComparison.OrdinalIgnoreCase))
            return secondary1; // Both secondaries agree → override primary

        // No consensus: all differ, or only one secondary available and it disagrees.
        // Keep the primary — a single dissenting secondary is 1-vs-1 with no tiebreaker,
        // and Deepgram (medical model) is more likely correct for medical terms.
        return primary;
    }

    private void AddConfidence(double conf)
    {
        // Not thread-critical — approximate is fine for display
        _totalConfidence += conf;
    }

    /// <summary>
    /// Commit this study's stats to all-time totals (call at study change when report was signed).
    /// </summary>
    public void CommitStudyToAlltime()
    {
        if (_studyWords == 0 && _studyCorrected == 0 &&
            _studyConfSum == 0 && _studyMerges == 0) return;
        _config.SttEnsembleAlltimeWords += _studyWords;
        _config.SttEnsembleAlltimeCorrected += _studyCorrected;
        _config.SttEnsembleAlltimeConfidenceSum += _studyConfSum;
        _config.SttEnsembleAlltimeMerges += _studyMerges;
        _config.Save();
        Logger.Trace($"Ensemble alltime committed: +{_studyWords} words, +{_studyCorrected} corrected, +{_studyMerges} merges");
    }

    /// <summary>
    /// Reset per-study counters for next study (call after CommitStudyToAlltime or on discard).
    /// </summary>
    public void ResetStudyCounters()
    {
        _studyWords = 0;
        _studyCorrected = 0;
        _studyConfSum = 0;
        _studyMerges = 0;
    }

    /// <summary>
    /// Flush study stats to all-time on shutdown (crash safety — even unsigned work is preserved).
    /// </summary>
    public void FlushAlltimeStats()
    {
        CommitStudyToAlltime();
        ResetStudyCounters();
    }

    // Validation counters (set by ActionController after study-change verification)
    internal int SessionValidated;
    internal int SessionRejected;

    public void FireStatsUpdated()
    {
        // All-time = persisted config values only (committed at study change, not during dictation)
        var atWords = _config.SttEnsembleAlltimeWords;
        var atCorrected = _config.SttEnsembleAlltimeCorrected;
        var atConfSum = _config.SttEnsembleAlltimeConfidenceSum;
        var atMerges = _config.SttEnsembleAlltimeMerges;
        var atAvgConf = atWords > 0 ? atConfSum / atWords : 0;

        StatsUpdated?.Invoke(new EnsembleStats(
            _totalWords, _lowConfidenceWords, _correctedWords,
            AverageConfidence, _mergeCount,
            _s1Arrivals, _s2Arrivals,
            _s1Corrections, _s2Corrections, _consensusCorrections,
            _s1Confirms, _s2Confirms,
            _lastCorrectionDetail,
            atWords, atCorrected, atAvgConf, atMerges,
            SessionValidated, SessionRejected,
            _config.SttEnsembleAlltimeValidated, _config.SttEnsembleAlltimeRejected));
    }
}
