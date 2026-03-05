// [CustomSTT] Ensemble merger: collects finals from all 3 providers and produces a merged result.
namespace MosaicTools.Services;

/// <summary>
/// Snapshot of ensemble performance stats for live display.
/// </summary>
public record EnsembleStats(
    int TotalWords, int LowConfidenceWords, int CorrectedWords,
    double AverageConfidence, int MergeCount,
    int S1Arrivals, int S2Arrivals, int S1MergeParticipations, int S2MergeParticipations,
    int S1Corrections, int S2Corrections, int ConsensusCorrections,
    int S1Confirms, int S2Confirms,
    string? LastCorrectionDetail,
    List<string> RecentCorrections,
    // All-time (persisted across sessions)
    int AlltimeWords, int AlltimeCorrected,
    double AlltimeAverageConfidence, int AlltimeMerges,
    // Correction validation (verified against signed reports)
    int SessionValidated, int SessionRejected,
    int AlltimeValidated, int AlltimeRejected,
    // Per-provider accuracy (bag-matched against signed report)
    int SessionEnsembleMatched, int SessionEnsembleTotal,
    int SessionDgMatched, int SessionDgTotal,
    int SessionS1Matched, int SessionS1Total,
    int SessionS2Matched, int SessionS2Total,
    int AlltimeEnsembleMatched, int AlltimeEnsembleTotal,
    int AlltimeDgMatched, int AlltimeDgTotal,
    int AlltimeS1Matched, int AlltimeS1Total,
    int AlltimeS2Matched, int AlltimeS2Total);

/// <summary>
/// Records a single ensemble correction for later validation against the signed report.
/// </summary>
public record CorrectionRecord(string Original, string Replacement, string Source);

/// <summary>
/// Records what each provider said at a single word position during ensemble merge.
/// </summary>
internal record WordRecord(string MergedWord, string DgWord, string? S1Word, string? S2Word);

/// <summary>
/// Result of bag-matching a provider's words against the signed report.
/// </summary>
internal record ProviderAccuracyResult(int Matched, int Total);

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
    private int _s1MergeParticipations; // Merges where Soniox data was available
    private int _s2MergeParticipations; // Merges where SM data was available
    private int _s1Corrections;        // Times Soniox word replaced Deepgram
    private int _s2Corrections;        // Times SM word replaced Deepgram
    private int _consensusCorrections; // Both secondaries agreed → overruled Deepgram
    private int _s1Confirms;           // Soniox agreed with low-conf Deepgram word (validated it)
    private int _s2Confirms;           // SM agreed with low-conf Deepgram word (validated it)
    private string? _lastCorrectionDetail; // e.g. "pole → poll (SNX+SM)"
    private readonly List<string> _recentCorrections = new(); // last 3 corrections

    // Per-study delta — only committed to all-time when report is signed
    private int _studyWords;
    private int _studyCorrected;
    private double _studyConfSum;
    private int _studyMerges;

    // Per-study word records for provider accuracy validation
    private readonly List<WordRecord> _studyWordRecords = new();

    // Per-provider accuracy (session, updated at each study change)
    private int _sessionEnsembleMatched, _sessionEnsembleTotal;
    private int _sessionDgMatched, _sessionDgTotal;
    private int _sessionS1Matched, _sessionS1Total;
    private int _sessionS2Matched, _sessionS2Total;

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
        "soniox" => "Soniox", "speechmatics" => "Speechmatics", "assemblyai" => "AssemblyAI", "none" => "\u2014",
        _ => n
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
                // Purge secondary entries from previous utterance (arrived >5s before this primary).
                // Wide window to keep word-by-word providers (SM) alive on long phrases;
                // FilterStaleAudioWords in EmitMerged handles precise cleanup via audio timestamps.
                PurgeStaleEntries(_secondary1Buffer, now - 5000);
                PurgeStaleEntries(_secondary2Buffer, now - 5000);
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
        _s1MergeParticipations = 0;
        _s2MergeParticipations = 0;
        _s1Corrections = 0;
        _s2Corrections = 0;
        _consensusCorrections = 0;
        _s1Confirms = 0;
        _s2Confirms = 0;
        _lastCorrectionDetail = null;
    }

    /// <summary>
    /// Reset ALL stats: session counters, validation counters, recent corrections, and per-study deltas.
    /// Called from "Clear All-Time Stats" button. Caller should also clear config alltime values.
    /// </summary>
    public void ResetAllStats()
    {
        ResetStats();
        SessionValidated = 0;
        SessionRejected = 0;
        _studyWords = 0;
        _studyCorrected = 0;
        _studyConfSum = 0;
        _studyMerges = 0;
        _studyWordRecords.Clear();
        _sessionEnsembleMatched = 0; _sessionEnsembleTotal = 0;
        _sessionDgMatched = 0; _sessionDgTotal = 0;
        _sessionS1Matched = 0; _sessionS1Total = 0;
        _sessionS2Matched = 0; _sessionS2Total = 0;
        lock (_recentCorrections) { _recentCorrections.Clear(); }
        FireStatsUpdated();
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

    /// <summary>
    /// Remove secondary words whose audio EndTime is before the primary's start.
    /// Handles slow providers (e.g., SNX 3s behind) whose wall-clock arrival passes
    /// the purge but whose audio data belongs to a previous merge segment.
    /// </summary>
    private static SttWord[] FilterStaleAudioWords(SttWord[] words, double audioFloor, string providerName)
    {
        if (words.Length == 0) return words;

        // Find first word whose midpoint is within the primary's time range.
        // Using midpoint instead of EndTime prevents boundary words from bleeding through
        // (e.g., "peritoneum" ending at 4.3s but centered at 4.0s when floor is 4.05s).
        int firstValid = 0;
        for (int i = 0; i < words.Length; i++)
        {
            var mid = (words[i].StartTime + words[i].EndTime) / 2;
            if (mid >= audioFloor)
            {
                firstValid = i;
                break;
            }
            firstValid = i + 1;
        }

        if (firstValid > 0)
        {
            Logger.Trace($"Ensemble: audio-time filter removed {firstValid} stale {providerName} words (before {audioFloor:F2}s)");
            return words[firstValid..];
        }
        return words;
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

        if (s1 != null) Interlocked.Increment(ref _s1MergeParticipations);
        if (s2 != null) Interlocked.Increment(ref _s2MergeParticipations);

        var s1Status = s1 != null ? $"{s1.Words?.Length ?? 0} words" : "missing";
        var s2Status = s2 != null ? $"{s2.Words?.Length ?? 0} words" : "missing";
        Logger.Trace($"Ensemble merge #{_mergeCount}: DG={primary.Words?.Length ?? 0} words, {Short(_s1Name)}={s1Status}, {Short(_s2Name)}={s2Status}");

        // If no secondaries arrived, just post-process and emit primary
        if (s1 == null && s2 == null)
        {
            Logger.Trace("Ensemble merge: no secondaries arrived, emitting primary only");
            var processed = SttTextProcessor.ProcessTranscript(primary.Transcript, _config);
            // Remove consecutive duplicate words
            processed = RemoveConsecutiveDuplicateWords(processed);
            var processedWords = processed.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var wc = primary.Words?.Length ?? 0;
            Interlocked.Add(ref _totalWords, wc);
            _studyWords += wc;
            if (primary.Words != null)
                foreach (var w in primary.Words)
                {
                    AddConfidence(w.Confidence);
                    _studyConfSum += w.Confidence;
                }
            // Record word positions for accuracy validation (no secondaries)
            foreach (var w in processedWords)
                _studyWordRecords.Add(new WordRecord(w, w, null, null));
            MergedResultReady?.Invoke(primary with
            {
                Transcript = processed,
                ProviderName = "ensemble"
            });
            FireStatsUpdated();
            return;
        }

        // Use raw DG word array for merge — ProcessTranscript can change word count
        // (spoken punctuation, radiology cleanup), which causes index misalignment between
        // processed words and the raw word/timestamp arrays used for alignment and context anchors.
        // Text processing is applied to the final merged result instead.
        // DG's smart_format puts punctuation in the transcript but NOT in individual words,
        // so we use transcript tokens for word text (preserving commas/periods) and the word
        // array for timestamps, confidence, and alignment.
        var s1WordArr = s1?.Words ?? Array.Empty<SttWord>();
        var s2WordArr = s2?.Words ?? Array.Empty<SttWord>();
        var originalWords = primary.Words ?? Array.Empty<SttWord>();
        var transcriptTokens = primary.Transcript.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        // Filter out stale secondary words from previous merge segments.
        // The wall-clock purge can miss slow providers (e.g., SNX 3s behind real-time)
        // whose data arrives "recently" but covers old audio timestamps.
        if (originalWords.Length > 0)
        {
            var primaryStartTime = originalWords[0].StartTime;
            if (primaryStartTime > 0)
            {
                // Allow 500ms overlap tolerance for boundary words
                var audioFloor = primaryStartTime - 0.5;
                s1WordArr = FilterStaleAudioWords(s1WordArr, audioFloor, _s1Name);
                s2WordArr = FilterStaleAudioWords(s2WordArr, audioFloor, _s2Name);
            }
        }

        // Align secondary words to primary using monotonic nearest-midpoint matching.
        // This handles the ~200ms timestamp offset between providers (SM/AAI vs DG)
        // that caused per-word overlap matching to consistently hit the adjacent word.
        var s1Aligned = AlignWordsByTimestamp(originalWords, s1WordArr);
        var s2Aligned = AlignWordsByTimestamp(originalWords, s2WordArr);

        // Merge: use primary timing structure, replace low-confidence words via voting
        var mergedWords = new List<string>();
        var mergedConfTier = new List<byte>(); // parallel to mergedWords: 0=normal, 1=medium, 2=low
        var correctionRecords = new List<CorrectionRecord>();
        int corrections = 0;
        int lowConf = 0;

        for (int i = 0; i < originalWords.Length; i++)
        {
            var origWord = originalWords[i];
            // Use transcript token (has punctuation from smart_format), fall back to word array
            var word = i < transcriptTokens.Length ? transcriptTokens[i] : origWord.Text;
            var confidence = origWord.Confidence;

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
                mergedConfTier.Add(0);
                _studyWordRecords.Add(new WordRecord(word, word, null, null));
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
            var timeTag = origWord.StartTime > 0 ? $" {origWord.StartTime:F2}-{origWord.EndTime:F2}s" : "";
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
                    if (s1Trusted && s1Word != null && s1Word.Confidence >= 0.85f && (s2Match == null || !s2Trusted) &&
                        !string.Equals(word, s1Match, StringComparison.OrdinalIgnoreCase))
                    {
                        Logger.Trace($"Ensemble single-secondary override: s1 \"{s1Match}\" ({s1Word.Confidence:F2}) vs DG \"{word}\" ({confidence:F2})");
                        winner = s1Match!;
                    }
                    else if (s2Trusted && s2Word != null && s2Word.Confidence >= 0.85f && (s1Match == null || !s1Trusted) &&
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
                    mergedConfTier.Add(0);
                    _studyWordRecords.Add(new WordRecord(word, word, s1Match, s2Match));
                    continue;
                }
            }

            // If the only difference is trailing punctuation, keep DG's version
            // (DG has smart_format which handles punctuation correctly; secondaries often omit it)
            if (string.Equals(StripTrailingPunctuation(winner), StripTrailingPunctuation(word), StringComparison.OrdinalIgnoreCase))
                winner = word;

            // Defense-in-depth: before accepting any correction, verify that the provider
            // supplying the replacement has at least one immediate neighbor (±1) matching.
            // This catches shifted alignments that pass the wider ±2 context window.
            if (!string.Equals(winner, word, StringComparison.OrdinalIgnoreCase) && contextPositions > 0)
            {
                bool fromS1 = s1Match != null && string.Equals(winner, s1Match, StringComparison.OrdinalIgnoreCase);
                bool fromS2 = s2Match != null && string.Equals(winner, s2Match, StringComparison.OrdinalIgnoreCase);
                bool neighborOk = false;
                if (fromS1 && HasImmediateNeighborMatch(originalWords, s1Aligned, i)) neighborOk = true;
                if (fromS2 && HasImmediateNeighborMatch(originalWords, s2Aligned, i)) neighborOk = true;
                if (!neighborOk)
                {
                    Logger.Trace($"Ensemble: BLOCKED correction \"{word}\" → \"{winner}\" — no immediate neighbor match (shifted alignment)");
                    winner = word;
                }
            }

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
                lock (_recentCorrections)
                {
                    _recentCorrections.Add(_lastCorrectionDetail);
                    if (_recentCorrections.Count > 3)
                        _recentCorrections.RemoveAt(0);
                }
                correctionRecords.Add(new CorrectionRecord(word, winner, source));

                Logger.Trace($"Ensemble CORRECTION: \"{word}\" ({confidence:F2} {confTag}) → \"{winner}\" [s1={s1Match ?? "?"}, s2={s2Match ?? "?"}] anchors[s1:{s1Anchors} s2:{s2Anchors}] source={source}");
            }
            else
            {
                // Primary word kept — track which secondaries confirmed it
                bool s1Confirmed = s1Match != null && string.Equals(word, s1Match, StringComparison.OrdinalIgnoreCase);
                bool s2Confirmed = s2Match != null && string.Equals(word, s2Match, StringComparison.OrdinalIgnoreCase);
                if (s1Confirmed) Interlocked.Increment(ref _s1Confirms);
                if (s2Confirmed) Interlocked.Increment(ref _s2Confirms);
            }

            // Confidence tier: 0=normal, 1=medium (0.60-threshold), 2=low (<0.60)
            // Only mark uncertain if DG was low-confidence and ensemble couldn't correct it
            byte confTier = 0;
            if (isLowConf && string.Equals(winner, word, StringComparison.OrdinalIgnoreCase))
                confTier = confidence < 0.60f ? (byte)2 : (byte)1;
            mergedWords.Add(winner);
            mergedConfTier.Add(confTier);
            _studyWordRecords.Add(new WordRecord(winner, word, s1Match, s2Match));
        }

        // Remove consecutive duplicate words (can happen when Deepgram sends overlapping finals)
        for (int i = mergedWords.Count - 1; i > 0; i--)
        {
            if (string.Equals(mergedWords[i], mergedWords[i - 1], StringComparison.OrdinalIgnoreCase))
            {
                Logger.Trace($"Ensemble: removed consecutive duplicate word \"{mergedWords[i]}\" at position {i}");
                mergedWords.RemoveAt(i);
                mergedConfTier.RemoveAt(i);
            }
        }

        var mergedRaw = string.Join(" ", mergedWords);

        // Apply text processing (spoken punctuation, radiology cleanup, custom replacements)
        // to the final merged result, keeping merge indices aligned with raw word arrays.
        var mergedTranscript = SttTextProcessor.ProcessTranscript(mergedRaw, _config);
        mergedTranscript = RemoveConsecutiveDuplicateWords(mergedTranscript);

        // Extract confidence-tier word indices (only if ProcessTranscript didn't change word count,
        // otherwise indices would be misaligned — skip confidence marking in that case)
        var finalWordCount = mergedTranscript.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
        int[]? mediumArr = null, lowArr = null;
        if (finalWordCount == mergedWords.Count)
        {
            var medium = new List<int>();
            var low = new List<int>();
            for (int idx = 0; idx < mergedConfTier.Count; idx++)
            {
                if (mergedConfTier[idx] == 1) medium.Add(idx);
                else if (mergedConfTier[idx] == 2) low.Add(idx);
            }
            if (medium.Count > 0) mediumArr = medium.ToArray();
            if (low.Count > 0) lowArr = low.ToArray();
        }

        if (corrections > 0)
            Logger.Trace($"Ensemble merge: {corrections} corrections out of {lowConf} low-confidence words");

        MergedResultReady?.Invoke(primary with
        {
            Transcript = mergedTranscript,
            ProviderName = "ensemble",
            MediumConfWordIndices = mediumArr,
            LowConfWordIndices = lowArr
        });
        if (correctionRecords.Count > 0)
            CorrectionsEmitted?.Invoke(correctionRecords);
        FireStatsUpdated();
    }

    private static string[] TokenizeProcessed(string text)
    {
        return text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
    }

    private static string RemoveConsecutiveDuplicateWords(string text)
    {
        var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length <= 1) return text;
        var result = new List<string> { words[0] };
        for (int i = 1; i < words.Length; i++)
        {
            if (!string.Equals(words[i], words[i - 1], StringComparison.OrdinalIgnoreCase))
                result.Add(words[i]);
            else
                Logger.Trace($"Ensemble: removed consecutive duplicate word \"{words[i]}\"");
        }
        return result.Count < words.Length ? string.Join(" ", result) : text;
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
    /// Strict alignment check: verify at least one immediate neighbor (±1) matches.
    /// A 1-position shift always fails this because direct neighbors are wrong,
    /// even if the wider ±2 window has coincidental matches.
    /// </summary>
    private static bool HasImmediateNeighborMatch(
        SttWord[] originalWords, SttWord?[] secondaryAligned, int targetIdx)
    {
        for (int d = -1; d <= 1; d += 2)
        {
            int pi = targetIdx + d;
            if (pi < 0 || pi >= originalWords.Length || pi >= secondaryAligned.Length) continue;
            if (secondaryAligned[pi] == null) continue;
            var pWord = StripTrailingPunctuation(originalWords[pi].Text);
            var sWord = StripTrailingPunctuation(secondaryAligned[pi]?.Text);
            if (pWord != null && sWord != null &&
                string.Equals(pWord, sWord, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
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
    /// Validate per-provider accuracy by bag-matching each provider's words against the signed report.
    /// Returns (ensemble, deepgram, s1, s2) accuracy results.
    /// </summary>
    internal (ProviderAccuracyResult ens, ProviderAccuracyResult dg,
              ProviderAccuracyResult s1, ProviderAccuracyResult s2)
        ValidateAccuracyAgainstReport(string reportText)
    {
        var (findings, impression) = CorrelationService.ExtractSections(reportText);
        var body = findings + " " + impression;
        if (string.IsNullOrWhiteSpace(body))
        {
            Logger.Trace($"Ensemble accuracy: ExtractSections returned empty (report={reportText.Length} chars)");
            return (new(0, 0), new(0, 0), new(0, 0), new(0, 0));
        }

        // Build word-frequency bag from report
        var reportBag = BuildWordBag(body);
        Logger.Trace($"Ensemble accuracy: body={body.Length} chars, bag={reportBag.Count} unique words, records={_studyWordRecords.Count}");

        // Need minimum data: at least 10 report words and 10 dictated word records
        if (reportBag.Count < 10 || _studyWordRecords.Count < 10)
        {
            Logger.Trace($"Ensemble accuracy: skipped — insufficient data (bag={reportBag.Count}, records={_studyWordRecords.Count})");
            return (new(0, 0), new(0, 0), new(0, 0), new(0, 0));
        }

        // Match each provider independently against a fresh copy of the bag
        var ensResult = MatchAgainstBag(reportBag, _studyWordRecords, r => r.MergedWord);
        var dgResult = MatchAgainstBag(reportBag, _studyWordRecords, r => r.DgWord);
        var s1Result = MatchAgainstBag(reportBag, _studyWordRecords, r => r.S1Word);
        var s2Result = MatchAgainstBag(reportBag, _studyWordRecords, r => r.S2Word);

        // Sanity check: if ensemble match rate < 30%, the report probably doesn't
        // correspond to the dictation (template text, stale report, or heavily edited).
        if (ensResult.Total > 0 && (double)ensResult.Matched / ensResult.Total < 0.30)
        {
            Logger.Trace($"Ensemble accuracy: skipped — match rate too low ({ensResult.Matched}/{ensResult.Total} = {100.0 * ensResult.Matched / ensResult.Total:F0}%), report likely doesn't match dictation");
            return (new(0, 0), new(0, 0), new(0, 0), new(0, 0));
        }

        return (ensResult, dgResult, s1Result, s2Result);
    }

    /// <summary>
    /// Add validated accuracy results to session counters.
    /// </summary>
    internal void AddSessionAccuracy(ProviderAccuracyResult ens, ProviderAccuracyResult dg,
        ProviderAccuracyResult s1, ProviderAccuracyResult s2)
    {
        _sessionEnsembleMatched += ens.Matched; _sessionEnsembleTotal += ens.Total;
        _sessionDgMatched += dg.Matched; _sessionDgTotal += dg.Total;
        _sessionS1Matched += s1.Matched; _sessionS1Total += s1.Total;
        _sessionS2Matched += s2.Matched; _sessionS2Total += s2.Total;
    }

    private static Dictionary<string, int> BuildWordBag(string text)
    {
        var bag = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var token in text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
        {
            var norm = NormalizeForMatch(token);
            if (norm == null || norm.Length <= 1) continue;
            bag.TryGetValue(norm, out int count);
            bag[norm] = count + 1;
        }
        return bag;
    }

    private static ProviderAccuracyResult MatchAgainstBag(
        Dictionary<string, int> reportBag, List<WordRecord> records,
        Func<WordRecord, string?> wordSelector)
    {
        // Clone the bag so each provider gets independent matching
        var bag = new Dictionary<string, int>(reportBag, StringComparer.OrdinalIgnoreCase);
        int matched = 0, total = 0;
        foreach (var rec in records)
        {
            var raw = wordSelector(rec);
            if (raw == null) continue; // Provider didn't have this position
            var norm = NormalizeForMatch(raw);
            if (norm == null || norm.Length <= 1) continue;
            total++;
            if (bag.TryGetValue(norm, out int count) && count > 0)
            {
                matched++;
                bag[norm] = count - 1;
            }
        }
        return new ProviderAccuracyResult(matched, total);
    }

    private static string? NormalizeForMatch(string? word)
    {
        if (string.IsNullOrEmpty(word)) return null;
        var trimmed = word.TrimEnd('.', ',', ';', ':', '!', '?', ')', '(', '"', '\'');
        trimmed = trimmed.TrimStart('(', '"', '\'');
        return string.IsNullOrEmpty(trimmed) ? null : trimmed.ToLowerInvariant();
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

        var atWords = _config.SttEnsembleAlltimeWords;
        var atCorrected = _config.SttEnsembleAlltimeCorrected;
        var atAvgConf = atWords > 0 ? _config.SttEnsembleAlltimeConfidenceSum / atWords : 0;
        var atVal = _config.SttEnsembleAlltimeValidated;
        var atRej = _config.SttEnsembleAlltimeRejected;
        var precStr = (atVal + atRej) > 0 ? $", precision={100.0 * atVal / (atVal + atRej):F0}% ({atVal}v/{atRej}r)" : "";
        Logger.Trace($"Ensemble alltime committed: +{_studyWords} words, +{_studyCorrected} corrected, +{_studyMerges} merges (total: {atWords}w, {atCorrected}corr, avgConf={atAvgConf:F3}{precStr})");
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
        _studyWordRecords.Clear();
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
            _s1Arrivals, _s2Arrivals, _s1MergeParticipations, _s2MergeParticipations,
            _s1Corrections, _s2Corrections, _consensusCorrections,
            _s1Confirms, _s2Confirms,
            _lastCorrectionDetail,
            new List<string>(_recentCorrections),
            atWords, atCorrected, atAvgConf, atMerges,
            SessionValidated, SessionRejected,
            _config.SttEnsembleAlltimeValidated, _config.SttEnsembleAlltimeRejected,
            _sessionEnsembleMatched, _sessionEnsembleTotal,
            _sessionDgMatched, _sessionDgTotal,
            _sessionS1Matched, _sessionS1Total,
            _sessionS2Matched, _sessionS2Total,
            _config.SttEnsembleAlltimeAccEnsembleMatched, _config.SttEnsembleAlltimeAccEnsembleTotal,
            _config.SttEnsembleAlltimeAccDgMatched, _config.SttEnsembleAlltimeAccDgTotal,
            _config.SttEnsembleAlltimeAccS1Matched, _config.SttEnsembleAlltimeAccS1Total,
            _config.SttEnsembleAlltimeAccS2Matched, _config.SttEnsembleAlltimeAccS2Total));
    }
}
