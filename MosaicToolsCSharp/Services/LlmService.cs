using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace MosaicTools.Services;

/// <summary>
/// Minimal Google Gemini REST client for Custom Process Report.
/// Stateless, async. No SDK — just HttpClient + JSON.
/// </summary>
internal class LlmService : IDisposable
{
    private readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(30) };
    private string? _apiKey;
    private string _model = "gemini-2.5-flash-lite";

    // ═══════ PROMPTS ═══════

    // Full prompt: used when radiologist did NOT dictate an impression (single call, LLM generates impression)
    private const string FullReportPrompt = @"<role>Radiology report formatter. Merge dictated TRANSCRIPT into TEMPLATE. Output FINDINGS and IMPRESSION.</role>

<critical-rules>
1. PRESERVE THE RADIOLOGIST'S EXACT WORDS in FINDINGS. No paraphrasing, no synonym substitution.
   ""the heart is enlarged"" → ""the heart is enlarged"", NOT ""cardiomegaly"".
   Do not add words not dictated (""stable"", ""unchanged"", ""chronic"").
2. NEVER fabricate findings or add clinical information not in the transcript.
3. NEVER drop any dictated content — every sentence/clause must appear in the output.
</critical-rules>

<findings-rules>
- Keep every subsection header from the template on its own line with colon (e.g. ""LIVER:"").
- All findings within a subsection go on ONE contiguous line (sentences separated by spaces, not line breaks).
- Replace ONLY the contradicted clause. Keep uncontradicted template normals.
  Example: Template ""No renal masses or calculi."" + Dictated ""4mm calculus in the left kidney.""
  Result: ""No renal masses. There is a 4mm calculus in the left kidney.""
- Unmentioned subsections: keep entire template text.
- Infer sentence boundaries from context. Ambiguous fragments → previous subsection.
</findings-rules>

<subsection-routing>
pulmonary edema, pulmonary vascular congestion, pneumonia, atelectasis, consolidation, air trapping, emphysema, pleural effusion, pneumothorax → LUNGS AND PLEURA
cardiomegaly, pericardial effusion, aortic calcification, mediastinal widening → HEART AND MEDIASTINUM
</subsection-routing>

<impression-rules>
- Numbered list of clinically significant diagnoses. Proper diagnostic terminology.
- Answer clinical question first when clinical history provided.
- Group related findings. Order by clinical significance.
- Exclude incidentals (small cysts, mild degenerative changes, atherosclerotic calcifications, old healed fractures).
- Normal exam: ""1. No acute findings."" or ""1. No significant findings.""
- Never invent diagnoses unsupported by findings.
</impression-rules>

<formatting>
- Every sentence: capital letter, ends with period.
- Fix STT errors but do NOT change medical terminology.
- Spoken punctuation → symbols: ""period""→. ""comma""→, ""semicolon""→; ""question mark""→? ""new line""→line break. Do NOT replace ""colon"" (anatomical term).
- Follow reasonable in-transcript formatting instructions (e.g. ""new section for radiation dose"").
- Output ONLY FINDINGS + IMPRESSION. No explanations, no markdown.
</formatting>";

    // Findings-only prompt: used when impression is handled separately
    private const string FindingsOnlyPrompt = @"<role>Radiology report formatter. Merge dictated TRANSCRIPT into TEMPLATE subsections. Output ONLY the FINDINGS section — do NOT generate an IMPRESSION.</role>

<rules>
- PRESERVE THE RADIOLOGIST'S EXACT WORDS. No paraphrasing, no synonym substitution.
  ""the heart is enlarged"" → ""the heart is enlarged"", NOT ""cardiomegaly"".
  Do not add words not dictated (""stable"", ""unchanged"", ""chronic"").
- Replace ONLY the contradicted clause. Keep uncontradicted template normals.
- All findings within a subsection go on ONE contiguous line (sentences separated by spaces, not line breaks).
- Keep every subsection header. Unmentioned subsections: keep template text.
- Infer sentence boundaries from context. Ambiguous fragments → previous subsection.
- Fix STT errors but do NOT change medical terminology.
- Spoken punctuation → symbols: ""period""→. ""comma""→, Do NOT replace ""colon"".
- Every sentence: capital letter, ends with period.
- NEVER fabricate or drop content.
- Output FINDINGS ONLY. No IMPRESSION section. No explanations.
</rules>

<subsection-routing>
pulmonary edema, pulmonary vascular congestion, pneumonia, atelectasis, consolidation, air trapping, emphysema, pleural effusion, pneumothorax → LUNGS AND PLEURA
cardiomegaly, pericardial effusion, aortic calcification, mediastinal widening → HEART AND MEDIASTINUM
</subsection-routing>";

    // Impression formatting prompt: used to format multi-sentence dictated impressions
    private const string ImpressionFormatPrompt = @"<role>Format these dictated radiology impression items into a numbered list.</role>

<rules>
- Each item: complete sentence, capital letter, period at end.
- Fix obvious STT errors. Use proper diagnostic terminology where clearly intended.
- Group closely related items into one. Order by clinical significance.
- CRITICAL: Do NOT add any items not present in the input. Do NOT remove any dictated items.
- Output ONLY the numbered list. No header, no explanations.
</rules>";

    internal void Configure(string apiKey, string model)
    {
        _apiKey = string.IsNullOrWhiteSpace(apiKey) ? null : apiKey.Trim();
        _model = string.IsNullOrWhiteSpace(model) ? "gemini-2.5-flash-lite" : model.Trim();
    }

    internal bool IsConfigured => !string.IsNullOrEmpty(_apiKey);

    /// <summary>
    /// Send transcript + template to Gemini and return the formatted report.
    /// Returns null on any failure.
    /// </summary>
    internal async Task<string?> ProcessReportAsync(string transcript, string template, string? clinicalHistory = null, CancellationToken ct = default)
    {
        if (string.IsNullOrEmpty(_apiKey)) return null;

        // Pre-process transcript: normalize "impression" divider
        Logger.Trace($"LLM: Raw transcript ({transcript.Length} chars): {transcript.Replace("\r", "\\r").Replace("\n", "\\n")}");
        transcript = NormalizeImpressionDivider(transcript);
        Logger.Trace($"LLM: After normalization ({transcript.Length} chars): {transcript.Replace("\r", "\\r").Replace("\n", "\\n")}");
        Logger.Trace($"LLM: Template ({template.Length} chars): {template[..Math.Min(template.Length, 200)].Replace("\r", "\\r").Replace("\n", "\\n")}...");
        if (!string.IsNullOrWhiteSpace(clinicalHistory))
            Logger.Trace($"LLM: Clinical history: {clinicalHistory}");

        // Split flow for 2.5-flash-lite: it can't follow "don't add impression items" rules,
        // so we split findings and impression into separate calls.
        const string divider = "\n\nIMPRESSION:\n";
        var divIdx = transcript.IndexOf(divider, StringComparison.Ordinal);
        bool isLiteModel = _model.Contains("lite", StringComparison.OrdinalIgnoreCase);
        bool useSplitFlow = divIdx >= 0 && isLiteModel;

        if (useSplitFlow)
        {
            var findingsTranscript = transcript[..divIdx].Trim();
            var dictatedImpression = transcript[(divIdx + divider.Length)..].Trim();
            Logger.Trace($"LLM: Split mode — findings ({findingsTranscript.Length} chars), impression ({dictatedImpression.Length} chars): \"{dictatedImpression}\"");

            var historyBlock = !string.IsNullOrWhiteSpace(clinicalHistory)
                ? $"\n\nCLINICAL HISTORY:\n{clinicalHistory}" : "";

            // Call 1: Findings only (always runs)
            var findingsPrompt = $"{FindingsOnlyPrompt}\n\nTEMPLATE:\n{template}{historyBlock}\n\nTRANSCRIPT:\n{findingsTranscript}";
            var findingsTask = CallGeminiAsync(findingsPrompt, ct);

            // Call 2: Impression — LLM for multi-sentence, local for single sentence
            var wordCount = dictatedImpression.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
            Task<string?> impressionTask;
            if (wordCount > 10)
            {
                Logger.Trace($"LLM: Impression has {wordCount} words — sending to LLM for formatting");
                var impPrompt = $"{ImpressionFormatPrompt}\n\nDICTATED IMPRESSION:\n{dictatedImpression}";
                impressionTask = CallGeminiAsync(impPrompt, ct);
            }
            else
            {
                Logger.Trace($"LLM: Impression has {wordCount} words — formatting locally");
                impressionTask = Task.FromResult<string?>(FormatSimpleImpression(dictatedImpression));
            }

            // Run in parallel
            await Task.WhenAll(findingsTask, impressionTask);

            var findings = findingsTask.Result;
            if (findings == null) return null;

            // Strip any IMPRESSION the findings call might have sneaked in
            var impSneak = findings.IndexOf("\nIMPRESSION", StringComparison.OrdinalIgnoreCase);
            if (impSneak >= 0)
            {
                Logger.Trace("LLM: Stripping sneaked IMPRESSION from findings-only response");
                findings = findings[..impSneak].TrimEnd();
            }

            var impression = impressionTask.Result ?? FormatSimpleImpression(dictatedImpression);

            Logger.Trace($"LLM: Split result — findings ({findings.Length} chars), impression ({impression.Length} chars)");
            return $"{findings}\n\nIMPRESSION:\n{impression}";
        }
        else
        {
            // Single call: full prompt (no dictated impression, or model handles it fine)
            var historyBlock = !string.IsNullOrWhiteSpace(clinicalHistory)
                ? $"\n\nCLINICAL HISTORY:\n{clinicalHistory}" : "";
            var prompt = $"{FullReportPrompt}\n\nTEMPLATE:\n{template}{historyBlock}\n\nTRANSCRIPT:\n{transcript}";
            return await CallGeminiAsync(prompt, ct);
        }
    }

    /// <summary>
    /// Format a short dictated impression locally (single sentence — no LLM needed).
    /// </summary>
    private static string FormatSimpleImpression(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "1. No significant findings.";

        // Split on newlines to handle multiple dictated impression items
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        var sb = new System.Text.StringBuilder();
        int n = 1;
        foreach (var line in lines)
        {
            var item = line.Trim().TrimEnd('.');
            if (item.Length == 0) continue;
            item = char.ToUpper(item[0]) + item[1..];
            sb.AppendLine($"{n}. {item}.");
            n++;
        }
        return sb.ToString().TrimEnd();
    }

    /// <summary>
    /// Send a prompt to Gemini and return the response text. Shared by all call paths.
    /// </summary>
    private async Task<string?> CallGeminiAsync(string prompt, CancellationToken ct)
    {
        var url = $"https://generativelanguage.googleapis.com/v1beta/models/{_model}:generateContent?key={_apiKey}";

        var requestBody = new
        {
            contents = new[]
            {
                new
                {
                    parts = new[]
                    {
                        new { text = prompt }
                    }
                }
            },
            generationConfig = new
            {
                temperature = 0.1,
                maxOutputTokens = 4096
            }
        };

        var json = JsonSerializer.Serialize(requestBody);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");

        try
        {
            var response = await _http.PostAsync(url, content, ct);
            var responseText = await response.Content.ReadAsStringAsync(ct);

            if (!response.IsSuccessStatusCode)
            {
                Logger.Trace($"LLM: HTTP {(int)response.StatusCode} — {responseText[..Math.Min(responseText.Length, 300)]}");
                return null;
            }

            using var doc = JsonDocument.Parse(responseText);
            var text = doc.RootElement
                .GetProperty("candidates")[0]
                .GetProperty("content")
                .GetProperty("parts")[0]
                .GetProperty("text")
                .GetString();

            var trimmed = text?.Trim();
            Logger.Trace($"LLM: Response ({trimmed?.Length ?? 0} chars): {trimmed?.Replace("\r", "\\r").Replace("\n", "\\n")}");
            return trimmed;
        }
        catch (OperationCanceledException)
        {
            Logger.Trace("LLM: Request cancelled");
            return null;
        }
        catch (Exception ex)
        {
            Logger.Trace($"LLM: Error — {ex.Message}");
            return null;
        }
    }

    // ═══════ PHI SCRUBBING ═══════

    // Patterns compiled once for performance
    private static readonly Regex RxMrn = new(@"MRN:\s*\S+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RxAccession = new(@"\b\d{8,}\w*\b", RegexOptions.Compiled);
    private static readonly Regex RxDob = new(@"DOB:\s*\S+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RxDateSlash = new(@"\b\d{1,2}/\d{1,2}/\d{2,4}\b", RegexOptions.Compiled);
    private static readonly Regex RxDateDash = new(@"\b(?:0?[1-9]|1[0-2])-(?:0?[1-9]|[12]\d|3[01])-\d{2,4}\b", RegexOptions.Compiled);
    private static readonly Regex RxDateWritten = new(
        @"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RxReferring = new(
        @"(?:Referring|Ordered\s+by|Ref\.?\s*Phys):\s*[^\r\n]+",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RxDrName = new(
        @"Dr\.?\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*",
        RegexOptions.Compiled);

    // Known section headers that are ALL-CAPS but NOT patient names
    private static readonly HashSet<string> SectionHeaders = new(StringComparer.OrdinalIgnoreCase)
    {
        "EXAM", "EXAMINATION", "CLINICAL HISTORY", "CLINICAL INFORMATION",
        "HISTORY", "INDICATION", "INDICATIONS", "TECHNIQUE", "COMPARISON",
        "FINDINGS", "IMPRESSION", "RECOMMENDATION", "RECOMMENDATIONS",
        "CONCLUSION", "ADDENDUM", "LOWER CHEST", "HEPATOBILIARY",
        "PANCREAS", "SPLEEN", "ADRENALS", "KIDNEYS AND URETERS",
        "PELVIS", "GASTROINTESTINAL", "VASCULATURE", "LYMPH NODES",
        "MUSCULOSKELETAL", "SOFT TISSUES", "OTHER", "BONE WINDOWS",
        "LUNGS", "HEART AND MEDIASTINUM", "PLEURA",
        "BRAIN PARENCHYMA", "VENTRICLES", "EXTRA-AXIAL SPACES",
        "ORBITS", "PARANASAL SINUSES", "MASTOIDS", "CALVARIUM",
        "NECK", "THYROID", "CHEST WALL", "BREAST"
    };

    /// <summary>
    /// Scrub PHI from text before sending to LLM. Preserves age, gender, and all medical content.
    /// </summary>
    internal static string ScrubPhi(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        text = RxMrn.Replace(text, "[REDACTED]");
        text = RxAccession.Replace(text, "[REDACTED]");
        text = RxDob.Replace(text, "[REDACTED]");
        text = RxDateSlash.Replace(text, "[REDACTED]");
        text = RxDateDash.Replace(text, "[REDACTED]");
        text = RxDateWritten.Replace(text, "[REDACTED]");
        text = RxReferring.Replace(text, "[REDACTED]");
        text = RxDrName.Replace(text, "[REDACTED]");

        // ALL-CAPS multi-word lines that aren't section headers (likely patient names like "SMITH, JOHN")
        var lines = text.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            var trimmed = lines[i].Trim();
            if (trimmed.Length < 3) continue;
            // Check if line is all uppercase letters, spaces, commas, periods
            if (Regex.IsMatch(trimmed, @"^[A-Z\s,.\-']+$") && trimmed.Contains(' '))
            {
                // Strip trailing colon/punctuation for header check
                var headerCheck = trimmed.TrimEnd(':', ' ');
                if (!SectionHeaders.Contains(headerCheck))
                    lines[i] = "[REDACTED]";
            }
        }

        return string.Join('\n', lines);
    }

    // ═══════ TRANSCRIPT PRE-PROCESSING ═══════

    // Match "impression" as a standalone section divider word, optionally followed by STT artifacts
    // like ", colon," or "colon" or ":". Excludes medical uses via lookbehind/lookahead:
    // - Lookbehind: skip "clinical impression", "initial impression", "overall impression", "vascular impression"
    // - Lookahead: skip "impression fracture/defect/on/of/from" (anatomical usage)
    private static readonly Regex RxImpressionDivider = new(
        @"(?<!clinical\s)(?<!initial\s)(?<!overall\s)(?<!vascular\s)(?<!my\s)(?<!\w)(?i:impression)(?!\s*(?:fracture|defect|on\s+the|of\s+the|from\s+the))\s*(?:[,.]?\s*(?:colon|:)\s*[,.]?\s*|[,.:]\s*)?",
        RegexOptions.Compiled);

    /// <summary>
    /// Find the word "impression" used as a section divider in the transcript and normalize it
    /// to a clear structural marker so the LLM recognizes the FINDINGS/IMPRESSION boundary.
    /// Skips medical uses like "impression fracture" or "clinical impression".
    /// Only normalizes the FIRST matching occurrence.
    /// </summary>
    internal static string NormalizeImpressionDivider(string transcript)
    {
        if (string.IsNullOrEmpty(transcript)) return transcript;

        var match = RxImpressionDivider.Match(transcript);
        if (!match.Success) return transcript;

        var before = transcript[..match.Index].TrimEnd();
        var after = transcript[(match.Index + match.Length)..].TrimStart();

        return $"{before}\n\nIMPRESSION:\n{after}";
    }

    public void Dispose()
    {
        _http.Dispose();
    }
}
