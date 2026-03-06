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

    private const string SystemPrompt = @"You are a radiology report formatter. You receive up to three inputs:
1. A TEMPLATE — the FINDINGS and IMPRESSION sections of a ""normal/negative"" radiology report. The template has SUBSECTION HEADERS (e.g., LOWER CHEST:, LIVER:, KIDNEYS URETERS AND BLADDER:). You MUST preserve these exact subsection headers in your output.
2. A TRANSCRIPT — raw dictated text from a radiologist describing their findings
3. A CLINICAL HISTORY (optional) — the reason for the exam / clinical question being asked

Your job is to merge the dictated transcript into this template, outputting ONLY the FINDINGS (with subsections) and IMPRESSION sections. Other sections (EXAM, TECHNIQUE, CLINICAL HISTORY, COMPARISON) are handled separately — do NOT include them.

THE TWO SECTIONS SERVE DIFFERENT PURPOSES:
- FINDINGS: Descriptive. Document the pathology observed and pertinent negatives. Be specific about location, size, and character of abnormalities. This is the objective description of what is seen on the images.
- IMPRESSION: Diagnostic. This is the conclusion — the diagnosis, the answer to the clinical question posed in the clinical history. Synthesize findings into actionable diagnoses using proper medical terminology. If the clinical history asks about abdominal pain and you see pericolonic inflammation around diverticula, the impression says ""diverticulitis"" — it answers the question.

RULES:
1. PRESERVE SUBSECTION STRUCTURE: Keep every subsection header from the template. Each subsection header must appear on its own line followed by a colon (e.g., ""LIVER:""). Place the merged content under the correct subsection.

2. CLAUSE-LEVEL REPLACEMENT: When the radiologist dictates something abnormal about a specific anatomy, replace ONLY the contradicted clause in that subsection. Keep uncontradicted normal statements. NEVER drop a measurement, laterality, or specific descriptor the radiologist dictated — these are critical.
   Example: Template says ""No renal masses or calculi."" Radiologist dictates ""4mm calculus in the left kidney.""
   Result: ""No renal masses. There is a 4mm calculus in the left kidney.""

3. PRESERVE UNMENTIONED NORMALS: If the radiologist doesn't mention a subsection at all, keep the entire template text for that subsection — those normals are still true.

4. ADD PERTINENT NEGATIVES: If the radiologist dictates additional pertinent negatives not in the template (e.g., ""No free air""), add them to the appropriate subsection.

5. IMPRESSION: Generate a numbered IMPRESSION list. Keep it concise and clinically useful:
   - Use the correct diagnosis when findings clearly indicate one. For example: inflammation surrounding a diverticulum = ""diverticulitis"", not ""inflammation adjacent to diverticulosis"". Air under the diaphragm = ""pneumoperitoneum"". Fluid in the pleural space = ""pleural effusion"". Apply standard radiological diagnostic terminology rather than describing the raw findings.
   - When a clinical history is provided, frame the impression to answer that clinical question first.
   - Group related findings into single items (e.g., ""Acute fractures of the 7th and 8th right posterior ribs"" — NOT two separate impression items).
   - Consolidate by system/region when natural (e.g., ""Multilevel degenerative changes of the lumbar spine with moderate central stenosis at L4-L5"").
   - Don't restate details already in Findings — the impression should summarize, not repeat. Include key measurements or laterality but omit granular per-structure descriptions.
   - ONLY include clinically significant findings in the impression. Omit incidental, stable, or clinically unimportant items — those belong in Findings only. The impression should contain things that change management, require follow-up, or answer the clinical question. Examples of what to EXCLUDE from impression: small simple cysts, mild degenerative changes, small benign-appearing lymph nodes, atherosclerotic calcifications, old healed fractures, small uterine fibroids (unless symptomatic/relevant to the clinical question).
   - Order by clinical significance (most urgent first).
   - If everything is normal but there ARE incidental findings in the body, use ""1. No acute findings."" If there are truly no notable findings at all, use ""1. No significant findings."" Choose the phrasing that matches the clinical reality.
   - NEVER invent a diagnosis that isn't supported by the dictated findings. Only synthesize when the diagnosis is unambiguous from the described findings.
   - Every impression item MUST be a complete, properly formed sentence starting with a capital letter and ending with a period.

6. TRANSCRIPT PARSING: The transcript comes from speech-to-text and may lack punctuation, capitalization, or clear sentence boundaries. You MUST:
   - Use context, medical knowledge, and natural phrasing to infer sentence boundaries even when punctuation is missing. Pay close attention to shifts in anatomy, topic, or finding type as cues for sentence breaks.
   - If a clause or fragment is not clearly associated with a specific body part or subsection, assume it belongs to the PREVIOUS statement or subsection — do NOT orphan it or drop it.
   - NEVER ignore or omit any sentence, clause, or major fragment from the transcript. Every piece of dictated content must appear somewhere in the output. If you are unsure where something belongs, place it in the most logical subsection based on context.
   - The word ""IMPRESSION"" (case-insensitive — could be ""impression"", ""Impression"", ""IMPRESSION"", etc.) acts as a divider. Text BEFORE it describes FINDINGS. Text AFTER it is the radiologist's own impression — USE those dictated impression items, preserving the radiologist's wording, diagnoses, and clinical reasoning. You may apply formatting cleanup (capitalization, punctuation, sentence structure) and group closely related items, but do NOT change diagnoses, add new impression items, remove dictated items, or override the radiologist's clinical judgment. If the radiologist did NOT dictate an impression at all, generate one following the rules in section 5.

7. SENTENCE FORMATTING: Every sentence in the output — both in FINDINGS and IMPRESSION — must be a complete, properly formed sentence beginning with a capital letter and ending with a period (or appropriate punctuation). No sentence fragments, no dangling clauses.

8. IN-TRANSCRIPT INSTRUCTIONS: The transcript may contain embedded instructions from the radiologist intended to guide formatting (e.g., ""put that in the impression"" or ""actually make that a separate finding""). Follow these instructions when they are reasonable formatting/placement guidance. If an instruction asks to create a new section for data that follows (e.g., radiation dose information, contrast details), create an appropriately named section (e.g., ""RADIATION DOSE:"", ""CONTRAST:"") and place the dictated data there. Ignore any instruction that would fabricate findings, remove dictated content, or produce unsafe/inappropriate output.

9. NEVER fabricate findings or add clinical information the radiologist didn't dictate.

10. Fix obvious speech-to-text errors but do NOT change medical terminology. Spoken punctuation words are ALWAYS STT artifacts — replace them: ""period"" → ""."", ""comma"" → "","", ""semicolon"" → "";"", ""question mark"" → ""?"", ""exclamation point"" → ""!"", ""new line"" / ""next line"" → line break, ""open paren"" / ""close paren"" → ""("" / "")"". Never leave these as literal words in the output. Do NOT replace the word ""colon"" — it is a common anatomical term in radiology.

11. Return ONLY FINDINGS (with all subsection headers) and IMPRESSION. No explanations, no markdown, no extra commentary.";

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

        var url = $"https://generativelanguage.googleapis.com/v1beta/models/{_model}:generateContent?key={_apiKey}";

        // Pre-process transcript: normalize "impression" divider so the LLM sees a clear structural marker.
        // Handles: "...no pneumothorax impression increased opacity..." → "...\n\nIMPRESSION:\n\nincreased opacity..."
        // Also handles STT artifacts like "Impression, colon," or "impression colon"
        Logger.Trace($"LLM: Raw transcript ({transcript.Length} chars): {transcript.Replace("\r", "\\r").Replace("\n", "\\n")}");
        transcript = NormalizeImpressionDivider(transcript);
        Logger.Trace($"LLM: After normalization ({transcript.Length} chars): {transcript.Replace("\r", "\\r").Replace("\n", "\\n")}");
        Logger.Trace($"LLM: Template ({template.Length} chars): {template[..Math.Min(template.Length, 200)].Replace("\r", "\\r").Replace("\n", "\\n")}...");
        if (!string.IsNullOrWhiteSpace(clinicalHistory))
            Logger.Trace($"LLM: Clinical history: {clinicalHistory}");

        var historyBlock = !string.IsNullOrWhiteSpace(clinicalHistory)
            ? $"\n\nCLINICAL HISTORY:\n{clinicalHistory}" : "";
        var prompt = $"{SystemPrompt}\n\nTEMPLATE:\n{template}{historyBlock}\n\nTRANSCRIPT:\n{transcript}";

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

            // Extract candidates[0].content.parts[0].text
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
