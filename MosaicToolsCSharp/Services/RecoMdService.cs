using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace MosaicTools.Services;

/// <summary>
/// RecoMD integration — sends report data to the local RecoMD agent for best practice recommendations.
/// RecoMD runs a Go agent on localhost:7771 with a REST API.
/// </summary>
public class RecoMdService : IDisposable
{
    private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(10) };
    private const string BaseUrl = "http://localhost:7771/api";
    private string? _currentAccession;

    /// <summary>
    /// Check if the RecoMD agent is running (GET /api/alive → 200 or 202).
    /// </summary>
    public async Task<bool> IsAliveAsync()
    {
        try
        {
            var resp = await _http.GetAsync($"{BaseUrl}/alive");
            return (int)resp.StatusCode == 200 || (int)resp.StatusCode == 202;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// POST /api/dictation/report/open with exam metadata.
    /// </summary>
    public async Task<bool> OpenReportAsync(string accession, string? description,
        string? patientName, string? gender, string? mrn, int age)
    {
        try
        {
            // Extract modality from first word of description
            string modality = "CT";
            if (!string.IsNullOrEmpty(description))
            {
                var firstWord = description.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                    .FirstOrDefault()?.ToUpperInvariant() ?? "CT";
                if (new[] { "CT", "MR", "MRI", "US", "XR", "CR", "NM", "PET", "FL", "DX", "MG", "RF" }.Contains(firstWord))
                    modality = firstWord == "MRI" ? "MR" : firstWord;
            }

            // Gender mapping: "Male" → "M", "Female" → "F"
            var genderCode = gender?.ToUpperInvariant() switch
            {
                "MALE" => "M",
                "FEMALE" => "F",
                _ => ""
            };

            // Split patient name (format: "Smith John" → last="Smith", first="John")
            string firstName = "", lastName = "";
            if (!string.IsNullOrEmpty(patientName))
            {
                var parts = patientName.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2)
                {
                    lastName = parts[0];
                    firstName = parts[1];
                }
                else if (parts.Length == 1)
                {
                    lastName = parts[0];
                }
            }

            // Build nested JSON payload per RecoMD integration guide
            // Top-level + Exam fields: PascalCase; Patient fields: snake_case
            var payload = new
            {
                Accession = accession,
                IsAddendum = false,
                Exam = new
                {
                    AccessionNumber = accession,
                    Modality = modality,
                    ExamDescription = description ?? "",
                    Patient = new
                    {
                        name = $"{firstName} {lastName}".Trim(),
                        first_name = firstName,
                        last_name = lastName,
                        mrn = mrn ?? "",
                        age_at_exam = age,
                        gender = genderCode
                    }
                }
            };

            var json = JsonSerializer.Serialize(payload);
            Logger.Trace($"RecoMD: Open payload: {json}");
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var resp = await _http.PostAsync($"{BaseUrl}/dictation/report/open", content);

            _currentAccession = accession;
            Logger.Trace($"RecoMD: Open report {accession} → {(int)resp.StatusCode}");
            return resp.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            Logger.Trace($"RecoMD: Open report error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// POST /api/dictation/report/{accession} with report text (text/plain).
    /// </summary>
    public async Task<bool> SendReportTextAsync(string accession, string reportText)
    {
        try
        {
            var content = new StringContent(reportText, Encoding.UTF8, "text/plain");
            var resp = await _http.PostAsync($"{BaseUrl}/dictation/report/{accession}", content);

            if (!resp.IsSuccessStatusCode)
                Logger.Trace($"RecoMD: Send report text for {accession} ({reportText.Length} chars) → {(int)resp.StatusCode}");
            return resp.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            Logger.Trace($"RecoMD: Send report text error: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// POST /api/dictation/report/close — no body.
    /// </summary>
    public async Task CloseReportAsync()
    {
        try
        {
            var content = new StringContent("", Encoding.UTF8, "application/json");
            var resp = await _http.PostAsync($"{BaseUrl}/dictation/report/close", content);
            Logger.Trace($"RecoMD: Close report → {(int)resp.StatusCode}");
            _currentAccession = null;
        }
        catch (Exception ex)
        {
            Logger.Trace($"RecoMD: Close report error: {ex.Message}");
        }
    }

    /// <summary>
    /// Clean report text for sending to RecoMD: sanitize special chars, convert pipe separators
    /// from ProseMirror scrape to newlines, and collapse blank lines.
    /// RecoMD needs section structure (FINDINGS:, IMPRESSION:, etc) on separate lines.
    /// </summary>
    public static string CleanReportText(string text)
    {
        // Sanitize special chars (U+FFFC, zero-width spaces, etc.)
        text = MosaicTools.UI.ReportPopupForm.SanitizeText(text);

        // ProseMirror Name property joins text runs with " | " separators.
        // Convert these to newlines so section headers appear on their own lines.
        // Only replace " | " (with surrounding spaces) to avoid touching pipes inside content.
        text = text.Replace(" | ", "\n");
        // Also handle "| " at start and " |" at end of text
        if (text.StartsWith("| ")) text = text.Substring(2);
        if (text.EndsWith(" |")) text = text.Substring(0, text.Length - 2);

        // Collapse runs of blank lines into a single blank line
        text = System.Text.RegularExpressions.Regex.Replace(text, @"(\r?\n\s*){3,}", "\n\n");

        // Collapse multiple spaces into one
        text = System.Text.RegularExpressions.Regex.Replace(text, @"  +", " ");

        return text.Trim();
    }

    public void Dispose()
    {
        // HttpClient is static, don't dispose it
    }
}
