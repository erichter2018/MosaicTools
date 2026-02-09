using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace MosaicTools.Services;

/// <summary>
/// A single stored template entry for a study type.
/// </summary>
public class TemplateEntry
{
    [JsonPropertyName("hash")]
    public string Hash { get; set; } = "";

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";

    [JsonPropertyName("count")]
    public int Count { get; set; } = 1;

    [JsonPropertyName("first_seen")]
    public string FirstSeen { get; set; } = "";

    [JsonPropertyName("last_seen")]
    public string LastSeen { get; set; } = "";
}

/// <summary>
/// Root structure for the template database JSON file.
/// </summary>
public class TemplateDatabaseData
{
    [JsonPropertyName("version")]
    public int Version { get; set; } = 1;

    [JsonPropertyName("templates")]
    public Dictionary<string, List<TemplateEntry>> Templates { get; set; } = new();
}

/// <summary>
/// Persistent database of clean report templates keyed by study description.
/// Non-drafted studies provide templates for free (they stay as-is). Over time,
/// repeated observations of the same template text for a study type reinforce
/// confidence that it's a true default. When real-time capture fails on a drafted
/// study, fall back to the stored template.
///
/// Only stores FINDINGS+IMPRESSION sections (not patient-specific header content).
/// </summary>
public class TemplateDatabase
{
    private static readonly object _saveLock = new();
    private TemplateDatabaseData _data = new();
    private static readonly string DbPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MosaicTools", "TemplateDatabase.json");

    private const int ConfidenceThreshold = 5;
    private const int MaxTemplatesPerStudy = 3;
    private const int StaleDays = 90;
    private const int LowConfidenceStaleDays = 30;

    /// <summary>
    /// Measurement pattern - reports with measurements are dictated, not templates.
    /// </summary>
    private static readonly Regex MeasurementPattern = new(
        @"\d+\.?\d*\s*(cm|mm)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Load database from JSON file. Falls back to empty on any error.
    /// </summary>
    public void Load()
    {
        try
        {
            if (File.Exists(DbPath))
            {
                var json = File.ReadAllText(DbPath);
                var data = JsonSerializer.Deserialize<TemplateDatabaseData>(json);
                if (data != null)
                {
                    _data = data;
                    int totalTemplates = _data.Templates.Values.Sum(list => list.Count);
                    Logger.Trace($"TemplateDatabase loaded: {_data.Templates.Count} study types, {totalTemplates} templates");
                    return;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"TemplateDatabase load error: {ex.Message}");
        }

        _data = new TemplateDatabaseData();
    }

    /// <summary>
    /// Save database to JSON file.
    /// </summary>
    public void Save()
    {
        try
        {
            var dir = Path.GetDirectoryName(DbPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(_data, options);
            lock (_saveLock)
            {
                File.WriteAllText(DbPath, json);
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"TemplateDatabase save error: {ex.Message}");
        }
    }

    /// <summary>
    /// Record a template observation. Non-drafted studies fully reinforce confidence;
    /// drafted studies keep templates alive (update last_seen) but don't build confidence,
    /// since drafted reports may have been modified by a prior reader.
    /// Extracts FINDINGS+IMPRESSION sections and stores only that portion.
    /// </summary>
    public void RecordTemplate(string studyDescription, string reportText, bool isDrafted = false)
    {
        if (string.IsNullOrWhiteSpace(studyDescription) || string.IsNullOrWhiteSpace(reportText))
            return;

        // Extract FINDINGS+IMPRESSION only
        var (findings, impression) = CorrelationService.ExtractSections(reportText);
        if (string.IsNullOrWhiteSpace(findings) || string.IsNullOrWhiteSpace(impression))
        {
            Logger.Trace($"TemplateDB: Skipping record for '{studyDescription}' - missing FINDINGS or IMPRESSION");
            return;
        }

        var sectionText = SanitizeForStorage($"FINDINGS:\n{findings}\nIMPRESSION:\n{impression}");

        if (!IsSafeToStore(sectionText))
        {
            Logger.Trace($"TemplateDB: Skipping record for '{studyDescription}' - failed safety check");
            return;
        }

        var normalized = Normalize(sectionText);
        var hash = ComputeHash(normalized);
        var key = studyDescription.Trim().ToUpperInvariant();
        var today = DateTime.Now.ToString("yyyy-MM-dd");

        if (!_data.Templates.TryGetValue(key, out var entries))
        {
            entries = new List<TemplateEntry>();
            _data.Templates[key] = entries;
        }

        // Check if this template already exists (by hash)
        var existing = entries.FirstOrDefault(e => e.Hash == hash);
        if (existing != null)
        {
            existing.LastSeen = today;
            if (!isDrafted)
            {
                existing.Count++;
                Logger.Trace($"TemplateDB: Template reinforced for '{studyDescription}' (count={existing.Count})");
            }
            else
            {
                Logger.Trace($"TemplateDB: Template seen in drafted study for '{studyDescription}' (count stays {existing.Count}, last_seen updated)");
            }
        }
        else
        {
            // New variant detected — apply conservative decay to existing templates
            // High-confidence templates (count >= 30) decay by 5% (minimal but allows eventual updates)
            // Medium-confidence templates (count 10-29) decay by 25%
            // Low-confidence templates (count < 10) decay by 50%
            // Drafted observations don't trigger decay (they're less trustworthy signals)
            if (!isDrafted)
            {
                foreach (var entry in entries)
                {
                    var oldCount = entry.Count;

                    if (entry.Count >= 30)
                    {
                        // Very high confidence - minimal 5% decay (allows gradual template updates)
                        entry.Count = Math.Max(1, (entry.Count * 19) / 20);
                        if (oldCount != entry.Count)
                            Logger.Trace($"TemplateDB: High-confidence template minimally decayed for '{studyDescription}' ({oldCount} → {entry.Count})");
                    }
                    else if (entry.Count >= 10)
                    {
                        // Medium confidence - conservative 25% decay
                        entry.Count = Math.Max(1, (entry.Count * 3) / 4);
                        if (oldCount != entry.Count)
                            Logger.Trace($"TemplateDB: Medium-confidence template decayed for '{studyDescription}' ({oldCount} → {entry.Count})");
                    }
                    else
                    {
                        // Low confidence - standard 50% decay
                        entry.Count = Math.Max(1, entry.Count / 2);
                        if (oldCount != entry.Count)
                            Logger.Trace($"TemplateDB: Low-confidence template decayed for '{studyDescription}' ({oldCount} → {entry.Count})");
                    }
                }
            }

            // Drafted studies start at count=0 (need non-drafted observations to build confidence)
            var initialCount = isDrafted ? 0 : 1;
            entries.Add(new TemplateEntry
            {
                Hash = hash,
                Text = sectionText,
                Count = initialCount,
                FirstSeen = today,
                LastSeen = today
            });
            Logger.Trace($"TemplateDB: New template variant recorded for '{studyDescription}' ({sectionText.Length} chars, count={initialCount}, drafted={isDrafted}), decayed {entries.Count - 1} existing");
        }

        // Enforce max templates per study type (keep top by count)
        if (entries.Count > MaxTemplatesPerStudy)
        {
            var sorted = entries.OrderByDescending(e => e.Count).ThenByDescending(e => e.LastSeen).ToList();
            _data.Templates[key] = sorted.Take(MaxTemplatesPerStudy).ToList();
        }

        Save();
    }

    /// <summary>
    /// Get a high-confidence fallback template for a study description.
    /// Returns the FINDINGS+IMPRESSION text, or null if no confident template exists.
    /// </summary>
    public string? GetFallbackTemplate(string studyDescription)
    {
        if (string.IsNullOrWhiteSpace(studyDescription))
            return null;

        var key = studyDescription.Trim().ToUpperInvariant();

        if (!_data.Templates.TryGetValue(key, out var entries) || entries.Count == 0)
            return null;

        // Select highest confidence template (count >= threshold)
        var candidates = entries
            .Where(e => e.Count >= ConfidenceThreshold)
            .OrderByDescending(e => e.Count)
            .ThenByDescending(e => e.LastSeen)
            .ToList();

        if (candidates.Count == 0)
            return null;

        var best = candidates[0];
        Logger.Trace($"TemplateDB: Fallback for '{studyDescription}' (count={best.Count}, lastSeen={best.LastSeen})");
        return best.Text;
    }

    /// <summary>
    /// Lightweight sanity checks to ensure report text is a clean template.
    /// </summary>
    private bool IsSafeToStore(string text)
    {
        // Reject if it contains measurement patterns (indicates dictated findings)
        if (MeasurementPattern.IsMatch(text))
            return false;

        // Both FINDINGS and IMPRESSION must be present
        if (!text.Contains("FINDINGS:", StringComparison.OrdinalIgnoreCase))
            return false;
        if (!text.Contains("IMPRESSION:", StringComparison.OrdinalIgnoreCase))
            return false;

        return true;
    }

    /// <summary>
    /// Strip invisible Unicode characters (U+FFFC, zero-width spaces, format chars, etc.)
    /// and collapse resulting empty lines. Keeps stored text clean and hash-stable.
    /// </summary>
    private static string SanitizeForStorage(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;

        var sb = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            if (c == '\r' || c == '\n' || c == '\t')
            {
                sb.Append(c);
                continue;
            }
            if (c == '\u00A0') { sb.Append(' '); continue; }

            var category = char.GetUnicodeCategory(c);
            if (category == System.Globalization.UnicodeCategory.Format ||
                category == System.Globalization.UnicodeCategory.Control ||
                category == System.Globalization.UnicodeCategory.Surrogate ||
                category == System.Globalization.UnicodeCategory.PrivateUse ||
                category == System.Globalization.UnicodeCategory.OtherNotAssigned ||
                c == '\uFFFC' || c == '\uFFFD')
                continue;

            sb.Append(c);
        }

        // Collapse runs of blank lines into single newlines
        var result = Regex.Replace(sb.ToString(), @"(\r?\n\s*){2,}", "\n");
        return result.Trim();
    }

    /// <summary>
    /// Normalize text for comparison: trim lines, collapse whitespace, normalize newlines.
    /// </summary>
    private string Normalize(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";

        text = text.Replace("\r\n", "\n").Replace("\r", "\n");

        var lines = text.Split('\n')
            .Select(line => Regex.Replace(line.Trim(), @"\s+", " "))
            .Where(line => !string.IsNullOrWhiteSpace(line));

        return string.Join("\n", lines);
    }

    /// <summary>
    /// Compute a truncated SHA256 hash of normalized text.
    /// </summary>
    private string ComputeHash(string text)
    {
        var bytes = Encoding.UTF8.GetBytes(text);
        var hashBytes = SHA256.HashData(bytes);
        // Truncate to first 12 hex chars (6 bytes) - sufficient for dedup
        return Convert.ToHexString(hashBytes, 0, 6).ToLowerInvariant();
    }

    /// <summary>
    /// Prune stale and low-confidence entries.
    /// - Remove templates not seen in 90+ days
    /// - Remove count=1 entries older than 30 days
    /// </summary>
    public void Cleanup()
    {
        var now = DateTime.Now;
        bool changed = false;

        foreach (var key in _data.Templates.Keys.ToList())
        {
            var entries = _data.Templates[key];
            int before = entries.Count;

            entries.RemoveAll(e =>
            {
                if (!DateTime.TryParse(e.LastSeen, out var lastSeen))
                    return true; // Remove entries with unparseable dates

                var age = (now - lastSeen).TotalDays;

                // Remove stale entries (not seen in 90+ days)
                if (age > StaleDays)
                    return true;

                // Remove low-confidence entries older than 30 days
                if (e.Count <= 1 && age > LowConfidenceStaleDays)
                    return true;

                return false;
            });

            if (entries.Count != before)
                changed = true;

            // Remove empty study types
            if (entries.Count == 0)
            {
                _data.Templates.Remove(key);
                changed = true;
            }
        }

        if (changed)
        {
            int totalTemplates = _data.Templates.Values.Sum(list => list.Count);
            Logger.Trace($"TemplateDatabase cleanup: {_data.Templates.Count} study types, {totalTemplates} templates remaining");
            Save();
        }
    }
}
