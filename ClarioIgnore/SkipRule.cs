using System;
using System.Linq;

namespace ClarioIgnore;

public class SkipRule
{
    public bool Enabled { get; set; } = true;
    public string Name { get; set; } = "";
    public string CriteriaRequired { get; set; } = "";
    public string CriteriaAnyOf { get; set; } = "";
    public string CriteriaExclude { get; set; } = "";
    public bool IncludePriority { get; set; } = false;

    /// <summary>
    /// Check if a study procedure name matches this rule.
    /// - CriteriaRequired: ALL terms must be present (comma-separated)
    /// - CriteriaAnyOf: At least ONE term must be present (comma-separated)
    /// - CriteriaExclude: NONE of these terms can be present (comma-separated)
    /// </summary>
    public bool MatchesStudy(string procedureName, string priority = "")
    {
        if (!Enabled) return false;
        if (string.IsNullOrWhiteSpace(procedureName)) return false;

        // Combine priority with procedure if IncludePriority is enabled
        var text = IncludePriority && !string.IsNullOrWhiteSpace(priority)
            ? $"{priority} | {procedureName}".ToUpperInvariant()
            : procedureName.ToUpperInvariant();

        // Check Required - ALL must match
        if (!string.IsNullOrWhiteSpace(CriteriaRequired))
        {
            var requiredTerms = CriteriaRequired
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.TrimStart().ToUpperInvariant())  // TrimStart preserves trailing spaces for exact matching
                .Where(t => !string.IsNullOrEmpty(t));

            foreach (var term in requiredTerms)
            {
                if (!text.Contains(term))
                    return false;
            }
        }

        // Check AnyOf - at least ONE must match
        if (!string.IsNullOrWhiteSpace(CriteriaAnyOf))
        {
            var anyOfTerms = CriteriaAnyOf
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.TrimStart().ToUpperInvariant())  // TrimStart preserves trailing spaces for exact matching
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList();

            if (anyOfTerms.Count > 0)
            {
                bool foundAny = anyOfTerms.Any(term => text.Contains(term));
                if (!foundAny)
                    return false;
            }
        }

        // Check Exclude - NONE must match
        if (!string.IsNullOrWhiteSpace(CriteriaExclude))
        {
            var excludeTerms = CriteriaExclude
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.TrimStart().ToUpperInvariant())  // TrimStart preserves trailing spaces for exact matching
                .Where(t => !string.IsNullOrEmpty(t));

            foreach (var term in excludeTerms)
            {
                if (text.Contains(term))
                    return false;
            }
        }

        // If we have no criteria at all, don't match anything
        if (string.IsNullOrWhiteSpace(CriteriaRequired) &&
            string.IsNullOrWhiteSpace(CriteriaAnyOf) &&
            string.IsNullOrWhiteSpace(CriteriaExclude))
        {
            return false;
        }

        return true;
    }

    public override string ToString()
    {
        return $"{Name} (Enabled: {Enabled})";
    }
}
