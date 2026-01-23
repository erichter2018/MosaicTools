using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ClarioIgnore;

public class Configuration
{
    private static readonly string SettingsFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ClarioIgnore");

    private static readonly string SettingsPath = Path.Combine(SettingsFolder, "settings.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    // Singleton
    private static Configuration? _instance;
    public static Configuration Instance => _instance ??= Load();

    // Settings
    public int PollIntervalSeconds { get; set; } = 10;
    public bool IsPaused { get; set; } = false;
    public bool ShowNotifications { get; set; } = true;
    public List<SkipRule> SkipRules { get; set; } = new();

    public Configuration() { }

    public static Configuration Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                var config = JsonSerializer.Deserialize<Configuration>(json, JsonOptions);
                if (config != null)
                {
                    _instance = config;
                    return config;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"Error loading configuration: {ex.Message}");
        }

        // Return default configuration with example rules
        var defaultConfig = new Configuration();
        defaultConfig.SkipRules.Add(new SkipRule
        {
            Enabled = false,
            Name = "Example: Skip Venous US",
            CriteriaAnyOf = "VENOUS, DOPPLER",
            CriteriaExclude = "ARTERIAL"
        });

        _instance = defaultConfig;

        // Save the default config so users can see the example
        defaultConfig.Save();

        return defaultConfig;
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(SettingsFolder);
            var json = JsonSerializer.Serialize(this, JsonOptions);
            File.WriteAllText(SettingsPath, json);
            Logger.Log($"Configuration saved to {SettingsPath} ({SkipRules.Count} rules)");
        }
        catch (Exception ex)
        {
            Logger.Log($"Error saving configuration: {ex.Message}");
        }
    }

    public SkipRule? FindMatchingRule(string procedureName)
    {
        foreach (var rule in SkipRules)
        {
            if (rule.MatchesStudy(procedureName))
                return rule;
        }
        return null;
    }
}
