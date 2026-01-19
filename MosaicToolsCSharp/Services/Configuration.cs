using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MosaicTools.Services;

/// <summary>
/// Configuration storage matching MosaicToolsSettings.json format.
/// </summary>
public class Configuration
{
    public static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MosaicTools", "MosaicToolsSettings.json");

    /// <summary>
    /// Hotkeys reserved by Mosaic that shouldn't be used as triggers to avoid feedback loops.
    /// </summary>
    public static readonly HashSet<string> RestrictedHotkeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "alt+n", "alt+r", "alt+p", "alt+s", "alt+t", "alt+c", "alt+f", "alt+1", "alt+2", "ctrl+/"
    };

    public static bool IsHotkeyRestricted(string hotkey)
    {
        if (string.IsNullOrWhiteSpace(hotkey)) return false;
        var normalized = hotkey.ToLowerInvariant().Replace(" ", "");
        return RestrictedHotkeys.Contains(normalized);
    }
    
    // General
    [JsonPropertyName("doctor_name")]
    public string DoctorName { get; set; } = "Radiologist";

    [JsonPropertyName("auto_update_enabled")]
    public bool AutoUpdateEnabled { get; set; } = true;

    // Beep Settings
    [JsonPropertyName("start_beep_enabled")]
    public bool StartBeepEnabled { get; set; } = true;
    
    [JsonPropertyName("stop_beep_enabled")]
    public bool StopBeepEnabled { get; set; } = true;
    
    [JsonPropertyName("start_beep_volume")]
    public double StartBeepVolume { get; set; } = 0.04;
    
    [JsonPropertyName("stop_beep_volume")]
    public double StopBeepVolume { get; set; } = 0.04;
    
    [JsonPropertyName("dictation_pause_ms")]
    public int DictationPauseMs { get; set; } = 1000;
    
    // UI Options
    [JsonPropertyName("floating_toolbar_enabled")]
    public bool FloatingToolbarEnabled { get; set; } = false;
    
    [JsonPropertyName("indicator_enabled")]
    public bool IndicatorEnabled { get; set; } = false;
    
    [JsonPropertyName("auto_stop_dictation")]
    public bool AutoStopDictation { get; set; } = false;
    
    [JsonPropertyName("dead_man_switch")]
    public bool DeadManSwitch { get; set; } = false;
    
    [JsonPropertyName("iv_report_hotkey")]
    public string IvReportHotkey { get; set; } = "v";

    [JsonPropertyName("critical_findings_template")]
    public string CriticalFindingsTemplate { get; set; } = "Critical findings were discussed with and acknowledged by {name} at {time} on {date}.";

    [JsonPropertyName("series_image_template")]
    public string SeriesImageTemplate { get; set; } = "(series {series}, image {image})";
    
    [JsonPropertyName("restore_focus_after_action")]
    public bool RestoreFocusAfterAction { get; set; } = true;
    
    [JsonPropertyName("scroll_to_bottom_on_process")]
    public bool ScrollToBottomOnProcess { get; set; } = false;

    [JsonPropertyName("scrape_mosaic_enabled")]
    public bool ScrapeMosaicEnabled { get; set; } = false;

    [JsonPropertyName("scrape_interval_seconds")]
    public int ScrapeIntervalSeconds { get; set; } = 3;

    [JsonPropertyName("show_clinical_history")]
    public bool ShowClinicalHistory { get; set; } = false;

    [JsonPropertyName("auto_fix_clinical_history")]
    public bool AutoFixClinicalHistory { get; set; } = false;

    [JsonPropertyName("show_drafted_indicator")]
    public bool ShowDraftedIndicator { get; set; } = false;

    [JsonPropertyName("show_template_mismatch")]
    public bool ShowTemplateMismatch { get; set; } = false;

    [JsonPropertyName("gender_check_enabled")]
    public bool GenderCheckEnabled { get; set; } = false;

    [JsonPropertyName("clinical_history_x")]
    public int ClinicalHistoryX { get; set; } = 100;

    [JsonPropertyName("clinical_history_y")]
    public int ClinicalHistoryY { get; set; } = 200;

    [JsonPropertyName("show_impression")]
    public bool ShowImpression { get; set; } = false;

    [JsonPropertyName("impression_x")]
    public int ImpressionX { get; set; } = 300;

    [JsonPropertyName("impression_y")]
    public int ImpressionY { get; set; } = 200;

    // Smart Scroll Thresholds
    [JsonPropertyName("scroll_threshold_1")]
    public int ScrollThreshold1 { get; set; } = 10; // 0 -> 1 PgDn
    
    [JsonPropertyName("scroll_threshold_2")]
    public int ScrollThreshold2 { get; set; } = 30; // 1 -> 2 PgDn
    
    [JsonPropertyName("scroll_threshold_3")]
    public int ScrollThreshold3 { get; set; } = 50; // 2 -> 3 PgDn

    [JsonPropertyName("show_line_count_toast")]
    public bool ShowLineCountToast { get; set; } = false;
    
    // Window Positions
    [JsonPropertyName("window_x")]
    public int WindowX { get; set; } = 100;
    
    [JsonPropertyName("window_y")]
    public int WindowY { get; set; } = 100;
    
    [JsonPropertyName("floating_toolbar_x")]
    public int FloatingToolbarX { get; set; } = 100;
    
    [JsonPropertyName("floating_toolbar_y")]
    public int FloatingToolbarY { get; set; } = 100;
    
    [JsonPropertyName("indicator_x")]
    public int IndicatorX { get; set; } = 100;
    
    [JsonPropertyName("indicator_y")]
    public int IndicatorY { get; set; } = 550;
    
    [JsonPropertyName("report_popup_x")]
    public int ReportPopupX { get; set; } = 300;
    
    [JsonPropertyName("report_popup_y")]
    public int ReportPopupY { get; set; } = 300;
    
    [JsonPropertyName("report_popup_width")]
    public int ReportPopupWidth { get; set; } = 750;

    [JsonPropertyName("settings_x")]
    public int SettingsX { get; set; } = 200;

    [JsonPropertyName("settings_y")]
    public int SettingsY { get; set; } = 200;
    
    // Action Mappings (action name -> {hotkey, mic_button})
    [JsonPropertyName("action_mappings")]
    public Dictionary<string, ActionMapping> ActionMappings { get; set; } = new();
    
    // Floating Buttons Configuration
    [JsonPropertyName("floating_buttons")]
    public FloatingButtonsConfig FloatingButtons { get; set; } = new();
    
    /// <summary>
    /// Load configuration from JSON file, or show onboarding if not found.
    /// </summary>
    public static Configuration Load()
    {
        // 1. Migration: If new path doesn't exist, but old path DOES, move it.
        string oldPath = Path.Combine(AppContext.BaseDirectory, "MosaicToolsSettings.json");
        if (!File.Exists(SettingsPath) && File.Exists(oldPath))
        {
            try
            {
                var dir = Path.GetDirectoryName(SettingsPath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.Move(oldPath, SettingsPath);
                Logger.Trace("Settings migrated to AppData location.");
            }
            catch (Exception ex)
            {
                Logger.Trace($"Migration failed: {ex.Message}");
            }
        }

        if (File.Exists(SettingsPath))
        {
            try
            {
                var json = File.ReadAllText(SettingsPath);
                var config = JsonSerializer.Deserialize<Configuration>(json);
                if (config != null)
                {
                    config.EnsureDefaults();
                    return config;
                }
            }
            catch (Exception ex)
            {
                Logger.Trace($"Error loading settings: {ex.Message}");
            }
        }
        
        // First run - show onboarding
        return ShowOnboarding();
    }
    
    private static Configuration ShowOnboarding()
    {
        System.Windows.Forms.MessageBox.Show(
            "Welcome to Mosaic Tools!\n\n" +
            "The tool needs to know your last name so it can correctly identify " +
            "and filter your name out of clinical statements when parsing Clario exam notes.\n\n" +
            "Example: If the note says 'Transferred Smith to Jones', and you are Dr. Smith, " +
            "the tool will know that Jones is the contact person.",
            "First Run Onboarding",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);
        
        var name = InputBox.Show(
            "Enter your last name (e.g., Smith):",
            "Doctor Name",
            "Radiologist");
        
        if (string.IsNullOrWhiteSpace(name))
            name = "Radiologist";
        
        var config = new Configuration { DoctorName = name.Trim() };
        config.EnsureDefaults();
        config.Save();
        
        return config;
    }
    
    /// <summary>
    /// Ensure all action mappings have defaults.
    /// </summary>
    private void EnsureDefaults()
    {
        // Default action mappings if empty
        if (ActionMappings.Count == 0)
        {
            ActionMappings = new Dictionary<string, ActionMapping>
            {
                [Actions.GetPrior] = new() { Hotkey = "", MicButton = "Left Button" },
                [Actions.CriticalFindings] = new() { Hotkey = "", MicButton = "Skip Forward" },
                [Actions.SystemBeep] = new() { Hotkey = "", MicButton = "" },
                [Actions.ShowReport] = new() { Hotkey = "", MicButton = "" },
                [Actions.CaptureSeries] = new() { Hotkey = "", MicButton = "Right Button" },
                [Actions.ToggleRecord] = new() { Hotkey = "", MicButton = "" },
                [Actions.ProcessReport] = new() { Hotkey = "", MicButton = "Skip Back" },
                [Actions.SignReport] = new() { Hotkey = "", MicButton = "Checkmark" }
            };
        }
        
        // Ensure floating buttons have defaults
        if (FloatingButtons.Buttons.Count == 0)
        {
            FloatingButtons = FloatingButtonsConfig.Default;
        }
    }
    
    public void Save()
    {
        try
        {
            var dir = Path.GetDirectoryName(SettingsPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(this, options);
            File.WriteAllText(SettingsPath, json);
        }
        catch (Exception ex)
        {
            Logger.Trace($"Error saving settings: {ex.Message}");
        }
    }
}

public class ActionMapping
{
    [JsonPropertyName("hotkey")]
    public string Hotkey { get; set; } = "";
    
    [JsonPropertyName("mic_button")]
    public string MicButton { get; set; } = "";
}

public class FloatingButtonsConfig
{
    [JsonPropertyName("columns")]
    public int Columns { get; set; } = 2;
    
    [JsonPropertyName("buttons")]
    public List<FloatingButtonDef> Buttons { get; set; } = new();
    
    public static FloatingButtonsConfig Default => new()
    {
        Columns = 2,
        Buttons = new List<FloatingButtonDef>
        {
            new() { Type = "square", Icon = "↕", Label = "", Keystroke = "ctrl+v" },
            new() { Type = "square", Icon = "↔", Label = "", Keystroke = "ctrl+h" },
            new() { Type = "square", Icon = "↺", Label = "", Keystroke = "," },
            new() { Type = "square", Icon = "↻", Label = "", Keystroke = "." },
            new() { Type = "wide", Icon = "", Label = "Zoom Out", Keystroke = "-" }
        }
    };
}

public class FloatingButtonDef
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "square";
    
    [JsonPropertyName("icon")]
    public string Icon { get; set; } = "";
    
    [JsonPropertyName("label")]
    public string Label { get; set; } = "";
    
    [JsonPropertyName("keystroke")]
    public string Keystroke { get; set; } = "";
}

/// <summary>
/// Action name constants matching Python's ACTION_* constants.
/// </summary>
public static class Actions
{
    public const string None = "None";
    public const string SystemBeep = "System Beep";
    public const string GetPrior = "Get Prior";
    public const string CriticalFindings = "Critical Findings";
    public const string ShowReport = "Show Report";
    public const string CaptureSeries = "Capture Series/Image";
    public const string ToggleRecord = "Start/Stop Recording";
    public const string ProcessReport = "Process Report";
    public const string SignReport = "Sign Report";

    public static readonly string[] All = {
        None, SystemBeep, GetPrior, CriticalFindings,
        ShowReport, CaptureSeries, ToggleRecord, ProcessReport, SignReport
    };
}
