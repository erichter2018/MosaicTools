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
    private static readonly object _saveLock = new();

    public static readonly string SettingsPath = Path.Combine(
        string.IsNullOrEmpty(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
            ? AppContext.BaseDirectory
            : Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
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

    [JsonPropertyName("target_timezone")]
    public string? TargetTimezone { get; set; } = null;  // null = keep original timezone from note

    [JsonPropertyName("preferred_microphone")]
    public string PreferredMicrophone { get; set; } = "Auto";  // Auto, PowerMic, or SpeechMike

    [JsonPropertyName("auto_update_enabled")]
    public bool AutoUpdateEnabled { get; set; } = true;

    [JsonPropertyName("hide_clinical_history_when_no_study")]
    public bool HideClinicalHistoryWhenNoStudy { get; set; } = true;

    [JsonPropertyName("hide_indicator_when_no_study")]
    public bool HideIndicatorWhenNoStudy { get; set; } = true;

    // Beep Settings
    [JsonPropertyName("start_beep_enabled")]
    public bool StartBeepEnabled { get; set; } = true;
    
    [JsonPropertyName("stop_beep_enabled")]
    public bool StopBeepEnabled { get; set; } = true;
    
    [JsonPropertyName("start_beep_volume")]
    public double StartBeepVolume { get; set; } = 0.08;

    [JsonPropertyName("stop_beep_volume")]
    public double StopBeepVolume { get; set; } = 0.08;
    
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

    [JsonPropertyName("comparison_template")]
    public string ComparisonTemplate { get; set; } = "COMPARISON: {date} {time} {description}. {noimages}";

    [JsonPropertyName("separate_pasted_items")]
    public bool SeparatePastedItems { get; set; } = true;

    [JsonPropertyName("restore_focus_after_action")]
    public bool RestoreFocusAfterAction { get; set; } = true;  // Always true, not user-configurable
    
    [JsonPropertyName("scroll_to_bottom_on_process")]
    public bool ScrollToBottomOnProcess { get; set; } = false;

    [JsonPropertyName("scrape_mosaic_enabled")]
    public bool ScrapeMosaicEnabled { get; set; } = true;  // Always true, not user-configurable

    [JsonPropertyName("scrape_interval_seconds")]
    public int ScrapeIntervalSeconds { get; set; } = 3;

    [JsonPropertyName("show_clinical_history")]
    public bool ShowClinicalHistory { get; set; } = false;

    [JsonPropertyName("always_show_clinical_history")]
    public bool AlwaysShowClinicalHistory { get; set; } = true;

    [JsonPropertyName("auto_fix_clinical_history")]
    public bool AutoFixClinicalHistory { get; set; } = false;

    [JsonPropertyName("show_drafted_indicator")]
    public bool ShowDraftedIndicator { get; set; } = false;

    [JsonPropertyName("show_template_mismatch")]
    public bool ShowTemplateMismatch { get; set; } = false;

    [JsonPropertyName("gender_check_enabled")]
    public bool GenderCheckEnabled { get; set; } = false;

    [JsonPropertyName("aidoc_scrape_enabled")]
    public bool AidocScrapeEnabled { get; set; } = false;  // Off by default (not everyone has Aidoc)

    [JsonPropertyName("recomd_enabled")]
    public bool RecoMdEnabled { get; set; } = false;

    [JsonPropertyName("recomd_auto_on_process")]
    public bool RecoMdAutoOnProcess { get; set; } = false;

    [JsonPropertyName("stroke_detection_enabled")]
    public bool StrokeDetectionEnabled { get; set; } = false;

    [JsonPropertyName("stroke_detection_use_clinical_history")]
    public bool StrokeDetectionUseClinicalHistory { get; set; } = false;

    [JsonPropertyName("stroke_click_to_create_note")]
    public bool StrokeClickToCreateNote { get; set; } = false;

    [JsonPropertyName("stroke_auto_create_note")]
    public bool StrokeAutoCreateNote { get; set; } = false;

    [JsonPropertyName("track_critical_studies")]
    public bool TrackCriticalStudies { get; set; } = true;

    // Ignore Inpatient Drafted - auto-select all text after insertions for matching studies
    [JsonPropertyName("ignore_inpatient_drafted")]
    public bool IgnoreInpatientDrafted { get; set; } = false;

    // 0 = All Inpatient XR, 1 = Inpatient Chest XR only
    [JsonPropertyName("ignore_inpatient_drafted_mode")]
    public int IgnoreInpatientDraftedMode { get; set; } = 0;

    [JsonPropertyName("stroke_clinical_history_keywords")]
    public List<string> StrokeClinicalHistoryKeywords { get; set; } = new()
    {
        "stroke", "CVA", "TIA", "hemiparesis", "hemiplegia", "aphasia",
        "dysarthria", "facial droop", "weakness", "numbness", "code stroke",
        "NIH stroke scale", "NIHSS"
    };

    [JsonPropertyName("clinical_history_x")]
    public int ClinicalHistoryX { get; set; } = 100;

    [JsonPropertyName("clinical_history_y")]
    public int ClinicalHistoryY { get; set; } = 200;

    [JsonPropertyName("show_impression")]
    public bool ShowImpression { get; set; } = false;

    [JsonPropertyName("impression_deletable_points")]
    public bool ImpressionDeletablePoints { get; set; } = false;

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

    [JsonPropertyName("report_popup_font_family")]
    public string ReportPopupFontFamily { get; set; } = "Consolas";

    [JsonPropertyName("report_popup_font_size")]
    public float ReportPopupFontSize { get; set; } = 11f;

    [JsonPropertyName("settings_x")]
    public int SettingsX { get; set; } = 200;

    [JsonPropertyName("settings_y")]
    public int SettingsY { get; set; } = 200;

    // Macros
    [JsonPropertyName("macros_enabled")]
    public bool MacrosEnabled { get; set; } = false;

    [JsonPropertyName("macros_blank_lines_before")]
    public bool MacrosBlankLinesBefore { get; set; } = false;

    [JsonPropertyName("macros")]
    public List<MacroConfig> Macros { get; set; } = new();

    // RVUCounter Integration
    [JsonPropertyName("rvucounter_enabled")]
    public bool RvuCounterEnabled { get; set; } = true;

    [JsonPropertyName("rvucounter_path")]
    public string RvuCounterPath { get; set; } = "";

    [JsonPropertyName("rvu_display_mode")]
    public RvuDisplayMode RvuDisplayMode { get; set; } = RvuDisplayMode.Total;

    [JsonPropertyName("rvu_metrics")]
    public RvuMetric RvuMetrics { get; set; } = RvuMetric.None;

    [JsonPropertyName("rvu_metrics_migrated")]
    public bool RvuMetricsMigrated { get; set; } = false;

    [JsonPropertyName("rvu_overflow_layout")]
    public RvuOverflowLayout RvuOverflowLayout { get; set; } = RvuOverflowLayout.Horizontal;

    [JsonPropertyName("rvu_goal_enabled")]
    public bool RvuGoalEnabled { get; set; } = false;

    [JsonPropertyName("rvu_goal_per_hour")]
    public double RvuGoalPerHour { get; set; } = 10.0;

    [JsonPropertyName("pace_car_enabled")]
    public bool PaceCarEnabled { get; set; } = false;

    [JsonPropertyName("pace_car_alternate_seconds")]
    public int PaceCarAlternateSeconds { get; set; } = 8;

    // Distraction Alert (beep volume when RVUCounter sends alert)
    [JsonPropertyName("distraction_alert_volume")]
    public double DistractionAlertVolume { get; set; } = 0.15;

    // Report Changes Highlighting
    [JsonPropertyName("show_report_changes")]
    public bool ShowReportChanges { get; set; } = false;

    [JsonPropertyName("report_changes_color")]
    public string ReportChangesColor { get; set; } = "#90EE90"; // Light green

    [JsonPropertyName("report_changes_alpha")]
    public int ReportChangesAlpha { get; set; } = 25; // 0-100 percent opacity

    [JsonPropertyName("correlation_enabled")]
    public bool CorrelationEnabled { get; set; } = false;

    [JsonPropertyName("orphan_findings_enabled")]
    public bool OrphanFindingsEnabled { get; set; } = false;

    [JsonPropertyName("report_popup_transparent")]
    public bool ReportPopupTransparent { get; set; } = true;

    [JsonPropertyName("report_popup_transparency")]
    public int ReportPopupTransparency { get; set; } = 70; // 0-100%, maps to background alpha

    [JsonPropertyName("show_report_after_process")]
    public bool ShowReportAfterProcess { get; set; } = false;

    [JsonPropertyName("template_database_enabled")]
    public bool TemplateDatabaseEnabled { get; set; } = true;

    // Version tracking for What's New popup
    [JsonPropertyName("last_seen_version")]
    public string? LastSeenVersion { get; set; } = null;

    // Pick Lists
    [JsonPropertyName("pick_lists_enabled")]
    public bool PickListsEnabled { get; set; } = false;

    [JsonPropertyName("pick_lists")]
    public List<PickListConfig> PickLists { get; set; } = new();

    [JsonPropertyName("pick_list_skip_single_match")]
    public bool PickListSkipSingleMatch { get; set; } = true;

    [JsonPropertyName("pick_list_keep_open")]
    public bool PickListKeepOpen { get; set; } = false;

    [JsonPropertyName("pick_list_popup_x")]
    public int PickListPopupX { get; set; } = 400;

    [JsonPropertyName("pick_list_popup_y")]
    public int PickListPopupY { get; set; } = 300;

    // [CustomSTT] Custom STT Mode settings
    [JsonPropertyName("custom_stt_enabled")]
    public bool CustomSttEnabled { get; set; } = false;

    [JsonPropertyName("stt_api_key")]
    public string SttApiKey { get; set; } = ""; // Deepgram API key (kept for backward compat)

    [JsonPropertyName("stt_assemblyai_api_key")]
    public string SttAssemblyAIApiKey { get; set; } = "";

    [JsonPropertyName("stt_corti_client_id")]
    public string SttCortiClientId { get; set; } = "";

    [JsonPropertyName("stt_corti_client_secret")]
    public string SttCortiClientSecret { get; set; } = "";

    [JsonPropertyName("stt_corti_environment")]
    public string SttCortiEnvironment { get; set; } = "us"; // "us" or "eu"

    [JsonPropertyName("stt_speechmatics_api_key")]
    public string SttSpeechmaticsApiKey { get; set; } = "";

    [JsonPropertyName("stt_speechmatics_region")]
    public string SttSpeechmaticsRegion { get; set; } = "us"; // "us" or "eu"

    [JsonPropertyName("stt_provider")]
    public string SttProvider { get; set; } = "deepgram";

    [JsonPropertyName("stt_model")]
    public string SttModel { get; set; } = "nova-3-medical";

    [JsonPropertyName("stt_audio_device_name")]
    public string SttAudioDeviceName { get; set; } = "";

    [JsonPropertyName("stt_auto_punctuate")]
    public bool SttAutoPunctuate { get; set; } = false;

    [JsonPropertyName("stt_start_beep_enabled")]
    public bool SttStartBeepEnabled { get; set; } = true;

    [JsonPropertyName("stt_stop_beep_enabled")]
    public bool SttStopBeepEnabled { get; set; } = true;

    [JsonPropertyName("stt_start_beep_volume")]
    public double SttStartBeepVolume { get; set; } = 0.08;

    [JsonPropertyName("stt_stop_beep_volume")]
    public double SttStopBeepVolume { get; set; } = 0.08;

    [JsonPropertyName("stt_show_indicator")]
    public bool SttShowIndicator { get; set; } = true;

    [JsonPropertyName("transcription_form_x")]
    public int TranscriptionFormX { get; set; } = 400;

    [JsonPropertyName("transcription_form_y")]
    public int TranscriptionFormY { get; set; } = 400;

    [JsonPropertyName("transcription_form_width")]
    public int TranscriptionFormWidth { get; set; } = 500;

    [JsonPropertyName("transcription_form_height")]
    public int TranscriptionFormHeight { get; set; } = 300;

    [JsonPropertyName("radai_auto_on_process")]  // [RadAI]
    public bool RadAiAutoOnProcess { get; set; } = false;

    [JsonPropertyName("radai_popup_x")]  // [RadAI]
    public int RadAiPopupX { get; set; } = -1;

    [JsonPropertyName("radai_popup_y")]  // [RadAI]
    public int RadAiPopupY { get; set; } = -1;

    [JsonPropertyName("pick_list_editor_width")]
    public int PickListEditorWidth { get; set; } = 900;

    [JsonPropertyName("pick_list_editor_height")]
    public int PickListEditorHeight { get; set; } = 600;

    // Connectivity Monitor Settings
    [JsonPropertyName("connectivity_monitor_enabled")]
    public bool ConnectivityMonitorEnabled { get; set; } = false;

    [JsonPropertyName("connectivity_check_interval_seconds")]
    public int ConnectivityCheckIntervalSeconds { get; set; } = 30;

    [JsonPropertyName("connectivity_timeout_ms")]
    public int ConnectivityTimeoutMs { get; set; } = 5000;

    [JsonPropertyName("connectivity_servers")]
    public List<ServerConfig> ConnectivityServers { get; set; } = new();

    // Experimental insertion mode toggle:
    // false = clipboard + Ctrl+V (default)
    // true = direct SendInput text insertion (no Ctrl+V)
    [JsonPropertyName("experimental_use_sendinput_insert")]
    public bool ExperimentalUseSendInputInsert { get; set; } = false;

    // UI Options
    [JsonPropertyName("show_tooltips")]
    public bool ShowTooltips { get; set; } = true;

    // InteleViewer Window/Level Cycle Keystrokes
    // Keys sent to InteleViewer when Cycle Window/Level action is triggered
    [JsonPropertyName("window_level_keys")]
    public List<string> WindowLevelKeys { get; set; } = new() { "F4", "F5", "F7", "F6" };

    // Action Mappings for PowerMic (action name -> {hotkey, mic_button})
    [JsonPropertyName("action_mappings")]
    public Dictionary<string, ActionMapping> ActionMappings { get; set; } = new();

    // Action Mappings for SpeechMike (action name -> {hotkey, mic_button})
    [JsonPropertyName("speechmike_action_mappings")]
    public Dictionary<string, ActionMapping> SpeechMikeActionMappings { get; set; } = new();

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
                // Back up corrupt file before overwriting
                try
                {
                    var backupPath = SettingsPath + ".corrupt";
                    File.Copy(SettingsPath, backupPath, overwrite: true);
                    Logger.Trace($"Backed up corrupt settings to {backupPath}");
                }
                catch { }
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
        // Default action mappings for PowerMic if empty
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

        // Default action mappings for SpeechMike if empty
        if (SpeechMikeActionMappings.Count == 0)
        {
            SpeechMikeActionMappings = new Dictionary<string, ActionMapping>
            {
                [Actions.GetPrior] = new() { Hotkey = "", MicButton = "Left Button" },
                [Actions.CriticalFindings] = new() { Hotkey = "", MicButton = "Skip Forward" },
                [Actions.SystemBeep] = new() { Hotkey = "", MicButton = "" },
                [Actions.ShowReport] = new() { Hotkey = "", MicButton = "" },
                [Actions.CaptureSeries] = new() { Hotkey = "", MicButton = "T Button" },
                [Actions.ToggleRecord] = new() { Hotkey = "", MicButton = "" },
                [Actions.ProcessReport] = new() { Hotkey = "", MicButton = "Skip Back" },
                [Actions.SignReport] = new() { Hotkey = "", MicButton = "Checkmark" }
            };
        }

        // Ensure macros have IDs (migration for pre-ID macros)
        foreach (var macro in Macros)
        {
            if (string.IsNullOrEmpty(macro.Id))
                macro.Id = Guid.NewGuid().ToString("N")[..8];
        }

        // Migrate pick list category options from legacy format
        foreach (var pickList in PickLists)
        {
            foreach (var category in pickList.Categories)
            {
                category.MigrateOptions();
            }
        }

        // Migrate legacy RvuDisplayMode to RvuMetrics flags (one-time)
        // RvuMetricsMigrated prevents re-triggering when user unchecks all metrics (None == 0)
        if (!RvuMetricsMigrated)
        {
            // Only overwrite if RvuMetrics is still at default (None) — don't clobber
            // existing selections from users who already used the new multi-select UI
            if (RvuMetrics == RvuMetric.None)
            {
                RvuMetrics = RvuDisplayMode switch
                {
                    RvuDisplayMode.Total => RvuMetric.Total,
                    RvuDisplayMode.PerHour => RvuMetric.PerHour,
                    RvuDisplayMode.Both => RvuMetric.Total | RvuMetric.PerHour,
                    _ => RvuMetric.Total
                };
            }
            RvuMetricsMigrated = true;
        }

        // Force mandatory settings (not user-configurable)
        ScrapeMosaicEnabled = true;
        RestoreFocusAfterAction = true;

        // Ensure floating buttons have defaults
        if (FloatingButtons.Buttons.Count == 0)
        {
            FloatingButtons = FloatingButtonsConfig.Default;
        }

        // Ensure connectivity servers have defaults
        // Using public DNS servers as placeholders until real IPs are configured
        var defaultServers = new List<(string Name, string Host)>
        {
            ("Mirth", "8.8.8.8"),           // Google DNS (placeholder)
            ("Mosaic", "1.1.1.1"),          // Cloudflare DNS (placeholder)
            ("Clario", "208.67.222.222"),   // OpenDNS (placeholder)
            ("InteleViewer", "9.9.9.9")     // Quad9 DNS (placeholder)
        };

        if (ConnectivityServers.Count == 0)
        {
            ConnectivityServers = defaultServers.Select(s => new ServerConfig
            {
                Name = s.Name,
                Host = s.Host,
                Port = 0,
                Enabled = true
            }).ToList();
        }
        else
        {
            // Ensure existing servers have hosts (migrate from old empty configs)
            foreach (var def in defaultServers)
            {
                var existing = ConnectivityServers.FirstOrDefault(s => s.Name == def.Name);
                if (existing != null && string.IsNullOrWhiteSpace(existing.Host))
                {
                    existing.Host = def.Host;
                }
            }
        }

        // Migrate "RecoMD" action name to "Trigger RecoMD" (v3.6+)
        void MigrateRecoMdAction(Dictionary<string, ActionMapping> mappings)
        {
            if (mappings.TryGetValue("RecoMD", out var mapping))
            {
                mappings.Remove("RecoMD");
                mappings[Actions.TriggerRecoMd] = mapping;
            }
        }
        MigrateRecoMdAction(ActionMappings);
        MigrateRecoMdAction(SpeechMikeActionMappings);

    }

    /// <summary>
    /// Check if a study matches the Ignore Inpatient Drafted criteria.
    /// Returns true if the study is inpatient XR (or inpatient chest XR if mode=1).
    /// </summary>
    public bool ShouldIgnoreInpatientDrafted(string? clarioClass, string? description)
    {
        if (!IgnoreInpatientDrafted) return false;
        if (string.IsNullOrEmpty(clarioClass) || string.IsNullOrEmpty(description)) return false;

        // Check inpatient
        var classLower = clarioClass.ToLowerInvariant();
        if (!classLower.Contains("inpatient") && !classLower.Contains(" ip") && !classLower.StartsWith("ip ") && classLower != "ip") return false;

        // Check XR
        var descLower = description.ToLowerInvariant();
        bool isXr = descLower.Contains("xr") || descLower.Contains("x-ray") || descLower.Contains("xray");
        if (!isXr) return false;

        // Mode 0 = All Inpatient XR, Mode 1 = Chest only
        if (IgnoreInpatientDraftedMode == 1)
        {
            return descLower.Contains("chest");
        }
        return true;
    }

    public void Save()
    {
        try
        {
            var dir = Path.GetDirectoryName(SettingsPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(this, options);
            lock (_saveLock)
            {
                File.WriteAllText(SettingsPath, json);
            }
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

/// <summary>
/// Configuration for a single macro - text snippet to insert for matching studies.
/// </summary>
public class MacroConfig
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = Guid.NewGuid().ToString("N")[..8];

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; } = true;

    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    /// <summary>
    /// Comma-separated required terms. ALL must match the study description.
    /// Empty = matches all studies (global macro).
    /// </summary>
    [JsonPropertyName("criteria_required")]
    public string CriteriaRequired { get; set; } = "";

    /// <summary>
    /// Comma-separated optional terms. At least ONE must match (if specified).
    /// Empty = no additional OR requirement.
    /// </summary>
    [JsonPropertyName("criteria_any_of")]
    public string CriteriaAnyOf { get; set; } = "";

    /// <summary>
    /// Comma-separated exclusion terms. NONE must match (if specified).
    /// Empty = no exclusion filter.
    /// </summary>
    [JsonPropertyName("criteria_exclude")]
    public string CriteriaExclude { get; set; } = "";

    /// <summary>
    /// Voice-triggered macro — activated by speaking "insert/input/macro {name}" during STT dictation.
    /// Voice macros ignore study criteria (always available when STT is active).
    /// </summary>
    [JsonPropertyName("voice")]
    public bool Voice { get; set; } = false;

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";

    /// <summary>
    /// Check if this macro matches the given study description.
    /// Logic: (all Required match) AND (at least one AnyOf, if specified) AND (none of Exclude match)
    /// All empty = global macro (matches all).
    /// </summary>
    public bool MatchesStudy(string? studyDescription)
    {
        bool hasRequired = !string.IsNullOrWhiteSpace(CriteriaRequired);
        bool hasAnyOf = !string.IsNullOrWhiteSpace(CriteriaAnyOf);
        bool hasExclude = !string.IsNullOrWhiteSpace(CriteriaExclude);

        // Global macro - matches all (but still check exclusions)
        if (!hasRequired && !hasAnyOf && !hasExclude)
            return true;

        if (string.IsNullOrWhiteSpace(studyDescription))
            return false;

        var description = studyDescription.ToUpperInvariant();

        // Check Exclude terms first - if any match, reject
        if (hasExclude)
        {
            var excludeTerms = CriteriaExclude.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var term in excludeTerms)
            {
                if (description.Contains(term.ToUpperInvariant()))
                    return false;
            }
        }

        // Check Required terms (AND logic) - all must match
        if (hasRequired)
        {
            var requiredTerms = CriteriaRequired.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var term in requiredTerms)
            {
                if (!description.Contains(term.ToUpperInvariant()))
                    return false;
            }
        }

        // Check AnyOf terms (OR logic) - at least one must match
        if (hasAnyOf)
        {
            var anyOfTerms = CriteriaAnyOf.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            bool anyMatch = false;
            foreach (var term in anyOfTerms)
            {
                if (description.Contains(term.ToUpperInvariant()))
                {
                    anyMatch = true;
                    break;
                }
            }
            if (!anyMatch)
                return false;
        }

        return true;
    }

    /// <summary>
    /// Returns a display string for the criteria (for list preview).
    /// </summary>
    public string GetCriteriaDisplayString()
    {
        if (Voice)
            return "Voice triggered";

        if (string.IsNullOrWhiteSpace(CriteriaRequired) &&
            string.IsNullOrWhiteSpace(CriteriaAnyOf) &&
            string.IsNullOrWhiteSpace(CriteriaExclude))
            return "All studies";

        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(CriteriaRequired))
            parts.Add(CriteriaRequired.Trim());
        if (!string.IsNullOrWhiteSpace(CriteriaAnyOf))
            parts.Add($"({CriteriaAnyOf.Trim()})");
        if (!string.IsNullOrWhiteSpace(CriteriaExclude))
            parts.Add($"-{CriteriaExclude.Trim()}");

        return string.Join(" + ", parts);
    }
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

    [JsonPropertyName("action")]
    public string Action { get; set; } = "";
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
    public const string CreateImpression = "Create Impression";
    public const string DiscardStudy = "Discard Study";
    public const string ShowPickLists = "Show Pick Lists";
    public const string CycleWindowLevel = "Cycle Window/Level";
    public const string CreateCriticalNote = "Create Critical Note";
    public const string RadAiImpression = "RadAI Impression";  // [RadAI] — remove when RadAI integration is retired
    public const string TriggerRecoMd = "Trigger RecoMD";
    public const string PasteRecoMd = "Paste RecoMD";

    public static readonly string[] All = {
        None, SystemBeep, GetPrior, CriticalFindings,
        ShowReport, CaptureSeries, ToggleRecord, ProcessReport, SignReport, CreateImpression, DiscardStudy, ShowPickLists, CycleWindowLevel, CreateCriticalNote, RadAiImpression, TriggerRecoMd, PasteRecoMd
    };
}

/// <summary>
/// Pick list mode - Tree for hierarchical navigation, Builder for sentence construction.
/// </summary>
public enum PickListMode
{
    Tree,
    Builder
}

/// <summary>
/// RVU display mode - what to show in the main window (legacy, kept for migration).
/// </summary>
public enum RvuDisplayMode
{
    Total,
    PerHour,
    Both
}

/// <summary>
/// Flags enum for which RVU metrics to display.
/// </summary>
[Flags]
public enum RvuMetric
{
    None = 0,
    Total = 1,
    PerHour = 2,
    CurrentHour = 4,
    PriorHour = 8,
    EstimatedTotal = 16,
    RvuPerStudy = 32,
    AvgPerHour = 64
}

/// <summary>
/// Layout mode when 3+ RVU metrics are selected.
/// </summary>
public enum RvuOverflowLayout
{
    Horizontal = 0,
    VerticalStack = 1,
    HoverPopup = 2,
    Carousel = 3
}

/// <summary>
/// Tree pick list style - controls how selections are formatted when pasted.
/// </summary>
public enum TreePickListStyle
{
    /// <summary>
    /// Each selection pastes immediately with a leading space (if text doesn't start with one).
    /// Result: "The lungs are clear. Heart size is normal."
    /// </summary>
    Freeform,

    /// <summary>
    /// Accumulates selections per top-level category, formats output as uppercase headings.
    /// Result: "LUNGS: RLL Opacity\nHEART: Mildly Enlarged Heart"
    /// </summary>
    Structured
}

/// <summary>
/// Controls text placement in structured tree mode output.
/// </summary>
public enum StructuredTextPlacement
{
    /// <summary>
    /// Text on same line as heading: "LUNGS: Clear"
    /// </summary>
    Inline,

    /// <summary>
    /// Text on line below heading with indent: "LUNGS:\n  Clear"
    /// </summary>
    BelowHeading
}

/// <summary>
/// A builder option that can be plain text, a macro reference, or a tree pick list reference.
/// </summary>
public class BuilderOption
{
    [JsonPropertyName("text")]
    public string Text { get; set; } = "";

    [JsonPropertyName("macro_id")]
    public string? MacroId { get; set; }

    [JsonPropertyName("tree_list_id")]
    public string? TreeListId { get; set; }

    [JsonIgnore]
    public bool IsMacroRef => !string.IsNullOrEmpty(MacroId);

    [JsonIgnore]
    public bool IsTreeRef => !string.IsNullOrEmpty(TreeListId);

    [JsonIgnore]
    public bool IsPlainText => !IsMacroRef && !IsTreeRef;
}

/// <summary>
/// A category in Builder mode pick lists. Each category has options to choose from.
/// </summary>
public class PickListCategory
{
    /// <summary>
    /// Display name for this category (e.g., "Degree of Severity", "Location").
    /// </summary>
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    /// <summary>
    /// Legacy plain-text options. Kept for backwards compatibility during deserialization.
    /// On load, if options_v2 is empty but options has items, they are migrated to options_v2.
    /// </summary>
    [JsonPropertyName("options")]
    public List<string> OptionsLegacy { get; set; } = new();

    /// <summary>
    /// Available options for this category (v2 format with reference support).
    /// Keys 1-9 select first 9 options; additional options selectable by click/arrow.
    /// </summary>
    [JsonPropertyName("options_v2")]
    public List<BuilderOption> Options { get; set; } = new();

    /// <summary>
    /// Text appended after the selected option. Default is a single space.
    /// Use empty string for no separator, or custom text like ", " or " and ".
    /// </summary>
    [JsonPropertyName("separator")]
    public string Separator { get; set; } = " ";

    /// <summary>
    /// Indices of options that complete the sentence immediately (terminal options).
    /// When a terminal option is selected, the sentence is finalized without continuing to subsequent categories.
    /// </summary>
    [JsonPropertyName("terminal_options")]
    public List<int> TerminalOptions { get; set; } = new();

    /// <summary>
    /// Check if the option at the given index is terminal.
    /// </summary>
    public bool IsTerminal(int index) => TerminalOptions.Contains(index);

    /// <summary>
    /// Migrate legacy options to v2 format if needed.
    /// </summary>
    public void MigrateOptions()
    {
        if (Options.Count == 0 && OptionsLegacy.Count > 0)
        {
            Options = OptionsLegacy.Select(text => new BuilderOption { Text = text }).ToList();
            OptionsLegacy.Clear();
        }
    }
}

/// <summary>
/// Configuration for a pick list - a category of text snippets.
/// </summary>
public class PickListConfig
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = Guid.NewGuid().ToString("N")[..8];

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; } = true;

    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    /// <summary>
    /// Pick list mode - Tree (hierarchical) or Builder (sentence construction).
    /// </summary>
    [JsonPropertyName("mode")]
    public PickListMode Mode { get; set; } = PickListMode.Tree;

    /// <summary>
    /// Tree pick list style - Freeform (immediate paste) or Structured (accumulated headings).
    /// Only applies when Mode is Tree.
    /// </summary>
    [JsonPropertyName("tree_style")]
    public TreePickListStyle TreeStyle { get; set; } = TreePickListStyle.Freeform;

    /// <summary>
    /// Text placement in structured mode - Inline (LUNGS: Clear) or BelowHeading (LUNGS:\n  Clear).
    /// Only applies when Mode is Tree and TreeStyle is Structured.
    /// </summary>
    [JsonPropertyName("structured_text_placement")]
    public StructuredTextPlacement StructuredTextPlacement { get; set; } = StructuredTextPlacement.Inline;

    /// <summary>
    /// Whether to add blank lines between sections in structured mode.
    /// Only applies when Mode is Tree and TreeStyle is Structured.
    /// </summary>
    [JsonPropertyName("structured_blank_lines")]
    public bool StructuredBlankLines { get; set; } = false;

    /// <summary>
    /// Comma-separated required terms. ALL must match the study description.
    /// Empty = matches all studies (global pick list).
    /// </summary>
    [JsonPropertyName("criteria_required")]
    public string CriteriaRequired { get; set; } = "";

    /// <summary>
    /// Comma-separated optional terms. At least ONE must match (if specified).
    /// Empty = no additional OR requirement.
    /// </summary>
    [JsonPropertyName("criteria_any_of")]
    public string CriteriaAnyOf { get; set; } = "";

    /// <summary>
    /// Comma-separated exclusion terms. NONE must match (if specified).
    /// Empty = no exclusion filter.
    /// </summary>
    [JsonPropertyName("criteria_exclude")]
    public string CriteriaExclude { get; set; } = "";

    /// <summary>
    /// Root nodes of this pick list (Tree mode).
    /// Each node can have nested children for arbitrary depth.
    /// </summary>
    [JsonPropertyName("nodes")]
    public List<PickListNode> Nodes { get; set; } = new();

    /// <summary>
    /// Categories for Builder mode. Each category has options to select from.
    /// Selections are concatenated in order to build a sentence.
    /// </summary>
    [JsonPropertyName("categories")]
    public List<PickListCategory> Categories { get; set; } = new();

    /// <summary>
    /// Check if this pick list matches the given study description.
    /// Uses same logic as MacroConfig for consistency.
    /// Logic: (all Required match) AND (at least one AnyOf, if specified) AND (none of Exclude match)
    /// </summary>
    public bool MatchesStudy(string? studyDescription)
    {
        bool hasRequired = !string.IsNullOrWhiteSpace(CriteriaRequired);
        bool hasAnyOf = !string.IsNullOrWhiteSpace(CriteriaAnyOf);
        bool hasExclude = !string.IsNullOrWhiteSpace(CriteriaExclude);

        // Global pick list - matches all (but still check exclusions)
        if (!hasRequired && !hasAnyOf && !hasExclude)
            return true;

        if (string.IsNullOrWhiteSpace(studyDescription))
            return false;

        var description = studyDescription.ToUpperInvariant();

        // Check Exclude terms first - if any match, reject
        if (hasExclude)
        {
            var excludeTerms = CriteriaExclude.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var term in excludeTerms)
            {
                if (description.Contains(term.ToUpperInvariant()))
                    return false;
            }
        }

        // Check Required terms (AND logic) - all must match
        if (hasRequired)
        {
            var requiredTerms = CriteriaRequired.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var term in requiredTerms)
            {
                if (!description.Contains(term.ToUpperInvariant()))
                    return false;
            }
        }

        // Check AnyOf terms (OR logic) - at least one must match
        if (hasAnyOf)
        {
            var anyOfTerms = CriteriaAnyOf.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            bool anyMatch = false;
            foreach (var term in anyOfTerms)
            {
                if (description.Contains(term.ToUpperInvariant()))
                {
                    anyMatch = true;
                    break;
                }
            }
            if (!anyMatch)
                return false;
        }

        return true;
    }

    /// <summary>
    /// Returns a display string for the criteria (for list preview).
    /// </summary>
    public string GetCriteriaDisplayString()
    {
        if (string.IsNullOrWhiteSpace(CriteriaRequired) &&
            string.IsNullOrWhiteSpace(CriteriaAnyOf) &&
            string.IsNullOrWhiteSpace(CriteriaExclude))
            return "All studies";

        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(CriteriaRequired))
            parts.Add(CriteriaRequired.Trim());
        if (!string.IsNullOrWhiteSpace(CriteriaAnyOf))
            parts.Add($"({CriteriaAnyOf.Trim()})");
        if (!string.IsNullOrWhiteSpace(CriteriaExclude))
            parts.Add($"-{CriteriaExclude.Trim()}");

        return string.Join(" + ", parts);
    }
}

/// <summary>
/// A node in the pick list tree. Can be a branch (has children), leaf (inserts text),
/// or builder-ref (launches a builder pick list).
/// Supports arbitrary nesting depth for complex pick lists.
/// </summary>
public class PickListNode
{
    /// <summary>
    /// Display label for this node (e.g., "Heart", "Enlarged", "Mildly").
    /// </summary>
    [JsonPropertyName("label")]
    public string Label { get; set; } = "";

    /// <summary>
    /// Text to insert when this node is selected.
    /// For branch nodes, this can be empty (user must drill down).
    /// For leaf nodes, this is the text that gets pasted.
    /// For builder-ref nodes, this is optional prefix text inserted before builder output.
    /// </summary>
    [JsonPropertyName("text")]
    public string Text { get; set; } = "";

    /// <summary>
    /// Child nodes. If empty, this is a leaf node.
    /// User presses 1-9 to select children.
    /// </summary>
    [JsonPropertyName("children")]
    public List<PickListNode> Children { get; set; } = new();

    /// <summary>
    /// Optional reference to a builder-mode pick list ID.
    /// When set, selecting this node launches that builder instead of using Text directly.
    /// The builder's output is appended to any prefix Text.
    /// </summary>
    [JsonPropertyName("builder_list_id")]
    public string? BuilderListId { get; set; }

    /// <summary>
    /// Optional reference to a macro ID.
    /// When set, selecting this node resolves the macro's current Text at runtime.
    /// </summary>
    [JsonPropertyName("macro_id")]
    public string? MacroId { get; set; }

    /// <summary>
    /// Returns true if this node has children (is a branch).
    /// </summary>
    [JsonIgnore]
    public bool HasChildren => Children.Count > 0;

    /// <summary>
    /// Returns true if this node is a leaf (no children, just text, no builder or macro reference).
    /// </summary>
    [JsonIgnore]
    public bool IsLeaf => Children.Count == 0 && string.IsNullOrEmpty(BuilderListId) && string.IsNullOrEmpty(MacroId);

    /// <summary>
    /// Returns true if this node references a builder pick list.
    /// </summary>
    [JsonIgnore]
    public bool IsBuilderRef => !string.IsNullOrEmpty(BuilderListId);

    /// <summary>
    /// Returns true if this node references a macro.
    /// </summary>
    [JsonIgnore]
    public bool IsMacroRef => !string.IsNullOrEmpty(MacroId);

    /// <summary>
    /// Deep clone this node and all children.
    /// </summary>
    public PickListNode Clone()
    {
        return new PickListNode
        {
            Label = Label,
            Text = Text,
            BuilderListId = BuilderListId,
            MacroId = MacroId,
            Children = Children.Select(c => c.Clone()).ToList()
        };
    }

    /// <summary>
    /// Count total descendants (for display).
    /// </summary>
    public int CountDescendants()
    {
        return Children.Count + Children.Sum(c => c.CountDescendants());
    }
}

/// <summary>
/// Configuration for a server endpoint to monitor connectivity.
/// </summary>
public class ServerConfig
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("host")]
    public string Host { get; set; } = "";

    /// <summary>
    /// Port to test TCP connectivity. 0 = ping only (ICMP).
    /// </summary>
    [JsonPropertyName("port")]
    public int Port { get; set; } = 0;

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; } = true;
}

/// <summary>
/// Legacy item format - kept for backwards compatibility during migration.
/// </summary>
public class PickListItem
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("options")]
    public List<string> Options { get; set; } = new();

    /// <summary>
    /// Convert legacy item to new node format.
    /// </summary>
    public PickListNode ToNode()
    {
        return new PickListNode
        {
            Label = Name,
            Text = "",
            Children = Options.Select(opt => new PickListNode
            {
                Label = string.IsNullOrWhiteSpace(opt) ? "(empty)" : (opt.Length > 30 ? opt.Substring(0, 30) + "..." : opt),
                Text = opt,
                Children = new()
            }).ToList()
        };
    }
}
