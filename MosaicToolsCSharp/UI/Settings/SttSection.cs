// [CustomSTT] Settings section for custom speech-to-text
using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Settings section for Custom STT Mode configuration.
/// Supports Deepgram, Speechmatics, and AssemblyAI providers.
/// </summary>
public class SttSection : SettingsSection
{
    public override string SectionId => "stt";

    private readonly CheckBox _enabledCheck;
    private readonly ComboBox _providerCombo;
    private readonly Label _apiKeyLabel;
    private readonly TextBox _apiKeyBox;
    private readonly Button _getKeyButton;
    private readonly Button _helpButton;
    private readonly Label _pricingHint;

    // Speechmatics region selector
    private readonly Label _regionLabel;
    private readonly ComboBox _regionCombo;

    // Deepgram keyterms
    private readonly Label _deepgramKeytermsLabel;
    private readonly TextBox _deepgramKeytermsBox;
    private readonly Label _deepgramKeytermsCount;
    private readonly Button _deepgramKeytermsSortButton;

    // Keyterm auto-learning
    private readonly CheckBox _keytermLearningCheck;
    private readonly Label _keytermLearningStats;
    private readonly Button _keytermLearningViewButton;
    private readonly Label _keytermLearningHint;

    private readonly ComboBox _audioDeviceCombo;
    private readonly CheckBox _autoPunctuateCheck;
    private readonly CheckBox _startBeepCheck;
    private readonly CheckBox _stopBeepCheck;
    private readonly TrackBar _startBeepVolume;
    private readonly TrackBar _stopBeepVolume;
    private readonly Label _startBeepVolLabel;
    private readonly Label _stopBeepVolLabel;
    private readonly CheckBox _showIndicatorCheck;
    private readonly CheckBox _autoStartCheck;

    // Provider dropdown indices
    private const int ProviderDeepgramMedical = 0;
    private const int ProviderSpeechmatics = 1;
    private const int ProviderAssemblyAI = 2;

    // Per-provider credential storage (swapped on provider change)
    private string _deepgramKey = "";
    private string _speechmaticsKey = "";
    private string _assemblyAIKey = "";
    private string _speechmaticsRegion = "us";
    private int _previousProviderIndex;
    private bool _loading; // Suppress OnProviderChanged during LoadSettings

    public SttSection(ToolTip toolTip) : base("Speech-to-Text", toolTip)
    {
        _searchTerms.AddRange(new[] { "stt", "speech", "transcription", "deepgram", "speechmatics", "assemblyai", "dictation", "custom stt", "beep", "punctuation" });

        _enabledCheck = AddCheckBox("Enable Custom STT Mode", LeftMargin, _nextY,
            "Use cloud STT instead of Mosaic's built-in speech recognition. Cancel Mosaic's WebHID prompt when enabled.");
        _enabledCheck.CheckedChanged += (_, _) => UpdateControlStates();
        _nextY += RowHeight;

        AddHintLabel("Cancel Mosaic's WebHID prompt when this is enabled", LeftMargin + 25);

        // Provider
        AddSectionDivider("Provider");

        AddLabel("Provider:", LeftMargin + 25, _nextY + 3);
        _providerCombo = AddComboBox(LeftMargin + 110, _nextY, 200,
            new[] { "Deepgram Nova-3 Medical", "Speechmatics Medical", "AssemblyAI" });
        _providerCombo.SelectedIndexChanged += (_, _) => OnProviderChanged();
        _nextY += RowHeight;

        // API Key
        _apiKeyLabel = AddLabel("API Key:", LeftMargin + 25, _nextY + 3);
        _apiKeyBox = AddTextBox(LeftMargin + 110, _nextY, 200);
        _apiKeyBox.UseSystemPasswordChar = true;

        _getKeyButton = AddButton("Get API Key", LeftMargin + 320, _nextY - 1, 90, 24, OnGetKeyClick,
            "Open signup page to create an account and get credentials.");

        _helpButton = AddButton("?", LeftMargin + 415, _nextY - 1, 24, 24, OnApiKeyHelpClick,
            "How to set up this provider");
        _helpButton.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        _nextY += SubRowHeight;

        // Dynamic pricing/info hint
        _pricingHint = new Label
        {
            Text = "Free tier: $200 credit. Nova-3 Medical: $0.0077/min streaming",
            Location = new Point(LeftMargin + 25, _nextY),
            AutoSize = true,
            ForeColor = Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", 7.5f)
        };
        Controls.Add(_pricingHint);
        _nextY += SubRowHeight;

        // Region selector (Speechmatics only)
        _regionLabel = AddLabel("Region:", LeftMargin + 25, _nextY + 3);
        _regionCombo = AddComboBox(LeftMargin + 120, _nextY, 100,
            new[] { "US", "EU" });
        _nextY += RowHeight;

        // Audio Device
        AddSectionDivider("Audio");

        AddLabel("Audio Device:", LeftMargin + 25, _nextY + 3);
        _audioDeviceCombo = AddComboBox(LeftMargin + 120, _nextY, 250, new[] { "(Auto-detect PowerMic)" });
        _nextY += RowHeight;

        // Populate audio devices
        try
        {
            var devices = SttService.GetAudioDevices();
            foreach (var device in devices)
            {
                _audioDeviceCombo.Items.Add(device);
            }
        }
        catch { /* NAudio not available in design mode */ }

        // Punctuation
        AddSectionDivider("Punctuation");

        _autoPunctuateCheck = AddCheckBox("Auto-punctuate", LeftMargin + 25, _nextY,
            "Automatically insert punctuation. When off, say \"period\", \"comma\", etc. to punctuate.");
        _nextY += SubRowHeight;

        AddHintLabel("Off = dictate punctuation (say \"period\", \"comma\"). On = auto-inserted.", LeftMargin + 25);

        // Deepgram keyterms (shown only for Deepgram provider)
        _deepgramKeytermsLabel = AddLabel("Keyterms:", LeftMargin + 25, _nextY + 3);
        _deepgramKeytermsBox = new TextBox
        {
            Location = new Point(LeftMargin + 110, _nextY),
            Width = 300,
            Height = 60,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Segoe UI", 9f),
            PlaceholderText = "One per line. Boosts recognition of specific words (e.g. periportal, echogenicity)"
        };
        _deepgramKeytermsBox.TextChanged += (_, _) => UpdateKeytermCount();
        Controls.Add(_deepgramKeytermsBox);

        _deepgramKeytermsSortButton = AddButton("Sort A\u2013Z", LeftMargin + 25, _nextY + 18, 55, 22, OnSortKeytermsClick,
            "Sort keyterms alphabetically");
        _deepgramKeytermsSortButton.Font = new Font("Segoe UI", 7.5f);

        _deepgramKeytermsCount = new Label
        {
            Location = new Point(LeftMargin + 110, _nextY + 62),
            AutoSize = true,
            ForeColor = Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", 7.5f),
            Text = "0 / 100 terms"
        };
        Controls.Add(_deepgramKeytermsCount);
        _nextY += 80;

        // Keyterm auto-learning (Deepgram only)
        _keytermLearningCheck = AddCheckBox("Auto-learn keyterms from reports", LeftMargin + 25, _nextY,
            "Tracks low-confidence words during dictation and verifies them against the final signed report.");
        _nextY += SubRowHeight;

        _keytermLearningHint = new Label
        {
            Text = "Compares dictation to final report. Fills remaining keyterm slots automatically.",
            Location = new Point(LeftMargin + 45, _nextY),
            AutoSize = true,
            ForeColor = Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", 7.5f)
        };
        Controls.Add(_keytermLearningHint);
        _nextY += SubRowHeight;

        _keytermLearningStats = new Label
        {
            Text = "No auto-learned terms yet",
            Location = new Point(LeftMargin + 45, _nextY + 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(170, 170, 170),
            Font = new Font("Segoe UI", 8f)
        };
        Controls.Add(_keytermLearningStats);

        _keytermLearningViewButton = AddButton("View", LeftMargin + 300, _nextY - 1, 50, 24, OnViewKeytermLearningClick,
            "View auto-learned keyterms with details");
        _keytermLearningViewButton.Font = new Font("Segoe UI", 8f);
        _nextY += RowHeight;

        // Beep Settings
        AddSectionDivider("Beeps");

        _startBeepCheck = AddCheckBox("Start beep", LeftMargin + 25, _nextY,
            "Play beep when recording starts.");
        _startBeepVolLabel = AddLabel("Vol:", LeftMargin + 160, _nextY + 3);
        _startBeepVolume = AddTrackBar(LeftMargin + 190, _nextY, 120, 0, 100, 8,
            "Start beep volume");
        _nextY += RowHeight;

        _stopBeepCheck = AddCheckBox("Stop beep", LeftMargin + 25, _nextY,
            "Play beep when recording stops.");
        _stopBeepVolLabel = AddLabel("Vol:", LeftMargin + 160, _nextY + 3);
        _stopBeepVolume = AddTrackBar(LeftMargin + 190, _nextY, 120, 0, 100, 8,
            "Stop beep volume");
        _nextY += RowHeight;

        // Display options
        AddSectionDivider("Display");

        _showIndicatorCheck = AddCheckBox("Show dictation indicator", LeftMargin + 25, _nextY,
            "Show a floating indicator with live transcription text while dictating.");
        _nextY += RowHeight;

        _autoStartCheck = AddCheckBox("Auto-start on case open", LeftMargin + 25, _nextY,
            "Automatically start recording when a new study opens. Restarts after Process Report, stops on Sign Report.");
        _nextY += RowHeight;

        UpdateHeight();
    }

    /// <summary>
    /// Save current UI key fields to the previous provider's local storage,
    /// then load the new provider's saved credentials into the UI.
    /// </summary>
    private void OnProviderChanged()
    {
        if (_loading) return; // Skip during LoadSettings — keys loaded directly

        int newIdx = _providerCombo.SelectedIndex;

        // Save current UI values to previous provider's storage
        SaveKeysFromUI(_previousProviderIndex);

        // Load new provider's saved values into UI
        LoadKeysToUI(newIdx);

        _previousProviderIndex = newIdx;

        // Update region selector visibility (Speechmatics only)
        bool isSpeechmatics = newIdx == ProviderSpeechmatics;
        _regionLabel.Visible = isSpeechmatics;
        _regionCombo.Visible = isSpeechmatics;

        // Update keyterms visibility (Deepgram only)
        bool isDeepgram = newIdx == ProviderDeepgramMedical;
        _deepgramKeytermsLabel.Visible = isDeepgram;
        _deepgramKeytermsBox.Visible = isDeepgram;
        _deepgramKeytermsCount.Visible = isDeepgram;
        _deepgramKeytermsSortButton.Visible = isDeepgram;
        _keytermLearningCheck.Visible = isDeepgram;
        _keytermLearningStats.Visible = isDeepgram;
        _keytermLearningViewButton.Visible = isDeepgram;
        _keytermLearningHint.Visible = isDeepgram;

        // Update pricing hint
        _pricingHint.Text = newIdx switch
        {
            ProviderDeepgramMedical => "Free tier: $200 credit. Nova-3 Medical: $0.0077/min streaming",
            ProviderSpeechmatics => "Free tier: 480 min/month. Medical model: $0.004/min streaming",
            ProviderAssemblyAI => "Free tier: 330 hours. $0.0025/min streaming",
            _ => ""
        };
    }

    private void SaveKeysFromUI(int providerIndex)
    {
        switch (providerIndex)
        {
            case ProviderDeepgramMedical:
                _deepgramKey = _apiKeyBox.Text;
                break;
            case ProviderSpeechmatics:
                _speechmaticsKey = _apiKeyBox.Text;
                _speechmaticsRegion = _regionCombo.SelectedIndex == 1 ? "eu" : "us";
                break;
            case ProviderAssemblyAI:
                _assemblyAIKey = _apiKeyBox.Text;
                break;
        }
    }

    private void LoadKeysToUI(int providerIndex)
    {
        switch (providerIndex)
        {
            case ProviderDeepgramMedical:
                _apiKeyBox.Text = _deepgramKey;
                break;
            case ProviderSpeechmatics:
                _apiKeyBox.Text = _speechmaticsKey;
                _regionCombo.SelectedIndex = _speechmaticsRegion == "eu" ? 1 : 0;
                break;
            case ProviderAssemblyAI:
                _apiKeyBox.Text = _assemblyAIKey;
                break;
        }
    }

    private void UpdateControlStates()
    {
        bool enabled = _enabledCheck.Checked;
        _providerCombo.Enabled = enabled;
        _apiKeyBox.Enabled = enabled;
        _getKeyButton.Enabled = enabled;
        _helpButton.Enabled = enabled;
        _regionCombo.Enabled = enabled;
        _audioDeviceCombo.Enabled = enabled;
        _autoPunctuateCheck.Enabled = enabled;
        _deepgramKeytermsBox.Enabled = enabled;
        _deepgramKeytermsSortButton.Enabled = enabled;
        _keytermLearningCheck.Enabled = enabled;
        _keytermLearningViewButton.Enabled = enabled;
        _startBeepCheck.Enabled = enabled;
        _stopBeepCheck.Enabled = enabled;
        _startBeepVolume.Enabled = enabled;
        _stopBeepVolume.Enabled = enabled;
        _showIndicatorCheck.Enabled = enabled;
    }

    private void RefreshKeytermLearningStats()
    {
        try
        {
            var svc = new KeytermLearningService();
            svc.Load();
            var count = svc.EntryCount;
            _keytermLearningStats.Text = count == 0
                ? "No auto-learned terms yet"
                : $"{count} auto-learned term{(count == 1 ? "" : "s")}";
        }
        catch
        {
            _keytermLearningStats.Text = "No auto-learned terms yet";
        }
    }

    private void UpdateKeytermCount()
    {
        var count = 0;
        if (!string.IsNullOrWhiteSpace(_deepgramKeytermsBox.Text))
        {
            var terms = _deepgramKeytermsBox.Text.Split(new[] { '\n', '\r', ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var t in terms)
                if (t.Trim().Length > 0) count++;
        }
        _deepgramKeytermsCount.Text = $"{count} / 100 terms";
        _deepgramKeytermsCount.ForeColor = count > 100 ? Color.IndianRed : Color.FromArgb(150, 150, 150);
    }

    private void OnSortKeytermsClick(object? sender, EventArgs e)
    {
        var terms = _deepgramKeytermsBox.Text
            .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(t => t.Trim())
            .Where(t => t.Length > 0)
            .OrderBy(t => t, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        _deepgramKeytermsBox.Text = string.Join("\r\n", terms);
    }

    private void OnViewKeytermLearningClick(object? sender, EventArgs e)
    {
        var svc = new KeytermLearningService();
        svc.Load();
        var entries = svc.GetAllEntries();

        if (entries.Count == 0)
        {
            MessageBox.Show("No auto-learned keyterms yet.\n\nDictate reports with Custom STT enabled, " +
                "then sign/advance studies. Terms will appear after the first verification cycle.",
                "Auto-Learned Keyterms", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Determine how many auto-learned terms are actually sent to Deepgram
        var manualTerms = _deepgramKeytermsBox.Text
            .Split(new[] { '\n', '\r', ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(t => t.Trim()).Where(t => t.Length > 0).ToList();
        var manualSet = new HashSet<string>(manualTerms, StringComparer.OrdinalIgnoreCase);
        var autoSlots = Math.Max(0, 100 - manualTerms.Count);

        // Build a detailed view with columns
        var lines = new System.Text.StringBuilder();
        lines.AppendLine($"{"Term",-24} {"Score",7} {"Conf",6} {"Seen",5}  {"Source",-10}");
        lines.AppendLine(new string('\u2500', 59));

        int sentCount = 0;
        bool separatorPlaced = false;
        foreach (var (word, score, avgConf, source, occurrences) in entries)
        {
            bool isDuplicate = manualSet.Contains(word);
            bool isSent = !isDuplicate && sentCount < autoSlots;

            if (isSent)
                sentCount++;
            else if (!separatorPlaced && sentCount > 0)
            {
                // Place separator between sent and not-sent terms
                separatorPlaced = true;
                lines.AppendLine($"{"--- not sent " + new string('-', 46)}");
            }

            var confStr = $"{avgConf:P0}";
            var sourceLabel = source == "corrected" ? "corrected" : "survived";
            lines.AppendLine($"{Truncate(word, 23),-24} {score,7:F2} {confStr,6} {occurrences,5}  {sourceLabel,-10}");
        }

        lines.AppendLine();
        lines.AppendLine($"Total: {entries.Count} auto-learned terms ({sentCount} sent to Deepgram, {manualTerms.Count} manual)");
        lines.AppendLine();
        lines.AppendLine("Conf  = average Deepgram confidence when the word was heard");
        lines.AppendLine("Score = log(1 + occurrences) \u00d7 (1 \u2212 conf)\u00b2 \u00d7 weight");
        lines.AppendLine("  survived=1.0x, corrected=2.0x (corrections are more valuable)");
        lines.AppendLine("  Log dampens volume; squared gap prioritizes truly problematic words.");
        lines.AppendLine();
        lines.AppendLine("Top-scoring terms are sent to Deepgram as keyterms.");

        var form = new Form
        {
            Text = "Auto-Learned Keyterms",
            Size = new Size(520, 420),
            StartPosition = FormStartPosition.CenterParent,
            MinimizeBox = false,
            MaximizeBox = false,
            ShowIcon = false
        };

        var textBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Consolas", 9f),
            Dock = DockStyle.Fill,
            Text = lines.ToString(),
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.FromArgb(220, 220, 220)
        };
        form.Controls.Add(textBox);
        form.Shown += (_, _) => { textBox.SelectionStart = 0; textBox.SelectionLength = 0; };
        form.ShowDialog(FindForm());
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s[..(max - 1)] + "\u2026";

    private void OnGetKeyClick(object? sender, EventArgs e)
    {
        var url = _providerCombo.SelectedIndex switch
        {
            ProviderSpeechmatics => "https://portal.speechmatics.com/signup",
            ProviderAssemblyAI => "https://www.assemblyai.com/dashboard/signup",
            _ => "https://console.deepgram.com/signup"
        };

        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch { }
    }

    private void OnApiKeyHelpClick(object? sender, EventArgs e)
    {
        var (text, title) = _providerCombo.SelectedIndex switch
        {
            ProviderSpeechmatics => (
                "How to get a Speechmatics API key:\n\n" +
                "1. Click \"Get API Key\" to open the Speechmatics Portal\n" +
                "2. Create a free account\n" +
                "3. After signing in, go to Manage > API Keys\n" +
                "4. Click \"Create API Key\"\n" +
                "5. Copy the generated key\n" +
                "6. Paste the key into the API Key field above\n" +
                "7. Select your preferred region (US or EU)\n\n" +
                "Free tier includes 480 minutes/month.\n" +
                "Medical model streaming: $0.004/min ($0.24/hour).\n" +
                "Dedicated medical model trained on 16 billion words of clinical data.\n\n" +
                "Keep your API key private \u2014 do not share it.",
                "Speechmatics Setup"),
            ProviderAssemblyAI => (
                "How to get an AssemblyAI API key:\n\n" +
                "1. Click \"Get API Key\" to open the AssemblyAI signup page\n" +
                "2. Create a free account (email or Google/GitHub login)\n" +
                "3. After signing in, you'll see your Dashboard\n" +
                "4. Your API key is displayed on the main dashboard page\n" +
                "5. Click the copy icon to copy it\n" +
                "6. Paste the key into the API Key field above\n\n" +
                "Free tier includes 330 hours of streaming transcription.\n" +
                "Streaming costs $0.15/hour ($0.0025/min).\n\n" +
                "Keep your API key private \u2014 do not share it.",
                "AssemblyAI Setup"),
            _ => (
                "How to get a Deepgram API key:\n\n" +
                "1. Click \"Get API Key\" to open the Deepgram signup page\n" +
                "2. Create a free account (email or Google/GitHub login)\n" +
                "3. After signing in, you'll land on the Deepgram Console\n" +
                "4. Click \"API Keys\" in the left sidebar (or go to Settings > API Keys)\n" +
                "5. Click \"Create a New API Key\"\n" +
                "6. Give it a name (e.g. \"MosaicTools\"), leave permissions as default\n" +
                "7. Click \"Create Key\" and copy the key that appears\n" +
                "8. Paste the key into the API Key field above\n\n" +
                "Free tier includes $200 in credit \u2014 enough for ~430 hours of\n" +
                "Nova-3 Medical dictation at $0.0077/min.\n\n" +
                "The key starts with \"dg_\" or a similar prefix.\n" +
                "Keep it private \u2014 do not share it.",
                "Deepgram API Key Setup")
        };

        MessageBox.Show(text, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public override void LoadSettings(Configuration config)
    {
        _enabledCheck.Checked = config.CustomSttEnabled;

        // Load per-provider credentials into local storage
        _deepgramKey = config.SttApiKey;
        _speechmaticsKey = config.SttSpeechmaticsApiKey;
        _speechmaticsRegion = config.SttSpeechmaticsRegion;
        _assemblyAIKey = config.SttAssemblyAIApiKey;

        // Map provider to dropdown index (corti falls back to deepgram)
        var idx = config.SttProvider switch
        {
            "deepgram" => ProviderDeepgramMedical,
            "speechmatics" => ProviderSpeechmatics,
            "assemblyai" => ProviderAssemblyAI,
            _ => ProviderDeepgramMedical
        };

        _previousProviderIndex = idx;
        _loading = true;
        _providerCombo.SelectedIndex = idx;
        _loading = false;

        // Manually load keys and update visibility (OnProviderChanged was suppressed)
        LoadKeysToUI(idx);
        bool isSpeechmatics = idx == ProviderSpeechmatics;
        _regionLabel.Visible = isSpeechmatics;
        _regionCombo.Visible = isSpeechmatics;
        bool isDeepgram = idx == ProviderDeepgramMedical;
        _deepgramKeytermsLabel.Visible = isDeepgram;
        _deepgramKeytermsBox.Visible = isDeepgram;
        _deepgramKeytermsCount.Visible = isDeepgram;
        _deepgramKeytermsSortButton.Visible = isDeepgram;
        _keytermLearningCheck.Visible = isDeepgram;
        _keytermLearningStats.Visible = isDeepgram;
        _keytermLearningViewButton.Visible = isDeepgram;
        _keytermLearningHint.Visible = isDeepgram;
        _pricingHint.Text = idx switch
        {
            ProviderDeepgramMedical => "Free tier: $200 credit. Nova-3 Medical: $0.0077/min streaming",
            ProviderSpeechmatics => "Free tier: 480 min/month. Medical model: $0.004/min streaming",
            ProviderAssemblyAI => "Free tier: 330 hours. $0.0025/min streaming",
            _ => ""
        };

        // Set region combo for Speechmatics
        _regionCombo.SelectedIndex = _speechmaticsRegion == "eu" ? 1 : 0;

        // Select audio device
        if (string.IsNullOrEmpty(config.SttAudioDeviceName))
        {
            _audioDeviceCombo.SelectedIndex = 0; // Auto-detect
        }
        else
        {
            int devIdx = _audioDeviceCombo.FindStringExact(config.SttAudioDeviceName);
            _audioDeviceCombo.SelectedIndex = devIdx >= 0 ? devIdx : 0;
        }

        _autoPunctuateCheck.Checked = config.SttAutoPunctuate;
        _deepgramKeytermsBox.Text = config.SttDeepgramKeyterms.Replace(",", "\r\n");
        _startBeepCheck.Checked = config.SttStartBeepEnabled;
        _stopBeepCheck.Checked = config.SttStopBeepEnabled;
        _startBeepVolume.Value = Math.Clamp((int)(config.SttStartBeepVolume * 100), 0, 100);
        _stopBeepVolume.Value = Math.Clamp((int)(config.SttStopBeepVolume * 100), 0, 100);
        _showIndicatorCheck.Checked = config.SttShowIndicator;
        _autoStartCheck.Checked = config.SttAutoStartOnCase;

        // Keyterm learning
        _keytermLearningCheck.Checked = config.SttKeytermLearningEnabled;
        RefreshKeytermLearningStats();

        UpdateControlStates();
    }

    public override void SaveSettings(Configuration config)
    {
        config.CustomSttEnabled = _enabledCheck.Checked;

        // Save current UI values to the active provider's local storage
        SaveKeysFromUI(_providerCombo.SelectedIndex);

        // Write all per-provider credentials to config
        config.SttApiKey = _deepgramKey.Trim();
        config.SttSpeechmaticsApiKey = _speechmaticsKey.Trim();
        config.SttSpeechmaticsRegion = _speechmaticsRegion;
        config.SttAssemblyAIApiKey = _assemblyAIKey.Trim();

        // Map dropdown index to provider+model
        (config.SttProvider, config.SttModel) = _providerCombo.SelectedIndex switch
        {
            ProviderDeepgramMedical => ("deepgram", "nova-3-medical"),
            ProviderSpeechmatics => ("speechmatics", ""),
            ProviderAssemblyAI => ("assemblyai", ""),
            _ => ("deepgram", "nova-3-medical")
        };

        // Save audio device (empty = auto-detect)
        if (_audioDeviceCombo.SelectedIndex <= 0)
        {
            config.SttAudioDeviceName = "";
        }
        else
        {
            config.SttAudioDeviceName = _audioDeviceCombo.SelectedItem?.ToString() ?? "";
        }

        config.SttAutoPunctuate = _autoPunctuateCheck.Checked;
        config.SttDeepgramKeyterms = _deepgramKeytermsBox.Text.Trim();
        config.SttStartBeepEnabled = _startBeepCheck.Checked;
        config.SttStopBeepEnabled = _stopBeepCheck.Checked;
        config.SttStartBeepVolume = _startBeepVolume.Value / 100.0;
        config.SttStopBeepVolume = _stopBeepVolume.Value / 100.0;
        config.SttShowIndicator = _showIndicatorCheck.Checked;
        config.SttAutoStartOnCase = _autoStartCheck.Checked;
        config.SttKeytermLearningEnabled = _keytermLearningCheck.Checked;
    }
}
