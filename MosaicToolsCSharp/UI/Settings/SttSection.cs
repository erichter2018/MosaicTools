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

    // Keyterm auto-learning (per-provider)
    private readonly CheckBox _keytermLearningCheck;
    private readonly Label _keytermLearningHint;
    private readonly Label _dgLearningStats;
    private readonly Button _dgLearningViewButton;
    private readonly Label _snxLearningStats;
    private readonly Button _snxLearningViewButton;
    private readonly Label _smLearningStats;
    private readonly Button _smLearningViewButton;
    private readonly Label _aaiLearningStats;
    private readonly Button _aaiLearningViewButton;

    // Ensemble mode
    private readonly CheckBox _ensembleCheck;
    private readonly Label _ensembleHint;
    private readonly Label _ensembleKeyStatus;
    private readonly Label _ensembleS1Label;
    private readonly ComboBox _ensembleS1Combo;
    private readonly Label _ensembleS2Label;
    private readonly ComboBox _ensembleS2Combo;
    private readonly Label _ensembleWaitLabel;
    private readonly NumericUpDown _ensembleWaitUpDown;
    private readonly Label _ensembleThreshLabel;
    private readonly NumericUpDown _ensembleThreshUpDown;
    private readonly CheckBox _ensembleShowMetricsCheck;

    // Text processing
    private readonly CheckBox _newlineAfterSentenceCheck;
    private readonly CheckBox _expandContractionsCheck;
    private readonly CheckBox _radiologyCleanupCheck;
    private readonly DataGridView _replacementsGrid;
    private readonly Button _addReplacementButton;
    private readonly Button _removeReplacementButton;

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
    private const int ProviderSoniox = 3;

    // Per-provider credential storage (swapped on provider change)
    private string _deepgramKey = "";
    private string _speechmaticsKey = "";
    private string _assemblyAIKey = "";
    private string _sonioxKey = "";
    private string _speechmaticsRegion = "us";
    private int _previousProviderIndex;
    private bool _loading; // Suppress OnProviderChanged during LoadSettings

    public SttSection(ToolTip toolTip) : base("Speech-to-Text", toolTip)
    {
        _searchTerms.AddRange(new[] { "stt", "speech", "transcription", "deepgram", "speechmatics", "assemblyai", "soniox", "dictation", "custom stt", "beep", "punctuation", "contraction", "replacement", "correction", "radiology cleanup", "ensemble" });

        _enabledCheck = AddCheckBox("Enable Custom STT Mode", LeftMargin, _nextY,
            "Use cloud STT instead of Mosaic's built-in speech recognition. Cancel Mosaic's WebHID prompt when enabled.");
        _enabledCheck.CheckedChanged += (_, _) => UpdateControlStates();
        _nextY += RowHeight;

        AddHintLabel("Cancel Mosaic's WebHID prompt when this is enabled", LeftMargin + 25);

        // Provider
        AddSectionDivider("Provider");

        AddLabel("Provider:", LeftMargin + 25, _nextY + 3);
        _providerCombo = AddComboBox(LeftMargin + 110, _nextY, 200,
            new[] { "Deepgram Nova-3 Medical", "Speechmatics Medical", "AssemblyAI", "Soniox" });
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

        _newlineAfterSentenceCheck = AddCheckBox("New line after every sentence", LeftMargin + 25, _nextY,
            "Insert a line break after each sentence-ending period.");
        _nextY += SubRowHeight;

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

        // Keyterm auto-learning (per-provider rows)
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

        // Per-provider stats rows
        var statsFont = new Font("Segoe UI", 8f);
        var statsColor = Color.FromArgb(170, 170, 170);

        AddLabel("Deepgram:", LeftMargin + 45, _nextY + 2).Font = statsFont;
        _dgLearningStats = new Label { Text = "0 terms", Location = new Point(LeftMargin + 130, _nextY + 2), AutoSize = true, ForeColor = statsColor, Font = statsFont };
        Controls.Add(_dgLearningStats);
        _dgLearningViewButton = AddButton("View", LeftMargin + 210, _nextY - 1, 45, 22, (s, e) => OnViewKeytermLearningClick("deepgram"), "View Deepgram auto-learned keyterms");
        _dgLearningViewButton.Font = new Font("Segoe UI", 7.5f);
        _nextY += SubRowHeight;

        AddLabel("Soniox:", LeftMargin + 45, _nextY + 2).Font = statsFont;
        _snxLearningStats = new Label { Text = "0 terms", Location = new Point(LeftMargin + 130, _nextY + 2), AutoSize = true, ForeColor = statsColor, Font = statsFont };
        Controls.Add(_snxLearningStats);
        _snxLearningViewButton = AddButton("View", LeftMargin + 210, _nextY - 1, 45, 22, (s, e) => OnViewKeytermLearningClick("soniox"), "View Soniox auto-learned keyterms");
        _snxLearningViewButton.Font = new Font("Segoe UI", 7.5f);
        _nextY += SubRowHeight;

        AddLabel("Speechmatics:", LeftMargin + 45, _nextY + 2).Font = statsFont;
        _smLearningStats = new Label { Text = "0 terms", Location = new Point(LeftMargin + 130, _nextY + 2), AutoSize = true, ForeColor = statsColor, Font = statsFont };
        Controls.Add(_smLearningStats);
        _smLearningViewButton = AddButton("View", LeftMargin + 210, _nextY - 1, 45, 22, (s, e) => OnViewKeytermLearningClick("speechmatics"), "View Speechmatics auto-learned keyterms");
        _smLearningViewButton.Font = new Font("Segoe UI", 7.5f);
        _nextY += SubRowHeight;

        AddLabel("AssemblyAI:", LeftMargin + 45, _nextY + 2).Font = statsFont;
        _aaiLearningStats = new Label { Text = "0 terms", Location = new Point(LeftMargin + 130, _nextY + 2), AutoSize = true, ForeColor = statsColor, Font = statsFont };
        Controls.Add(_aaiLearningStats);
        _aaiLearningViewButton = AddButton("View", LeftMargin + 210, _nextY - 1, 45, 22, (s, e) => OnViewKeytermLearningClick("assemblyai"), "View AssemblyAI auto-learned keyterms");
        _aaiLearningViewButton.Font = new Font("Segoe UI", 7.5f);
        _nextY += RowHeight;

        // Ensemble Mode
        AddSectionDivider("Ensemble Mode");

        _ensembleCheck = AddCheckBox("Enable Ensemble Mode", LeftMargin + 25, _nextY,
            "Run all 3 providers simultaneously. Deepgram drives, secondaries correct low-confidence words.");
        _ensembleCheck.CheckedChanged += (_, _) => UpdateEnsembleStates();
        _nextY += SubRowHeight;

        _ensembleHint = new Label
        {
            Text = "Requires API keys for all 3 providers. ~$0.014/min total.",
            Location = new Point(LeftMargin + 45, _nextY),
            AutoSize = true,
            ForeColor = Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", 7.5f)
        };
        Controls.Add(_ensembleHint);
        _nextY += SubRowHeight;

        _ensembleKeyStatus = new Label
        {
            Text = "",
            Location = new Point(LeftMargin + 45, _nextY),
            AutoSize = true,
            Font = new Font("Segoe UI", 7.5f)
        };
        Controls.Add(_ensembleKeyStatus);
        _nextY += SubRowHeight;

        var primaryLabel = new Label
        {
            Text = "Primary:",
            Location = new Point(LeftMargin + 45, _nextY + 3),
            AutoSize = true,
            ForeColor = Color.FromArgb(200, 200, 200),
            Font = new Font("Segoe UI", 8.25f)
        };
        Controls.Add(primaryLabel);
        var primaryValue = new Label
        {
            Text = "Deepgram (driver)",
            Location = new Point(LeftMargin + 140, _nextY + 3),
            AutoSize = true,
            ForeColor = Color.FromArgb(76, 175, 80),
            Font = new Font("Segoe UI", 8.25f, FontStyle.Bold)
        };
        Controls.Add(primaryValue);
        _nextY += SubRowHeight;

        var secondaryProviders = new[] { "Soniox", "Speechmatics", "AssemblyAI", "None" };
        _ensembleS1Label = AddLabel("Secondary 1:", LeftMargin + 45, _nextY + 3);
        _ensembleS1Combo = AddComboBox(LeftMargin + 140, _nextY, 130, secondaryProviders);
        _ensembleS1Combo.SelectedIndexChanged += (_, _) => UpdateEnsembleStates();
        _nextY += SubRowHeight;

        _ensembleS2Label = AddLabel("Secondary 2:", LeftMargin + 45, _nextY + 3);
        _ensembleS2Combo = AddComboBox(LeftMargin + 140, _nextY, 130, secondaryProviders);
        _ensembleS2Combo.SelectedIndexChanged += (_, _) => UpdateEnsembleStates();
        _nextY += SubRowHeight;

        _ensembleWaitLabel = AddLabel("Merge wait:", LeftMargin + 45, _nextY + 3);
        _ensembleWaitUpDown = AddNumericUpDown(LeftMargin + 130, _nextY, 60, 100, 2000, 500);
        AddLabel("ms", LeftMargin + 195, _nextY + 3);

        _ensembleThreshLabel = AddLabel("Correct below:", LeftMargin + 240, _nextY + 3);
        _ensembleThreshUpDown = AddNumericUpDown(LeftMargin + 340, _nextY, 50, 50, 99, 80);
        AddLabel("%", LeftMargin + 395, _nextY + 3);
        _nextY += SubRowHeight;

        _ensembleShowMetricsCheck = AddCheckBox("Show live metrics popup", LeftMargin + 45, _nextY,
            "Show a draggable overlay with confidence, corrections, and provider stats during dictation.");
        _nextY += RowHeight;

        // Text Processing
        AddSectionDivider("Text Processing");

        _expandContractionsCheck = AddCheckBox("Expand contractions", LeftMargin + 25, _nextY,
            "Convert contractions to full words (e.g. \"there's\" → \"there is\").");
        _nextY += SubRowHeight;

        _radiologyCleanupCheck = AddCheckBox("Radiology cleanup", LeftMargin + 25, _nextY,
            "Convert spoken spine levels, units, dimensions, and dates to standard notation.");
        _nextY += RowHeight;

        // Custom replacements label
        AddLabel("Custom word replacements:", LeftMargin + 25, _nextY + 2);
        _nextY += SubRowHeight;

        // DataGridView for custom replacements
        _replacementsGrid = new DataGridView
        {
            Location = new Point(LeftMargin + 25, _nextY),
            Width = 390,
            Height = 120,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            ScrollBars = ScrollBars.Vertical,
            BorderStyle = BorderStyle.FixedSingle,
            BackgroundColor = Color.FromArgb(30, 30, 30),
            GridColor = Color.FromArgb(60, 60, 60),
            Font = new Font("Segoe UI", 9f),
            DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.FromArgb(220, 220, 220),
                SelectionBackColor = Color.FromArgb(60, 80, 110),
                SelectionForeColor = Color.FromArgb(220, 220, 220),
            },
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(45, 45, 45),
                ForeColor = Color.FromArgb(200, 200, 200),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleLeft,
            },
            EnableHeadersVisualStyles = false,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight = 26,
        };

        var enabledCol = new DataGridViewCheckBoxColumn
        {
            Name = "Enabled",
            HeaderText = "",
            Width = 30,
            FlatStyle = FlatStyle.Flat,
        };
        var findCol = new DataGridViewTextBoxColumn
        {
            Name = "Find",
            HeaderText = "Find",
            Width = 160,
        };
        var replaceCol = new DataGridViewTextBoxColumn
        {
            Name = "Replace",
            HeaderText = "Replace",
            Width = 160,
        };
        _replacementsGrid.Columns.AddRange(enabledCol, findCol, replaceCol);
        Controls.Add(_replacementsGrid);
        _nextY += 124;

        // +/- buttons
        _addReplacementButton = AddButton("+", LeftMargin + 25, _nextY, 30, 24, OnAddReplacementClick,
            "Add a new replacement entry");
        _addReplacementButton.Font = new Font("Segoe UI", 10f, FontStyle.Bold);

        _removeReplacementButton = AddButton("\u2212", LeftMargin + 58, _nextY, 30, 24, OnRemoveReplacementClick,
            "Remove the selected replacement entry");
        _removeReplacementButton.Font = new Font("Segoe UI", 10f, FontStyle.Bold);

        var replacementHint = new Label
        {
            Text = "Word-for-word replacements. Case is preserved automatically.",
            Location = new Point(LeftMargin + 95, _nextY + 5),
            AutoSize = true,
            ForeColor = Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", 7.5f)
        };
        Controls.Add(replacementHint);
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

        // Keyterms and auto-learning are now supported by all providers
        _deepgramKeytermsLabel.Visible = true;
        _deepgramKeytermsBox.Visible = true;
        _deepgramKeytermsCount.Visible = true;
        _deepgramKeytermsSortButton.Visible = true;
        _keytermLearningCheck.Visible = true;
        _dgLearningStats.Visible = true;
        _dgLearningViewButton.Visible = true;
        _snxLearningStats.Visible = true;
        _snxLearningViewButton.Visible = true;
        _smLearningStats.Visible = true;
        _smLearningViewButton.Visible = true;
        _aaiLearningStats.Visible = true;
        _aaiLearningViewButton.Visible = true;
        _keytermLearningHint.Visible = true;

        // Update pricing hint
        _pricingHint.Text = newIdx switch
        {
            ProviderDeepgramMedical => "Free tier: $200 credit. Nova-3 Medical: $0.0077/min streaming",
            ProviderSpeechmatics => "Free tier: 480 min/month. Medical model: $0.004/min streaming",
            ProviderAssemblyAI => "Free tier: 330 hours. $0.0025/min streaming",
            ProviderSoniox => "Pay-as-you-go. Real-time model: $0.002/min streaming",
            _ => ""
        };

        UpdateEnsembleStates();
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
            case ProviderSoniox:
                _sonioxKey = _apiKeyBox.Text;
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
            case ProviderSoniox:
                _apiKeyBox.Text = _sonioxKey;
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
        _newlineAfterSentenceCheck.Enabled = enabled;
        _deepgramKeytermsBox.Enabled = enabled;
        _deepgramKeytermsSortButton.Enabled = enabled;
        _keytermLearningCheck.Enabled = enabled;
        _dgLearningViewButton.Enabled = enabled;
        _snxLearningViewButton.Enabled = enabled;
        _smLearningViewButton.Enabled = enabled;
        _aaiLearningViewButton.Enabled = enabled;
        _expandContractionsCheck.Enabled = enabled;
        _radiologyCleanupCheck.Enabled = enabled;
        _replacementsGrid.Enabled = enabled;
        _addReplacementButton.Enabled = enabled;
        _removeReplacementButton.Enabled = enabled;
        _startBeepCheck.Enabled = enabled;
        _stopBeepCheck.Enabled = enabled;
        _startBeepVolume.Enabled = enabled;
        _stopBeepVolume.Enabled = enabled;
        _showIndicatorCheck.Enabled = enabled;
        _ensembleCheck.Enabled = enabled;
        UpdateEnsembleStates();
    }

    private void UpdateEnsembleStates()
    {
        bool ensembleEnabled = _enabledCheck.Checked && _ensembleCheck.Checked;
        _ensembleWaitUpDown.Enabled = ensembleEnabled;
        _ensembleThreshUpDown.Enabled = ensembleEnabled;
        _ensembleShowMetricsCheck.Enabled = ensembleEnabled;
        _ensembleS1Combo.Enabled = _enabledCheck.Checked;
        _ensembleS2Combo.Enabled = _enabledCheck.Checked;

        // Dynamic key status based on selected secondaries
        var missing = new List<string>();
        if (string.IsNullOrWhiteSpace(_deepgramKey)) missing.Add("Deepgram");
        var s1Name = EnsembleComboToProvider(_ensembleS1Combo.SelectedIndex);
        var s2Name = EnsembleComboToProvider(_ensembleS2Combo.SelectedIndex);
        if (s1Name != "none" && !HasKeyForSelectedProvider(s1Name)) missing.Add(EnsembleProviderDisplayName(s1Name));
        if (s2Name != "none" && !HasKeyForSelectedProvider(s2Name)) missing.Add(EnsembleProviderDisplayName(s2Name));

        var activeCount = 1 + (s1Name != "none" ? 1 : 0) + (s2Name != "none" ? 1 : 0);
        if (missing.Count == 0)
        {
            _ensembleKeyStatus.Text = activeCount == 1
                ? "Only Deepgram — select secondaries to enable ensemble"
                : $"All {activeCount} API keys configured";
            _ensembleKeyStatus.ForeColor = activeCount == 1 ? Color.FromArgb(200, 180, 80) : Color.FromArgb(100, 180, 100);
        }
        else
        {
            _ensembleKeyStatus.Text = $"Missing: {string.Join(", ", missing)}";
            _ensembleKeyStatus.ForeColor = Color.IndianRed;
        }
    }

    private static string EnsembleComboToProvider(int idx) => idx switch
    {
        0 => "soniox",
        1 => "speechmatics",
        2 => "assemblyai",
        3 => "none",
        _ => "soniox"
    };

    private static int ProviderToEnsembleCombo(string provider) => provider switch
    {
        "soniox" => 0,
        "speechmatics" => 1,
        "assemblyai" => 2,
        "none" => 3,
        _ => 0
    };

    private static string EnsembleProviderDisplayName(string provider) => provider switch
    {
        "soniox" => "Soniox",
        "speechmatics" => "Speechmatics",
        "assemblyai" => "AssemblyAI",
        "none" => "None",
        _ => provider
    };

    private bool HasKeyForSelectedProvider(string provider) => provider switch
    {
        "soniox" => !string.IsNullOrWhiteSpace(_sonioxKey),
        "speechmatics" => !string.IsNullOrWhiteSpace(_speechmaticsKey),
        "assemblyai" => !string.IsNullOrWhiteSpace(_assemblyAIKey),
        "none" => true,
        _ => false
    };

    private void RefreshKeytermLearningStats()
    {
        RefreshProviderStats("deepgram", _dgLearningStats);
        RefreshProviderStats("soniox", _snxLearningStats);
        RefreshProviderStats("speechmatics", _smLearningStats);
        RefreshProviderStats("assemblyai", _aaiLearningStats);
    }

    private static void RefreshProviderStats(string providerName, Label label)
    {
        try
        {
            var svc = new KeytermLearningService(providerName);
            svc.Load();
            var count = svc.EntryCount;
            label.Text = $"{count} term{(count == 1 ? "" : "s")}";
        }
        catch
        {
            label.Text = "0 terms";
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

    private void OnViewKeytermLearningClick(string providerName)
    {
        var svc = new KeytermLearningService(providerName);
        svc.Load();
        var entries = svc.GetAllEntries();

        if (entries.Count == 0)
        {
            var provLabel = providerName switch { "soniox" => "Soniox", "speechmatics" => "Speechmatics", "assemblyai" => "AssemblyAI", _ => "Deepgram" };
            MessageBox.Show($"No auto-learned keyterms for {provLabel} yet.\n\nDictate reports with Custom STT enabled, " +
                "then sign/advance studies. Terms will appear after the first verification cycle.",
                $"Auto-Learned Keyterms ({provLabel})", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        var provDisplayName = providerName switch { "soniox" => "Soniox", "speechmatics" => "Speechmatics", "assemblyai" => "AssemblyAI", _ => "Deepgram" };
        lines.AppendLine();
        lines.AppendLine($"Total: {entries.Count} auto-learned terms ({sentCount} sent to {provDisplayName}, {manualTerms.Count} manual)");
        lines.AppendLine();
        lines.AppendLine($"Conf  = average {provDisplayName} confidence when the word was heard");
        lines.AppendLine("Score = log(1 + occurrences) \u00d7 (1 \u2212 conf)\u00b2 \u00d7 weight");
        lines.AppendLine("  survived=1.0x, corrected=2.0x (corrections are more valuable)");
        lines.AppendLine("  Log dampens volume; squared gap prioritizes truly problematic words.");
        lines.AppendLine();
        lines.AppendLine($"Top-scoring terms are sent to {provDisplayName} as keyterms.");

        var form = new Form
        {
            Text = $"Auto-Learned Keyterms ({provDisplayName})",
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

    private void OnAddReplacementClick(object? sender, EventArgs e)
    {
        _replacementsGrid.Rows.Add(true, "", "");
        // Select the new row's Find cell for immediate editing
        var newRow = _replacementsGrid.Rows[_replacementsGrid.Rows.Count - 1];
        _replacementsGrid.CurrentCell = newRow.Cells["Find"];
        _replacementsGrid.BeginEdit(true);
    }

    private void OnRemoveReplacementClick(object? sender, EventArgs e)
    {
        if (_replacementsGrid.CurrentRow != null)
        {
            _replacementsGrid.Rows.Remove(_replacementsGrid.CurrentRow);
        }
    }

    private void OnGetKeyClick(object? sender, EventArgs e)
    {
        var url = _providerCombo.SelectedIndex switch
        {
            ProviderSpeechmatics => "https://portal.speechmatics.com/signup",
            ProviderAssemblyAI => "https://www.assemblyai.com/dashboard/signup",
            ProviderSoniox => "https://console.soniox.com",
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
            ProviderSoniox => (
                "How to get a Soniox API key:\n\n" +
                "1. Click \"Get API Key\" to open the Soniox Console\n" +
                "2. Create an account\n" +
                "3. After signing in, go to API Keys\n" +
                "4. Create a new API key and copy it\n" +
                "5. Paste the key into the API Key field above\n\n" +
                "Real-time streaming: $0.002/min ($0.12/hour).\n" +
                "Token-by-token delivery with word-level confidence.\n\n" +
                "Keep your API key private \u2014 do not share it.",
                "Soniox Setup"),
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
        _sonioxKey = config.SttSonioxApiKey;

        // Map provider to dropdown index (corti falls back to deepgram)
        var idx = config.SttProvider switch
        {
            "deepgram" => ProviderDeepgramMedical,
            "speechmatics" => ProviderSpeechmatics,
            "assemblyai" => ProviderAssemblyAI,
            "soniox" => ProviderSoniox,
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
        // Keyterms and auto-learning are now supported by all providers
        _deepgramKeytermsLabel.Visible = true;
        _deepgramKeytermsBox.Visible = true;
        _deepgramKeytermsCount.Visible = true;
        _deepgramKeytermsSortButton.Visible = true;
        _keytermLearningCheck.Visible = true;
        _dgLearningStats.Visible = true;
        _dgLearningViewButton.Visible = true;
        _snxLearningStats.Visible = true;
        _snxLearningViewButton.Visible = true;
        _smLearningStats.Visible = true;
        _smLearningViewButton.Visible = true;
        _aaiLearningStats.Visible = true;
        _aaiLearningViewButton.Visible = true;
        _keytermLearningHint.Visible = true;
        _pricingHint.Text = idx switch
        {
            ProviderDeepgramMedical => "Free tier: $200 credit. Nova-3 Medical: $0.0077/min streaming",
            ProviderSpeechmatics => "Free tier: 480 min/month. Medical model: $0.004/min streaming",
            ProviderAssemblyAI => "Free tier: 330 hours. $0.0025/min streaming",
            ProviderSoniox => "Pay-as-you-go. Real-time model: $0.002/min streaming",
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
        _newlineAfterSentenceCheck.Checked = config.SttNewlineAfterSentence;
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

        // Ensemble
        _ensembleCheck.Checked = config.SttEnsembleEnabled;
        _ensembleS1Combo.SelectedIndex = ProviderToEnsembleCombo(config.SttEnsembleSecondary1);
        _ensembleS2Combo.SelectedIndex = ProviderToEnsembleCombo(config.SttEnsembleSecondary2);
        _ensembleWaitUpDown.Value = Math.Clamp(config.SttEnsembleWaitMs, 100, 2000);
        _ensembleThreshUpDown.Value = Math.Clamp((int)(config.SttEnsembleConfidenceThreshold * 100), 50, 99);
        _ensembleShowMetricsCheck.Checked = config.SttEnsembleShowMetrics;
        UpdateEnsembleStates();

        // Text processing
        _expandContractionsCheck.Checked = config.SttExpandContractions;
        _radiologyCleanupCheck.Checked = config.SttRadiologyCleanup;
        _replacementsGrid.Rows.Clear();
        foreach (var entry in config.SttCustomReplacements)
        {
            _replacementsGrid.Rows.Add(entry.Enabled, entry.Find, entry.Replace);
        }

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
        config.SttSonioxApiKey = _sonioxKey.Trim();

        // Map dropdown index to provider+model
        (config.SttProvider, config.SttModel) = _providerCombo.SelectedIndex switch
        {
            ProviderDeepgramMedical => ("deepgram", "nova-3-medical"),
            ProviderSpeechmatics => ("speechmatics", ""),
            ProviderAssemblyAI => ("assemblyai", ""),
            ProviderSoniox => ("soniox", ""),
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
        config.SttNewlineAfterSentence = _newlineAfterSentenceCheck.Checked;
        config.SttDeepgramKeyterms = _deepgramKeytermsBox.Text.Trim();
        config.SttStartBeepEnabled = _startBeepCheck.Checked;
        config.SttStopBeepEnabled = _stopBeepCheck.Checked;
        config.SttStartBeepVolume = _startBeepVolume.Value / 100.0;
        config.SttStopBeepVolume = _stopBeepVolume.Value / 100.0;
        config.SttShowIndicator = _showIndicatorCheck.Checked;
        config.SttAutoStartOnCase = _autoStartCheck.Checked;
        config.SttKeytermLearningEnabled = _keytermLearningCheck.Checked;

        // Ensemble
        config.SttEnsembleEnabled = _ensembleCheck.Checked;
        config.SttEnsembleSecondary1 = EnsembleComboToProvider(_ensembleS1Combo.SelectedIndex);
        config.SttEnsembleSecondary2 = EnsembleComboToProvider(_ensembleS2Combo.SelectedIndex);
        config.SttEnsembleWaitMs = (int)_ensembleWaitUpDown.Value;
        config.SttEnsembleConfidenceThreshold = (double)_ensembleThreshUpDown.Value / 100.0;
        config.SttEnsembleShowMetrics = _ensembleShowMetricsCheck.Checked;

        // Text processing
        config.SttExpandContractions = _expandContractionsCheck.Checked;
        config.SttRadiologyCleanup = _radiologyCleanupCheck.Checked;
        config.SttCustomReplacements.Clear();
        foreach (DataGridViewRow row in _replacementsGrid.Rows)
        {
            var find = row.Cells["Find"].Value?.ToString()?.Trim() ?? "";
            var replace = row.Cells["Replace"].Value?.ToString()?.Trim() ?? "";
            if (find.Length == 0) continue; // Skip empty entries
            config.SttCustomReplacements.Add(new SttReplacementEntry
            {
                Enabled = row.Cells["Enabled"].Value is true,
                Find = find,
                Replace = replace
            });
        }
    }
}
