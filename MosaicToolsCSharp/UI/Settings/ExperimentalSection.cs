using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI.Settings;

/// <summary>
/// Experimental settings: Network monitor.
/// </summary>
public class ExperimentalSection : SettingsSection
{
    public override string SectionId => "experimental";

    private readonly CheckBox _connectivityMonitorEnabledCheck;
    private readonly NumericUpDown _connectivityIntervalUpDown;
    private readonly NumericUpDown _connectivityTimeoutUpDown;
    private readonly CheckBox _useSendInputInsertCheck;
    private readonly CheckBox _cdpEnabledCheck;
    private readonly CheckBox _cdpScrollFixCheck;
    private readonly CheckBox _cdpAutoScrollCheck;
    private readonly CheckBox _cdpHideDragHandlesCheck;
    private readonly CheckBox _cdpFlashingAlertTextCheck;
    private readonly CheckBox _sttHighlightDictatedCheck;
    private readonly CheckBox _cdpVisualEnhancementsCheck;
    private readonly CheckBox _impressionFixerInEditorCheck;
    private readonly CheckBox _clarioCdpEnabledCheck;
    private readonly TextBox _clarioCdpUrlBox;
    private readonly Button _clarioCdpSetupButton;
    private readonly Label _clarioCdpStatusLabel;
    private readonly Label _mosaicCdpDot;
    private readonly Label _mosaicCdpLabel;
    private readonly Label _clarioCdpDot;
    private readonly Label _clarioCdpLabel;
    private readonly CheckBox _llmProcessEnabledCheck;
    private readonly ComboBox _llmProviderCombo;
    // Gemini controls
    private readonly TextBox _llmApiKeyBox;
    private readonly ComboBox _llmVersionCombo;
    private readonly ComboBox _llmTierCombo;
    private readonly Label _llmPreviewLabel;
    private readonly ComboBox _llmModeCombo;
    private readonly List<Control> _geminiControls = new();
    // GPT controls
    private readonly TextBox _gptApiKeyBox;
    private readonly ComboBox _gptModelCombo;
    private readonly ComboBox _gptModeCombo;
    private readonly List<Control> _gptControls = new();
    // Groq controls
    private readonly TextBox _groqApiKeyBox;
    private readonly ComboBox _groqModelCombo;
    private readonly ComboBox _groqModeCombo;
    private readonly List<Control> _groqControls = new();
    // Grok (xAI) controls
    private readonly TextBox _grokApiKeyBox;
    private readonly ComboBox _grokModelCombo;
    private readonly ComboBox _grokModeCombo;
    private readonly List<Control> _grokControls = new();
    // Quad Compare controls
    private readonly ComboBox _quadCombo1;
    private readonly ComboBox _quadCombo2;
    private readonly ComboBox _quadCombo3;
    private readonly ComboBox _quadCombo4;
    private readonly List<Control> _quadControls = new();
    // Shared
    private readonly Label _llmCostLabel;

    private static readonly (string Version, string Tier, string ModelId, bool Preview,
        decimal InputPer1M, decimal OutputPer1M)[] LlmModels = {
        ("2.5", "Flash Lite", "gemini-2.5-flash-lite",          false, 0.10m, 0.40m),
        ("2.5", "Flash",      "gemini-2.5-flash",               false, 0.30m, 2.50m),
        ("2.5", "Pro",        "gemini-2.5-pro",                 false, 1.25m, 10.00m),
        ("3.0", "Flash",      "gemini-3-flash-preview",         true,  0.50m, 3.00m),
        ("3.1", "Flash Lite", "gemini-3.1-flash-lite-preview",  true,  0.25m, 1.50m),
        ("3.1", "Pro",        "gemini-3.1-pro-preview",         true,  2.00m, 12.00m),
    };

    private static readonly (string Name, string ModelId, decimal InputPer1M, decimal OutputPer1M)[] GptModels = {
        ("GPT-4.1 Mini", "gpt-4.1-mini", 0.16m, 0.64m),
        ("GPT-5 Nano",   "gpt-5-nano",   0.05m, 0.40m),
        ("GPT-5 Mini",   "gpt-5-mini",   0.25m, 2.00m),
    };

    private static readonly (string Name, string ModelId, decimal InputPer1M, decimal OutputPer1M)[] GroqModels = {
        ("GPT-OSS 120B",     "openai/gpt-oss-120b",                           0.15m, 0.75m),
        ("Llama 4 Maverick", "meta-llama/llama-4-maverick-17b-128e-instruct",  0.50m, 0.77m),
        ("Llama 3.3 70B",    "llama-3.3-70b-versatile",                        0.59m, 0.79m),
    };

    private static readonly (string Name, string ModelId, decimal InputPer1M, decimal OutputPer1M)[] GrokModels = {
        ("Grok 3 Mini",      "grok-3-mini",                  0.30m, 0.50m),
        ("Grok 4.1 Fast",    "grok-4-1-fast-non-reasoning",  0.20m, 0.50m),
    };

    public ExperimentalSection(ToolTip toolTip) : base("Experimental", toolTip)
    {

        var warningLabel = AddLabel("⚠ These features may change or be removed", LeftMargin, _nextY);
        warningLabel.ForeColor = Color.FromArgb(255, 180, 50);
        warningLabel.Font = new Font("Segoe UI", 8, FontStyle.Italic);
        _nextY += RowHeight;

        AddSectionDivider("Insertion");

        _useSendInputInsertCheck = AddCheckBox("Use SendInput instead of Ctrl+V", LeftMargin, _nextY,
            "Experimental: insert text via SendInput (no Ctrl+V). Default remains clipboard + Ctrl+V.");
        _nextY += RowHeight;

        // Network Monitor
        AddSectionDivider("Network Monitor");

        _connectivityMonitorEnabledCheck = AddCheckBox("Show connectivity status dots", LeftMargin, _nextY,
            "Show connectivity dots in widget bar for network status monitoring.");
        _connectivityMonitorEnabledCheck.CheckedChanged += (s, e) => UpdateNetworkSettingsStates();
        _nextY += RowHeight;

        AddLabel("Check every:", LeftMargin + 25, _nextY + 3);
        _connectivityIntervalUpDown = AddNumericUpDown(LeftMargin + 110, _nextY, 50, 10, 120, 30);
        AddLabel("sec", LeftMargin + 165, _nextY + 3);

        AddLabel("Timeout:", LeftMargin + 210, _nextY + 3);
        _connectivityTimeoutUpDown = AddNumericUpDown(LeftMargin + 270, _nextY, 50, 1, 10, 5);
        AddLabel("sec", LeftMargin + 325, _nextY + 3);
        _nextY += SubRowHeight;

        AddHintLabel("Shows 4 status dots: Mirth, Mosaic, Clario, InteleViewer", LeftMargin + 25);

        // CDP (Direct DOM Access)
        AddSectionDivider("Mosaic CDP (Direct DOM Access)");

        _cdpEnabledCheck = AddCheckBox("Use CDP for Mosaic interaction", LeftMargin, _nextY,
            "Falls back to UI Automation if CDP unavailable. Requires Mosaic restart on first enable.");
        _cdpEnabledCheck.CheckedChanged += (s, e) => UpdateCdpSettingsStates();
        _nextY += RowHeight;

        AddHintLabel("Reads Mosaic DOM directly via Chrome DevTools Protocol — faster, no COM leaks", LeftMargin + 25);

        _cdpScrollFixCheck = AddCheckBox("Independent column scrolling", LeftMargin + 25, _nextY,
            "Makes Transcript, Report, and sidebar columns scroll independently instead of the whole page.");
        _nextY += RowHeight;

        _cdpAutoScrollCheck = AddCheckBox("Auto-scroll to cursor during dictation", LeftMargin + 25, _nextY,
            "Keeps the cursor visible by scrolling the editor area when text is inserted near the bottom.");
        _nextY += RowHeight;

        _cdpHideDragHandlesCheck = AddCheckBox("Hide report drag/delete handles", LeftMargin + 25, _nextY,
            "Hides the drag-to-reorder and delete icons on paragraphs in the report editor.");
        _nextY += RowHeight;

        _cdpFlashingAlertTextCheck = AddCheckBox("Flashing Alert Text", LeftMargin + 25, _nextY,
            "When gender mismatch or findings/impression mismatch is active, matching final report text flashes via CDP until corrected.");
        _nextY += RowHeight;

        AddHintLabel("Alert colors:", LeftMargin + 25);
        var genderColorLabel = AddLabel("Gender mismatch = red (#DC0000 / #780000)", LeftMargin + 40, _nextY);
        genderColorLabel.ForeColor = Color.FromArgb(255, 120, 120);
        genderColorLabel.Font = new Font("Segoe UI", 8);
        _nextY += 18;

        var fimColorLabel = AddLabel("FIM mismatch = orange (#FFA000 / #B46E00)", LeftMargin + 40, _nextY);
        fimColorLabel.ForeColor = Color.FromArgb(255, 190, 110);
        fimColorLabel.Font = new Font("Segoe UI", 8);
        _nextY += 18;

        _sttHighlightDictatedCheck = AddCheckBox("Highlight dictated text", LeftMargin + 25, _nextY,
            "Applies a subtle background tint to text inserted via custom STT, to distinguish it from template and typed text. Requires CDP.");
        _nextY += RowHeight;

        _cdpVisualEnhancementsCheck = AddCheckBox("Visual report enhancements", LeftMargin + 25, _nextY,
            "Carolina blue change-tracking highlight, inline subsection headers, and slightly smaller content text for visual hierarchy.");
        _nextY += RowHeight;

        _impressionFixerInEditorCheck = AddCheckBox("Impression fixer buttons in editor", LeftMargin + 25, _nextY,
            "Show impression fixer quick-insert/replace buttons directly in the Mosaic report editor next to IMPRESSION.");
        _nextY += RowHeight;

        // Clario CDP
        AddSectionDivider("Clario CDP (Direct DOM Access)");

        _clarioCdpEnabledCheck = AddCheckBox("Use CDP for Clario interaction", LeftMargin, _nextY,
            "Reads Clario data via Chrome DevTools Protocol instead of UI Automation. Requires Chrome launched with CDP shortcut.");
        _nextY += RowHeight;

        AddHintLabel("Priority/Class, exam notes, critical note creation —", LeftMargin + 25);
        AddHintLabel("all via instant JS eval instead of FlaUI tree walks", LeftMargin + 25);

        AddLabel("Clario URL:", LeftMargin + 25, _nextY + 3);
        _clarioCdpUrlBox = new TextBox
        {
            Location = new System.Drawing.Point(LeftMargin + 100, _nextY),
            Width = 220,
            Font = new Font("Segoe UI", 9)
        };
        _toolTip.SetToolTip(_clarioCdpUrlBox, "The Clario web app URL (used for the CDP shortcut).");
        Controls.Add(_clarioCdpUrlBox);
        _nextY += RowHeight;

        _clarioCdpSetupButton = AddButton("Setup Chrome CDP Shortcut", LeftMargin + 25, _nextY, 200, 26,
            OnSetupClarioCdpClick, "Creates a desktop shortcut and directory junction so Chrome launches with CDP enabled while keeping your normal profile.");
        _clarioCdpStatusLabel = AddLabel("", LeftMargin + 235, _nextY + 4);
        _clarioCdpStatusLabel.Font = new Font("Segoe UI", 8);
        _clarioCdpStatusLabel.AutoSize = true;
        UpdateClarioCdpStatus();
        _nextY += RowHeight;

        AddHintLabel("Creates 'Clario CDP' shortcut on desktop. Close Chrome before launching it.", LeftMargin + 25);

        // CDP Status Dashboard
        AddSectionDivider("CDP Status");

        var dotFont = new Font("Segoe UI", 9, FontStyle.Bold);
        var labelFont = new Font("Segoe UI", 8);

        _mosaicCdpDot = AddLabel("●", LeftMargin + 10, _nextY);
        _mosaicCdpDot.Font = dotFont;
        _mosaicCdpDot.ForeColor = Color.Gray;
        _mosaicCdpDot.AutoSize = true;
        _mosaicCdpLabel = AddLabel("Mosaic: checking...", LeftMargin + 28, _nextY + 1);
        _mosaicCdpLabel.Font = labelFont;
        _mosaicCdpLabel.ForeColor = Color.Gray;
        _mosaicCdpLabel.AutoSize = true;
        _nextY += 20;

        _clarioCdpDot = AddLabel("●", LeftMargin + 10, _nextY);
        _clarioCdpDot.Font = dotFont;
        _clarioCdpDot.ForeColor = Color.Gray;
        _clarioCdpDot.AutoSize = true;
        _clarioCdpLabel = AddLabel("Clario: checking...", LeftMargin + 28, _nextY + 1);
        _clarioCdpLabel.Font = labelFont;
        _clarioCdpLabel.ForeColor = Color.Gray;
        _clarioCdpLabel.AutoSize = true;
        _nextY += 20;

        // Beyond Experimental — LLM Custom Process Report
        _nextY += 10;
        var separatorLabel = AddLabel("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", LeftMargin, _nextY);
        separatorLabel.ForeColor = Color.FromArgb(255, 100, 60);
        separatorLabel.Font = new Font("Segoe UI", 8);
        _nextY += 18;
        var beyondLabel = AddLabel("BEYOND EXPERIMENTAL", LeftMargin, _nextY);
        beyondLabel.ForeColor = Color.FromArgb(255, 100, 60);
        beyondLabel.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        _nextY += 18;
        var separatorLabel2 = AddLabel("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", LeftMargin, _nextY);
        separatorLabel2.ForeColor = Color.FromArgb(255, 100, 60);
        separatorLabel2.Font = new Font("Segoe UI", 8);
        _nextY += RowHeight;

        _llmProcessEnabledCheck = AddCheckBox("Custom Process Report (LLM)", LeftMargin, _nextY,
            "Reads transcript + template via CDP, sends to LLM, writes merged report back. " +
            "Bypasses Mosaic's built-in LLM. Requires CDP enabled. PHI is scrubbed before sending.");
        _llmProcessEnabledCheck.CheckedChanged += (s, e) => UpdateLlmSettingsStates();
        _nextY += RowHeight;

        // Provider selector
        AddLabel("Provider:", LeftMargin + 25, _nextY + 3);
        _llmProviderCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 120,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.FromArgb(100, 200, 255),
            FlatStyle = FlatStyle.Flat
        };
        _llmProviderCombo.Items.AddRange(new object[] { "Google Gemini", "OpenAI GPT", "Groq", "xAI Grok", "Quad Compare" });
        _llmProviderCombo.SelectedIndex = 0;
        _llmProviderCombo.SelectedIndexChanged += (_, _) => OnProviderChanged();
        Controls.Add(_llmProviderCombo);
        _nextY += RowHeight;

        // We'll build Gemini and GPT controls at the same Y positions (overlapping), then toggle visibility.
        int providerStartY = _nextY;

        // ═══ GEMINI CONTROLS ═══
        var gApiKeyLabel = AddLabel("API Key:", LeftMargin + 25, _nextY + 3);
        _geminiControls.Add(gApiKeyLabel);
        _llmApiKeyBox = new TextBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 200,
            Font = new Font("Segoe UI", 9),
            UseSystemPasswordChar = true
        };
        _toolTip.SetToolTip(_llmApiKeyBox, "Google Gemini API key (from aistudio.google.com)");
        Controls.Add(_llmApiKeyBox);
        _geminiControls.Add(_llmApiKeyBox);

        var getGeminiKeyBtn = AddButton("Get Key", LeftMargin + 295, _nextY - 1, 55, 22,
            (_, _) => { try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://aistudio.google.com/apikey") { UseShellExecute = true }); } catch { } },
            "Open Google AI Studio to create an API key");
        getGeminiKeyBtn.Font = new Font("Segoe UI", 7);
        _geminiControls.Add(getGeminiKeyBtn);
        _nextY += SubRowHeight;

        _geminiControls.Add(CreateHintLabel("1. Click 'Get Key' → sign in with Google → 'Create API Key'", LeftMargin + 50));
        _geminiControls.Add(CreateHintLabel("2. Select 'Create API key in new project' → copy the key → paste above", LeftMargin + 50));

        var gVersionLabel = AddLabel("Version:", LeftMargin + 25, _nextY + 3);
        _geminiControls.Add(gVersionLabel);
        _llmVersionCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 55,
            Font = new Font("Segoe UI", 9),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _llmVersionCombo.Items.AddRange(LlmModels.Select(m => m.Version).Distinct().Cast<object>().ToArray());
        _llmVersionCombo.SelectedIndex = 0;
        _llmVersionCombo.SelectedIndexChanged += (s, e) => OnVersionChanged();
        Controls.Add(_llmVersionCombo);
        _geminiControls.Add(_llmVersionCombo);

        var gTierLabel = AddLabel("Tier:", LeftMargin + 160, _nextY + 3);
        _geminiControls.Add(gTierLabel);
        _llmTierCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 195, _nextY),
            Width = 100,
            Font = new Font("Segoe UI", 9),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _llmTierCombo.SelectedIndexChanged += TierChangedHandler;
        Controls.Add(_llmTierCombo);
        _geminiControls.Add(_llmTierCombo);

        _llmPreviewLabel = new Label
        {
            Text = "(preview)",
            Location = new Point(LeftMargin + 300, _nextY + 3),
            AutoSize = true,
            ForeColor = Color.FromArgb(140, 140, 140),
            Font = new Font("Segoe UI", 8, FontStyle.Italic),
            Visible = false
        };
        Controls.Add(_llmPreviewLabel);
        _geminiControls.Add(_llmPreviewLabel);
        _nextY += SubRowHeight;

        // Skip a line for cost (shared — positioned later)
        _nextY += 20;

        // Initialize tier combo for the default version
        OnVersionChanged();

        var gModeLabel = AddLabel("Mode:", LeftMargin + 25, _nextY + 3);
        _geminiControls.Add(gModeLabel);
        _llmModeCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 90,
            Font = new Font("Segoe UI", 9),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _llmModeCombo.Items.AddRange(new object[] { "Single", "Dual", "Triple" });
        _llmModeCombo.SelectedIndex = 0;
        _llmModeCombo.SelectedIndexChanged += (_, _) => UpdateCostDisplay();
        Controls.Add(_llmModeCombo);
        _geminiControls.Add(_llmModeCombo);
        _toolTip.SetToolTip(_llmModeCombo,
            "Single: selected model only.\n" +
            "Dual: 2.5 Lite + 3.1 Lite in parallel — 3.1 preferred if within 1s.\n" +
            "Triple: Dual + 3.0 Flash — Flash replaces Lite if it arrives in time.");
        _nextY += SubRowHeight;

        int geminiEndY = _nextY;

        // ═══ GPT CONTROLS ═══ (built at same Y range, toggled visibility)
        _nextY = providerStartY;

        var oApiKeyLabel = AddLabel("API Key:", LeftMargin + 25, _nextY + 3);
        _gptControls.Add(oApiKeyLabel);
        _gptApiKeyBox = new TextBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 200,
            Font = new Font("Segoe UI", 9),
            UseSystemPasswordChar = true
        };
        _toolTip.SetToolTip(_gptApiKeyBox, "OpenAI API key (from platform.openai.com)");
        Controls.Add(_gptApiKeyBox);
        _gptControls.Add(_gptApiKeyBox);

        var getGptKeyBtn = AddButton("Get Key", LeftMargin + 295, _nextY - 1, 55, 22,
            (_, _) => { try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://platform.openai.com/api-keys") { UseShellExecute = true }); } catch { } },
            "Open OpenAI platform to create an API key");
        getGptKeyBtn.Font = new Font("Segoe UI", 7);
        _gptControls.Add(getGptKeyBtn);
        _nextY += SubRowHeight;

        _gptControls.Add(CreateHintLabel("1. Click 'Get Key' → sign in → 'Create new secret key'", LeftMargin + 50));
        _gptControls.Add(CreateHintLabel("2. Copy the key → paste above (starts with sk-)", LeftMargin + 50));

        var oModelLabel = AddLabel("Model:", LeftMargin + 25, _nextY + 3);
        _gptControls.Add(oModelLabel);
        _gptModelCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 100,
            Font = new Font("Segoe UI", 9),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _gptModelCombo.Items.AddRange(GptModels.Select(m => (object)m.Name).ToArray());
        _gptModelCombo.SelectedIndex = 0;
        _gptModelCombo.SelectedIndexChanged += (_, _) => UpdateCostDisplay();
        Controls.Add(_gptModelCombo);
        _gptControls.Add(_gptModelCombo);
        _nextY += SubRowHeight;

        // Skip a line for cost (shared)
        _nextY += 20;

        var oModeLabel = AddLabel("Mode:", LeftMargin + 25, _nextY + 3);
        _gptControls.Add(oModeLabel);
        _gptModeCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 90,
            Font = new Font("Segoe UI", 9),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _gptModeCombo.Items.AddRange(new object[] { "Single", "Triple" });
        _gptModeCombo.SelectedIndex = 0;
        _gptModeCombo.SelectedIndexChanged += (_, _) => UpdateCostDisplay();
        Controls.Add(_gptModeCombo);
        _gptControls.Add(_gptModeCombo);
        _toolTip.SetToolTip(_gptModeCombo,
            "Single: selected model only.\n" +
            "Triple: 4.1 Mini + 5 Nano + 5 Mini — fastest wins, 5 Mini upgrade if it arrives within 8s.");
        _nextY += SubRowHeight;

        int gptEndY = _nextY;

        // ═══ GROQ CONTROLS ═══ (built at same Y range, toggled visibility)
        _nextY = providerStartY;

        var qApiKeyLabel = AddLabel("API Key:", LeftMargin + 25, _nextY + 3);
        _groqControls.Add(qApiKeyLabel);
        _groqApiKeyBox = new TextBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 200,
            Font = new Font("Segoe UI", 9),
            UseSystemPasswordChar = true
        };
        _toolTip.SetToolTip(_groqApiKeyBox, "Groq API key (from console.groq.com)");
        Controls.Add(_groqApiKeyBox);
        _groqControls.Add(_groqApiKeyBox);

        var getGroqKeyBtn = AddButton("Get Key", LeftMargin + 295, _nextY - 1, 55, 22,
            (_, _) => { try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://console.groq.com/keys") { UseShellExecute = true }); } catch { } },
            "Open Groq console to create an API key");
        getGroqKeyBtn.Font = new Font("Segoe UI", 7);
        _groqControls.Add(getGroqKeyBtn);
        _nextY += SubRowHeight;

        _groqControls.Add(CreateHintLabel("1. Click 'Get Key' → sign in → 'Create API Key'", LeftMargin + 50));
        _groqControls.Add(CreateHintLabel("2. Copy the key → paste above (starts with gsk_)", LeftMargin + 50));

        var qModelLabel = AddLabel("Model:", LeftMargin + 25, _nextY + 3);
        _groqControls.Add(qModelLabel);
        _groqModelCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 140,
            Font = new Font("Segoe UI", 9),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _groqModelCombo.Items.AddRange(GroqModels.Select(m => (object)m.Name).ToArray());
        _groqModelCombo.SelectedIndex = 0;
        _groqModelCombo.SelectedIndexChanged += (_, _) => UpdateCostDisplay();
        Controls.Add(_groqModelCombo);
        _groqControls.Add(_groqModelCombo);
        _nextY += SubRowHeight;

        // Skip a line for cost (shared)
        _nextY += 20;

        var qModeLabel = AddLabel("Mode:", LeftMargin + 25, _nextY + 3);
        _groqControls.Add(qModeLabel);
        _groqModeCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 90,
            Font = new Font("Segoe UI", 9),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _groqModeCombo.Items.AddRange(new object[] { "Single", "Triple" });
        _groqModeCombo.SelectedIndex = 0;
        _groqModeCombo.SelectedIndexChanged += (_, _) => UpdateCostDisplay();
        Controls.Add(_groqModeCombo);
        _groqControls.Add(_groqModeCombo);
        _toolTip.SetToolTip(_groqModeCombo,
            "Single: selected model only.\n" +
            "Triple: GPT-OSS 20B + Maverick + 3.3 70B — fastest wins, 70B upgrade if it arrives within 8s.");
        _nextY += SubRowHeight;

        int groqEndY = _nextY;

        // ═══ GROK (xAI) CONTROLS ═══ (built at same Y range, toggled visibility)
        _nextY = providerStartY;

        var kApiKeyLabel = AddLabel("API Key:", LeftMargin + 25, _nextY + 3);
        _grokControls.Add(kApiKeyLabel);
        _grokApiKeyBox = new TextBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 200,
            Font = new Font("Segoe UI", 9),
            UseSystemPasswordChar = true
        };
        _toolTip.SetToolTip(_grokApiKeyBox, "xAI Grok API key (from console.x.ai)");
        Controls.Add(_grokApiKeyBox);
        _grokControls.Add(_grokApiKeyBox);

        var getGrokKeyBtn = AddButton("Get Key", LeftMargin + 295, _nextY - 1, 55, 22,
            (_, _) => { try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://console.x.ai/") { UseShellExecute = true }); } catch { } },
            "Open xAI console to create an API key");
        getGrokKeyBtn.Font = new Font("Segoe UI", 7);
        _grokControls.Add(getGrokKeyBtn);
        _nextY += SubRowHeight;

        _grokControls.Add(CreateHintLabel("1. Click 'Get Key' → sign in → 'Create API Key'", LeftMargin + 50));
        _grokControls.Add(CreateHintLabel("2. Copy the key → paste above", LeftMargin + 50));

        var kModelLabel = AddLabel("Model:", LeftMargin + 25, _nextY + 3);
        _grokControls.Add(kModelLabel);
        _grokModelCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 140,
            Font = new Font("Segoe UI", 9),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _grokModelCombo.Items.AddRange(GrokModels.Select(m => (object)m.Name).ToArray());
        _grokModelCombo.SelectedIndex = 0;
        _grokModelCombo.SelectedIndexChanged += (_, _) => UpdateCostDisplay();
        Controls.Add(_grokModelCombo);
        _grokControls.Add(_grokModelCombo);
        _nextY += SubRowHeight;

        // Skip a line for cost (shared)
        _nextY += 20;

        var kModeLabel = AddLabel("Mode:", LeftMargin + 25, _nextY + 3);
        _grokControls.Add(kModeLabel);
        _grokModeCombo = new ComboBox
        {
            Location = new Point(LeftMargin + 90, _nextY),
            Width = 90,
            Font = new Font("Segoe UI", 9),
            DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        _grokModeCombo.Items.AddRange(new object[] { "Single", "Triple" });
        _grokModeCombo.SelectedIndex = 0;
        _grokModeCombo.SelectedIndexChanged += (_, _) => UpdateCostDisplay();
        Controls.Add(_grokModeCombo);
        _grokControls.Add(_grokModeCombo);
        _toolTip.SetToolTip(_grokModeCombo,
            "Single: selected model only.\n" +
            "Triple: Grok 3 Mini + Grok 4.1 Fast — fastest wins, 4.1 Fast upgrade if it arrives within 8s.");
        _nextY += SubRowHeight;

        int grokEndY = _nextY;

        _nextY = providerStartY;
        // Quad Compare: 4 numbered dropdowns, each with all models from all providers
        // Slot 1 leads with Gemini, Slot 2 with OpenAI, Slot 3 with Groq, Slot 4 with Grok
        var allModels = new List<(string Group, string Name, string ModelId)>();
        foreach (var m in LlmModels)
            allModels.Add(("Gemini", $"Gemini {m.Version} {m.Tier}", m.ModelId));
        foreach (var m in GptModels)
            allModels.Add(("OpenAI", m.Name, m.ModelId));
        foreach (var m in GroqModels)
            allModels.Add(("Groq", m.Name, m.ModelId));
        foreach (var m in GrokModels)
            allModels.Add(("Grok", m.Name, m.ModelId));

        string[] groupOrder = { "Gemini", "OpenAI", "Groq", "Grok" };
        ComboBox[] quadCombos = new ComboBox[4];
        for (int slot = 0; slot < 4; slot++)
        {
            var lead = groupOrder[slot];
            var sorted = allModels.Where(m => m.Group == lead)
                .Concat(allModels.Where(m => m.Group != lead))
                .ToList();

            var lbl = AddLabel($"Slot {slot + 1}:", LeftMargin + 25, _nextY + 3);
            _quadControls.Add(lbl);

            var combo = new ComboBox
            {
                Location = new Point(LeftMargin + 80, _nextY),
                Width = 270,
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            foreach (var m in sorted)
            {
                combo.Items.Add(m.Name);
            }
            combo.Tag = sorted.Select(m => m.ModelId).ToArray();
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;
            combo.SelectedIndexChanged += (_, _) => UpdateCostDisplay();
            Controls.Add(combo);
            _quadControls.Add(combo);
            quadCombos[slot] = combo;
            _nextY += SubRowHeight;
        }
        _quadCombo1 = quadCombos[0];
        _quadCombo2 = quadCombos[1];
        _quadCombo3 = quadCombos[2];
        _quadCombo4 = quadCombos[3];

        int quadEndY = _nextY;

        // Use the tallest of the provider sections
        _nextY = Math.Max(Math.Max(Math.Max(Math.Max(geminiEndY, gptEndY), groqEndY), grokEndY), quadEndY);

        // Shared cost label (positioned below both provider sections)
        _llmCostLabel = new Label
        {
            Location = new Point(LeftMargin + 50, _nextY),
            AutoSize = true,
            ForeColor = Color.FromArgb(120, 120, 120),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        Controls.Add(_llmCostLabel);
        _nextY += SubRowHeight;

        AddHintLabel("Map 'Custom Process Report' in Key Mappings after enabling", LeftMargin + 25);

        // Hide GPT, Groq, Grok, and Quad controls by default
        foreach (var c in _gptControls) c.Visible = false;
        foreach (var c in _groqControls) c.Visible = false;
        foreach (var c in _grokControls) c.Visible = false;
        foreach (var c in _quadControls) c.Visible = false;

        UpdateHeight();
    }

    private void UpdateNetworkSettingsStates()
    {
        bool enabled = _connectivityMonitorEnabledCheck.Checked;
        _connectivityIntervalUpDown.Enabled = enabled;
        _connectivityTimeoutUpDown.Enabled = enabled;
    }

    private void OnSetupClarioCdpClick(object? sender, EventArgs e)
    {
        var url = _clarioCdpUrlBox.Text.Trim();
        if (string.IsNullOrEmpty(url))
        {
            MessageBox.Show("Please enter the Clario URL first.", "Setup", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var (success, message) = ClarioCdpSetup.Setup(url);
        MessageBox.Show(message, success ? "Setup Complete" : "Setup Failed",
            MessageBoxButtons.OK, success ? MessageBoxIcon.Information : MessageBoxIcon.Error);
        UpdateClarioCdpStatus();
    }

    private void UpdateClarioCdpStatus()
    {
        if (ClarioCdpSetup.IsJunctionConfigured())
        {
            _clarioCdpStatusLabel.Text = "✓ Junction configured";
            _clarioCdpStatusLabel.ForeColor = Color.FromArgb(100, 200, 100);
        }
        else
        {
            _clarioCdpStatusLabel.Text = "Not configured";
            _clarioCdpStatusLabel.ForeColor = Color.Gray;
        }
    }

    private void ProbeCdpStatus()
    {
        // Mosaic: find DevToolsActivePort file → get port → HTTP probe
        Task.Run(() =>
        {
            try
            {
                int? port = FindMosaicCdpPort();
                if (port == null)
                {
                    SetDot(_mosaicCdpDot, _mosaicCdpLabel, false, "Mosaic: not detected (no DevTools port)");
                    return;
                }
                var targets = ProbeCdpTargets(port.Value);
                if (targets == null)
                {
                    SetDot(_mosaicCdpDot, _mosaicCdpLabel, false, $"Mosaic: port {port} not responding");
                    return;
                }
                // Look for SlimHub or iframe target
                bool hasSlimHub = false, hasIframe = false;
                foreach (var t in targets.Value.EnumerateArray())
                {
                    var title = t.TryGetProperty("title", out var ti) ? ti.GetString() ?? "" : "";
                    var type = t.TryGetProperty("type", out var ty) ? ty.GetString() ?? "" : "";
                    if (title.Equals("SlimHub", StringComparison.OrdinalIgnoreCase)) hasSlimHub = true;
                    if (type == "iframe") hasIframe = true;
                }
                var detail = hasSlimHub && hasIframe ? "connected (SlimHub + report editor)"
                           : hasSlimHub ? "connected (SlimHub only — no report open)"
                           : $"port {port} open, no Mosaic targets";
                SetDot(_mosaicCdpDot, _mosaicCdpLabel, hasSlimHub || hasIframe, $"Mosaic: {detail}");
            }
            catch
            {
                SetDot(_mosaicCdpDot, _mosaicCdpLabel, false, "Mosaic: probe failed");
            }
        });

        // Clario: HTTP probe on port 9224
        Task.Run(() =>
        {
            try
            {
                var targets = ProbeCdpTargets(9224);
                if (targets == null)
                {
                    SetDot(_clarioCdpDot, _clarioCdpLabel, false, "Clario: port 9224 not responding");
                    return;
                }
                string? clarioTitle = null;
                foreach (var t in targets.Value.EnumerateArray())
                {
                    var title = t.TryGetProperty("title", out var ti) ? ti.GetString() ?? "" : "";
                    if (title.Contains("Clario", StringComparison.OrdinalIgnoreCase))
                    {
                        clarioTitle = title;
                        break;
                    }
                }
                if (clarioTitle != null)
                    SetDot(_clarioCdpDot, _clarioCdpLabel, true, $"Clario: connected ({clarioTitle})");
                else
                    SetDot(_clarioCdpDot, _clarioCdpLabel, false, "Clario: port 9224 open, no Clario target");
            }
            catch
            {
                SetDot(_clarioCdpDot, _clarioCdpLabel, false, "Clario: probe failed");
            }
        });
    }

    private void SetDot(Label dot, Label label, bool ok, string text)
    {
        if (IsDisposed) return;
        try
        {
            BeginInvoke(() =>
            {
                if (IsDisposed) return;
                var color = ok ? Color.FromArgb(80, 200, 80) : Color.FromArgb(200, 80, 80);
                dot.ForeColor = color;
                label.ForeColor = ok ? Color.FromArgb(160, 210, 160) : Color.FromArgb(160, 160, 160);
                label.Text = text;
            });
        }
        catch { }
    }

    private static int? FindMosaicCdpPort()
    {
        // Mosaic's WebView2 writes DevToolsActivePort in its UWP package data dir
        var packagesDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Packages");
        if (!Directory.Exists(packagesDir)) return null;

        var mosaicDirs = Directory.GetDirectories(packagesDir, "MosaicInfoHub_*");
        if (mosaicDirs.Length == 0) return null;

        var portFile = Path.Combine(mosaicDirs[0], "LocalState", "EBWebView", "DevToolsActivePort");
        if (!File.Exists(portFile)) return null;

        try
        {
            var lines = File.ReadAllLines(portFile);
            if (lines.Length > 0 && int.TryParse(lines[0].Trim(), out int port) && port > 0)
                return port;
        }
        catch { }
        return null;
    }

    private static JsonElement? ProbeCdpTargets(int port)
    {
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
            var json = http.GetStringAsync($"http://127.0.0.1:{port}/json/list").GetAwaiter().GetResult();
            return JsonSerializer.Deserialize<JsonElement>(json);
        }
        catch { return null; }
    }

    private void UpdateCdpSettingsStates()
    {
        _cdpScrollFixCheck.Enabled = _cdpEnabledCheck.Checked;
        _cdpAutoScrollCheck.Enabled = _cdpEnabledCheck.Checked;
        _cdpHideDragHandlesCheck.Enabled = _cdpEnabledCheck.Checked;
        _cdpFlashingAlertTextCheck.Enabled = _cdpEnabledCheck.Checked;
        _sttHighlightDictatedCheck.Enabled = _cdpEnabledCheck.Checked;
        _cdpVisualEnhancementsCheck.Enabled = _cdpEnabledCheck.Checked;
        _impressionFixerInEditorCheck.Enabled = _cdpEnabledCheck.Checked;
    }

    private void UpdateLlmSettingsStates()
    {
        var enabled = _llmProcessEnabledCheck.Checked;
        _llmProviderCombo.Enabled = enabled;
        _llmApiKeyBox.Enabled = enabled;
        _llmVersionCombo.Enabled = enabled;
        _llmTierCombo.Enabled = enabled;
        _llmModeCombo.Enabled = enabled;
        _gptApiKeyBox.Enabled = enabled;
        _gptModelCombo.Enabled = enabled;
        _gptModeCombo.Enabled = enabled;
        _groqApiKeyBox.Enabled = enabled;
        _groqModelCombo.Enabled = enabled;
        _groqModeCombo.Enabled = enabled;
        _grokApiKeyBox.Enabled = enabled;
        _grokModelCombo.Enabled = enabled;
        _grokModeCombo.Enabled = enabled;
    }

    private Label CreateHintLabel(string text, int x)
    {
        var label = new Label
        {
            Text = text,
            Location = new Point(x, _nextY),
            AutoSize = true,
            ForeColor = Color.FromArgb(120, 120, 120),
            Font = new Font("Segoe UI", 8, FontStyle.Italic)
        };
        Controls.Add(label);
        _nextY += 18;
        return label;
    }

    private void OnProviderChanged()
    {
        int idx = _llmProviderCombo.SelectedIndex; // 0=Gemini, 1=GPT, 2=Groq, 3=Grok, 4=Quad
        foreach (var c in _geminiControls) c.Visible = idx == 0;
        foreach (var c in _gptControls) c.Visible = idx == 1;
        foreach (var c in _groqControls) c.Visible = idx == 2;
        foreach (var c in _grokControls) c.Visible = idx == 3;
        foreach (var c in _quadControls) c.Visible = idx == 4;
        _llmCostLabel.Visible = true;
        _llmPreviewLabel.Visible = idx == 0 && _llmPreviewLabel.Tag != null;
        if (idx == 0) UpdateModelDisplay();
        UpdateCostDisplay();
    }

    private void OnVersionChanged()
    {
        var version = _llmVersionCombo.SelectedItem?.ToString();
        if (version == null) return;

        var tiers = LlmModels.Where(m => m.Version == version).Select(m => m.Tier).ToArray();
        _llmTierCombo.SelectedIndexChanged -= TierChangedHandler;
        _llmTierCombo.Items.Clear();
        _llmTierCombo.Items.AddRange(tiers.Cast<object>().ToArray());
        if (_llmTierCombo.Items.Count > 0) _llmTierCombo.SelectedIndex = 0;
        _llmTierCombo.SelectedIndexChanged += TierChangedHandler;
        UpdateModelDisplay();
    }

    private void TierChangedHandler(object? s, EventArgs e) => UpdateModelDisplay();

    private void UpdateModelDisplay()
    {
        var version = _llmVersionCombo.SelectedItem?.ToString();
        var tier = _llmTierCombo.SelectedItem?.ToString();
        if (version == null || tier == null) return;

        var match = LlmModels.FirstOrDefault(m => m.Version == version && m.Tier == tier);
        if (match.ModelId == null) return;

        _llmPreviewLabel.Visible = match.Preview;
        UpdateCostDisplay();
    }

    private void SelectQuadCombo(ComboBox combo, string modelId)
    {
        var ids = combo.Tag as string[];
        if (ids == null) return;
        var idx = Array.IndexOf(ids, modelId);
        combo.SelectedIndex = idx >= 0 ? idx : 0;
    }

    private string GetQuadModelId(ComboBox combo)
    {
        var ids = combo.Tag as string[];
        var idx = combo.SelectedIndex;
        if (ids != null && idx >= 0 && idx < ids.Length)
            return ids[idx];
        return "";
    }

    private void UpdateCostDisplay()
    {
        if (_llmCostLabel == null) return; // constructor not finished

        const decimal estInput = 1500m;
        const decimal estOutput = 500m;
        const int studiesPerShift = 200;

        int providerIdx = _llmProviderCombo?.SelectedIndex ?? 0; // 0=Gemini, 1=GPT, 2=Groq, 3=Grok
        decimal perStudy;
        string modeLabel;

        if (providerIdx == 0)
        {
            var mode = _llmModeCombo?.SelectedIndex ?? 0; // 0=Single, 1=Dual, 2=Triple
            if (mode >= 1)
            {
                var lite25 = LlmModels.First(m => m.ModelId == "gemini-2.5-flash-lite");
                var lite31 = LlmModels.First(m => m.ModelId == "gemini-3.1-flash-lite-preview");
                perStudy = (lite25.InputPer1M * estInput + lite25.OutputPer1M * estOutput) / 1_000_000m
                         + (lite31.InputPer1M * estInput + lite31.OutputPer1M * estOutput) / 1_000_000m;
                if (mode == 2)
                {
                    var flash = LlmModels.First(m => m.ModelId == "gemini-3-flash-preview");
                    perStudy += (flash.InputPer1M * estInput + flash.OutputPer1M * estOutput) / 1_000_000m;
                }
            }
            else
            {
                var version = _llmVersionCombo?.SelectedItem?.ToString();
                var tier = _llmTierCombo?.SelectedItem?.ToString();
                var match = LlmModels.FirstOrDefault(m => m.Version == version && m.Tier == tier);
                if (match.ModelId == null) { _llmCostLabel.Text = ""; return; }
                perStudy = (match.InputPer1M * estInput + match.OutputPer1M * estOutput) / 1_000_000m;
            }
            modeLabel = mode switch { 1 => " [dual]", 2 => " [triple]", _ => "" };
        }
        else if (providerIdx == 1)
        {
            var gptMode = _gptModeCombo?.SelectedIndex ?? 0; // 0=Single, 1=Triple
            if (gptMode == 1)
            {
                perStudy = 0;
                foreach (var m in GptModels)
                    perStudy += (m.InputPer1M * estInput + m.OutputPer1M * estOutput) / 1_000_000m;
            }
            else
            {
                var idx = _gptModelCombo?.SelectedIndex ?? 0;
                var model = GptModels[idx];
                perStudy = (model.InputPer1M * estInput + model.OutputPer1M * estOutput) / 1_000_000m;
            }
            modeLabel = gptMode == 1 ? " [triple]" : "";
        }
        else if (providerIdx == 2)
        {
            var groqMode = _groqModeCombo?.SelectedIndex ?? 0; // 0=Single, 1=Triple
            if (groqMode == 1)
            {
                perStudy = 0;
                foreach (var m in GroqModels)
                    perStudy += (m.InputPer1M * estInput + m.OutputPer1M * estOutput) / 1_000_000m;
            }
            else
            {
                var idx = _groqModelCombo?.SelectedIndex ?? 0;
                var model = GroqModels[idx];
                perStudy = (model.InputPer1M * estInput + model.OutputPer1M * estOutput) / 1_000_000m;
            }
            modeLabel = groqMode == 1 ? " [triple]" : "";
        }
        else if (providerIdx == 3)
        {
            var grokMode = _grokModeCombo?.SelectedIndex ?? 0; // 0=Single, 1=Triple
            if (grokMode == 1)
            {
                perStudy = 0;
                foreach (var m in GrokModels)
                    perStudy += (m.InputPer1M * estInput + m.OutputPer1M * estOutput) / 1_000_000m;
            }
            else
            {
                var idx = _grokModelCombo?.SelectedIndex ?? 0;
                var model = GrokModels[idx];
                perStudy = (model.InputPer1M * estInput + model.OutputPer1M * estOutput) / 1_000_000m;
            }
            modeLabel = grokMode == 1 ? " [triple]" : "";
        }
        else // Quad Compare
        {
            // Build a lookup of all model pricing by ModelId
            var allPricing = new Dictionary<string, (decimal InputPer1M, decimal OutputPer1M)>();
            foreach (var m in LlmModels) allPricing[m.ModelId] = (m.InputPer1M, m.OutputPer1M);
            foreach (var m in GptModels) allPricing[m.ModelId] = (m.InputPer1M, m.OutputPer1M);
            foreach (var m in GroqModels) allPricing[m.ModelId] = (m.InputPer1M, m.OutputPer1M);
            foreach (var m in GrokModels) allPricing[m.ModelId] = (m.InputPer1M, m.OutputPer1M);

            perStudy = 0;
            foreach (var combo in new[] { _quadCombo1, _quadCombo2, _quadCombo3, _quadCombo4 })
            {
                var modelId = GetQuadModelId(combo);
                if (allPricing.TryGetValue(modelId, out var p))
                    perStudy += (p.InputPer1M * estInput + p.OutputPer1M * estOutput) / 1_000_000m;
            }
            modeLabel = " [quad]";
        }

        var perShift = perStudy * studiesPerShift;
        _llmCostLabel.Text = $"Est. cost{modeLabel}: ~${perStudy:F4}/study \u00b7 ~${perShift:F2}/shift ({studiesPerShift} studies)";
    }

    public override void LoadSettings(Configuration config)
    {
        _connectivityMonitorEnabledCheck.Checked = config.ConnectivityMonitorEnabled;
        _connectivityIntervalUpDown.Value = config.ConnectivityCheckIntervalSeconds;
        // Config stores ms, UI shows seconds
        _connectivityTimeoutUpDown.Value = Math.Max(1, (config.ConnectivityTimeoutMs + 500) / 1000);
        _useSendInputInsertCheck.Checked = config.ExperimentalUseSendInputInsert;
        _cdpEnabledCheck.Checked = config.CdpEnabled;
        _cdpScrollFixCheck.Checked = config.CdpIndependentScrolling;
        _cdpAutoScrollCheck.Checked = config.CdpAutoScrollEnabled;
        _cdpHideDragHandlesCheck.Checked = config.CdpHideDragHandles;
        _cdpFlashingAlertTextCheck.Checked = config.CdpFlashingAlertText;
        _sttHighlightDictatedCheck.Checked = config.SttHighlightDictated;
        _cdpVisualEnhancementsCheck.Checked = config.CdpVisualEnhancements;
        _impressionFixerInEditorCheck.Checked = config.ImpressionFixerInEditor;
        _clarioCdpEnabledCheck.Checked = config.ClarioCdpEnabled;
        _clarioCdpUrlBox.Text = config.ClarioCdpUrl;
        _llmProcessEnabledCheck.Checked = config.LlmProcessEnabled;

        // Provider selector
        _llmProviderCombo.SelectedIndex = config.LlmProvider switch { "openai" => 1, "groq" => 2, "grok" => 3, "quad" => 4, _ => 0 };

        // Gemini settings
        _llmApiKeyBox.Text = config.LlmApiKey;
        var saved = LlmModels.FirstOrDefault(m => m.ModelId == config.LlmModel);
        var version = saved.ModelId != null ? saved.Version : "2.5";
        var tier = saved.ModelId != null ? saved.Tier : "Flash Lite";
        var vIdx = _llmVersionCombo.Items.IndexOf(version);
        _llmVersionCombo.SelectedIndex = vIdx >= 0 ? vIdx : 0;
        OnVersionChanged();
        var tIdx = _llmTierCombo.Items.IndexOf(tier);
        _llmTierCombo.SelectedIndex = tIdx >= 0 ? tIdx : 0;
        var modeIdx = (config.LlmProcessMode ?? "single") switch
        {
            "dual" => 1, "triple" => 2, _ => 0
        };
        _llmModeCombo.SelectedIndex = modeIdx;

        // GPT settings
        _gptApiKeyBox.Text = config.LlmOpenAiApiKey;
        var gptSaved = GptModels.FirstOrDefault(m => m.ModelId == config.LlmOpenAiModel);
        var gptIdx = Array.FindIndex(GptModels, m => m.ModelId == (gptSaved.ModelId ?? "gpt-5-nano"));
        _gptModelCombo.SelectedIndex = gptIdx >= 0 ? gptIdx : 0;
        _gptModeCombo.SelectedIndex = config.LlmOpenAiProcessMode == "triple" ? 1 : 0;

        // Groq settings
        _groqApiKeyBox.Text = config.LlmGroqApiKey;
        var groqSaved = GroqModels.FirstOrDefault(m => m.ModelId == config.LlmGroqModel);
        var groqIdx = Array.FindIndex(GroqModels, m => m.ModelId == (groqSaved.ModelId ?? GroqModels[0].ModelId));
        _groqModelCombo.SelectedIndex = groqIdx >= 0 ? groqIdx : 0;
        _groqModeCombo.SelectedIndex = config.LlmGroqProcessMode == "triple" ? 1 : 0;

        // Grok (xAI) settings
        _grokApiKeyBox.Text = config.LlmGrokApiKey;
        var grokSaved = GrokModels.FirstOrDefault(m => m.ModelId == config.LlmGrokModel);
        var grokIdx = Array.FindIndex(GrokModels, m => m.ModelId == (grokSaved.ModelId ?? GrokModels[0].ModelId));
        _grokModelCombo.SelectedIndex = grokIdx >= 0 ? grokIdx : 0;
        _grokModeCombo.SelectedIndex = config.LlmGrokProcessMode == "triple" ? 1 : 0;

        // Quad Compare settings
        SelectQuadCombo(_quadCombo1, config.LlmQuadModel1);
        SelectQuadCombo(_quadCombo2, config.LlmQuadModel2);
        SelectQuadCombo(_quadCombo3, config.LlmQuadModel3);
        SelectQuadCombo(_quadCombo4, config.LlmQuadModel4);

        OnProviderChanged();

        UpdateNetworkSettingsStates();
        UpdateCdpSettingsStates();
        UpdateLlmSettingsStates();
        ProbeCdpStatus();
    }

    public override void SaveSettings(Configuration config)
    {
        config.ConnectivityMonitorEnabled = _connectivityMonitorEnabledCheck.Checked;
        config.ConnectivityCheckIntervalSeconds = (int)_connectivityIntervalUpDown.Value;
        // UI shows seconds, config stores ms
        config.ConnectivityTimeoutMs = (int)_connectivityTimeoutUpDown.Value * 1000;
        config.ExperimentalUseSendInputInsert = _useSendInputInsertCheck.Checked;
        config.CdpEnabled = _cdpEnabledCheck.Checked;
        config.CdpIndependentScrolling = _cdpScrollFixCheck.Checked;
        config.CdpAutoScrollEnabled = _cdpAutoScrollCheck.Checked;
        config.CdpHideDragHandles = _cdpHideDragHandlesCheck.Checked;
        config.CdpFlashingAlertText = _cdpFlashingAlertTextCheck.Checked;
        config.SttHighlightDictated = _sttHighlightDictatedCheck.Checked;
        config.CdpVisualEnhancements = _cdpVisualEnhancementsCheck.Checked;
        config.ImpressionFixerInEditor = _impressionFixerInEditorCheck.Checked;
        config.ClarioCdpEnabled = _clarioCdpEnabledCheck.Checked;
        var url = _clarioCdpUrlBox.Text.Trim();
        if (!string.IsNullOrEmpty(url))
            config.ClarioCdpUrl = url;
        config.LlmProcessEnabled = _llmProcessEnabledCheck.Checked;
        config.LlmProvider = _llmProviderCombo.SelectedIndex switch { 1 => "openai", 2 => "groq", 3 => "grok", 4 => "quad", _ => "gemini" };

        // Gemini settings
        config.LlmApiKey = _llmApiKeyBox.Text.Trim();
        var version = _llmVersionCombo.SelectedItem?.ToString();
        var tier = _llmTierCombo.SelectedItem?.ToString();
        var match = LlmModels.FirstOrDefault(m => m.Version == version && m.Tier == tier);
        config.LlmModel = match.ModelId ?? "gemini-2.5-flash-lite";
        config.LlmProcessMode = _llmModeCombo.SelectedIndex switch
        {
            1 => "dual", 2 => "triple", _ => "single"
        };

        // GPT settings
        config.LlmOpenAiApiKey = _gptApiKeyBox.Text.Trim();
        var gptIdx = _gptModelCombo.SelectedIndex;
        config.LlmOpenAiModel = gptIdx >= 0 && gptIdx < GptModels.Length ? GptModels[gptIdx].ModelId : "gpt-5-nano";
        config.LlmOpenAiProcessMode = _gptModeCombo.SelectedIndex == 1 ? "triple" : "single";

        // Groq settings
        config.LlmGroqApiKey = _groqApiKeyBox.Text.Trim();
        var groqIdx = _groqModelCombo.SelectedIndex;
        config.LlmGroqModel = groqIdx >= 0 && groqIdx < GroqModels.Length ? GroqModels[groqIdx].ModelId : GroqModels[0].ModelId;
        config.LlmGroqProcessMode = _groqModeCombo.SelectedIndex == 1 ? "triple" : "single";

        // Grok (xAI) settings
        config.LlmGrokApiKey = _grokApiKeyBox.Text.Trim();
        var grokIdx = _grokModelCombo.SelectedIndex;
        config.LlmGrokModel = grokIdx >= 0 && grokIdx < GrokModels.Length ? GrokModels[grokIdx].ModelId : GrokModels[0].ModelId;
        config.LlmGrokProcessMode = _grokModeCombo.SelectedIndex == 1 ? "triple" : "single";

        // Quad Compare settings
        config.LlmQuadModel1 = GetQuadModelId(_quadCombo1);
        config.LlmQuadModel2 = GetQuadModelId(_quadCombo2);
        config.LlmQuadModel3 = GetQuadModelId(_quadCombo3);
        config.LlmQuadModel4 = GetQuadModelId(_quadCombo4);
    }
}
