using System;
using System.Drawing;
using System.IO;
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
    private readonly CheckBox _clarioCdpEnabledCheck;
    private readonly TextBox _clarioCdpUrlBox;
    private readonly Button _clarioCdpSetupButton;
    private readonly Label _clarioCdpStatusLabel;
    private readonly Label _mosaicCdpDot;
    private readonly Label _mosaicCdpLabel;
    private readonly Label _clarioCdpDot;
    private readonly Label _clarioCdpLabel;

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
        _clarioCdpEnabledCheck.Checked = config.ClarioCdpEnabled;
        _clarioCdpUrlBox.Text = config.ClarioCdpUrl;

        UpdateNetworkSettingsStates();
        UpdateCdpSettingsStates();
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
        config.ClarioCdpEnabled = _clarioCdpEnabledCheck.Checked;
        var url = _clarioCdpUrlBox.Text.Trim();
        if (!string.IsNullOrEmpty(url))
            config.ClarioCdpUrl = url;
    }
}
