using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace MosaicTools.Services;

/// <summary>
/// RadAI Impression API integration — generates impressions from report findings.
/// Self-contained service: delete this file and remove references to cleanly remove.
/// Integration points marked with [RadAI] in ActionController.cs and Configuration.cs.
/// </summary>
public class RadAiService : IDisposable
{
    private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(15) };
    private const string BaseUrl = "https://ioapi.radai-systems.com";
    private const string ConfigPath = @"C:\ProgramData\RadAI\Storage\api.config";

    private readonly string _id;
    private readonly string _key;
    private string? _token;
    private DateTime _tokenExpiry;

    private RadAiService(string id, string key)
    {
        _id = id;
        _key = key;
    }

    /// <summary>
    /// Try to create a RadAiService. Returns null if RadAI is not installed (no config file).
    /// </summary>
    public static RadAiService? TryCreate()
    {
        try
        {
            if (!File.Exists(ConfigPath))
            {
                Logger.Trace("RadAI: Config file not found, service disabled.");
                return null;
            }

            var doc = XDocument.Load(ConfigPath);
            var id = doc.Root?.Element("Id")?.Value;
            var key = doc.Root?.Element("Key")?.Value;

            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(key))
            {
                Logger.Trace("RadAI: Config file missing Id or Key.");
                return null;
            }

            Logger.Trace($"RadAI: Service initialized (Id={id[..Math.Min(6, id.Length)]}...)");
            return new RadAiService(id, key);
        }
        catch (Exception ex)
        {
            Logger.Trace($"RadAI: Failed to load config: {ex.Message}");
            return null;
        }
    }

    /// <summary>Whether RadAI config exists on this workstation.</summary>
    public static bool IsAvailable => File.Exists(ConfigPath);

    private async Task EnsureAuthenticatedAsync()
    {
        if (_token != null && DateTime.UtcNow < _tokenExpiry)
            return;

        var content = new FormUrlEncodedContent(new[]
        {
            new KeyValuePair<string, string>("username", _id),
            new KeyValuePair<string, string>("password", _key)
        });

        var resp = await _http.PostAsync($"{BaseUrl}/v2/auth", content);
        resp.EnsureSuccessStatusCode();

        var json = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
        _token = json.RootElement.GetProperty("access_token").GetString();
        var expiresIn = json.RootElement.GetProperty("expires_in").GetInt32();
        _tokenExpiry = DateTime.UtcNow.AddSeconds(expiresIn - 30);

        Logger.Trace("RadAI: Authenticated successfully.");
    }

    /// <summary>
    /// Call the FTOI endpoint to generate an impression from report text.
    /// </summary>
    public async Task<RadAiResult> GetImpressionAsync(
        string reportText,
        string? procedureDesc = null,
        string? gender = null)
    {
        try
        {
            await EnsureAuthenticatedAsync();

            // Extract modality from procedure description (first word)
            string modality = "CT";
            if (!string.IsNullOrEmpty(procedureDesc))
            {
                var firstWord = procedureDesc.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                    .FirstOrDefault()?.ToUpperInvariant() ?? "CT";
                if (new[] { "CT", "MR", "MRI", "US", "XR", "CR", "NM", "PET" }.Contains(firstWord))
                    modality = firstWord == "MRI" ? "MR" : firstWord;
            }

            // Map gender format: "Male"/"Female" -> "M"/"F"
            var genderCode = gender?.ToUpperInvariant() switch
            {
                "MALE" => "M",
                "FEMALE" => "F",
                _ => ""
            };

            var requestBody = new
            {
                radiologist = Environment.UserName,
                report_id = Guid.NewGuid().ToString(),
                ftoi_trigger = "hot_key",
                impression_template = "",
                report = reportText,
                order_data = new
                {
                    age = 0,
                    gender = genderCode,
                    other = new[]
                    {
                        new { procedures = modality, procedure_desc = procedureDesc ?? "" }
                    }
                },
                integration_type = "fluency"
            };

            var jsonString = JsonSerializer.Serialize(requestBody);

            using var reqMsg = new HttpRequestMessage(HttpMethod.Post,
                $"{BaseUrl}/v2/ftoi/inference/unstructured_report");
            reqMsg.Content = new StringContent(jsonString, Encoding.UTF8, "application/json");
            reqMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _token);

            var resp = await _http.SendAsync(reqMsg);

            // Handle 401 — re-auth and retry once
            if (resp.StatusCode == System.Net.HttpStatusCode.Unauthorized)
            {
                _token = null;
                await EnsureAuthenticatedAsync();

                using var retryMsg = new HttpRequestMessage(HttpMethod.Post,
                    $"{BaseUrl}/v2/ftoi/inference/unstructured_report");
                retryMsg.Content = new StringContent(jsonString, Encoding.UTF8, "application/json");
                retryMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _token);
                resp = await _http.SendAsync(retryMsg);
            }

            resp.EnsureSuccessStatusCode();

            var result = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
            var items = result.RootElement.GetProperty("impression_items");
            var impression = string.Join("\n", items.EnumerateArray().Select(i => i.GetString()));

            var itemsArray = items.EnumerateArray().Select(i => i.GetString() ?? "").ToArray();

            Logger.Trace($"RadAI: Impression generated ({impression.Length} chars, {itemsArray.Length} items)");
            return new RadAiResult { Success = true, Impression = impression, ImpressionItems = itemsArray };
        }
        catch (Exception ex)
        {
            Logger.Trace($"RadAI: Error: {ex.Message}");
            return new RadAiResult { Success = false, Error = ex.Message };
        }
    }

    /// <summary>
    /// Show a non-modal popup with the RadAI impression result. Call on UI thread.
    /// Returns the Form reference for tracking (close on accession change).
    /// </summary>
    public static Form ShowResultPopup(string[] impressionItems, Action onInsertClicked, Configuration config)
    {
        // Format display text: number items if 2+, each on its own line
        var displayText = impressionItems.Length >= 2
            ? string.Join("\r\n", impressionItems.Select((item, i) => $"{i + 1}. {item}"))
            : impressionItems.Length == 1 ? impressionItems[0] : "";

        var form = new Form();
        var textBox = new TextBox();
        var insertBtn = new Button();
        var copyBtn = new Button();
        var closeBtn = new Button();

        form.Text = "RadAI Impression";
        form.ClientSize = new Size(500, 300);
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
        form.TopMost = true;
        form.BackColor = Color.FromArgb(45, 45, 45);
        form.ForeColor = Color.White;
        form.KeyPreview = true;
        form.KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) form.Close(); };

        // Restore saved position or center on screen
        if (config.RadAiPopupX >= 0 && config.RadAiPopupY >= 0)
        {
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(config.RadAiPopupX, config.RadAiPopupY);
        }
        else
        {
            form.StartPosition = FormStartPosition.CenterScreen;
        }

        // Save position on move
        form.LocationChanged += (s, e) =>
        {
            if (form.WindowState == FormWindowState.Normal)
            {
                config.RadAiPopupX = form.Location.X;
                config.RadAiPopupY = form.Location.Y;
                config.Save();
            }
        };

        textBox.Multiline = true;
        textBox.ReadOnly = true;
        textBox.ScrollBars = ScrollBars.Vertical;
        textBox.WordWrap = true;
        textBox.Text = displayText;
        textBox.SelectionStart = 0;
        textBox.SelectionLength = 0;
        textBox.SetBounds(12, 12, 476, 240);
        textBox.BackColor = Color.FromArgb(60, 60, 60);
        textBox.ForeColor = Color.White;
        textBox.Font = new Font("Segoe UI", 10);
        textBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

        var hintLabel = new Label();
        hintLabel.Text = "Press RadAI button or Insert to apply";
        hintLabel.SetBounds(12, 266, 230, 20);
        hintLabel.ForeColor = Color.FromArgb(140, 140, 140);
        hintLabel.Font = new Font("Segoe UI", 8);
        hintLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

        insertBtn.Text = "Insert";
        insertBtn.SetBounds(247, 262, 75, 28);
        insertBtn.FlatStyle = FlatStyle.Flat;
        insertBtn.BackColor = Color.FromArgb(70, 70, 70);
        insertBtn.ForeColor = Color.White;
        insertBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        insertBtn.Click += (s, e) => { onInsertClicked(); };

        copyBtn.Text = "Copy";
        copyBtn.SetBounds(330, 262, 75, 28);
        copyBtn.FlatStyle = FlatStyle.Flat;
        copyBtn.BackColor = Color.FromArgb(70, 70, 70);
        copyBtn.ForeColor = Color.White;
        copyBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        copyBtn.Click += (s, e) =>
        {
            try { Clipboard.SetText(displayText); copyBtn.Text = "Copied!"; }
            catch { }
        };

        closeBtn.Text = "Close";
        closeBtn.SetBounds(413, 262, 75, 28);
        closeBtn.FlatStyle = FlatStyle.Flat;
        closeBtn.BackColor = Color.FromArgb(70, 70, 70);
        closeBtn.ForeColor = Color.White;
        closeBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        closeBtn.Click += (s, e) => form.Close();

        form.Controls.AddRange(new Control[] { textBox, hintLabel, insertBtn, copyBtn, closeBtn });
        form.FormClosed += (s, e) => form.Dispose();

        form.Show();
        return form;
    }

    public void Dispose() { }
}

/// <summary>
/// Result from a RadAI FTOI API call.
/// </summary>
public record RadAiResult
{
    public bool Success { get; init; }
    public string Impression { get; init; } = "";
    public string[] ImpressionItems { get; init; } = Array.Empty<string>();
    public string Error { get; init; } = "";
}
