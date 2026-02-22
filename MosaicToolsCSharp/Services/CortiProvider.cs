// [CustomSTT] Corti Solo WebSocket streaming provider with WebM/Opus encoding
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace MosaicTools.Services;

/// <summary>
/// Streams audio to Corti via WebSocket using WebM/Opus encoding.
/// Uses a WebmOpusMuxer that produces clusters every ~240ms,
/// matching what Chrome's MediaRecorder produces natively.
/// </summary>
public class CortiProvider : ISttProvider
{
    private readonly string _clientId;
    private readonly string _clientSecret;
    private readonly string _environment; // "us" or "eu"
    private readonly bool _autoPunctuate;
    private ClientWebSocket? _ws;
    private CancellationTokenSource? _receiveCts;
    private Task? _receiveTask;
    private volatile bool _connected;
    private WebmOpusMuxer? _muxer;
    private byte[]? _bufferedHeader; // WebM header queued until first audio cluster

    // OAuth token caching
    private string? _accessToken;
    private DateTime _tokenExpiry = DateTime.MinValue;

    public string Name => "Corti Solo";
    public bool RequiresApiKey => true;
    public string? SignupUrl => "https://console.corti.app/signup";
    public bool IsConnected => _connected;
    public SttAudioFormat AudioFormat { get; } = new();

    public event Action<SttResult>? TranscriptionReceived;
    public event Action<string>? ErrorOccurred;
    public event Action<bool>? ConnectionStateChanged;

    public CortiProvider(string clientId, string clientSecret, string environment = "us", bool autoPunctuate = false)
    {
        _clientId = clientId;
        _clientSecret = clientSecret;
        _environment = environment;
        _autoPunctuate = autoPunctuate;
    }

    public async Task<bool> StartSessionAsync(CancellationToken ct = default)
    {
        try
        {
            // Step 1: Get OAuth token
            var token = await GetAccessTokenAsync(ct);
            if (token == null)
            {
                ErrorOccurred?.Invoke("Corti authentication failed. Check Client ID/Secret in Settings.");
                return false;
            }

            // Step 2: Connect WebSocket
            _ws = new ClientWebSocket();
            var uri = new Uri(
                $"wss://api.{_environment}.corti.app/audio-bridge/v2/transcribe" +
                $"?tenant-name=base&token=Bearer%20{Uri.EscapeDataString(token)}");

            await _ws.ConnectAsync(uri, ct);

            // Step 3: Send config (must be within 10 seconds)
            var config = JsonSerializer.Serialize(new
            {
                type = "config",
                configuration = new
                {
                    primaryLanguage = "en-US",
                    interimResults = true,
                    automaticPunctuation = _autoPunctuate,
                    spokenPunctuation = false
                }
            });
            Logger.Trace($"CortiProvider: Sending config: {config}");
            var configBytes = Encoding.UTF8.GetBytes(config);
            await _ws.SendAsync(new ArraySegment<byte>(configBytes), WebSocketMessageType.Text, true, ct);

            // Step 4: Wait for CONFIG_ACCEPTED
            var response = await ReceiveTextAsync(ct, TimeSpan.FromSeconds(10));
            Logger.Trace($"CortiProvider: Config response: {response}");
            if (response == null || !response.Contains("CONFIG_ACCEPTED"))
            {
                Logger.Trace($"CortiProvider: Config not accepted: {response}");
                ErrorOccurred?.Invoke("Corti rejected configuration.");
                try { await _ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "config rejected", CancellationToken.None); } catch { }
                return false;
            }

            // Step 5: Create WebM/Opus muxer — buffer header to send with first audio cluster.
            // Corti docs: "initial audio chunk must be sufficiently large to contain audio headers"
            _muxer = new WebmOpusMuxer(sampleRate: AudioFormat.SampleRate);
            var headerChunks = _muxer.TakePages();
            if (headerChunks.Length > 0)
            {
                using var ms = new MemoryStream();
                foreach (var chunk in headerChunks)
                    ms.Write(chunk, 0, chunk.Length);
                _bufferedHeader = ms.ToArray();
            }
            Logger.Trace($"CortiProvider: Buffered WebM header ({_bufferedHeader?.Length ?? 0} bytes)");

            _connected = true;
            ConnectionStateChanged?.Invoke(true);

            _receiveCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            _receiveTask = Task.Run(() => ReceiveLoop(_receiveCts.Token));

            Logger.Trace("CortiProvider: Connected");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CortiProvider: Connect failed: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection failed: {ex.Message}");
            return false;
        }
    }

    public void SendAudio(byte[] pcmData, int offset, int count)
    {
        var ws = _ws;
        if (!_connected || ws?.State != WebSocketState.Open || _muxer == null) return;

        try
        {
            // Encode PCM → Opus → WebM clusters
            _muxer.WritePcm16(pcmData, offset, count);
            var pages = _muxer.TakePages();
            if (pages.Length == 0) return;

            // Build a single buffer: [buffered header (if first send)] + [all clusters]
            var headers = _bufferedHeader;
            _bufferedHeader = null;

            int totalSize = (headers?.Length ?? 0);
            foreach (var p in pages) totalSize += p.Length;

            var combined = new byte[totalSize];
            int pos = 0;
            if (headers != null)
            {
                Buffer.BlockCopy(headers, 0, combined, 0, headers.Length);
                pos = headers.Length;
                Logger.Trace($"CortiProvider: Sending headers + first audio ({totalSize} bytes)");
            }
            foreach (var page in pages)
            {
                Buffer.BlockCopy(page, 0, combined, pos, page.Length);
                pos += page.Length;
            }

            // Single send per callback — no concurrent WebSocket writes
            _ = ws.SendAsync(new ArraySegment<byte>(combined),
                WebSocketMessageType.Binary, true, CancellationToken.None);
        }
        catch (Exception ex)
        {
            Logger.Trace($"CortiProvider: SendAudio error: {ex.Message}");
        }
    }

    public async Task EndSessionAsync()
    {
        // Tail buffer: audio keeps flowing from SttService while we wait,
        // capturing the last spoken word that may still be in the mic buffer.
        await Task.Delay(300);

        // Stop accepting new audio and let any in-flight fire-and-forget send drain.
        // ClientWebSocket only allows one outstanding SendAsync at a time.
        _connected = false;
        await Task.Delay(50);

        var ws = _ws;
        if (ws?.State != WebSocketState.Open) return;

        try
        {
            // Force-flush remaining audio in the WebM muxer
            if (_muxer != null)
            {
                _muxer.ForceFlush();
                var pages = _muxer.TakePages();
                if (pages.Length > 0)
                {
                    int totalSize = 0;
                    foreach (var p in pages) totalSize += p.Length;
                    var combined = new byte[totalSize];
                    int pos = 0;
                    foreach (var page in pages)
                    {
                        Buffer.BlockCopy(page, 0, combined, pos, page.Length);
                        pos += page.Length;
                    }
                    await ws.SendAsync(new ArraySegment<byte>(combined),
                        WebSocketMessageType.Binary, true, CancellationToken.None);
                    Logger.Trace($"CortiProvider: Flushed final audio ({totalSize} bytes)");
                }
            }

            // Tell Corti to process remaining audio
            var msg = Encoding.UTF8.GetBytes("{\"type\":\"flush\"}");
            await ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
            Logger.Trace("CortiProvider: Sent flush");
            // No disconnect — Corti charges by connection time, keep alive for next PTT press
        }
        catch (Exception ex)
        {
            Logger.Trace($"CortiProvider: EndSession error: {ex.Message}");
        }
    }

    public async Task ShutdownAsync()
    {
        _connected = false;
        ConnectionStateChanged?.Invoke(false);

        var ws = _ws; _ws = null;
        var cts = _receiveCts; _receiveCts = null;
        var task = _receiveTask; _receiveTask = null;
        var muxer = _muxer; _muxer = null;
        _bufferedHeader = null;

        muxer?.Dispose();

        try
        {
            if (ws?.State == WebSocketState.Open)
            {
                var msg = Encoding.UTF8.GetBytes("{\"type\":\"end\"}");
                await ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
                await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", CancellationToken.None);
            }
        }
        catch { }

        cts?.Cancel();
        try { if (task != null) await task; } catch { }

        ws?.Dispose();
        cts?.Dispose();

        Logger.Trace("CortiProvider: Shutdown");
    }

    private async Task<string?> GetAccessTokenAsync(CancellationToken ct)
    {
        if (_accessToken != null && DateTime.UtcNow < _tokenExpiry)
            return _accessToken;

        try
        {
            using var http = new HttpClient();
            var content = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["client_id"] = _clientId,
                ["client_secret"] = _clientSecret,
                ["grant_type"] = "client_credentials",
                ["scope"] = "openid"
            });

            var response = await http.PostAsync(
                $"https://auth.{_environment}.corti.app/realms/base/protocol/openid-connect/token",
                content, ct);

            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync(ct);
                Logger.Trace($"CortiProvider: Token request failed ({response.StatusCode}): {body}");
                return null;
            }

            var json = await response.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(json);
            _accessToken = doc.RootElement.GetProperty("access_token").GetString();
            var expiresIn = doc.RootElement.GetProperty("expires_in").GetInt32();
            _tokenExpiry = DateTime.UtcNow.AddSeconds(expiresIn - 30);

            Logger.Trace($"CortiProvider: Got access token (expires in {expiresIn}s)");
            return _accessToken;
        }
        catch (Exception ex)
        {
            Logger.Trace($"CortiProvider: Token error: {ex.Message}");
            return null;
        }
    }

    private async Task<string?> ReceiveTextAsync(CancellationToken ct, TimeSpan timeout)
    {
        var buffer = new byte[4096];
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeoutCts.CancelAfter(timeout);

        try
        {
            var result = await _ws!.ReceiveAsync(new ArraySegment<byte>(buffer), timeoutCts.Token);
            return Encoding.UTF8.GetString(buffer, 0, result.Count);
        }
        catch
        {
            return null;
        }
    }

    private async Task ReceiveLoop(CancellationToken ct)
    {
        var buffer = new byte[8192];
        var msgBuffer = new List<byte>();

        try
        {
            while (!ct.IsCancellationRequested)
            {
                var ws = _ws;
                if (ws?.State != WebSocketState.Open) break;

                var result = await ws.ReceiveAsync(new ArraySegment<byte>(buffer), ct);

                if (result.MessageType == WebSocketMessageType.Close)
                {
                    Logger.Trace("CortiProvider: Server closed connection");
                    break;
                }

                if (result.MessageType != WebSocketMessageType.Text) continue;

                msgBuffer.AddRange(new ArraySegment<byte>(buffer, 0, result.Count));

                if (result.EndOfMessage)
                {
                    var json = Encoding.UTF8.GetString(msgBuffer.ToArray());
                    msgBuffer.Clear();
                    Logger.Trace($"CortiProvider: Raw message: {json[..Math.Min(json.Length, 300)]}");
                    ParseResponse(json);
                }
            }
        }
        catch (OperationCanceledException) { }
        catch (WebSocketException ex)
        {
            Logger.Trace($"CortiProvider: WebSocket error: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection lost: {ex.Message}");
        }
        catch (Exception ex)
        {
            Logger.Trace($"CortiProvider: Receive error: {ex.Message}");
        }
        finally
        {
            _connected = false;
            ConnectionStateChanged?.Invoke(false);
        }
    }

    private void ParseResponse(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            var type = root.GetProperty("type").GetString();

            if (type == "error")
            {
                var errMsg = root.TryGetProperty("error", out var err)
                    ? (err.TryGetProperty("title", out var title) ? title.GetString() : "Unknown error")
                    : "Unknown error";
                Logger.Trace($"CortiProvider: Error: {errMsg}");
                ErrorOccurred?.Invoke(errMsg ?? "Corti error");
                return;
            }

            if (type == "usage")
            {
                if (root.TryGetProperty("credits", out var credits))
                    Logger.Trace($"CortiProvider: Credits used: {credits.GetDouble():F4}");
                return;
            }

            if (type == "ended" || type == "flushed")
            {
                Logger.Trace($"CortiProvider: Received {type}");
                return;
            }

            if (type != "transcript")
            {
                Logger.Trace($"CortiProvider: Unknown type={type}: {json[..Math.Min(json.Length, 200)]}");
                return;
            }

            var data = root.GetProperty("data");
            var text = data.GetProperty("text").GetString() ?? "";
            var isFinal = data.GetProperty("isFinal").GetBoolean();
            var start = data.TryGetProperty("start", out var s) ? s.GetDouble() : 0;
            var end = data.TryGetProperty("end", out var e) ? e.GetDouble() : 0;

            Logger.Trace($"CortiProvider: Transcript (final={isFinal}): \"{text[..Math.Min(text.Length, 80)]}\"");

            if (string.IsNullOrEmpty(text)) return;

            var words = new[] { new SttWord(text, text, 1f, start, end) };
            var duration = end - start;

            var result = new SttResult(text, words, 1f, isFinal, isFinal, duration);
            TranscriptionReceived?.Invoke(result);
        }
        catch (Exception ex)
        {
            Logger.Trace($"CortiProvider: Parse error: {ex.Message}");
        }
    }

    public void Dispose()
    {
        try { ShutdownAsync().GetAwaiter().GetResult(); } catch { }
    }
}
