// [CustomSTT] Deepgram Nova-3 Medical WebSocket provider
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace MosaicTools.Services;

/// <summary>
/// Streams audio to Deepgram via WebSocket and receives real-time transcription.
/// Supports Nova-3 and Nova-3 Medical models.
/// </summary>
public class DeepgramProvider : ISttProvider
{
    private readonly string _apiKey;
    private readonly string _model;
    private readonly bool _autoPunctuate;
    private ClientWebSocket? _ws;
    private CancellationTokenSource? _receiveCts;
    private Task? _receiveTask;
    private volatile bool _connected;
    private readonly SemaphoreSlim _connectLock = new(1, 1); // Serialize StartSessionAsync calls

    public string Name => _model == "nova-3-medical" ? "Deepgram Nova-3 Medical" : "Deepgram Nova-3";
    public bool RequiresApiKey => true;
    public string? SignupUrl => "https://console.deepgram.com/signup";
    public bool IsConnected => _connected;
    public SttAudioFormat AudioFormat { get; } = new();

    public event Action<SttResult>? TranscriptionReceived;
    public event Action<string>? ErrorOccurred;
    public event Action<bool>? ConnectionStateChanged;

    public DeepgramProvider(string apiKey, string model = "nova-3-medical", bool autoPunctuate = false)
    {
        _apiKey = apiKey;
        _model = model;
        _autoPunctuate = autoPunctuate;
    }

    public async Task<bool> StartSessionAsync(CancellationToken ct = default)
    {
        if (_connected) return true; // Already connected (e.g., pre-connect succeeded)

        // Serialize concurrent calls (pre-connect racing with PTT press)
        await _connectLock.WaitAsync(ct);
        try
        {
            if (_connected) return true; // Re-check after acquiring lock

            _ws = new ClientWebSocket();
            _ws.Options.SetRequestHeader("Authorization", $"Token {_apiKey}");

            // Auto-punctuate: Deepgram adds punctuation based on speech patterns.
            // Manual punctuation: no server-side punctuation — spoken punctuation
            // words ("period", "comma", etc.) are converted client-side in ActionController,
            // which lets us keep "colon" as a word (body part) instead of punctuation.
            var punctParams = _autoPunctuate
                ? "&punctuate=true&smart_format=true"
                : "&punctuate=false";

            var uri = new Uri(
                $"wss://api.deepgram.com/v1/listen" +
                $"?model={_model}" +
                $"&encoding=linear16&sample_rate={AudioFormat.SampleRate}&channels={AudioFormat.Channels}" +
                punctParams +
                $"&interim_results=true&endpointing=300");

            await _ws.ConnectAsync(uri, ct);
            _connected = true;
            ConnectionStateChanged?.Invoke(true);

            _receiveCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            _receiveTask = Task.Run(() => ReceiveLoop(_receiveCts.Token));

            Logger.Trace($"DeepgramProvider: Connected ({_model})");
            return true;
        }
        catch (WebSocketException ex) when (ex.Message.Contains("401") || ex.Message.Contains("Unauthorized"))
        {
            Logger.Trace($"DeepgramProvider: Auth failed: {ex.Message}");
            ErrorOccurred?.Invoke("Invalid API key. Check Settings.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Trace($"DeepgramProvider: Connect failed: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection failed: {ex.Message}");
            return false;
        }
        finally
        {
            _connectLock.Release();
        }
    }

    public void SendAudio(byte[] pcmData, int offset, int count)
    {
        if (!_connected || _ws?.State != WebSocketState.Open) return;
        try
        {
            // Fire-and-forget binary frame (audio is time-sensitive, don't await)
            _ = _ws.SendAsync(new ArraySegment<byte>(pcmData, offset, count),
                WebSocketMessageType.Binary, true, CancellationToken.None);
        }
        catch (Exception ex)
        {
            Logger.Trace($"DeepgramProvider: SendAudio error: {ex.Message}");
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

        if (_ws?.State != WebSocketState.Open) return;
        try
        {
            // Finalize flushes pending audio
            var msg = Encoding.UTF8.GetBytes("{\"type\":\"Finalize\"}");
            await _ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
            Logger.Trace("DeepgramProvider: Sent Finalize");

            // Brief wait for Deepgram to flush finals (~50-100ms typical)
            await Task.Delay(300);

            // Disconnect — Deepgram sessions are cheap to reconnect
            await ShutdownAsync();
        }
        catch (Exception ex)
        {
            Logger.Trace($"DeepgramProvider: EndSession error: {ex.Message}");
        }
    }

    public async Task ShutdownAsync()
    {
        _connected = false;
        ConnectionStateChanged?.Invoke(false);

        // Grab references and null them to prevent concurrent disconnect races
        var ws = _ws; _ws = null;
        var cts = _receiveCts; _receiveCts = null;
        var task = _receiveTask; _receiveTask = null;

        try
        {
            if (ws?.State == WebSocketState.Open)
            {
                // Send CloseStream message
                var msg = Encoding.UTF8.GetBytes("{\"type\":\"CloseStream\"}");
                await ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
                await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", CancellationToken.None);
            }
        }
        catch { }

        cts?.Cancel();
        try { if (task != null) await task; } catch { }

        ws?.Dispose();
        cts?.Dispose();

        Logger.Trace("DeepgramProvider: Shutdown");
    }

    private async Task ReceiveLoop(CancellationToken ct)
    {
        var buffer = new byte[8192];
        var msgBuffer = new List<byte>();

        try
        {
            while (!ct.IsCancellationRequested && _ws?.State == WebSocketState.Open)
            {
                var result = await _ws.ReceiveAsync(new ArraySegment<byte>(buffer), ct);

                if (result.MessageType == WebSocketMessageType.Close)
                {
                    Logger.Trace("DeepgramProvider: Server closed connection");
                    break;
                }

                msgBuffer.AddRange(new ArraySegment<byte>(buffer, 0, result.Count));

                if (result.EndOfMessage)
                {
                    var json = Encoding.UTF8.GetString(msgBuffer.ToArray());
                    msgBuffer.Clear();
                    ParseResponse(json);
                }
            }
        }
        catch (OperationCanceledException) { }
        catch (WebSocketException ex)
        {
            Logger.Trace($"DeepgramProvider: WebSocket error: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection lost: {ex.Message}");
        }
        catch (Exception ex)
        {
            Logger.Trace($"DeepgramProvider: Receive error: {ex.Message}");
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

            // Check for error responses
            if (root.TryGetProperty("type", out var typeEl) && typeEl.GetString() == "Error")
            {
                var errMsg = root.TryGetProperty("description", out var desc) ? desc.GetString() : "Unknown error";
                Logger.Trace($"DeepgramProvider: Error response: {errMsg}");
                ErrorOccurred?.Invoke(errMsg ?? "Deepgram error");
                return;
            }

            // Only process Results type
            if (!root.TryGetProperty("type", out var t) || t.GetString() != "Results")
                return;

            var channel = root.GetProperty("channel");
            var alternatives = channel.GetProperty("alternatives");
            if (alternatives.GetArrayLength() == 0) return;

            var alt = alternatives[0];
            var transcript = alt.GetProperty("transcript").GetString() ?? "";
            if (string.IsNullOrEmpty(transcript)) return;

            var confidence = alt.TryGetProperty("confidence", out var conf) ? conf.GetSingle() : 0f;
            var isFinal = root.GetProperty("is_final").GetBoolean();
            var speechFinal = root.TryGetProperty("speech_final", out var sf) && sf.GetBoolean();
            var duration = root.TryGetProperty("duration", out var dur) ? dur.GetDouble() : 0;

            // Parse per-word data
            var words = Array.Empty<SttWord>();
            if (alt.TryGetProperty("words", out var wordsEl))
            {
                var wordList = new List<SttWord>();
                foreach (var w in wordsEl.EnumerateArray())
                {
                    wordList.Add(new SttWord(
                        Text: w.GetProperty("word").GetString() ?? "",
                        PunctuatedText: w.TryGetProperty("punctuated_word", out var pw) ? pw.GetString() ?? "" : w.GetProperty("word").GetString() ?? "",
                        Confidence: w.TryGetProperty("confidence", out var wc) ? wc.GetSingle() : 1f,
                        StartTime: w.TryGetProperty("start", out var ws) ? ws.GetDouble() : 0,
                        EndTime: w.TryGetProperty("end", out var we) ? we.GetDouble() : 0
                    ));
                }
                words = wordList.ToArray();
            }

            var result = new SttResult(transcript, words, confidence, isFinal, speechFinal, duration);
            TranscriptionReceived?.Invoke(result);
        }
        catch (Exception ex)
        {
            Logger.Trace($"DeepgramProvider: Parse error: {ex.Message}");
        }
    }

    public void Dispose()
    {
        try { ShutdownAsync().GetAwaiter().GetResult(); } catch { }
    }
}
