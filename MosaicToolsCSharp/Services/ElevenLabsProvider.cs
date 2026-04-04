// [CustomSTT] ElevenLabs Scribe v2 Realtime WebSocket provider
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace MosaicTools.Services;

/// <summary>
/// Streams audio to ElevenLabs Scribe v2 Realtime via WebSocket.
/// Audio is sent as base64-encoded PCM in JSON frames. Uses VAD-based commit strategy
/// with word-level timestamps. Sub-150ms latency, 90+ languages.
/// </summary>
public class ElevenLabsProvider : ISttProvider
{
    private readonly string _apiKey;
    private readonly bool _autoPunctuate;
    private ClientWebSocket? _ws;
    private CancellationTokenSource? _receiveCts;
    private Task? _receiveTask;
    private volatile bool _connected;
    private readonly SemaphoreSlim _connectLock = new(1, 1);

    public string Name => "ElevenLabs Scribe v2";
    public bool RequiresApiKey => true;
    public string? SignupUrl => "https://elevenlabs.io/app/sign-up";
    public bool IsConnected => _connected;
    public SttAudioFormat AudioFormat { get; } = new();

    public event Action<SttResult>? TranscriptionReceived;
    public event Action<string>? ErrorOccurred;
    public event Action<bool>? ConnectionStateChanged;

    public ElevenLabsProvider(string apiKey, bool autoPunctuate = false)
    {
        _apiKey = apiKey;
        _autoPunctuate = autoPunctuate;
    }

    public async Task<bool> StartSessionAsync(CancellationToken ct = default)
    {
        if (_connected) return true;

        await _connectLock.WaitAsync(ct);
        try
        {
            if (_connected) return true;

            // Clean up previous session
            var oldCts = _receiveCts; _receiveCts = null;
            var oldTask = _receiveTask; _receiveTask = null;
            var oldWs = _ws; _ws = null;

            oldCts?.Cancel();
            try { if (oldTask != null) await oldTask; } catch { }
            oldWs?.Dispose();
            oldCts?.Dispose();

            _ws = new ClientWebSocket();
            _ws.Options.SetRequestHeader("xi-api-key", _apiKey);

            // Scribe v2 Realtime: VAD commit, PCM 16kHz, word timestamps
            var uri = new Uri(
                "wss://api.elevenlabs.io/v1/speech-to-text/realtime" +
                "?model_id=scribe_v2" +
                "&audio_format=pcm_16000" +
                "&language_code=en" +
                "&commit_strategy=vad" +
                "&include_timestamps=true" +
                "&vad_silence_threshold_secs=1.0" +
                "&vad_threshold=0.4" +
                "&min_speech_duration_ms=100" +
                "&min_silence_duration_ms=100");

            await _ws.ConnectAsync(uri, ct);
            _connected = true;
            ConnectionStateChanged?.Invoke(true);

            _receiveCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            _receiveTask = Task.Run(() => ReceiveLoop(_receiveCts.Token));

            Logger.Trace("ElevenLabsProvider: Connected (Scribe v2 Realtime)");
            return true;
        }
        catch (WebSocketException ex) when (ex.Message.Contains("401") || ex.Message.Contains("403") || ex.Message.Contains("Unauthorized"))
        {
            Logger.Trace($"ElevenLabsProvider: Auth failed: {ex.Message}");
            ErrorOccurred?.Invoke("Invalid API key. Check Settings.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Trace($"ElevenLabsProvider: Connect failed: {ex.Message}");
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
            // ElevenLabs requires base64-encoded audio in JSON frames
            var base64 = Convert.ToBase64String(pcmData, offset, count);
            var msg = JsonSerializer.Serialize(new
            {
                message_type = "input_audio_chunk",
                audio_base_64 = base64
            });
            var bytes = Encoding.UTF8.GetBytes(msg);
            _ = _ws.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None);
        }
        catch (Exception ex)
        {
            Logger.Trace($"ElevenLabsProvider: SendAudio error: {ex.Message}");
        }
    }

    public async Task EndSessionAsync()
    {
        // Tail buffer: let last audio drain
        await Task.Delay(300);

        _connected = false;
        await Task.Delay(50);

        if (_ws?.State != WebSocketState.Open) return;
        try
        {
            // Send a final commit to flush any pending audio
            var msg = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(new
            {
                message_type = "input_audio_chunk",
                audio_base_64 = "",
                commit = true
            }));
            await _ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
            Logger.Trace("ElevenLabsProvider: Sent final commit");

            // Wait for final transcript delivery
            await Task.Delay(300);
            await ShutdownAsync();
        }
        catch (Exception ex)
        {
            Logger.Trace($"ElevenLabsProvider: EndSession error: {ex.Message}");
        }
    }

    public async Task ShutdownAsync()
    {
        _connected = false;
        ConnectionStateChanged?.Invoke(false);

        var ws = _ws; _ws = null;
        var cts = _receiveCts; _receiveCts = null;
        var task = _receiveTask; _receiveTask = null;

        try
        {
            if (ws?.State == WebSocketState.Open)
                await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", CancellationToken.None);
        }
        catch { }

        cts?.Cancel();
        try { if (task != null) await task; } catch { }

        ws?.Dispose();
        cts?.Dispose();

        Logger.Trace("ElevenLabsProvider: Shutdown");
    }

    private async Task ReceiveLoop(CancellationToken ct)
    {
        var buffer = new byte[16384];
        var msgBuffer = new List<byte>();

        try
        {
            while (!ct.IsCancellationRequested && _ws?.State == WebSocketState.Open)
            {
                var result = await _ws.ReceiveAsync(new ArraySegment<byte>(buffer), ct);

                if (result.MessageType == WebSocketMessageType.Close)
                {
                    Logger.Trace("ElevenLabsProvider: Server closed connection");
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
            Logger.Trace($"ElevenLabsProvider: WebSocket error: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection lost: {ex.Message}");
        }
        catch (Exception ex)
        {
            Logger.Trace($"ElevenLabsProvider: Receive error: {ex.Message}");
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

            var msgType = root.TryGetProperty("message_type", out var mt) ? mt.GetString() : "";

            // Handle error responses
            if (msgType == "error" || msgType == "auth_error" || msgType == "quota_exceeded" ||
                msgType == "rate_limited" || msgType == "resource_exhausted")
            {
                var errMsg = root.TryGetProperty("error", out var err) ? err.GetString() : msgType;
                Logger.Trace($"ElevenLabsProvider: Error: {errMsg}");
                if (msgType == "auth_error")
                    ErrorOccurred?.Invoke("Invalid API key. Check Settings.");
                else
                    ErrorOccurred?.Invoke(errMsg ?? "ElevenLabs error");
                return;
            }

            if (msgType == "session_started")
            {
                Logger.Trace("ElevenLabsProvider: Session started");
                return;
            }

            // Partial transcript (interim)
            if (msgType == "partial_transcript")
            {
                var text = root.TryGetProperty("text", out var t) ? t.GetString() ?? "" : "";
                if (!string.IsNullOrEmpty(text))
                    TranscriptionReceived?.Invoke(new SttResult(text, Array.Empty<SttWord>(), 0f, false, false, 0));
                return;
            }

            // Committed transcript with timestamps (final)
            if (msgType == "committed_transcript_with_timestamps")
            {
                var text = root.TryGetProperty("text", out var t) ? t.GetString() ?? "" : "";
                if (string.IsNullOrEmpty(text)) return;

                var words = Array.Empty<SttWord>();
                if (root.TryGetProperty("words", out var wordsEl))
                {
                    var wordList = new List<SttWord>();
                    foreach (var w in wordsEl.EnumerateArray())
                    {
                        var wType = w.TryGetProperty("type", out var wt) ? wt.GetString() : "";
                        if (wType != "word") continue;

                        var wText = w.TryGetProperty("text", out var wtt) ? wtt.GetString() ?? "" : "";
                        var start = w.TryGetProperty("start", out var ws) ? ws.GetDouble() : 0;
                        var end = w.TryGetProperty("end", out var we) ? we.GetDouble() : 0;
                        // logprob is log-probability; convert to 0-1 confidence approximation
                        var logprob = w.TryGetProperty("logprob", out var lp) ? lp.GetDouble() : 0;
                        var confidence = (float)Math.Min(1.0, Math.Max(0.0, Math.Exp(logprob)));

                        if (!string.IsNullOrEmpty(wText))
                        {
                            wordList.Add(new SttWord(
                                Text: wText,
                                PunctuatedText: wText,
                                Confidence: confidence,
                                StartTime: start,
                                EndTime: end
                            ));
                        }
                    }
                    words = wordList.ToArray();
                }

                // Strip punctuation if auto-punctuate is off
                var transcript = text;
                if (!_autoPunctuate)
                    transcript = StripPunctuation(transcript);

                if (string.IsNullOrWhiteSpace(transcript)) return;

                var avgConf = words.Length > 0 ? words.Average(w => w.Confidence) : 0f;
                var duration = words.Length > 0 ? words.Max(w => w.EndTime) - words.Min(w => w.StartTime) : 0;

                TranscriptionReceived?.Invoke(new SttResult(transcript, words, avgConf, true, true, duration));
                return;
            }

            // Committed transcript without timestamps (fallback)
            if (msgType == "committed_transcript")
            {
                var text = root.TryGetProperty("text", out var t) ? t.GetString() ?? "" : "";
                if (string.IsNullOrEmpty(text)) return;

                if (!_autoPunctuate)
                    text = StripPunctuation(text);

                if (string.IsNullOrWhiteSpace(text)) return;

                TranscriptionReceived?.Invoke(new SttResult(text, Array.Empty<SttWord>(), 0f, true, true, 0));
            }
        }
        catch (Exception ex)
        {
            Logger.Trace($"ElevenLabsProvider: Parse error: {ex.Message}");
        }
    }

    private static string StripPunctuation(string text)
    {
        // Remove sentence-ending punctuation that the model adds
        return text.TrimEnd('.', ',', '!', '?', ';', ':');
    }

    public void Dispose()
    {
        try { ShutdownAsync().GetAwaiter().GetResult(); } catch { }
    }
}
