// [CustomSTT] Speechmatics real-time WebSocket provider (medical model)
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace MosaicTools.Services;

/// <summary>
/// Streams audio to Speechmatics via WebSocket and receives real-time transcription.
/// Uses the medical domain model for clinical terminology accuracy.
/// </summary>
public class SpeechmaticsProvider : ISttProvider
{
    private readonly string _apiKey;
    private readonly string _region; // "us" or "eu"
    private readonly bool _autoPunctuate;
    private ClientWebSocket? _ws;
    private CancellationTokenSource? _receiveCts;
    private Task? _receiveTask;
    private volatile bool _connected;
    private int _seqNo; // Audio sequence number for EndOfStream
    private TaskCompletionSource<bool>? _endOfTranscript; // Signals when server finishes delivering finals

    public string Name => "Speechmatics Medical";
    public bool RequiresApiKey => true;
    public string? SignupUrl => "https://portal.speechmatics.com/signup";
    public bool IsConnected => _connected;
    public SttAudioFormat AudioFormat { get; } = new();

    public event Action<SttResult>? TranscriptionReceived;
    public event Action<string>? ErrorOccurred;
    public event Action<bool>? ConnectionStateChanged;

    public SpeechmaticsProvider(string apiKey, string region = "us", bool autoPunctuate = false)
    {
        _apiKey = apiKey;
        _region = region;
        _autoPunctuate = autoPunctuate;
    }

    public async Task<bool> StartSessionAsync(CancellationToken ct = default)
    {
        try
        {
            // Clean up previous session (EndOfStream kills the session, so reconnect
            // needs to tear down the old receive loop before starting a new one)
            var oldCts = _receiveCts; _receiveCts = null;
            var oldTask = _receiveTask; _receiveTask = null;
            var oldWs = _ws; _ws = null;
            _connected = false;

            oldCts?.Cancel();
            try { if (oldTask != null) await oldTask; } catch { }
            oldWs?.Dispose();
            oldCts?.Dispose();

            _ws = new ClientWebSocket();
            _ws.Options.SetRequestHeader("Authorization", $"Bearer {_apiKey}");

            var host = _region == "eu" ? "eu.rt.speechmatics.com" : "us.rt.speechmatics.com";
            var uri = new Uri($"wss://{host}/v2");

            await _ws.ConnectAsync(uri, ct);

            // Send StartRecognition config
            var txConfig = new Dictionary<string, object>
            {
                ["language"] = "en",
                ["domain"] = "medical",
                ["enable_partials"] = true,
                ["enable_entities"] = true,
                ["max_delay"] = 2.0,
                ["operating_point"] = "enhanced"
            };
            if (!_autoPunctuate)
                txConfig["punctuation_overrides"] = new { permitted_marks = Array.Empty<string>() };

            var startMsg = JsonSerializer.Serialize(new
            {
                message = "StartRecognition",
                audio_format = new
                {
                    type = "raw",
                    encoding = "pcm_s16le",
                    sample_rate = AudioFormat.SampleRate
                },
                transcription_config = txConfig
            });
            var startBytes = Encoding.UTF8.GetBytes(startMsg);
            await _ws.SendAsync(new ArraySegment<byte>(startBytes), WebSocketMessageType.Text, true, ct);

            _seqNo = 0;
            _connected = true;
            ConnectionStateChanged?.Invoke(true);

            _receiveCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            _receiveTask = Task.Run(() => ReceiveLoop(_receiveCts.Token));

            Logger.Trace($"SpeechmaticsProvider: Connected ({_region})");
            return true;
        }
        catch (WebSocketException ex) when (ex.Message.Contains("401") || ex.Message.Contains("403") || ex.Message.Contains("Unauthorized"))
        {
            Logger.Trace($"SpeechmaticsProvider: Auth failed: {ex.Message}");
            ErrorOccurred?.Invoke("Invalid API key. Check Settings.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Trace($"SpeechmaticsProvider: Connect failed: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection failed: {ex.Message}");
            return false;
        }
    }

    public void SendAudio(byte[] pcmData, int offset, int count)
    {
        if (!_connected || _ws?.State != WebSocketState.Open) return;
        try
        {
            // Send raw PCM as binary frame (AddAudio)
            _ = _ws.SendAsync(new ArraySegment<byte>(pcmData, offset, count),
                WebSocketMessageType.Binary, true, CancellationToken.None);
            _seqNo++;
        }
        catch (Exception ex)
        {
            Logger.Trace($"SpeechmaticsProvider: SendAudio error: {ex.Message}");
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

        // EndOfStream flushes all pending audio but terminates the session.
        // After this, the server sends remaining AddTranscript + EndOfTranscript
        // then closes the WebSocket.
        if (_ws?.State != WebSocketState.Open) return;
        try
        {
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            _endOfTranscript = tcs;

            var msg = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(new
            {
                message = "EndOfStream",
                last_seq_no = _seqNo
            }));
            await _ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
            Logger.Trace("SpeechmaticsProvider: Sent EndOfStream, waiting for finals...");

            // Wait for the server to deliver all finals (EndOfTranscript).
            // Speechmatics can take 100ms–2s depending on buffered audio.
            var completed = await Task.WhenAny(tcs.Task, Task.Delay(2000)) == tcs.Task;
            _endOfTranscript = null;

            Logger.Trace($"SpeechmaticsProvider: EndOfStream {(completed ? "got EndOfTranscript" : "timed out")}");
        }
        catch (Exception ex)
        {
            _endOfTranscript = null;
            Logger.Trace($"SpeechmaticsProvider: EndOfStream error: {ex.Message}");
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
            {
                // Send EndOfStream to cleanly terminate the session
                var msg = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(new
                {
                    message = "EndOfStream",
                    last_seq_no = _seqNo
                }));
                await ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
                await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", CancellationToken.None);
            }
        }
        catch { }

        cts?.Cancel();
        try { if (task != null) await task; } catch { }

        ws?.Dispose();
        cts?.Dispose();

        Logger.Trace("SpeechmaticsProvider: Shutdown");
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
                    Logger.Trace("SpeechmaticsProvider: Server closed connection");
                    break;
                }

                // Skip binary messages (AudioAdded confirmations are text, but just in case)
                if (result.MessageType == WebSocketMessageType.Binary)
                    continue;

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
            Logger.Trace($"SpeechmaticsProvider: WebSocket error: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection lost: {ex.Message}");
        }
        catch (Exception ex)
        {
            Logger.Trace($"SpeechmaticsProvider: Receive error: {ex.Message}");
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

            var msgType = root.GetProperty("message").GetString();

            if (msgType == "RecognitionStarted" || msgType == "AudioAdded")
                return;

            if (msgType == "EndOfTranscript")
            {
                _endOfTranscript?.TrySetResult(true);
                return;
            }

            if (msgType == "Error")
            {
                var reason = root.TryGetProperty("reason", out var r) ? r.GetString() : "Unknown error";
                Logger.Trace($"SpeechmaticsProvider: Error: {reason}");
                ErrorOccurred?.Invoke(reason ?? "Speechmatics error");
                return;
            }

            if (msgType == "Warning")
            {
                var reason = root.TryGetProperty("reason", out var r) ? r.GetString() : "Unknown warning";
                Logger.Trace($"SpeechmaticsProvider: Warning: {reason}");
                return;
            }

            bool isFinal = msgType == "AddTranscript";
            bool isPartial = msgType == "AddPartialTranscript";
            if (!isFinal && !isPartial) return;

            // Transcript and timing are inside the metadata object
            var transcript = "";
            double startTime = 0, endTime = 0;
            if (root.TryGetProperty("metadata", out var meta))
            {
                transcript = meta.TryGetProperty("transcript", out var tr) ? tr.GetString() ?? "" : "";
                startTime = meta.TryGetProperty("start_time", out var st) ? st.GetDouble() : 0;
                endTime = meta.TryGetProperty("end_time", out var et) ? et.GetDouble() : 0;
            }
            if (string.IsNullOrEmpty(transcript)) return;

            // Parse per-word results
            var words = Array.Empty<SttWord>();
            if (root.TryGetProperty("results", out var resultsEl))
            {
                var wordList = new List<SttWord>();
                foreach (var r in resultsEl.EnumerateArray())
                {
                    var type = r.TryGetProperty("type", out var tp) ? tp.GetString() : "";
                    if (type != "word") continue;

                    var wStart = r.TryGetProperty("start_time", out var ws) ? ws.GetDouble() : 0;
                    var wEnd = r.TryGetProperty("end_time", out var we) ? we.GetDouble() : 0;

                    var content = "";
                    float confidence = 1f;
                    if (r.TryGetProperty("alternatives", out var alts) && alts.GetArrayLength() > 0)
                    {
                        var alt = alts[0];
                        content = alt.TryGetProperty("content", out var c) ? c.GetString() ?? "" : "";
                        confidence = alt.TryGetProperty("confidence", out var cf) ? cf.GetSingle() : 1f;
                    }

                    if (!string.IsNullOrEmpty(content))
                    {
                        wordList.Add(new SttWord(
                            Text: content,
                            PunctuatedText: content,
                            Confidence: confidence,
                            StartTime: wStart,
                            EndTime: wEnd
                        ));
                    }
                }
                words = wordList.ToArray();
            }

            var avgConfidence = words.Length > 0 ? words.Average(w => w.Confidence) : 0f;
            var duration = endTime - startTime;

            var result = new SttResult(transcript, words, avgConfidence, isFinal, isFinal, duration);
            TranscriptionReceived?.Invoke(result);
        }
        catch (Exception ex)
        {
            Logger.Trace($"SpeechmaticsProvider: Parse error: {ex.Message}");
        }
    }

    public void Dispose()
    {
        try { ShutdownAsync().GetAwaiter().GetResult(); } catch { }
    }
}
