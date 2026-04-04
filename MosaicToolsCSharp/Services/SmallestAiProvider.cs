// [CustomSTT] Smallest.ai Pulse STT real-time WebSocket provider
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace MosaicTools.Services;

/// <summary>
/// Streams audio to Smallest.ai Pulse STT via WebSocket and receives real-time transcription.
/// Sends raw PCM binary frames, receives JSON with word timestamps.
/// Sub-70ms latency, 39 languages, keyword boosting.
/// </summary>
public class SmallestAiProvider : ISttProvider
{
    private readonly string _apiKey;
    private readonly bool _autoPunctuate;
    private readonly string _keyterms;
    private ClientWebSocket? _ws;
    private CancellationTokenSource? _receiveCts;
    private Task? _receiveTask;
    private volatile bool _connected;
    private readonly SemaphoreSlim _connectLock = new(1, 1);
    private TaskCompletionSource<bool>? _finalizeComplete;

    public string Name => "Smallest.ai Pulse";
    public bool RequiresApiKey => true;
    public string? SignupUrl => "https://smallest.ai";
    public bool IsConnected => _connected;
    public SttAudioFormat AudioFormat { get; } = new();

    public event Action<SttResult>? TranscriptionReceived;
    public event Action<string>? ErrorOccurred;
    public event Action<bool>? ConnectionStateChanged;

    public SmallestAiProvider(string apiKey, bool autoPunctuate = false, string keyterms = "")
    {
        _apiKey = apiKey;
        _autoPunctuate = autoPunctuate;
        _keyterms = keyterms;
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
            _ws.Options.SetRequestHeader("Authorization", $"Bearer {_apiKey}");

            // Build URI with query parameters
            var uriBuilder = new StringBuilder(
                "wss://api.smallest.ai/waves/v1/pulse/get_text" +
                $"?language=en&encoding=linear16&sample_rate={AudioFormat.SampleRate}" +
                "&word_timestamps=true&full_transcript=false");

            // Add keyword boosting (comma-separated, max 100 terms)
            if (!string.IsNullOrWhiteSpace(_keyterms))
            {
                var terms = _keyterms.Split(new[] { '\n', '\r', ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Trim()).Where(t => t.Length > 0)
                    .Take(100).ToArray();
                if (terms.Length > 0)
                    uriBuilder.Append($"&keywords={Uri.EscapeDataString(string.Join(",", terms))}");
            }

            var uri = new Uri(uriBuilder.ToString());
            await _ws.ConnectAsync(uri, ct);

            _connected = true;
            ConnectionStateChanged?.Invoke(true);

            _receiveCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            _receiveTask = Task.Run(() => ReceiveLoop(_receiveCts.Token));

            Logger.Trace("SmallestAiProvider: Connected (Pulse STT)");
            return true;
        }
        catch (WebSocketException ex) when (ex.Message.Contains("401") || ex.Message.Contains("403") || ex.Message.Contains("Unauthorized"))
        {
            Logger.Trace($"SmallestAiProvider: Auth failed: {ex.Message}");
            ErrorOccurred?.Invoke("Invalid API key. Check Settings.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Trace($"SmallestAiProvider: Connect failed: {ex.Message}");
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
            // Pulse accepts raw binary PCM frames
            _ = _ws.SendAsync(new ArraySegment<byte>(pcmData, offset, count),
                WebSocketMessageType.Binary, true, CancellationToken.None);
        }
        catch (Exception ex)
        {
            Logger.Trace($"SmallestAiProvider: SendAudio error: {ex.Message}");
        }
    }

    public async Task EndSessionAsync()
    {
        // Tail buffer
        await Task.Delay(300);

        _connected = false;
        await Task.Delay(50);

        if (_ws?.State != WebSocketState.Open) return;
        try
        {
            // Send finalize to flush pending audio
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            _finalizeComplete = tcs;

            var msg = Encoding.UTF8.GetBytes("{\"type\":\"finalize\"}");
            await _ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
            Logger.Trace("SmallestAiProvider: Sent finalize, waiting for is_last...");

            // Wait for final response (is_last=true)
            var completed = await Task.WhenAny(tcs.Task, Task.Delay(1500)) == tcs.Task;
            _finalizeComplete = null;

            Logger.Trace($"SmallestAiProvider: Finalize {(completed ? "got is_last" : "timed out")}");
            await ShutdownAsync();
        }
        catch (Exception ex)
        {
            _finalizeComplete = null;
            Logger.Trace($"SmallestAiProvider: EndSession error: {ex.Message}");
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

        Logger.Trace("SmallestAiProvider: Shutdown");
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
                    Logger.Trace("SmallestAiProvider: Server closed connection");
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
            Logger.Trace($"SmallestAiProvider: WebSocket error: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection lost: {ex.Message}");
        }
        catch (Exception ex)
        {
            Logger.Trace($"SmallestAiProvider: Receive error: {ex.Message}");
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

            // Check for error
            if (root.TryGetProperty("status", out var status) && status.GetString() != "success")
            {
                var errMsg = root.TryGetProperty("error", out var err) ? err.GetString() : "Unknown error";
                Logger.Trace($"SmallestAiProvider: Error: {errMsg}");
                ErrorOccurred?.Invoke(errMsg ?? "Smallest.ai error");
                return;
            }

            var transcript = root.TryGetProperty("transcript", out var tr) ? tr.GetString() ?? "" : "";
            var isFinal = root.TryGetProperty("is_final", out var fin) && fin.GetBoolean();
            var isLast = root.TryGetProperty("is_last", out var last) && last.GetBoolean();

            if (isLast)
            {
                _finalizeComplete?.TrySetResult(true);
                // Still process the transcript if it has content
            }

            if (string.IsNullOrEmpty(transcript)) return;

            // Parse word-level data
            var words = Array.Empty<SttWord>();
            if (root.TryGetProperty("words", out var wordsEl))
            {
                var wordList = new List<SttWord>();
                foreach (var w in wordsEl.EnumerateArray())
                {
                    var wText = w.TryGetProperty("word", out var wt) ? wt.GetString() ?? "" : "";
                    var start = w.TryGetProperty("start", out var ws) ? ws.GetDouble() : 0;
                    var end = w.TryGetProperty("end", out var we) ? we.GetDouble() : 0;
                    var confidence = w.TryGetProperty("confidence", out var wc) ? wc.GetSingle() : 1f;

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
            var displayTranscript = transcript;
            if (!_autoPunctuate && isFinal)
            {
                displayTranscript = displayTranscript.TrimEnd('.', ',', '!', '?', ';', ':');
                if (string.IsNullOrWhiteSpace(displayTranscript)) return;
            }

            var avgConf = words.Length > 0 ? words.Average(w => w.Confidence) : 0f;
            var duration = words.Length > 0 ? words.Max(w => w.EndTime) - words.Min(w => w.StartTime) : 0;

            TranscriptionReceived?.Invoke(new SttResult(displayTranscript, words, avgConf, isFinal, isFinal, duration));
        }
        catch (Exception ex)
        {
            Logger.Trace($"SmallestAiProvider: Parse error: {ex.Message}");
        }
    }

    public void Dispose()
    {
        try { ShutdownAsync().GetAwaiter().GetResult(); } catch { }
    }
}
