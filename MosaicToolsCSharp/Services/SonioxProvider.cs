// [CustomSTT] Soniox v4 real-time WebSocket provider
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace MosaicTools.Services;

/// <summary>
/// Streams audio to Soniox via WebSocket and receives real-time transcription.
/// Uses the stt-rt-v4 model with raw PCM input and token-by-token delivery.
/// </summary>
public class SonioxProvider : ISttProvider
{
    private readonly string _apiKey;
    private readonly bool _autoPunctuate;
    private readonly string _keyterms;
    private ClientWebSocket? _ws;
    private CancellationTokenSource? _receiveCts;
    private Task? _receiveTask;
    private volatile bool _connected;
    private readonly SemaphoreSlim _connectLock = new(1, 1); // Serialize StartSessionAsync calls
    private TaskCompletionSource<bool>? _finalizeComplete;
    private string _accumulatedText = ""; // Growing transcript for interim display

    // Pending word accumulation: BPE tokens need to be merged into whole words.
    // A token with a leading space starts a new word; tokens without continue the previous.
    private string _pendingWordText = "";
    private float _pendingWordConfSum;
    private int _pendingWordTokenCount;
    private int _pendingWordStartMs;
    private int _pendingWordEndMs;

    public string Name => "Soniox";
    public bool RequiresApiKey => true;
    public string? SignupUrl => "https://console.soniox.com";
    public bool IsConnected => _connected;
    public SttAudioFormat AudioFormat { get; } = new();

    public event Action<SttResult>? TranscriptionReceived;
    public event Action<string>? ErrorOccurred;
    public event Action<bool>? ConnectionStateChanged;

    public SonioxProvider(string apiKey, bool autoPunctuate = false, string keyterms = "")
    {
        _apiKey = apiKey;
        _autoPunctuate = autoPunctuate;
        _keyterms = keyterms;
    }

    public async Task<bool> StartSessionAsync(CancellationToken ct = default)
    {
        if (_connected) return true; // Already connected (e.g., pre-connect succeeded)

        // Serialize concurrent calls (pre-connect racing with PTT press)
        await _connectLock.WaitAsync(ct);
        try
        {
            if (_connected) return true; // Re-check after acquiring lock

            // Clean up previous session
            var oldCts = _receiveCts; _receiveCts = null;
            var oldTask = _receiveTask; _receiveTask = null;
            var oldWs = _ws; _ws = null;

            oldCts?.Cancel();
            try { if (oldTask != null) await oldTask; } catch { }
            oldWs?.Dispose();
            oldCts?.Dispose();

            _ws = new ClientWebSocket();
            var uri = new Uri("wss://stt-rt.soniox.com/transcribe-websocket");
            await _ws.ConnectAsync(uri, ct);

            // Build context object for custom vocabulary
            object? context = null;
            if (!string.IsNullOrWhiteSpace(_keyterms))
            {
                var terms = _keyterms.Split(new[] { '\n', '\r', ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Trim()).Where(t => t.Length > 0)
                    .Take(500).ToArray();
                if (terms.Length > 0)
                    context = new { terms };
            }

            // Send initial configuration message
            var config = new Dictionary<string, object?>
            {
                ["api_key"] = _apiKey,
                ["model"] = "stt-rt-v4",
                ["audio_format"] = "pcm_s16le",
                ["sample_rate"] = AudioFormat.SampleRate,
                ["num_channels"] = AudioFormat.Channels,
                ["enable_endpoint_detection"] = true,
                ["max_endpoint_delay_ms"] = 300,
            };
            if (context != null)
                config["context"] = context;

            var configJson = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(config));
            await _ws.SendAsync(new ArraySegment<byte>(configJson), WebSocketMessageType.Text, true, ct);

            _connected = true;
            _accumulatedText = "";
            _pendingWordText = "";
            _pendingWordConfSum = 0;
            _pendingWordTokenCount = 0;
            ConnectionStateChanged?.Invoke(true);

            _receiveCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            _receiveTask = Task.Run(() => ReceiveLoop(_receiveCts.Token));

            Logger.Trace($"SonioxProvider: Connected, sent config with {(context != null ? "keyterms" : "no keyterms")}");
            return true;
        }
        catch (WebSocketException ex) when (ex.Message.Contains("401") || ex.Message.Contains("403"))
        {
            Logger.Trace($"SonioxProvider: Auth failed: {ex.Message}");
            ErrorOccurred?.Invoke("Invalid API key. Check Settings.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Trace($"SonioxProvider: Connect failed: {ex.Message}");
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
            _ = _ws.SendAsync(new ArraySegment<byte>(pcmData, offset, count),
                WebSocketMessageType.Binary, true, CancellationToken.None);
        }
        catch (Exception ex)
        {
            Logger.Trace($"SonioxProvider: SendAudio error: {ex.Message}");
        }
    }

    public async Task EndSessionAsync()
    {
        // Tail buffer: let last audio drain from SttService
        await Task.Delay(300);

        _connected = false;
        await Task.Delay(50);

        if (_ws?.State != WebSocketState.Open) return;
        try
        {
            // Send finalize to flush all pending tokens as final
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            _finalizeComplete = tcs;

            var msg = Encoding.UTF8.GetBytes("{\"type\":\"finalize\"}");
            await _ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
            Logger.Trace("SonioxProvider: Sent finalize, waiting for final tokens...");

            var completed = await Task.WhenAny(tcs.Task, Task.Delay(1000)) == tcs.Task;
            _finalizeComplete = null;

            Logger.Trace($"SonioxProvider: Finalize {(completed ? "got <fin>" : "timed out")}");

            // Send empty frame to signal end-of-audio, then disconnect
            await _ws.SendAsync(new ArraySegment<byte>(Array.Empty<byte>()), WebSocketMessageType.Binary, true, CancellationToken.None);

            // Wait briefly for finished message
            await Task.Delay(200);
            await ShutdownAsync();
        }
        catch (Exception ex)
        {
            _finalizeComplete = null;
            Logger.Trace($"SonioxProvider: EndSession error: {ex.Message}");
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
                // Send empty frame to signal end-of-audio
                await ws.SendAsync(new ArraySegment<byte>(Array.Empty<byte>()), WebSocketMessageType.Binary, true, CancellationToken.None);
                await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", CancellationToken.None);
            }
        }
        catch { }

        cts?.Cancel();
        try { if (task != null) await task; } catch { }

        ws?.Dispose();
        cts?.Dispose();

        Logger.Trace("SonioxProvider: Shutdown");
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
                    Logger.Trace("SonioxProvider: Server closed connection");
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
            Logger.Trace($"SonioxProvider: WebSocket error: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection lost: {ex.Message}");
        }
        catch (Exception ex)
        {
            Logger.Trace($"SonioxProvider: Receive error: {ex.Message}");
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

            // Check for errors
            if (root.TryGetProperty("error_code", out var errCode) && errCode.GetInt32() != 0)
            {
                var errMsg = root.TryGetProperty("error_message", out var em) ? em.GetString() : "Unknown error";
                Logger.Trace($"SonioxProvider: Error {errCode.GetInt32()}: {errMsg}");
                if (errCode.GetInt32() == 401)
                    ErrorOccurred?.Invoke("Invalid API key. Check Settings.");
                else
                    ErrorOccurred?.Invoke($"Soniox error: {errMsg}");
                return;
            }

            // Check for finished message
            if (root.TryGetProperty("finished", out var fin) && fin.GetBoolean())
            {
                Logger.Trace("SonioxProvider: Session finished");
                return;
            }

            if (!root.TryGetProperty("tokens", out var tokensEl)) return;

            // Process tokens: final tokens are confirmed words, non-final are provisional.
            // We emit finals for paste and build a growing display transcript for the overlay.
            var finalWords = new List<SttWord>();
            var finalTextParts = new List<string>();
            string nonFinalText = "";
            bool gotFinMarker = false;

            foreach (var tok in tokensEl.EnumerateArray())
            {
                var text = tok.TryGetProperty("text", out var t) ? t.GetString() ?? "" : "";
                var isFinalToken = tok.TryGetProperty("is_final", out var isf) && isf.GetBoolean();

                // Skip special marker tokens
                if (text == "<fin>" && isFinalToken)
                {
                    gotFinMarker = true;
                    continue;
                }
                if (text.StartsWith('<') && text.EndsWith('>'))
                    continue; // Skip <end>, <sil>, etc.

                var confidence = tok.TryGetProperty("confidence", out var c) ? c.GetSingle() : 1f;
                var startMs = tok.TryGetProperty("start_ms", out var sm) ? sm.GetInt32() : 0;
                var endMs = tok.TryGetProperty("end_ms", out var em2) ? em2.GetInt32() : 0;

                if (isFinalToken)
                {
                    // Accumulate final text into the persistent display buffer
                    _accumulatedText += text;
                    finalTextParts.Add(text);

                    // Merge BPE subword tokens into whole words.
                    // A token with a leading space starts a new word; without continues previous.
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        bool isWordStart = text.Length > 0 && text[0] == ' ';

                        if (isWordStart && _pendingWordText.Length > 0)
                        {
                            // Emit completed word
                            finalWords.Add(new SttWord(
                                Text: _pendingWordText,
                                PunctuatedText: _pendingWordText,
                                Confidence: _pendingWordConfSum / _pendingWordTokenCount,
                                StartTime: _pendingWordStartMs / 1000.0,
                                EndTime: _pendingWordEndMs / 1000.0
                            ));
                            _pendingWordText = "";
                        }

                        var trimmed = text.Trim();
                        if (_pendingWordText.Length == 0)
                        {
                            // Start new word
                            _pendingWordText = trimmed;
                            _pendingWordConfSum = confidence;
                            _pendingWordTokenCount = 1;
                            _pendingWordStartMs = startMs;
                            _pendingWordEndMs = endMs;
                        }
                        else
                        {
                            // Continue accumulating into current word
                            _pendingWordText += trimmed;
                            _pendingWordConfSum += confidence;
                            _pendingWordTokenCount++;
                            _pendingWordEndMs = endMs;
                        }
                    }
                }
                else
                {
                    // Non-final: provisional text for display only (will be replaced)
                    nonFinalText += text;
                }
            }

            // Flush pending word on finalize (session end) — otherwise carry across responses
            // so cross-message BPE splits like "perip"+"ancreatic" → "peripancreatic" merge properly
            if (gotFinMarker && _pendingWordText.Length > 0)
            {
                finalWords.Add(new SttWord(
                    Text: _pendingWordText,
                    PunctuatedText: _pendingWordText,
                    Confidence: _pendingWordConfSum / _pendingWordTokenCount,
                    StartTime: _pendingWordStartMs / 1000.0,
                    EndTime: _pendingWordEndMs / 1000.0
                ));
                _pendingWordText = "";
                _pendingWordConfSum = 0;
                _pendingWordTokenCount = 0;
            }

            if (gotFinMarker)
                _finalizeComplete?.TrySetResult(true);

            // Emit final words for paste
            if (finalWords.Count > 0)
            {
                var transcript = string.Join("", finalTextParts).Trim();
                var avgConf = finalWords.Average(w => w.Confidence);
                var duration = finalWords.Max(w => w.EndTime) - finalWords.Min(w => w.StartTime);
                TranscriptionReceived?.Invoke(new SttResult(transcript, finalWords.ToArray(), avgConf, true, false, duration));
            }

            // Always emit an interim for the display overlay showing the full growing transcript.
            // Without this, the overlay flickers between full text (interims) and short fragments (finals).
            var displayText = (_accumulatedText + nonFinalText).Trim();
            if (!string.IsNullOrEmpty(displayText))
                TranscriptionReceived?.Invoke(new SttResult(displayText, Array.Empty<SttWord>(), 0f, false, false, 0));
        }
        catch (Exception ex)
        {
            Logger.Trace($"SonioxProvider: Parse error: {ex.Message}");
        }
    }

    public void Dispose()
    {
        try { ShutdownAsync().GetAwaiter().GetResult(); } catch { }
    }
}
