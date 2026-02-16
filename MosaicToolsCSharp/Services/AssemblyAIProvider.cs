// [CustomSTT] AssemblyAI Universal Streaming v3 WebSocket provider
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace MosaicTools.Services;

/// <summary>
/// Streams audio to AssemblyAI via WebSocket and receives real-time transcription.
/// Uses the Universal Streaming v3 API with raw PCM input.
/// </summary>
public class AssemblyAIProvider : ISttProvider
{
    private readonly string _apiKey;
    private readonly bool _autoPunctuate;
    private ClientWebSocket? _ws;
    private CancellationTokenSource? _receiveCts;
    private Task? _receiveTask;
    private volatile bool _connected;

    public string Name => "AssemblyAI";
    public bool RequiresApiKey => true;
    public string? SignupUrl => "https://www.assemblyai.com/dashboard/signup";
    public decimal CostPerMinute => 0.0025m; // $0.15/hr
    public bool IsConnected => _connected;

    public event Action<SttResult>? TranscriptionReceived;
    public event Action<string>? ErrorOccurred;
    public event Action<bool>? ConnectionStateChanged;

    public AssemblyAIProvider(string apiKey, bool autoPunctuate = false)
    {
        _apiKey = apiKey;
        _autoPunctuate = autoPunctuate;
    }

    public async Task<bool> ConnectAsync(CancellationToken ct = default)
    {
        try
        {
            _ws = new ClientWebSocket();
            _ws.Options.SetRequestHeader("Authorization", _apiKey);

            // format_turns=true gives punctuated/capitalized output on end-of-turn
            var uri = new Uri(
                "wss://streaming.assemblyai.com/v3/ws" +
                "?sample_rate=16000&encoding=pcm_s16le" +
                "&format_turns=true");

            await _ws.ConnectAsync(uri, ct);
            _connected = true;
            ConnectionStateChanged?.Invoke(true);

            _receiveCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            _receiveTask = Task.Run(() => ReceiveLoop(_receiveCts.Token));

            Logger.Trace("AssemblyAIProvider: Connected");
            return true;
        }
        catch (WebSocketException ex) when (ex.Message.Contains("401") || ex.Message.Contains("403") || ex.Message.Contains("Unauthorized"))
        {
            Logger.Trace($"AssemblyAIProvider: Auth failed: {ex.Message}");
            ErrorOccurred?.Invoke("Invalid API key. Check Settings.");
            return false;
        }
        catch (Exception ex)
        {
            Logger.Trace($"AssemblyAIProvider: Connect failed: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection failed: {ex.Message}");
            return false;
        }
    }

    public void SendAudio(byte[] pcmData, int offset, int count)
    {
        if (_ws?.State != WebSocketState.Open) return;
        try
        {
            _ = _ws.SendAsync(new ArraySegment<byte>(pcmData, offset, count),
                WebSocketMessageType.Binary, true, CancellationToken.None);
        }
        catch (Exception ex)
        {
            Logger.Trace($"AssemblyAIProvider: SendAudio error: {ex.Message}");
        }
    }

    public Task SendKeepAliveAsync() => Task.CompletedTask; // AssemblyAI doesn't need keepalives

    public async Task FinalizeAsync()
    {
        if (_ws?.State != WebSocketState.Open) return;
        try
        {
            // ForceEndpoint immediately ends the current turn
            var msg = Encoding.UTF8.GetBytes("{\"type\":\"ForceEndpoint\"}");
            await _ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
            Logger.Trace("AssemblyAIProvider: Sent ForceEndpoint");
        }
        catch (Exception ex)
        {
            Logger.Trace($"AssemblyAIProvider: ForceEndpoint error: {ex.Message}");
        }
    }

    public async Task DisconnectAsync()
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
                var msg = Encoding.UTF8.GetBytes("{\"type\":\"Terminate\"}");
                await ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
                await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "done", CancellationToken.None);
            }
        }
        catch { }

        cts?.Cancel();
        try { if (task != null) await task; } catch { }

        ws?.Dispose();
        cts?.Dispose();

        Logger.Trace("AssemblyAIProvider: Disconnected");
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
                    Logger.Trace("AssemblyAIProvider: Server closed connection");
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
            Logger.Trace($"AssemblyAIProvider: WebSocket error: {ex.Message}");
            ErrorOccurred?.Invoke($"Connection lost: {ex.Message}");
        }
        catch (Exception ex)
        {
            Logger.Trace($"AssemblyAIProvider: Receive error: {ex.Message}");
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

            if (type == "Begin")
            {
                Logger.Trace("AssemblyAIProvider: Session started");
                return;
            }

            if (type == "Termination")
            {
                if (root.TryGetProperty("audio_duration_seconds", out var dur))
                    Logger.Trace($"AssemblyAIProvider: Session ended, audio={dur.GetDouble():F1}s");
                return;
            }

            if (type != "Turn") return;

            var endOfTurn = root.GetProperty("end_of_turn").GetBoolean();
            var turnIsFormatted = root.GetProperty("turn_is_formatted").GetBoolean();
            var transcript = root.GetProperty("transcript").GetString() ?? "";

            // Skip unformatted end-of-turn — wait for the formatted version
            if (endOfTurn && !turnIsFormatted) return;

            // Parse word-level data
            var words = Array.Empty<SttWord>();
            if (root.TryGetProperty("words", out var wordsEl))
            {
                var wordList = new List<SttWord>();
                foreach (var w in wordsEl.EnumerateArray())
                {
                    var text = w.GetProperty("text").GetString() ?? "";
                    wordList.Add(new SttWord(
                        Text: text,
                        PunctuatedText: text,
                        Confidence: w.TryGetProperty("confidence", out var wc) ? wc.GetSingle() : 1f,
                        // AssemblyAI times are in milliseconds
                        StartTime: w.TryGetProperty("start", out var ws) ? ws.GetInt32() / 1000.0 : 0,
                        EndTime: w.TryGetProperty("end", out var we) ? we.GetInt32() / 1000.0 : 0
                    ));
                }
                words = wordList.ToArray();
            }

            // Build display text
            string displayTranscript;
            if (endOfTurn && turnIsFormatted)
            {
                // Final formatted result — use the properly punctuated/capitalized transcript
                displayTranscript = transcript;
            }
            else
            {
                // Interim: join all words (including non-final) for live preview
                displayTranscript = string.Join(" ", words.Select(w => w.Text));
            }

            if (string.IsNullOrEmpty(displayTranscript) && words.Length == 0) return;

            var confidence = words.Length > 0 ? words.Average(w => w.Confidence) : 0f;
            var duration = words.Length > 0 ? words.Max(w => w.EndTime) - words.Min(w => w.StartTime) : 0;
            var isFinal = endOfTurn && turnIsFormatted;

            var result = new SttResult(displayTranscript, words, confidence, isFinal, isFinal, duration);
            TranscriptionReceived?.Invoke(result);
        }
        catch (Exception ex)
        {
            Logger.Trace($"AssemblyAIProvider: Parse error: {ex.Message}");
        }
    }

    public void Dispose()
    {
        try { DisconnectAsync().GetAwaiter().GetResult(); } catch { }
    }
}
