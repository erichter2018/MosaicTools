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
    private TaskCompletionSource<bool>? _finalizeComplete; // Signals when ForceEndpoint response arrives
    private DateTime _lastDisconnect = DateTime.MinValue; // Rate-limit protection

    // Radiology keyterms to boost recognition accuracy (max 100 terms, 50 chars each)
    private static readonly string[] MedicalKeyterms =
    [
        // Chest
        "atelectasis", "pneumothorax", "pneumomediastinum",       // 3
        "cardiomegaly", "hepatomegaly", "splenomegaly",           // 6
        "lymphadenopathy", "bronchiectasis",                       // 8
        "hemothorax", "hemopneumothorax",                          // 10
        "retrocardiac", "costophrenic", "cardiophrenic",           // 13
        "hilar", "perihilar", "basilar", "bibasilar",              // 17
        "peribronchial", "parenchymal", "mediastinal",             // 20
        // Tubes/procedures
        "nasogastric", "endotracheal", "tracheostomy",             // 23
        "thoracentesis", "paracentesis",                           // 25
        // Hepatobiliary
        "cholecystectomy", "cholecystitis",                        // 27
        "cholelithiasis", "choledocholithiasis",                   // 29
        // Renal
        "hydronephrosis", "nephrolithiasis", "ureterolithiasis",   // 32
        "pyelonephritis",                                          // 33
        // GI
        "diverticulitis", "diverticulosis", "pancreatitis",        // 36
        "pneumoperitoneum", "pneumobilia",                         // 38
        "intussusception", "volvulus", "ileus",                    // 41
        "mesenteric", "omentum", "retroperitoneal",                // 44
        // MSK
        "opacification", "radiolucency", "lucency",                // 47
        "osseous", "periosteal", "osteophyte", "osteophytic",      // 51
        "spondylosis", "spondylolisthesis",                        // 53
        "atherosclerotic",                                         // 54
        "patulous", "tortuous", "ectatic",                         // 57
        "lobulated", "spiculated", "infundibulum",                 // 60
        // Gyn
        "adnexal", "endometrial", "myometrial",                   // 63
        // Spine
        "paraspinal", "foraminal", "neuroforaminal",               // 66
        "laminectomy", "discectomy",                               // 68
        "vertebroplasty", "kyphoplasty", "arthroplasty",           // 71
        // Trauma
        "subluxation", "diastasis", "comminuted", "nondisplaced",  // 75
        "avulsion",                                                // 76
        // Neuro
        "intraparenchymal", "subarachnoid",                        // 78
        "falcine", "tentorial", "uncal",                           // 81
        "vasogenic", "cytotoxic",                                  // 83
        "encephalomalacia", "gliosis", "demyelination",            // 86
        "periventricular", "pneumocephalus",                       // 88
        // Misc descriptors
        "Hounsfield", "radiopacity",                               // 90
    ];  // 90 terms

    public string Name => "AssemblyAI";
    public bool RequiresApiKey => true;
    public string? SignupUrl => "https://www.assemblyai.com/dashboard/signup";
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
            // Rate-limit protection: AssemblyAI rejects rapid reconnections
            var sinceLast = DateTime.UtcNow - _lastDisconnect;
            if (sinceLast.TotalMilliseconds < 1500)
            {
                var waitMs = 1500 - (int)sinceLast.TotalMilliseconds;
                Logger.Trace($"AssemblyAIProvider: Waiting {waitMs}ms before reconnect (rate-limit protection)");
                await Task.Delay(waitMs, ct);
            }

            _ws = new ClientWebSocket();
            _ws.Options.SetRequestHeader("Authorization", _apiKey);

            // No format_turns — we handle punctuation client-side via spoken punctuation
            var uri = new Uri(
                "wss://streaming.assemblyai.com/v3/ws" +
                "?sample_rate=16000&encoding=pcm_s16le");

            await _ws.ConnectAsync(uri, ct);
            _connected = true;
            ConnectionStateChanged?.Invoke(true);

            _receiveCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            _receiveTask = Task.Run(() => ReceiveLoop(_receiveCts.Token));

            // Send medical keyterms to boost recognition accuracy
            var config = new { type = "UpdateConfiguration", keyterms_prompt = MedicalKeyterms };
            var configJson = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(config));
            await _ws.SendAsync(new ArraySegment<byte>(configJson), WebSocketMessageType.Text, true, ct);

            Logger.Trace($"AssemblyAIProvider: Connected, sent {MedicalKeyterms.Length} keyterms");
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
            // Set up a waiter so we can block until AssemblyAI responds with the final formatted turn
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            _finalizeComplete = tcs;

            // ForceEndpoint immediately ends the current turn
            var msg = Encoding.UTF8.GetBytes("{\"type\":\"ForceEndpoint\"}");
            await _ws.SendAsync(new ArraySegment<byte>(msg), WebSocketMessageType.Text, true, CancellationToken.None);
            Logger.Trace("AssemblyAIProvider: Sent ForceEndpoint, waiting for final turn...");

            // Wait briefly for end-of-turn response. Without format_turns, the text
            // is usually already delivered as interim results before ForceEndpoint responds,
            // so this is mostly a courtesy flush. Keep timeout short.
            var completed = await Task.WhenAny(tcs.Task, Task.Delay(500)) == tcs.Task;
            _finalizeComplete = null;

            Logger.Trace($"AssemblyAIProvider: ForceEndpoint {(completed ? "got final turn" : "timed out")}");
        }
        catch (Exception ex)
        {
            _finalizeComplete = null;
            Logger.Trace($"AssemblyAIProvider: ForceEndpoint error: {ex.Message}");
        }
    }

    public async Task DisconnectAsync()
    {
        _connected = false;
        _lastDisconnect = DateTime.UtcNow;
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
            var transcript = root.GetProperty("transcript").GetString() ?? "";

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

            // Build display text — use transcript for final, word join for interim
            var displayTranscript = endOfTurn
                ? transcript
                : string.Join(" ", words.Select(w => w.Text));

            if (string.IsNullOrEmpty(displayTranscript) && words.Length == 0) return;

            var confidence = words.Length > 0 ? words.Average(w => w.Confidence) : 0f;
            var duration = words.Length > 0 ? words.Max(w => w.EndTime) - words.Min(w => w.StartTime) : 0;
            var isFinal = endOfTurn;

            var result = new SttResult(displayTranscript, words, confidence, isFinal, isFinal, duration);
            TranscriptionReceived?.Invoke(result);

            // Signal FinalizeAsync that the final turn has been delivered
            if (isFinal)
                _finalizeComplete?.TrySetResult(true);
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
