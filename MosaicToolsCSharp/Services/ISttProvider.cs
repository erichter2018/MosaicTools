// [CustomSTT] Provider interface and data types for custom speech-to-text
namespace MosaicTools.Services;

/// <summary>
/// A single recognized word with confidence and timing.
/// </summary>
public record SttWord(string Text, string PunctuatedText, float Confidence, double StartTime, double EndTime);

/// <summary>
/// A transcription result (interim or final).
/// </summary>
public record SttResult(string Transcript, SttWord[] Words, float Confidence, bool IsFinal, bool SpeechFinal, double Duration);

/// <summary>
/// Provider-agnostic interface for streaming speech-to-text.
/// Implementations: DeepgramProvider (Deepgram Nova-3 Medical via WebSocket).
/// </summary>
public interface ISttProvider : IDisposable
{
    string Name { get; }
    bool RequiresApiKey { get; }
    string? SignupUrl { get; }
    bool IsConnected { get; }

    Task<bool> ConnectAsync(CancellationToken ct = default);
    void SendAudio(byte[] pcmData, int offset, int count);
    Task SendKeepAliveAsync();
    Task FinalizeAsync();
    Task DisconnectAsync();

    event Action<SttResult>? TranscriptionReceived;
    event Action<string>? ErrorOccurred;
    event Action<bool>? ConnectionStateChanged;
}
