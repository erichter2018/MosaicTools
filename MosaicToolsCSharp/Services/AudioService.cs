using System;
using System.IO;
using System.Media;
using System.Runtime.InteropServices;

namespace MosaicTools.Services;

/// <summary>
/// Audio services for beeps and notifications.
/// Matches Python's play_beep() function with volume control.
/// </summary>
public static class AudioService
{
    /// <summary>
    /// Play a beep with volume control (0.0 to 1.0).
    /// Uses generated WAV in memory for volume control.
    /// </summary>
    public static void PlayBeep(int frequency, int durationMs, double volume = 0.04)
    {
        try
        {
            Logger.Trace($"AudioService: Playing {frequency}Hz for {durationMs}ms (vol={volume:F2})");
            const int sampleRate = 44100;
            int samples = sampleRate * durationMs / 1000;
            
            // Generate sine wave samples
            var amplitude = (short)(32767 * Math.Clamp(volume, 0.0, 1.0));
            var buffer = new short[samples];
            
            for (int i = 0; i < samples; i++)
            {
                double t = (double)i / sampleRate;
                buffer[i] = (short)(amplitude * Math.Sin(2 * Math.PI * frequency * t));
            }
            
            // Create WAV in memory
            using var ms = new MemoryStream();
            using var writer = new BinaryWriter(ms);
            
            // WAV header
            int dataSize = samples * 2; // 16-bit = 2 bytes per sample
            int fileSize = 36 + dataSize;
            
            writer.Write("RIFF"u8.ToArray());
            writer.Write(fileSize);
            writer.Write("WAVE"u8.ToArray());
            writer.Write("fmt "u8.ToArray());
            writer.Write(16); // Subchunk1Size (PCM)
            writer.Write((short)1); // AudioFormat (PCM)
            writer.Write((short)1); // NumChannels (Mono)
            writer.Write(sampleRate); // SampleRate
            writer.Write(sampleRate * 2); // ByteRate
            writer.Write((short)2); // BlockAlign
            writer.Write((short)16); // BitsPerSample
            writer.Write("data"u8.ToArray());
            writer.Write(dataSize);
            
            // Write samples
            foreach (var sample in buffer)
            {
                writer.Write(sample);
            }
            
            // Play
            ms.Position = 0;
            using var player = new SoundPlayer(ms);
            player.PlaySync();
        }
        catch (Exception ex)
        {
            Logger.Trace($"Beep error: {ex.Message}");
            // Fallback to system beep
            Console.Beep(frequency, durationMs);
        }
    }
    
    /// <summary>
    /// Play beep asynchronously in a background thread.
    /// </summary>
    public static void PlayBeepAsync(int frequency, int durationMs, double volume = 0.04)
    {
        ThreadPool.QueueUserWorkItem(_ => PlayBeep(frequency, durationMs, volume));
    }
}
