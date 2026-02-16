// [CustomSTT] Minimal WebM/Opus muxer for Corti streaming.
// Produces WebM byte chunks (header + clusters) matching what Chrome's
// MediaRecorder produces, since Corti natively consumes this format.
using Concentus;
using Concentus.Enums;
using System.Text;

namespace MosaicTools.Services;

/// <summary>
/// Encodes 16-bit mono PCM to Opus and muxes into WebM (Matroska) clusters.
/// Completed chunks are queued for the caller to send over WebSocket.
/// </summary>
internal sealed class WebmOpusMuxer : IDisposable
{
    private const int FRAME_SIZE_MS = 20;
    private const int FRAMES_PER_CLUSTER = 12; // ~240ms per cluster (Corti recommends 250ms)

    private readonly IOpusEncoder _encoder;
    private readonly int _sampleRate;
    private readonly int _frameSamples;
    private readonly List<byte[]> _readyChunks = new();

    // PCM frame accumulator
    private readonly short[] _frameBuffer;
    private int _frameIndex;

    // Cluster accumulator
    private readonly List<byte[]> _clusterPackets = new();
    private int _totalFrameCount;
    private bool _disposed;

    public WebmOpusMuxer(int sampleRate = 16000)
    {
        _sampleRate = sampleRate;
        _frameSamples = sampleRate * FRAME_SIZE_MS / 1000; // 320 for 16kHz
        _frameBuffer = new short[_frameSamples];

        _encoder = OpusCodecFactory.CreateEncoder(sampleRate, 1, OpusApplication.OPUS_APPLICATION_VOIP);
        _encoder.Bitrate = 24000;

        EmitHeader();
    }

    /// <summary>
    /// Feed raw 16-bit LE mono PCM bytes. WebM clusters are auto-flushed to the queue.
    /// </summary>
    public void WritePcm16(byte[] pcmBytes, int offset, int count)
    {
        if (_disposed) return;
        int end = offset + count;
        for (int i = offset; i + 1 < end; i += 2)
        {
            _frameBuffer[_frameIndex++] = (short)(pcmBytes[i] | (pcmBytes[i + 1] << 8));
            if (_frameIndex == _frameSamples)
            {
                EncodeFrame();
                _frameIndex = 0;
            }
        }
    }

    /// <summary>
    /// Pad current partial frame with silence, encode, and flush the cluster.
    /// </summary>
    public void ForceFlush()
    {
        if (_disposed) return;
        if (_frameIndex > 0)
        {
            Array.Clear(_frameBuffer, _frameIndex, _frameSamples - _frameIndex);
            _frameIndex = _frameSamples;
            EncodeFrame();
            _frameIndex = 0;
        }
        if (_clusterPackets.Count > 0)
            EmitCluster();
    }

    /// <summary>
    /// Take all completed WebM chunks. Returns empty array if none ready.
    /// </summary>
    public byte[][] TakePages()
    {
        if (_readyChunks.Count == 0) return Array.Empty<byte[]>();
        var result = _readyChunks.ToArray();
        _readyChunks.Clear();
        return result;
    }

    private void EncodeFrame()
    {
        var encBuf = new byte[4000];
        int packetSize = _encoder.Encode(_frameBuffer.AsSpan(), _frameSamples,
            encBuf.AsSpan(), encBuf.Length);

        var packet = new byte[packetSize];
        Buffer.BlockCopy(encBuf, 0, packet, 0, packetSize);
        _clusterPackets.Add(packet);
        _totalFrameCount++;

        if (_clusterPackets.Count >= FRAMES_PER_CLUSTER)
            EmitCluster();
    }

    #region WebM Structure

    private void EmitHeader()
    {
        var buf = new List<byte>(256);

        // === EBML Header ===
        var ebml = new List<byte>();
        WriteElement(ebml, 0x4286, EUInt(1));     // EBMLVersion
        WriteElement(ebml, 0x42F7, EUInt(1));     // EBMLReadVersion
        WriteElement(ebml, 0x4282, EUInt(4));     // EBMLMaxIDLength
        WriteElement(ebml, 0x4283, EUInt(8));     // EBMLMaxSizeLength
        WriteElement(ebml, 0x4287, Encoding.ASCII.GetBytes("webm")); // DocType
        WriteElement(ebml, 0x4285, EUInt(2));     // DocTypeVersion
        WriteElement(ebml, 0x4289, EUInt(2));     // DocTypeReadVersion

        WriteId(buf, 0x1A45DFA3);
        WriteSize(buf, ebml.Count);
        buf.AddRange(ebml);

        // === Segment (unknown size for streaming) ===
        WriteId(buf, 0x18538067);
        buf.AddRange(new byte[] { 0x01, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF });

        // === Info ===
        var info = new List<byte>();
        WriteElement(info, 0x2AD7B1, EUInt(1000000));  // TimecodeScale = 1ms
        WriteElement(info, 0x4D80, Encoding.UTF8.GetBytes("MosaicTools")); // MuxingApp
        WriteElement(info, 0x5741, Encoding.UTF8.GetBytes("MosaicTools")); // WritingApp

        WriteId(buf, 0x1549A966);
        WriteSize(buf, info.Count);
        buf.AddRange(info);

        // === Tracks ===
        var audio = new List<byte>();
        WriteElement(audio, 0xB5, EFloat(48000.0f)); // SamplingFrequency (Opus output is always 48kHz)
        WriteElement(audio, 0x9F, EUInt(1));          // Channels

        var entry = new List<byte>();
        WriteElement(entry, 0xD7, EUInt(1));        // TrackNumber
        WriteElement(entry, 0x73C5, EUInt(1));      // TrackUID
        WriteElement(entry, 0x9C, EUInt(0));        // FlagLacing = 0
        WriteElement(entry, 0x83, EUInt(2));        // TrackType = audio
        WriteElement(entry, 0x86, Encoding.ASCII.GetBytes("A_OPUS")); // CodecID
        WriteElement(entry, 0x63A2, BuildOpusHead()); // CodecPrivate
        WriteElement(entry, 0x56AA, EUInt(0));      // CodecDelay = 0ns
        WriteElement(entry, 0x56BB, EUInt(80000000)); // SeekPreRoll = 80ms in ns
        WriteId(entry, 0xE1);  // Audio sub-element
        WriteSize(entry, audio.Count);
        entry.AddRange(audio);

        var tracks = new List<byte>();
        WriteId(tracks, 0xAE);  // TrackEntry
        WriteSize(tracks, entry.Count);
        tracks.AddRange(entry);

        WriteId(buf, 0x1654AE6B);
        WriteSize(buf, tracks.Count);
        buf.AddRange(tracks);

        _readyChunks.Add(buf.ToArray());
    }

    private void EmitCluster()
    {
        if (_clusterPackets.Count == 0) return;

        int clusterStartFrame = _totalFrameCount - _clusterPackets.Count;
        uint clusterTimecodeMs = (uint)(clusterStartFrame * FRAME_SIZE_MS);

        // Build cluster content
        var content = new List<byte>(2048);
        WriteElement(content, 0xE7, EUInt(clusterTimecodeMs)); // Cluster Timecode

        for (int i = 0; i < _clusterPackets.Count; i++)
        {
            short relativeMs = (short)(i * FRAME_SIZE_MS);
            var opusData = _clusterPackets[i];

            // SimpleBlock: [TrackNum VINT][Timecode int16 BE][Flags][OpusData]
            var block = new byte[4 + opusData.Length];
            block[0] = 0x81; // Track 1 as VINT
            block[1] = (byte)(relativeMs >> 8);
            block[2] = (byte)(relativeMs & 0xFF);
            block[3] = 0x80; // Flags: keyframe
            Buffer.BlockCopy(opusData, 0, block, 4, opusData.Length);

            WriteElement(content, 0xA3, block); // SimpleBlock
        }

        // Wrap in Cluster element with known size
        var cluster = new List<byte>(content.Count + 12);
        WriteId(cluster, 0x1F43B675);
        WriteSize(cluster, content.Count);
        cluster.AddRange(content);

        _readyChunks.Add(cluster.ToArray());
        _clusterPackets.Clear();
    }

    private byte[] BuildOpusHead()
    {
        // RFC 7845 Section 5.1 — OpusHead identification header
        var head = new byte[19];
        Encoding.ASCII.GetBytes("OpusHead", 0, 8, head, 0);
        head[8] = 1;  // Version
        head[9] = 1;  // Channel count
        // Pre-skip (16-bit LE) at [10-11]: 0
        // Input sample rate (32-bit LE) at [12-15]
        head[12] = (byte)(_sampleRate & 0xFF);
        head[13] = (byte)((_sampleRate >> 8) & 0xFF);
        head[14] = (byte)((_sampleRate >> 16) & 0xFF);
        head[15] = (byte)((_sampleRate >> 24) & 0xFF);
        // Output gain (16-bit LE) at [16-17]: 0
        // Channel mapping family at [18]: 0
        return head;
    }

    #endregion

    #region EBML Encoding

    private static void WriteElement(List<byte> buf, uint id, byte[] data)
    {
        WriteId(buf, id);
        WriteSize(buf, data.Length);
        buf.AddRange(data);
    }

    private static void WriteId(List<byte> buf, uint id)
    {
        if (id <= 0xFF) { buf.Add((byte)id); }
        else if (id <= 0xFFFF) { buf.Add((byte)(id >> 8)); buf.Add((byte)id); }
        else if (id <= 0xFFFFFF) { buf.Add((byte)(id >> 16)); buf.Add((byte)(id >> 8)); buf.Add((byte)id); }
        else { buf.Add((byte)(id >> 24)); buf.Add((byte)(id >> 16)); buf.Add((byte)(id >> 8)); buf.Add((byte)id); }
    }

    private static void WriteSize(List<byte> buf, int size)
    {
        // EBML variable-length size encoding (VINT)
        if (size < 0x7F)
        {
            buf.Add((byte)(0x80 | size));
        }
        else if (size < 0x3FFF)
        {
            buf.Add((byte)(0x40 | (size >> 8)));
            buf.Add((byte)(size & 0xFF));
        }
        else if (size < 0x1FFFFF)
        {
            buf.Add((byte)(0x20 | (size >> 16)));
            buf.Add((byte)((size >> 8) & 0xFF));
            buf.Add((byte)(size & 0xFF));
        }
        else
        {
            buf.Add((byte)(0x10 | (size >> 24)));
            buf.Add((byte)((size >> 16) & 0xFF));
            buf.Add((byte)((size >> 8) & 0xFF));
            buf.Add((byte)(size & 0xFF));
        }
    }

    /// <summary>Encode unsigned integer as big-endian, minimum bytes.</summary>
    private static byte[] EUInt(ulong val)
    {
        if (val == 0) return new byte[] { 0 };
        int len = 0;
        ulong v = val;
        while (v > 0) { len++; v >>= 8; }
        var result = new byte[len];
        for (int i = len - 1; i >= 0; i--) { result[i] = (byte)(val & 0xFF); val >>= 8; }
        return result;
    }

    /// <summary>Encode IEEE 754 float32 as big-endian.</summary>
    private static byte[] EFloat(float val)
    {
        var bytes = BitConverter.GetBytes(val);
        if (BitConverter.IsLittleEndian) Array.Reverse(bytes);
        return bytes;
    }

    #endregion

    public void Dispose() => _disposed = true;
}
