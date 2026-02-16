// [CustomSTT] Self-contained Ogg/Opus streaming writer with frequent page flushing.
// Forked from Concentus.OggFile's OpusOggWriteStream — page flush threshold reduced
// from 248 segments (~5s) to 12 segments (~240ms) for real-time streaming to Corti.
using Concentus;
using Concentus.Enums;
using IOpusEncoder = Concentus.IOpusEncoder;
using System.Text;

namespace MosaicTools.Services;

/// <summary>
/// Encodes 16-bit mono PCM to Opus and packages into Ogg pages, flushing every
/// ~240ms. Completed pages are queued for the caller to send over WebSocket.
/// </summary>
internal sealed class StreamingOggOpusWriter : IDisposable
{
    private const int FRAME_SIZE_MS = 20;
    private const int PAGE_FLUSH_THRESHOLD = 12; // ~240ms (12 × 20ms frames)

    private readonly IOpusEncoder _encoder;
    private readonly int _sampleRate;
    private readonly int _frameSamples; // samples per 20ms frame
    private readonly List<byte[]> _readyPages = new();

    // PCM frame accumulator
    private readonly short[] _frameBuffer;
    private int _frameIndex;

    // Ogg page buffers
    private readonly byte[] _header = new byte[400];
    private readonly byte[] _payload = new byte[65536];
    private int _headerIndex;
    private int _payloadIndex;
    private int _pageCounter;
    private readonly int _streamId;
    private long _granulePosition;
    private byte _lacingTableCount;
    private bool _disposed;

    // Fixed positions in Ogg page header
    private const int PAGE_FLAGS_POS = 5;
    private const int GRANULE_POS = 6;
    private const int CHECKSUM_POS = 22;
    private const int SEGMENT_COUNT_POS = 26;

    public StreamingOggOpusWriter(int sampleRate = 16000)
    {
        _sampleRate = sampleRate;
        _frameSamples = sampleRate * FRAME_SIZE_MS / 1000; // 320 for 16kHz
        _frameBuffer = new short[_frameSamples];
        _streamId = Random.Shared.Next();

        _encoder = OpusCodecFactory.CreateEncoder(sampleRate, 1, OpusApplication.OPUS_APPLICATION_VOIP);
        _encoder.Bitrate = 24000;

        // Write Ogg header pages (OpusHead + OpusTags) — queued for caller
        BeginNewPage();
        WriteOpusHeadPage();
        WriteOpusTagsPage();
    }

    /// <summary>
    /// Feed raw 16-bit LE mono PCM bytes. Ogg pages are auto-flushed to the queue.
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
    /// Pad current partial frame with silence, encode it, and flush the page.
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
        if (_lacingTableCount > 0)
            FinalizePage();
    }

    /// <summary>
    /// Take all completed Ogg pages. Returns empty array if none ready.
    /// </summary>
    public byte[][] TakePages()
    {
        if (_readyPages.Count == 0) return Array.Empty<byte[]>();
        var result = _readyPages.ToArray();
        _readyPages.Clear();
        return result;
    }

    private void EncodeFrame()
    {
        var encBuf = new byte[4000];
        int packetSize = _encoder.Encode(_frameBuffer.AsSpan(), _frameSamples,
            encBuf.AsSpan(), encBuf.Length);

        Buffer.BlockCopy(encBuf, 0, _payload, _payloadIndex, packetSize);
        _payloadIndex += packetSize;

        // Granule position is always in 48kHz units per Opus spec
        _granulePosition += FRAME_SIZE_MS * 48;

        // Lacing table: split packet into 255-byte segments
        int remaining = packetSize;
        while (remaining >= 255)
        {
            _header[_headerIndex++] = 0xFF;
            _lacingTableCount++;
            remaining -= 255;
        }
        _header[_headerIndex++] = (byte)remaining;
        _lacingTableCount++;

        if (_lacingTableCount >= PAGE_FLUSH_THRESHOLD)
            FinalizePage();
    }

    #region Ogg Page Construction

    private void BeginNewPage()
    {
        _headerIndex = 0;
        _payloadIndex = 0;
        _lacingTableCount = 0;

        _headerIndex += WriteStr("OggS", _header, _headerIndex);
        _header[_headerIndex++] = 0; // stream version
        _header[_headerIndex++] = 0; // flags placeholder
        _headerIndex += WriteI64(_granulePosition, _header, _headerIndex);
        _headerIndex += WriteI32(_streamId, _header, _headerIndex);
        _headerIndex += WriteI32(_pageCounter, _header, _headerIndex);
        _headerIndex += 4; // checksum placeholder (zeros)
        _header[_headerIndex++] = 0; // segment count placeholder

        _pageCounter++;
    }

    private void FinalizePage()
    {
        if (_disposed || _lacingTableCount == 0) return;

        _header[SEGMENT_COUNT_POS] = _lacingTableCount;
        WriteI64(_granulePosition, _header, GRANULE_POS);

        // Zero the checksum field before CRC computation
        _header[CHECKSUM_POS] = 0;
        _header[CHECKSUM_POS + 1] = 0;
        _header[CHECKSUM_POS + 2] = 0;
        _header[CHECKSUM_POS + 3] = 0;

        uint crc = ComputeOggCrc(_header, _headerIndex, _payload, _payloadIndex);
        WriteU32(crc, _header, CHECKSUM_POS);

        var page = new byte[_headerIndex + _payloadIndex];
        Buffer.BlockCopy(_header, 0, page, 0, _headerIndex);
        Buffer.BlockCopy(_payload, 0, page, _headerIndex, _payloadIndex);
        _readyPages.Add(page);

        BeginNewPage();
    }

    private void WriteOpusHeadPage()
    {
        _payloadIndex += WriteStr("OpusHead", _payload, _payloadIndex);
        _payload[_payloadIndex++] = 1; // version
        _payload[_payloadIndex++] = 1; // mono
        _payloadIndex += WriteI16(0, _payload, _payloadIndex); // pre-skip
        _payloadIndex += WriteI32(_sampleRate, _payload, _payloadIndex); // input sample rate
        _payloadIndex += WriteI16(0, _payload, _payloadIndex); // output gain
        _payload[_payloadIndex++] = 0; // channel mapping family 0

        // Single segment (payload < 255 bytes)
        _header[_headerIndex++] = (byte)_payloadIndex;
        _lacingTableCount++;
        _header[PAGE_FLAGS_POS] = 0x02; // BeginningOfStream
        FinalizePage();
    }

    private void WriteOpusTagsPage()
    {
        _payloadIndex += WriteStr("OpusTags", _payload, _payloadIndex);
        // Vendor string
        var vendor = Encoding.UTF8.GetBytes("MosaicTools");
        _payloadIndex += WriteI32(vendor.Length, _payload, _payloadIndex);
        Buffer.BlockCopy(vendor, 0, _payload, _payloadIndex, vendor.Length);
        _payloadIndex += vendor.Length;
        // Zero additional tags
        _payloadIndex += WriteI32(0, _payload, _payloadIndex);

        // Lacing for tags payload
        int segSize = _payloadIndex;
        while (segSize >= 255)
        {
            _header[_headerIndex++] = 255;
            _lacingTableCount++;
            segSize -= 255;
        }
        _header[_headerIndex++] = (byte)segSize;
        _lacingTableCount++;
        FinalizePage();
    }

    #endregion

    #region Little-Endian Byte Helpers

    private static int WriteStr(string val, byte[] buf, int off)
        => Encoding.ASCII.GetBytes(val, 0, val.Length, buf, off);

    private static int WriteI16(int val, byte[] buf, int off)
    {
        buf[off] = (byte)(val & 0xFF);
        buf[off + 1] = (byte)((val >> 8) & 0xFF);
        return 2;
    }

    private static int WriteI32(int val, byte[] buf, int off)
    {
        buf[off] = (byte)(val & 0xFF);
        buf[off + 1] = (byte)((val >> 8) & 0xFF);
        buf[off + 2] = (byte)((val >> 16) & 0xFF);
        buf[off + 3] = (byte)((val >> 24) & 0xFF);
        return 4;
    }

    private static void WriteU32(uint val, byte[] buf, int off)
    {
        buf[off] = (byte)(val & 0xFF);
        buf[off + 1] = (byte)((val >> 8) & 0xFF);
        buf[off + 2] = (byte)((val >> 16) & 0xFF);
        buf[off + 3] = (byte)((val >> 24) & 0xFF);
    }

    private static int WriteI64(long val, byte[] buf, int off)
    {
        for (int i = 0; i < 8; i++)
            buf[off + i] = (byte)((val >> (i * 8)) & 0xFF);
        return 8;
    }

    #endregion

    #region Ogg CRC32

    private static readonly uint[] CrcTable = BuildCrcTable();

    private static uint[] BuildCrcTable()
    {
        const uint poly = 0x04c11db7;
        var table = new uint[256];
        for (uint i = 0; i < 256; i++)
        {
            uint s = i << 24;
            for (int j = 0; j < 8; j++)
                s = (s << 1) ^ (s >= (1U << 31) ? poly : 0);
            table[i] = s;
        }
        return table;
    }

    private static uint ComputeOggCrc(byte[] header, int headerLen, byte[] payload, int payloadLen)
    {
        uint crc = 0;
        for (int i = 0; i < headerLen; i++)
            crc = (crc << 8) ^ CrcTable[header[i] ^ (crc >> 24)];
        for (int i = 0; i < payloadLen; i++)
            crc = (crc << 8) ^ CrcTable[payload[i] ^ (crc >> 24)];
        return crc;
    }

    #endregion

    public void Dispose()
    {
        _disposed = true;
    }
}
