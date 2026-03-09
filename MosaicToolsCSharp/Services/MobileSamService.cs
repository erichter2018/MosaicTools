using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;

namespace MosaicTools.Services;

/// <summary>
/// MobileSAM click-to-segment using ONNX Runtime with DirectML GPU acceleration.
/// Encoder runs once per image (~150ms GPU), decoder per click (~30ms).
/// Models auto-downloaded on first use to %LOCALAPPDATA%\MosaicTools\Models\.
/// </summary>
public class MobileSamService : IDisposable
{
    private InferenceSession? _encoder;
    private InferenceSession? _decoder;
    private DenseTensor<float>? _imageEmbedding;
    private int _origWidth, _origHeight;
    private float _scale;
    private bool _initialized;

    private static readonly string ModelsDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MosaicTools", "Models");

    private const string ModelZipUrl = "https://huggingface.co/vietanhdev/segment-anything-onnx-models/resolve/main/mobile_sam_20230629.zip";

    private static string EncoderPath => Path.Combine(ModelsDir, "mobile_sam.encoder.onnx");
    private static string DecoderPath => Path.Combine(ModelsDir, "sam_vit_h_4b8939.decoder.onnx");

    public bool EnsureInitialized()
    {
        if (_initialized) return true;

        try
        {
            if (!File.Exists(EncoderPath) || !File.Exists(DecoderPath))
            {
                Logger.Trace("MobileSAM: models not found, downloading...");
                DownloadModelsSync();
            }

            if (!File.Exists(EncoderPath) || !File.Exists(DecoderPath))
            {
                Logger.Trace("MobileSAM: model files missing after download");
                return false;
            }

            var opts = CreateSessionOptions();
            _encoder = new InferenceSession(EncoderPath, opts);
            _decoder = new InferenceSession(DecoderPath, opts);
            _initialized = true;
            Logger.Trace("MobileSAM: initialized successfully");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Trace($"MobileSAM: init failed: {ex.Message}");
            return false;
        }
    }

    private static SessionOptions CreateSessionOptions()
    {
        var opts = new SessionOptions();
        try
        {
            opts.AppendExecutionProvider_DML(0);
            Logger.Trace("MobileSAM: using DirectML GPU");
        }
        catch
        {
            Logger.Trace("MobileSAM: using CPU fallback");
        }
        return opts;
    }

    private static void DownloadModelsSync()
    {
        Directory.CreateDirectory(ModelsDir);
        using var http = new HttpClient();
        http.Timeout = TimeSpan.FromMinutes(5);

        Logger.Trace("MobileSAM: downloading model zip (~37MB)...");
        var zipData = http.GetByteArrayAsync(ModelZipUrl).GetAwaiter().GetResult();
        Logger.Trace($"MobileSAM: downloaded {zipData.Length / 1024 / 1024}MB");

        using var zipStream = new MemoryStream(zipData);
        using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read);
        foreach (var entry in archive.Entries)
        {
            if (entry.Name.EndsWith(".onnx", StringComparison.OrdinalIgnoreCase))
            {
                var destPath = Path.Combine(ModelsDir, entry.Name);
                entry.ExtractToFile(destPath, overwrite: true);
                Logger.Trace($"MobileSAM: extracted {entry.Name} ({entry.Length / 1024}KB)");
            }
        }
    }

    /// <summary>
    /// Run encoder on a bitmap. Call once per image, then call GetMask() for each click.
    /// The vietanhdev encoder expects [H,W,3] float32 RGB (handles resize/normalize internally).
    /// </summary>
    public void SetImage(Bitmap bitmap)
    {
        if (_encoder == null) throw new InvalidOperationException("Not initialized");

        _origWidth = bitmap.Width;
        _origHeight = bitmap.Height;
        _scale = 1024f / Math.Max(_origWidth, _origHeight);

        using var rgb = new Bitmap(_origWidth, _origHeight, PixelFormat.Format24bppRgb);
        using (var g = Graphics.FromImage(rgb))
            g.DrawImage(bitmap, 0, 0, _origWidth, _origHeight);

        var inputData = new float[_origHeight * _origWidth * 3];
        var bmpData = rgb.LockBits(
            new Rectangle(0, 0, _origWidth, _origHeight),
            ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

        unsafe
        {
            byte* ptr = (byte*)bmpData.Scan0;
            for (int y = 0; y < _origHeight; y++)
            {
                byte* row = ptr + y * bmpData.Stride;
                for (int x = 0; x < _origWidth; x++)
                {
                    int idx = (y * _origWidth + x) * 3;
                    inputData[idx + 0] = row[x * 3 + 2]; // R
                    inputData[idx + 1] = row[x * 3 + 1]; // G
                    inputData[idx + 2] = row[x * 3];     // B
                }
            }
        }
        rgb.UnlockBits(bmpData);

        var inputTensor = new DenseTensor<float>(inputData, new[] { _origHeight, _origWidth, 3 });
        string inputName = _encoder.InputMetadata.Keys.First();

        using var results = _encoder.Run(new List<NamedOnnxValue>
        {
            NamedOnnxValue.CreateFromTensor(inputName, inputTensor)
        });

        _imageEmbedding = results.First().AsTensor<float>().ToDenseTensor();
        Logger.Trace($"MobileSAM: encoded {_origWidth}x{_origHeight}, scale={_scale:F3}");
    }

    /// <summary>
    /// Get segmentation mask using a foreground point click only (pass 1 of two-pass strategy).
    /// </summary>
    public bool[,]? GetMask(int x, int y)
    {
        return RunDecoder(x, y, null);
    }

    /// <summary>
    /// Get segmentation mask using a box prompt + foreground point (pass 2 of two-pass strategy).
    /// Box coords are in original image space: (x1, y1, x2, y2).
    /// Research shows box prompts massively outperform point prompts for medical images.
    /// </summary>
    public bool[,]? GetMaskWithBox(int x, int y, int boxX1, int boxY1, int boxX2, int boxY2)
    {
        return RunDecoder(x, y, (boxX1, boxY1, boxX2, boxY2));
    }

    private bool[,]? RunDecoder(int x, int y, (int x1, int y1, int x2, int y2)? box)
    {
        if (_decoder == null || _imageEmbedding == null)
            return null;

        // Transform click point to model space (1024-based)
        float px = x * _scale;
        float py = y * _scale;

        float[] coords;
        float[] labels;
        int numPoints;

        if (box != null)
        {
            // Box prompt: foreground point + box top-left (label=2) + box bottom-right (label=3)
            float bx1 = box.Value.x1 * _scale;
            float by1 = box.Value.y1 * _scale;
            float bx2 = box.Value.x2 * _scale;
            float by2 = box.Value.y2 * _scale;
            numPoints = 3;
            coords = new float[] { px, py, bx1, by1, bx2, by2 };
            labels = new float[] { 1f, 2f, 3f }; // 1=fg, 2=box-TL, 3=box-BR
            Logger.Trace($"MobileSAM: box prompt [{box.Value.x1},{box.Value.y1}]-[{box.Value.x2},{box.Value.y2}] + point ({x},{y})");
        }
        else
        {
            // Point-only prompt: foreground click + 4 background corners
            int margin = 5;
            numPoints = 5;
            coords = new float[]
            {
                px, py,
                margin * _scale, margin * _scale,
                (_origWidth - margin) * _scale, margin * _scale,
                margin * _scale, (_origHeight - margin) * _scale,
                (_origWidth - margin) * _scale, (_origHeight - margin) * _scale
            };
            labels = new float[] { 1f, 0f, 0f, 0f, 0f };
        }

        var embeddingName = _decoder.InputMetadata.Keys.First(k => k.Contains("embed"));

        var inputs = new List<NamedOnnxValue>
        {
            NamedOnnxValue.CreateFromTensor(embeddingName, _imageEmbedding),
            NamedOnnxValue.CreateFromTensor("point_coords",
                new DenseTensor<float>(coords, new[] { 1, numPoints, 2 })),
            NamedOnnxValue.CreateFromTensor("point_labels",
                new DenseTensor<float>(labels, new[] { 1, numPoints })),
            NamedOnnxValue.CreateFromTensor("mask_input",
                new DenseTensor<float>(new float[256 * 256], new[] { 1, 1, 256, 256 })),
            NamedOnnxValue.CreateFromTensor("has_mask_input",
                new DenseTensor<float>(new[] { 0f }, new[] { 1 })),
            NamedOnnxValue.CreateFromTensor("orig_im_size",
                new DenseTensor<float>(new[] { (float)_origHeight, (float)_origWidth }, new[] { 2 })),
        };

        using var results = _decoder.Run(inputs);

        var masks = results.First(r => r.Name == "masks").AsTensor<float>();
        var iouResult = results.FirstOrDefault(r => r.Name == "iou_predictions");
        int numMasks = masks.Dimensions[1];

        // Pick best mask: smallest that contains click point, or highest IOU
        int bestIdx = 0;
        if (numMasks > 1)
        {
            int bestSize = int.MaxValue;
            var iou = iouResult?.AsTensor<float>();
            for (int i = 0; i < numMasks; i++)
            {
                if (masks[0, i, y, x] <= 0f) continue;
                int count = CountForeground(masks, i);
                float iouScore = iou != null ? iou[0, i] : 1f;
                if (count < bestSize && iouScore > 0.5f)
                {
                    bestSize = count;
                    bestIdx = i;
                }
            }
        }

        // Extract mask (logits > 0 = foreground)
        var result = new bool[_origHeight, _origWidth];
        int fgCount = 0;
        for (int row = 0; row < _origHeight; row++)
            for (int col = 0; col < _origWidth; col++)
                if (masks[0, bestIdx, row, col] > 0f)
                {
                    result[row, col] = true;
                    fgCount++;
                }

        float iouFinal = iouResult?.AsTensor<float>()[0, bestIdx] ?? 0f;
        Logger.Trace($"MobileSAM: mask[{bestIdx}] {fgCount}px, IOU={iouFinal:F3}" +
            (box != null ? " (box prompt)" : " (point prompt)"));
        return fgCount > 20 ? result : null;
    }

    private int CountForeground(Tensor<float> masks, int maskIdx)
    {
        int count = 0;
        for (int r = 0; r < _origHeight; r++)
            for (int c = 0; c < _origWidth; c++)
                if (masks[0, maskIdx, r, c] > 0f) count++;
        return count;
    }

    public void Dispose()
    {
        _encoder?.Dispose();
        _decoder?.Dispose();
        _encoder = null;
        _decoder = null;
        _imageEmbedding = null;
        _initialized = false;
    }
}
