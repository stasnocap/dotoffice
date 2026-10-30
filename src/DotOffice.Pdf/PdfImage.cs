namespace DotOffice.Pdf;

/// <summary>
/// An image XObject. JPEG data is embedded as-is (DCTDecode) without decoding; PNG data is decoded
/// and re-embedded as FlateDecode samples with an optional soft mask for alpha.
/// </summary>
public sealed class PdfImage
{
    private readonly byte[] _data;
    private readonly string _colorSpace;
    private readonly bool _isJpeg;
    private readonly bool _adobeInverted;
    private readonly byte[]? _alpha;

    private PdfImage(byte[] data, int width, int height, string colorSpace, bool isJpeg, bool adobeInverted = false, byte[]? alpha = null)
    {
        _data = data;
        Width = width;
        Height = height;
        _colorSpace = colorSpace;
        _isJpeg = isJpeg;
        _adobeInverted = adobeInverted;
        _alpha = alpha;
    }

    public int Width { get; }
    public int Height { get; }
    public bool HasAlpha => _alpha is not null;

    internal string ResourceName { get; set; } = "";

    public static PdfImage FromJpeg(byte[] data)
    {
        if (data.Length < 4 || data[0] != 0xFF || data[1] != 0xD8)
            throw new InvalidDataException("Not a JPEG file.");

        bool adobe = false;
        int pos = 2;
        while (pos + 4 <= data.Length)
        {
            if (data[pos] != 0xFF) { pos++; continue; }
            byte marker = data[pos + 1];
            if (marker == 0xFF) { pos++; continue; }
            if (marker is 0xD8 or 0x01 || marker is >= 0xD0 and <= 0xD7) { pos += 2; continue; }

            int length = (data[pos + 2] << 8) | data[pos + 3];
            if (marker == 0xEE && length >= 7 && data.AsSpan(pos + 4, 5).SequenceEqual("Adobe"u8))
                adobe = true;

            bool isStartOfFrame = marker is >= 0xC0 and <= 0xCF && marker is not (0xC4 or 0xC8 or 0xCC);
            if (isStartOfFrame && pos + 9 < data.Length)
            {
                int height = (data[pos + 5] << 8) | data[pos + 6];
                int width = (data[pos + 7] << 8) | data[pos + 8];
                int components = data[pos + 9];
                var colorSpace = components switch
                {
                    1 => "/DeviceGray",
                    3 => "/DeviceRGB",
                    4 => "/DeviceCMYK",
                    _ => throw new NotSupportedException($"JPEG with {components} components is not supported."),
                };
                if (width <= 0 || height <= 0) throw new InvalidDataException("JPEG has invalid dimensions.");
                return new PdfImage(data, width, height, colorSpace, isJpeg: true, adobe && components == 4);
            }
            pos += 2 + length;
        }
        throw new InvalidDataException("JPEG has no frame header.");
    }

    /// <summary>Decodes a PNG (any standard color type/bit depth, interlaced or not).</summary>
    public static PdfImage FromPng(byte[] data, long maxPixels = PngDecoder.DefaultMaxPixels)
    {
        var png = PngDecoder.Decode(data, maxPixels);
        return new PdfImage(png.Color, png.Width, png.Height, png.Channels == 1 ? "/DeviceGray" : "/DeviceRGB",
            isJpeg: false, alpha: png.Alpha);
    }

    internal void Write(PdfWriter writer, int id)
    {
        var dict = $"/Type /XObject /Subtype /Image /Width {Width} /Height {Height} /ColorSpace {_colorSpace} /BitsPerComponent 8";
        if (_isJpeg)
        {
            var decode = _adobeInverted ? " /Decode [1 0 1 0 1 0 1 0]" : "";
            writer.WriteStream(id, $"{dict} /Filter /DCTDecode{decode}", _data, compress: false);
            return;
        }

        if (_alpha is null)
        {
            writer.WriteStream(id, dict, _data);
            return;
        }
        int maskId = writer.Allocate();
        writer.WriteStream(id, $"{dict} /SMask {maskId} 0 R", _data);
        writer.WriteStream(maskId,
            $"/Type /XObject /Subtype /Image /Width {Width} /Height {Height} /ColorSpace /DeviceGray /BitsPerComponent 8",
            _alpha);
    }
}