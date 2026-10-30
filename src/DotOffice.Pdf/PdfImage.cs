namespace DotOffice.Pdf;

/// <summary>An image XObject. JPEG data is embedded as-is (DCTDecode) without decoding.</summary>
public sealed class PdfImage
{
    private readonly byte[] _data;
    private readonly int _components;
    private readonly bool _adobeInverted;

    private PdfImage(byte[] data, int width, int height, int components, bool adobeInverted)
    {
        _data = data;
        Width = width;
        Height = height;
        _components = components;
        _adobeInverted = adobeInverted;
    }

    public int Width { get; }
    public int Height { get; }

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
                if (components is not (1 or 3 or 4))
                    throw new NotSupportedException($"JPEG with {components} components is not supported.");
                return new PdfImage(data, width, height, components, adobe && components == 4);
            }
            pos += 2 + length;
        }
        throw new InvalidDataException("JPEG has no frame header.");
    }

    internal void Write(PdfWriter writer, int id)
    {
        var colorSpace = _components switch { 1 => "/DeviceGray", 3 => "/DeviceRGB", _ => "/DeviceCMYK" };
        var decode = _adobeInverted ? " /Decode [1 0 1 0 1 0 1 0]" : "";
        writer.WriteStream(id,
            $"/Type /XObject /Subtype /Image /Width {Width} /Height {Height} /ColorSpace {colorSpace} " +
            $"/BitsPerComponent 8 /Filter /DCTDecode{decode}",
            _data, compress: false);
    }
}
