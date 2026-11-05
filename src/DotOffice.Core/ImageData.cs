namespace DotOffice.Core;

public enum ImageFormat { Unknown, Png, Jpeg, Gif, Bmp, Tiff, Emf, Wmf }

/// <summary>Encoded image bytes shared by every run that references them.</summary>
public sealed class ImageData(byte[] data, string? name = null)
{
    public byte[] Data { get; } = data;
    public string? Name { get; } = name;
    public ImageFormat Format { get; } = Detect(data);

    /// <summary>
    /// The size of a PNG, JPEG, GIF, BMP or TIFF picture in pixels, and the resolution it declares (PNG's pHYs, a TIFF's XResolution; 96 dots per inch otherwise).
    /// Null for other formats and for data that cannot be read.
    /// </summary>
    public (int Width, int Height, double Dpi)? PixelSize => _size ??= ReadSize();

    private (int, int, double)? _size;

    private (int, int, double)? ReadSize()
    {
        var d = Data;
        static int Be16(byte[] b, int at) => (b[at] << 8) | b[at + 1];
        static int Be32(byte[] b, int at) => (b[at] << 24) | (b[at + 1] << 16) | (b[at + 2] << 8) | b[at + 3];
        try
        {
            switch (Format)
            {
                case ImageFormat.Png when d.Length >= 24:
                {
                    double dpi = 96;
                    // pHYs (pixels per meter) sits between IHDR and IDAT.
                    for (int at = 8; at + 12 <= d.Length;)
                    {
                        int length = Be32(d, at);
                        if (length < 0) break;
                        string type = System.Text.Encoding.ASCII.GetString(d, at + 4, 4);
                        if (type == "pHYs" && length >= 9 && at + 8 + 9 <= d.Length && d[at + 16] == 1 && Be32(d, at + 8) > 0) dpi = Be32(d, at + 8) * 0.0254;
                        if (type is "IDAT" or "IEND") break;
                        at += 12 + length;
                    }
                    int w = Be32(d, 16), h = Be32(d, 20);
                    return w > 0 && h > 0 ? (w, h, Math.Clamp(dpi, 24, 1200)) : null;
                }
                case ImageFormat.Gif when d.Length >= 10:
                    return (d[6] | d[7] << 8, d[8] | d[9] << 8, 96);
                case ImageFormat.Bmp when d.Length >= 26:
                {
                    int w = BitConverter.ToInt32(d, 18), h = Math.Abs(BitConverter.ToInt32(d, 22));
                    return w > 0 && h > 0 ? (w, h, 96) : null;
                }
                case ImageFormat.Tiff when d.Length >= 8:
                    return ReadTiffSize(d);
                case ImageFormat.Jpeg:
                    for (int at = 2; at + 9 < d.Length;)
                    {
                        if (d[at] != 0xFF) { at++; continue; }
                        byte marker = d[at + 1];
                        if (marker == 0xFF || marker is 0xD8 or 0x01 or >= 0xD0 and <= 0xD7) { at += marker == 0xFF ? 1 : 2; continue; }
                        if (marker is >= 0xC0 and <= 0xCF and not (0xC4 or 0xC8 or 0xCC))
                            return (Be16(d, at + 7), Be16(d, at + 5), 96);
                        at += 2 + Be16(d, at + 2);
                    }
                    return null;
            }
        }
        catch (IndexOutOfRangeException) { }
        return null;
    }

    /// <summary>The size of a TIFF from its first directory: width, length and the horizontal resolution (inches or centimeters; 96 when missing).</summary>
    private static (int, int, double)? ReadTiffSize(byte[] d)
    {
        bool le = d[0] == 'I';
        ushort U16(long p) => le ? (ushort)(d[p] | d[p + 1] << 8) : (ushort)(d[p] << 8 | d[p + 1]);
        uint U32(long p) => le ? (uint)(d[p] | d[p + 1] << 8 | d[p + 2] << 16 | d[p + 3] << 24) : (uint)(d[p] << 24 | d[p + 1] << 16 | d[p + 2] << 8 | d[p + 3]);
        long ifd = U32(4);
        int count = U16(ifd);
        uint width = 0, height = 0;
        double dpi = 96, resolution = 0;
        int unit = 2;
        for (int i = 0; i < count; i++)
        {
            long e = ifd + 2 + i * 12L;
            int tag = U16(e), type = U16(e + 2);
            switch (tag)
            {
                case 256: width = type == 3 ? U16(e + 8) : U32(e + 8); break;
                case 257: height = type == 3 ? U16(e + 8) : U32(e + 8); break;
                case 296: unit = U16(e + 8); break;
                case 282 when type == 5:
                {
                    long at = U32(e + 8);
                    double denominator = U32(at + 4);
                    if (denominator > 0) resolution = U32(at) / denominator;
                    break;
                }
            }
        }
        if (resolution > 0) dpi = unit == 3 ? resolution * 2.54 : unit == 2 ? resolution : 96;
        return width > 0 && height > 0 ? ((int)width, (int)height, Math.Clamp(dpi, 24, 1200)) : null;
    }

    public static ImageFormat Detect(ReadOnlySpan<byte> data)
    {
        if (data.StartsWith((ReadOnlySpan<byte>)[0x89, (byte)'P', (byte)'N', (byte)'G'])) return ImageFormat.Png;
        if (data.StartsWith((ReadOnlySpan<byte>)[0xFF, 0xD8])) return ImageFormat.Jpeg;
        if (data.StartsWith("GIF8"u8)) return ImageFormat.Gif;
        if (data.StartsWith("BM"u8)) return ImageFormat.Bmp;
        if (data.StartsWith("II*\0"u8) || data.StartsWith("MM\0*"u8)) return ImageFormat.Tiff;
        if (data.Length >= 44 && data[0] == 1 && data[1] == 0 && data[2] == 0 && data[3] == 0 && data.Slice(40, 4).SequenceEqual(" EMF"u8))
            return ImageFormat.Emf;
        if (data.StartsWith((ReadOnlySpan<byte>)[0xD7, 0xCD, 0xC6, 0x9A]) || data.StartsWith((ReadOnlySpan<byte>)[1, 0, 9, 0]))
            return ImageFormat.Wmf;
        return ImageFormat.Unknown;
    }
}