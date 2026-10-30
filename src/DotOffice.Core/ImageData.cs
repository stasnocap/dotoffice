namespace DotOffice.Core;

public enum ImageFormat { Unknown, Png, Jpeg, Gif, Bmp, Tiff, Emf, Wmf }

/// <summary>Encoded image bytes shared by every run that references them.</summary>
public sealed class ImageData(byte[] data, string? name = null)
{
    public byte[] Data { get; } = data;
    public string? Name { get; } = name;
    public ImageFormat Format { get; } = Detect(data);

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