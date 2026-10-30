using System.Buffers.Binary;
using System.IO.Compression;

namespace DotOffice.Pdf;

/// <summary>
/// Decoded PNG pixels normalized to 8 bits per sample: <see cref="Color"/> holds 1 (gray) or 3 (RGB)
/// samples per pixel; <see cref="Alpha"/> is null when the image is fully opaque.
/// </summary>
internal sealed class PngImage(int width, int height, int channels, byte[] color, byte[]? alpha)
{
    public int Width { get; } = width;
    public int Height { get; } = height;
    public int Channels { get; } = channels;
    public byte[] Color { get; } = color;
    public byte[]? Alpha { get; } = alpha;

    public (byte R, byte G, byte B, byte A) GetPixel(int x, int y)
    {
        int i = y * Width + x;
        byte a = Alpha?[i] ?? 255;
        return Channels == 1
            ? (Color[i], Color[i], Color[i], a)
            : (Color[i * 3], Color[i * 3 + 1], Color[i * 3 + 2], a);
    }
}

/// <summary>
/// Managed PNG decoder: color types 0/2/3/4/6, bit depths 1–16, all five filters, Adam7 interlacing
/// and tRNS transparency. Input is treated as untrusted: dimensions and decompressed size are capped.
/// </summary>
internal static class PngDecoder
{
    public const long DefaultMaxPixels = 40_000_000;

    private static readonly (int X, int Y, int Dx, int Dy)[] Adam7 =
        [(0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2)];

    public static bool IsPng(ReadOnlySpan<byte> data) =>
        data.StartsWith((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

    public static PngImage Decode(byte[] data, long maxPixels = DefaultMaxPixels)
    {
        if (!IsPng(data)) throw new InvalidDataException("Not a PNG file.");

        int width = 0, height = 0, bitDepth = 0, colorType = -1, interlace = 0;
        byte[]? palette = null, transparency = null;
        using var idat = new MemoryStream();
        int pos = 8;
        while (pos + 8 <= data.Length)
        {
            int length = BinaryPrimitives.ReadInt32BigEndian(data.AsSpan(pos));
            if (length < 0 || pos + 12L + length > data.Length) throw new InvalidDataException("Truncated PNG chunk.");
            var type = data.AsSpan(pos + 4, 4);
            var body = data.AsSpan(pos + 8, length);
            if (type.SequenceEqual("IHDR"u8))
            {
                if (length < 13) throw new InvalidDataException("Invalid IHDR.");
                width = BinaryPrimitives.ReadInt32BigEndian(body);
                height = BinaryPrimitives.ReadInt32BigEndian(body[4..]);
                bitDepth = body[8];
                colorType = body[9];
                interlace = body[12];
                if (body[10] != 0 || body[11] != 0 || interlace > 1)
                    throw new InvalidDataException("Unsupported PNG compression, filter or interlace method.");
            }
            else if (type.SequenceEqual("PLTE"u8)) palette = body.ToArray();
            else if (type.SequenceEqual("tRNS"u8)) transparency = body.ToArray();
            else if (type.SequenceEqual("IDAT"u8)) idat.Write(body);
            else if (type.SequenceEqual("IEND"u8)) break;
            pos += 12 + length;
        }

        int channels = colorType switch { 0 => 1, 2 => 3, 3 => 1, 4 => 2, 6 => 4, _ => 0 };
        bool validDepth = colorType switch
        {
            0 => bitDepth is 1 or 2 or 4 or 8 or 16,
            3 => bitDepth is 1 or 2 or 4 or 8,
            2 or 4 or 6 => bitDepth is 8 or 16,
            _ => false,
        };
        if (channels == 0 || !validDepth) throw new InvalidDataException($"Invalid PNG color type {colorType} / bit depth {bitDepth}.");
        if (width <= 0 || height <= 0 || (long)width * height > maxPixels)
            throw new InvalidDataException($"PNG dimensions {width}x{height} are invalid or exceed the limit.");
        if (colorType == 3 && (palette is null || palette.Length % 3 != 0))
            throw new InvalidDataException("Indexed PNG without a valid palette.");

        int bitsPerPixel = channels * bitDepth;
        var passes = interlace == 1 ? Adam7 : [(0, 0, 1, 1)];
        long expected = 0;
        foreach (var (px, py, dx, dy) in passes)
        {
            long pw = (width - px + dx - 1) / dx, ph = (height - py + dy - 1) / dy;
            if (pw > 0 && ph > 0) expected += ph * (1 + (pw * bitsPerPixel + 7) / 8);
        }

        var raw = Inflate(idat, expected);

        // Unpacked samples (16-bit samples rounded to 8 bits) before color conversion.
        var samples = new byte[(long)width * height * channels];
        var transparentKey = new int[channels];
        int offset = 0;
        foreach (var (px, py, dx, dy) in passes)
        {
            int pw = (width - px + dx - 1) / dx, ph = (height - py + dy - 1) / dy;
            if (pw <= 0 || ph <= 0) continue;
            int stride = (pw * bitsPerPixel + 7) / 8;
            int bpp = Math.Max(1, bitsPerPixel / 8);
            var previous = new byte[stride];
            var current = new byte[stride];
            for (int row = 0; row < ph; row++)
            {
                int filter = raw[offset];
                Array.Copy(raw, offset + 1, current, 0, stride);
                offset += 1 + stride;
                Unfilter(filter, current, previous, bpp);
                int y = py + row * dy;
                for (int col = 0; col < pw; col++)
                {
                    int x = px + col * dx;
                    long dst = ((long)y * width + x) * channels;
                    for (int c = 0; c < channels; c++)
                        samples[dst + c] = ReadSample(current, col * channels + c, bitDepth, scale: colorType != 3);
                }
                (previous, current) = (current, previous);
            }
        }

        return ToImage(width, height, colorType, bitDepth, samples, palette, transparency);
    }

    private static byte[] Inflate(MemoryStream compressed, long expected)
    {
        if (expected > int.MaxValue) throw new InvalidDataException("PNG is too large.");
        var raw = new byte[expected];
        compressed.Position = 0;
        using var zlib = new ZLibStream(compressed, CompressionMode.Decompress);
        int read = 0;
        try
        {
            while (read < raw.Length)
            {
                int n = zlib.Read(raw, read, raw.Length - read);
                if (n == 0) break;
                read += n;
            }
        }
        catch (InvalidDataException) when (read > 0)
        {
            // Truncated/corrupt tail: keep what decoded (remaining rows stay zero), like most viewers.
        }
        if (read == 0) throw new InvalidDataException("PNG image data is empty or corrupt.");
        return raw;
    }

    private static void Unfilter(int filter, byte[] cur, byte[] prev, int bpp)
    {
        switch (filter)
        {
            case 0: break;
            case 1:
                for (int i = bpp; i < cur.Length; i++) cur[i] += cur[i - bpp];
                break;
            case 2:
                for (int i = 0; i < cur.Length; i++) cur[i] += prev[i];
                break;
            case 3:
                for (int i = 0; i < cur.Length; i++)
                    cur[i] += (byte)(((i >= bpp ? cur[i - bpp] : 0) + prev[i]) >> 1);
                break;
            case 4:
                for (int i = 0; i < cur.Length; i++)
                {
                    int a = i >= bpp ? cur[i - bpp] : 0, b = prev[i], c = i >= bpp ? prev[i - bpp] : 0;
                    int p = a + b - c, pa = Math.Abs(p - a), pb = Math.Abs(p - b), pc = Math.Abs(p - c);
                    cur[i] += (byte)(pa <= pb && pa <= pc ? a : pb <= pc ? b : c);
                }
                break;
            default:
                throw new InvalidDataException($"Invalid PNG filter type {filter}.");
        }
    }

    /// <summary>Reads sample <paramref name="index"/> of a row; sub-byte gray values are scaled to 0–255.</summary>
    private static byte ReadSample(byte[] row, int index, int bitDepth, bool scale)
    {
        switch (bitDepth)
        {
            case 8: return row[index];
            case 16: return Round16((row[index * 2] << 8) | row[index * 2 + 1]);
            default:
                int perByte = 8 / bitDepth;
                int shift = 8 - bitDepth * (index % perByte + 1);
                int value = (row[index / perByte] >> shift) & ((1 << bitDepth) - 1);
                return scale ? (byte)(value * 255 / ((1 << bitDepth) - 1)) : (byte)value;
        }
    }

    private static byte Round16(int value) => (byte)((value + 128) / 257);

    private static PngImage ToImage(int width, int height, int colorType, int bitDepth, byte[] s, byte[]? palette, byte[]? trns)
    {
        int n = width * height;
        byte[] color;
        byte[]? alpha = null;
        switch (colorType)
        {
            case 0:
                color = s;
                if (trns is { Length: >= 2 })
                {
                    int key = BinaryPrimitives.ReadUInt16BigEndian(trns);
                    byte key8 = bitDepth == 16 ? Round16(key) : bitDepth == 8 ? (byte)key : (byte)(key * 255 / ((1 << bitDepth) - 1));
                    alpha = new byte[n];
                    for (int i = 0; i < n; i++) alpha[i] = s[i] == key8 ? (byte)0 : (byte)255;
                }
                return Finish(width, height, 1, color, alpha);
            case 2:
                color = s;
                if (trns is { Length: >= 6 })
                {
                    byte r = Key(trns, 0, bitDepth), g = Key(trns, 2, bitDepth), b = Key(trns, 4, bitDepth);
                    alpha = new byte[n];
                    for (int i = 0; i < n; i++)
                        alpha[i] = s[i * 3] == r && s[i * 3 + 1] == g && s[i * 3 + 2] == b ? (byte)0 : (byte)255;
                }
                return Finish(width, height, 3, color, alpha);
            case 3:
                color = new byte[n * 3];
                alpha = trns is { Length: > 0 } ? new byte[n] : null;
                int entries = palette!.Length / 3;
                for (int i = 0; i < n; i++)
                {
                    int idx = s[i];
                    if (idx < entries) Array.Copy(palette, idx * 3, color, i * 3, 3);
                    if (alpha is not null) alpha[i] = idx < trns!.Length ? trns[idx] : (byte)255;
                }
                return Finish(width, height, 3, color, alpha);
            case 4:
                color = new byte[n];
                alpha = new byte[n];
                for (int i = 0; i < n; i++) (color[i], alpha[i]) = (s[i * 2], s[i * 2 + 1]);
                return Finish(width, height, 1, color, alpha);
            default:
                color = new byte[n * 3];
                alpha = new byte[n];
                for (int i = 0; i < n; i++)
                {
                    Array.Copy(s, i * 4, color, i * 3, 3);
                    alpha[i] = s[i * 4 + 3];
                }
                return Finish(width, height, 3, color, alpha);
        }

        static byte Key(byte[] t, int at, int depth)
        {
            int v = BinaryPrimitives.ReadUInt16BigEndian(t.AsSpan(at));
            return depth == 16 ? Round16(v) : (byte)v;
        }
    }

    private static PngImage Finish(int width, int height, int channels, byte[] color, byte[]? alpha)
    {
        if (alpha is not null && Array.TrueForAll(alpha, a => a == 255)) alpha = null;
        return new PngImage(width, height, channels, color, alpha);
    }
}