using System.Buffers.Binary;
using System.IO.Compression;

namespace DotOffice.Tests;

/// <summary>
/// Minimal PNG encoder for tests. Samples are given at native bit depth, row-major,
/// <c>channels</c> values per pixel. Rows cycle through all five filter types.
/// </summary>
internal static class PngTestEncoder
{
    private static readonly (int X, int Y, int Dx, int Dy)[] Adam7 =
        [(0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2)];

    public static byte[] Encode(int width, int height, int colorType, int bitDepth, int[] samples,
        byte[]? palette = null, byte[]? trns = null, bool interlace = false)
    {
        int channels = colorType switch { 0 => 1, 2 => 3, 3 => 1, 4 => 2, _ => 4 };
        int bpp = Math.Max(1, channels * bitDepth / 8);
        using var raw = new MemoryStream();
        int filterCounter = 0;
        var passes = interlace ? Adam7 : [(0, 0, 1, 1)];
        foreach (var (px, py, dx, dy) in passes)
        {
            int pw = (width - px + dx - 1) / dx, ph = (height - py + dy - 1) / dy;
            if (pw <= 0 || ph <= 0) continue;
            int stride = (pw * channels * bitDepth + 7) / 8;
            var prev = new byte[stride];
            for (int r = 0; r < ph; r++)
            {
                var row = new byte[stride];
                var bits = new BitWriter(row);
                for (int c = 0; c < pw; c++)
                {
                    int x = px + c * dx, y = py + r * dy;
                    for (int ch = 0; ch < channels; ch++)
                        bits.Write(samples[(y * width + x) * channels + ch], bitDepth);
                }
                int filter = filterCounter++ % 5;
                raw.WriteByte((byte)filter);
                raw.Write(Filter(filter, row, prev, bpp));
                prev = row;
            }
        }

        using var png = new MemoryStream();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
        var ihdr = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(ihdr, width);
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(4), height);
        ihdr[8] = (byte)bitDepth;
        ihdr[9] = (byte)colorType;
        ihdr[12] = (byte)(interlace ? 1 : 0);
        Chunk(png, "IHDR", ihdr);
        if (palette is not null) Chunk(png, "PLTE", palette);
        if (trns is not null) Chunk(png, "tRNS", trns);
        using var compressed = new MemoryStream();
        using (var z = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true)) z.Write(raw.ToArray());
        var data = compressed.ToArray();
        // Split image data across two IDAT chunks to exercise concatenation.
        Chunk(png, "IDAT", data[..(data.Length / 2)]);
        Chunk(png, "IDAT", data[(data.Length / 2)..]);
        Chunk(png, "IEND", []);
        return png.ToArray();
    }

    private static byte[] Filter(int type, byte[] row, byte[] prev, int bpp)
    {
        var o = new byte[row.Length];
        for (int i = 0; i < row.Length; i++)
        {
            int a = i >= bpp ? row[i - bpp] : 0, b = prev[i], c = i >= bpp ? prev[i - bpp] : 0;
            int pred = type switch
            {
                0 => 0,
                1 => a,
                2 => b,
                3 => (a + b) >> 1,
                _ => Paeth(a, b, c),
            };
            o[i] = (byte)(row[i] - pred);
        }
        return o;
    }

    private static int Paeth(int a, int b, int c)
    {
        int p = a + b - c, pa = Math.Abs(p - a), pb = Math.Abs(p - b), pc = Math.Abs(p - c);
        return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
    }

    private static void Chunk(Stream s, string type, byte[] body)
    {
        Span<byte> len = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(len, body.Length);
        s.Write(len);
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        s.Write(typeBytes);
        s.Write(body);
        BinaryPrimitives.WriteUInt32BigEndian(len, Crc32([.. typeBytes, .. body]));
        s.Write(len);
    }

    private static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (byte b in data)
        {
            crc ^= b;
            for (int k = 0; k < 8; k++) crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
        }
        return ~crc;
    }

    private sealed class BitWriter(byte[] target)
    {
        private int _bit;

        public void Write(int value, int bits)
        {
            if (bits == 16)
            {
                target[_bit / 8] = (byte)(value >> 8);
                target[_bit / 8 + 1] = (byte)value;
                _bit += 16;
                return;
            }
            for (int i = bits - 1; i >= 0; i--, _bit++)
                if (((value >> i) & 1) != 0) target[_bit / 8] |= (byte)(0x80 >> (_bit % 8));
        }
    }
}