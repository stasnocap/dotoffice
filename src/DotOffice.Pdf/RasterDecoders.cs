using System.Buffers.Binary;

namespace DotOffice.Pdf;

/// <summary>Decoded 8-bit RGB pixels (3 bytes per pixel) with an optional alpha plane.</summary>
internal sealed record RasterImage(int Width, int Height, byte[] Rgb, byte[]? Alpha);

/// <summary>Managed BMP, GIF and TIFF decoders. Input is untrusted: dimensions are capped.</summary>
internal static class RasterDecoders
{
    private const long MaxPixels = 40_000_000;

    private static void CheckSize(long width, long height)
    {
        if (width <= 0 || height <= 0 || width * height > MaxPixels) throw new InvalidDataException("Image has invalid dimensions.");
    }

    // ---------------------------------------------------------------- BMP

    public static RasterImage DecodeBmp(byte[] d)
    {
        if (d.Length < 26 || d[0] != 'B' || d[1] != 'M') throw new InvalidDataException("Not a BMP file.");
        int dataOffset = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(10));
        int headerSize = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(14));
        int width, height, bpp, compression = 0, colors = 0, entry;
        if (headerSize == 12)
        {
            width = BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(18));
            height = BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(20));
            bpp = BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(24));
            entry = 3;
        }
        else if (headerSize >= 40 && d.Length >= 14 + 40)
        {
            width = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(18));
            height = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(22));
            bpp = BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(28));
            compression = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(30));
            colors = BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(46));
            entry = 4;
        }
        else throw new InvalidDataException("Unsupported BMP header.");

        bool topDown = height < 0;
        height = Math.Abs(height);
        CheckSize(width, height);
        if (compression is not (0 or 3)) throw new NotSupportedException($"BMP compression {compression} is not supported.");
        if (bpp is not (1 or 4 or 8 or 16 or 24 or 32)) throw new InvalidDataException($"BMP with {bpp} bits per pixel is not supported.");

        int paletteAt = 14 + headerSize;
        int paletteCount = bpp <= 8 ? (colors > 0 ? colors : 1 << bpp) : 0;
        var palette = new byte[paletteCount * 3];
        for (int i = 0; i < paletteCount; i++)
        {
            int p = paletteAt + i * entry;
            if (p + 3 > d.Length) break;
            palette[i * 3] = d[p + 2]; palette[i * 3 + 1] = d[p + 1]; palette[i * 3 + 2] = d[p];
        }

        // 16-bit bitfields and 32-bit masks (BI_BITFIELDS); default 5-5-5 and BGRA byte order.
        uint rMask = 0x7C00, gMask = 0x03E0, bMask = 0x001F;
        if (compression == 3 && bpp is 16 or 32 && headerSize >= 40 && d.Length >= 14 + 40 + 12)
        {
            rMask = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(54));
            gMask = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(58));
            bMask = BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(62));
        }

        long stride = ((long)width * bpp + 31) / 32 * 4;
        if (dataOffset < 0 || dataOffset + stride * height > d.Length) throw new InvalidDataException("BMP pixel data is truncated.");
        var rgb = new byte[width * height * 3];
        for (int y = 0; y < height; y++)
        {
            int row = topDown ? y : height - 1 - y;
            long src = dataOffset + stride * y;
            for (int x = 0; x < width; x++)
            {
                int o = (row * width + x) * 3;
                switch (bpp)
                {
                    case 24:
                    case 32 when compression == 0:
                    {
                        long p = src + (long)x * (bpp / 8);
                        rgb[o] = d[p + 2]; rgb[o + 1] = d[p + 1]; rgb[o + 2] = d[p];
                        break;
                    }
                    case 32:
                    case 16:
                    {
                        uint v = bpp == 16 ? BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan((int)(src + x * 2L)))
                                           : BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan((int)(src + x * 4L)));
                        rgb[o] = Scale(v, rMask); rgb[o + 1] = Scale(v, gMask); rgb[o + 2] = Scale(v, bMask);
                        break;
                    }
                    default:
                    {
                        int bit = x * bpp;
                        int v = (d[src + bit / 8] >> (8 - bpp - bit % 8)) & ((1 << bpp) - 1);
                        if (v >= paletteCount) v = 0;
                        rgb[o] = palette[v * 3]; rgb[o + 1] = palette[v * 3 + 1]; rgb[o + 2] = palette[v * 3 + 2];
                        break;
                    }
                }
            }
        }
        return new RasterImage(width, height, rgb, null);
    }

    private static byte Scale(uint value, uint mask)
    {
        if (mask == 0) return 0;
        int shift = System.Numerics.BitOperations.TrailingZeroCount(mask);
        uint max = mask >> shift;
        return (byte)(((value & mask) >> shift) * 255 / max);
    }

    // ---------------------------------------------------------------- GIF

    /// <summary>Decodes the first frame of a GIF (global/local palette, transparency, interlacing, LZW).</summary>
    public static RasterImage DecodeGif(byte[] d)
    {
        if (d.Length < 13 || !d.AsSpan(0, 4).SequenceEqual("GIF8"u8)) throw new InvalidDataException("Not a GIF file.");
        int screenW = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(6));
        int screenH = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(8));
        byte flags = d[10];
        int pos = 13;
        byte[] globalPalette = [];
        if ((flags & 0x80) != 0)
        {
            int n = 3 << ((flags & 7) + 1);
            if (pos + n > d.Length) throw new InvalidDataException("GIF palette is truncated.");
            globalPalette = d.AsSpan(pos, n).ToArray();
            pos += n;
        }
        int transparent = -1;
        while (pos < d.Length)
        {
            byte block = d[pos++];
            if (block == 0x3B) break;
            if (block == 0x21)
            {
                if (pos >= d.Length) break;
                byte label = d[pos++];
                if (label == 0xF9 && pos + 5 < d.Length && d[pos] >= 4)
                    transparent = (d[pos + 1] & 1) != 0 ? d[pos + 4] : -1;
                pos = SkipBlocks(d, pos);
                continue;
            }
            if (block != 0x2C) throw new InvalidDataException("Unexpected GIF block.");
            if (pos + 9 > d.Length) throw new InvalidDataException("GIF image is truncated.");
            int left = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(pos));
            int top = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(pos + 2));
            int w = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(pos + 4));
            int h = BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(pos + 6));
            byte f = d[pos + 8];
            pos += 9;
            var palette = globalPalette;
            if ((f & 0x80) != 0)
            {
                int n = 3 << ((f & 7) + 1);
                if (pos + n > d.Length) throw new InvalidDataException("GIF palette is truncated.");
                palette = d.AsSpan(pos, n).ToArray();
                pos += n;
            }
            CheckSize(w, h);
            if (pos >= d.Length) throw new InvalidDataException("GIF image is truncated.");
            int minCode = d[pos++];
            if (minCode is < 2 or > 11) throw new InvalidDataException("Invalid GIF code size.");
            var data = new List<byte>();
            while (pos < d.Length && d[pos] != 0)
            {
                int len = d[pos++];
                if (pos + len > d.Length) throw new InvalidDataException("GIF image is truncated.");
                for (int i = 0; i < len; i++) data.Add(d[pos + i]);
                pos += len;
            }
            var indices = Lzw(data.ToArray(), minCode, w * h);

            int width = Math.Max(screenW, left + w), height = Math.Max(screenH, top + h);
            if (left == 0 && top == 0) { width = w; height = h; }
            CheckSize(width, height);
            var rgb = new byte[width * height * 3];
            var alpha = new byte[width * height];
            bool interlaced = (f & 0x40) != 0;
            var rowOf = interlaced ? InterlaceRows(h) : null;
            bool anyTransparent = false;
            for (int i = 0; i < w * h; i++)
            {
                int row = (rowOf?[i / w] ?? i / w) + top, x = i % w + left;
                if (row >= height || x >= width) continue;
                int c = indices[i];
                int o = row * width + x;
                if (c == transparent) { anyTransparent = true; continue; }
                alpha[o] = 255;
                if (c * 3 + 2 < palette.Length)
                { rgb[o * 3] = palette[c * 3]; rgb[o * 3 + 1] = palette[c * 3 + 1]; rgb[o * 3 + 2] = palette[c * 3 + 2]; }
            }
            return new RasterImage(width, height, rgb, anyTransparent || left != 0 || top != 0 ? alpha : null);
        }
        throw new InvalidDataException("GIF has no image.");
    }

    private static int SkipBlocks(byte[] d, int pos)
    {
        while (pos < d.Length && d[pos] != 0) pos += d[pos] + 1;
        return pos + 1;
    }

    private static int[] InterlaceRows(int h)
    {
        var map = new int[h];
        int k = 0;
        foreach (var (start, step) in new[] { (0, 8), (4, 8), (2, 4), (1, 2) })
            for (int r = start; r < h; r += step) map[k++] = r;
        return map;
    }

    private static byte[] Lzw(byte[] data, int minCode, int count)
    {
        var output = new byte[count];
        int clear = 1 << minCode, end = clear + 1;
        var prefix = new short[4096];
        var suffix = new byte[4096];
        var stack = new byte[4097];
        int size = minCode + 1, next = end + 1, prev = -1, written = 0;
        int bits = 0, bitCount = 0, pos = 0;
        byte first = 0;
        while (written < count)
        {
            while (bitCount < size && pos < data.Length) { bits |= data[pos++] << bitCount; bitCount += 8; }
            if (bitCount < size) break;
            int code = bits & ((1 << size) - 1);
            bits >>= size; bitCount -= size;
            if (code == clear) { size = minCode + 1; next = end + 1; prev = -1; continue; }
            if (code == end) break;
            int sp = 0, cur = code;
            if (prev == -1)
            {
                if (code >= clear) throw new InvalidDataException("Corrupt GIF data.");
                output[written++] = first = (byte)code;
                prev = code;
                continue;
            }
            if (code >= next)
            {
                if (code > next) throw new InvalidDataException("Corrupt GIF data.");
                stack[sp++] = first; cur = prev;
            }
            while (cur >= clear) { stack[sp++] = suffix[cur]; cur = prefix[cur]; }
            first = (byte)cur;
            stack[sp++] = first;
            while (sp > 0 && written < count) output[written++] = stack[--sp];
            if (next < 4096)
            {
                prefix[next] = (short)prev; suffix[next] = first; next++;
                if (next == 1 << size && size < 12) size++;
            }
            prev = code;
        }
        return output;
    }

    // ---------------------------------------------------------------- TIFF

    /// <summary>Decodes a baseline TIFF: 1/2/4/8-bit gray or palette and 8-bit RGB(A), stored or PackBits strips.</summary>
    public static RasterImage DecodeTiff(byte[] d)
    {
        if (d.Length < 8) throw new InvalidDataException("Not a TIFF file.");
        bool le = d[0] == 'I';
        if (!(le && d[1] == 'I' || !le && d[0] == 'M' && d[1] == 'M')) throw new InvalidDataException("Not a TIFF file.");
        ushort U16(long p) => p + 2 > d.Length ? throw new InvalidDataException("TIFF is truncated.")
            : le ? BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan((int)p)) : BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan((int)p));
        uint U32(long p) => p + 4 > d.Length ? throw new InvalidDataException("TIFF is truncated.")
            : le ? BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan((int)p)) : BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan((int)p));

        long ifd = U32(4);
        int count = U16(ifd);
        var tags = new Dictionary<int, uint[]>();
        for (int i = 0; i < count; i++)
        {
            long e = ifd + 2 + i * 12L;
            int tag = U16(e), type = U16(e + 2);
            uint n = U32(e + 4);
            int size = type switch { 3 => 2, 4 => 4, 1 => 1, _ => 0 };
            if (size == 0 || n > 1_000_000) continue;
            long at = size * n <= 4 ? e + 8 : U32(e + 8);
            var values = new uint[n];
            for (int k = 0; k < n; k++) values[k] = size switch { 1 => d[at + k], 2 => U16(at + 2L * k), _ => U32(at + 4L * k) };
            tags[tag] = values;
        }
        uint One(int tag, uint def) => tags.TryGetValue(tag, out var v) && v.Length > 0 ? v[0] : def;

        int width = (int)One(256, 0), height = (int)One(257, 0);
        CheckSize(width, height);
        int bits = (int)One(258, 1), compression = (int)One(259, 1), photo = (int)One(262, 1);
        int spp = (int)One(277, 1), planar = (int)One(284, 1), predictor = (int)One(317, 1);
        if (compression is not (1 or 32773)) throw new NotSupportedException($"TIFF compression {compression} is not supported.");
        if (planar != 1 || predictor != 1) throw new NotSupportedException("Planar or predicted TIFF is not supported.");
        if (photo is not (0 or 1 or 2 or 3)) throw new NotSupportedException($"TIFF photometric {photo} is not supported.");
        bool color = photo == 2;
        if (color && (bits != 8 || spp is not (3 or 4))) throw new NotSupportedException("Only 8-bit RGB TIFF is supported.");
        if (!color && (bits is not (1 or 2 or 4 or 8) || spp != 1)) throw new NotSupportedException("Unsupported TIFF sample layout.");
        if (!tags.TryGetValue(273, out var offsets) || !tags.TryGetValue(279, out var counts) || offsets.Length != counts.Length)
            throw new InvalidDataException("TIFF has no strips.");

        int rowBytes = (width * spp * bits + 7) / 8;
        long total = (long)rowBytes * height;
        var raw = new byte[total];
        long filled = 0;
        for (int s = 0; s < offsets.Length && filled < total; s++)
        {
            long start = offsets[s], len = counts[s];
            if (start + len > d.Length) throw new InvalidDataException("TIFF strip is truncated.");
            var strip = d.AsSpan((int)start, (int)len);
            if (compression == 1)
            {
                int n = (int)Math.Min(len, total - filled);
                strip[..n].CopyTo(raw.AsSpan((int)filled));
                filled += n;
            }
            else
            {
                int i = 0;
                while (i < strip.Length && filled < total)
                {
                    sbyte c = (sbyte)strip[i++];
                    if (c >= 0)
                        for (int k = 0; k <= c && i < strip.Length && filled < total; k++) raw[filled++] = strip[i++];
                    else if (c != -128 && i < strip.Length)
                    {
                        byte v = strip[i++];
                        for (int k = 0; k < 1 - c && filled < total; k++) raw[filled++] = v;
                    }
                }
            }
        }
        if (filled < total) throw new InvalidDataException("TIFF pixel data is truncated.");

        var rgb = new byte[width * height * 3];
        byte[]? alpha = color && spp == 4 ? new byte[width * height] : null;
        tags.TryGetValue(320, out var map);
        for (int y = 0; y < height; y++)
            for (int x = 0; x < width; x++)
            {
                int o = (y * width + x) * 3;
                if (color)
                {
                    long p = (long)y * rowBytes + x * spp;
                    rgb[o] = raw[p]; rgb[o + 1] = raw[p + 1]; rgb[o + 2] = raw[p + 2];
                    if (alpha is not null) alpha[y * width + x] = raw[p + 3];
                    continue;
                }
                int bit = x * bits;
                int v = (raw[(long)y * rowBytes + bit / 8] >> (8 - bits - bit % 8)) & ((1 << bits) - 1);
                if (photo == 3)
                {
                    int n = 1 << bits;
                    if (map is null || map.Length < n * 3) throw new InvalidDataException("TIFF has no color map.");
                    rgb[o] = (byte)(map[v] >> 8); rgb[o + 1] = (byte)(map[n + v] >> 8); rgb[o + 2] = (byte)(map[2 * n + v] >> 8);
                }
                else
                {
                    int g = v * 255 / ((1 << bits) - 1);
                    if (photo == 0) g = 255 - g;
                    rgb[o] = rgb[o + 1] = rgb[o + 2] = (byte)g;
                }
            }
        return new RasterImage(width, height, rgb, alpha);
    }
}
