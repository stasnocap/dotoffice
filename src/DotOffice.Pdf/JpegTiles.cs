namespace DotOffice.Pdf;

/// <summary>
/// Joins the tiles of a JPEG-compressed TIFF into one baseline JPEG without decoding any pixels. Each tile is a complete JPEG of its own
/// (MCUs in raster order inside the tile), so the MCUs of a row of the whole image are spread over several tiles and cannot simply be
/// chained. The tiles are entropy-decoded into their DCT coefficients, laid out on the grid of the whole image and entropy-coded again
/// with the Huffman tables the tiles already use; quantization is untouched, so the result decodes to the same pixels.
/// </summary>
internal static class JpegTiles
{
    private sealed class Huffman
    {
        public readonly int[] MinCode = new int[17], MaxCode = new int[17], ValPtr = new int[17];
        public readonly byte[] Values;
        public readonly int[] Code = new int[256];
        public readonly byte[] Size = new byte[256];

        public Huffman(byte[] counts, byte[] values)
        {
            Values = values;
            int code = 0, k = 0;
            for (int length = 1; length <= 16; length++)
            {
                ValPtr[length] = k;
                MinCode[length] = code;
                for (int i = 0; i < counts[length - 1] && k < values.Length; i++, code++, k++) { Code[values[k]] = code; Size[values[k]] = (byte)length; }
                MaxCode[length] = counts[length - 1] == 0 ? -1 : code - 1;
                code <<= 1;
            }
        }
    }

    private sealed record Component(int Id, int H, int V, int DcTable, int AcTable);

    private sealed class Stream
    {
        public int FrameAt = -1, ScanAt = -1, DataAt = -1, Interval;
        public int Precision, Width, Height;
        public List<Component> Components = [];
        public Dictionary<int, Huffman> Tables = [];
        public List<(int At, int Length)> Segments = [];
    }

    /// <summary>The segments of a JPEG stream up to its first scan, and its baseline frame; null for anything else (progressive, several scans, arithmetic coding).</summary>
    private static Stream? Parse(byte[] j)
    {
        var s = new Stream();
        var selectors = new Dictionary<int, (int Dc, int Ac)>();
        int p = 2;
        while (p + 4 <= j.Length)
        {
            if (j[p] != 0xFF) return null;
            byte marker = j[p + 1];
            if (marker == 0xFF) { p++; continue; }
            int length = j[p + 2] << 8 | j[p + 3];
            if (length < 2 || p + 2 + length > j.Length) return null;
            int body = p + 4, end = p + 2 + length;
            switch (marker)
            {
                case 0xC0 or 0xC1:
                {
                    s.FrameAt = p;
                    s.Precision = j[body];
                    if (s.Precision != 8) return null;
                    (s.Height, s.Width) = (j[body + 1] << 8 | j[body + 2], j[body + 3] << 8 | j[body + 4]);
                    int n = j[body + 5];
                    if (n is < 1 or > 4 || length < 8 + 3 * n) return null;
                    for (int k = 0; k < n; k++) s.Components.Add(new Component(j[body + 6 + 3 * k], Math.Max(1, j[body + 7 + 3 * k] >> 4), Math.Max(1, j[body + 7 + 3 * k] & 15), 0, 0));
                    break;
                }
                case 0xC2 or 0xC3 or (>= 0xC5 and <= 0xCF and not 0xC8 and not 0xCC): return null;
                case 0xC4:
                    for (int q = body; q + 17 <= end;)
                    {
                        int info = j[q];
                        var counts = j[(q + 1)..(q + 17)];
                        int total = counts.Sum(c => c);
                        if (q + 17 + total > end) return null;
                        s.Tables[info] = new Huffman(counts, j[(q + 17)..(q + 17 + total)]);
                        q += 17 + total;
                    }
                    break;
                case 0xDD: s.Interval = j[body] << 8 | j[body + 1]; break;
                case 0xDA:
                {
                    if (s.FrameAt < 0) return null;
                    int n = j[body];
                    if (n != s.Components.Count) return null; // not interleaved: several scans
                    for (int k = 0; k < n; k++) selectors[j[body + 1 + 2 * k]] = (j[body + 2 + 2 * k] >> 4, j[body + 2 + 2 * k] & 15);
                    s.Components = s.Components.Select(c => selectors.TryGetValue(c.Id, out var t) ? c with { DcTable = t.Dc, AcTable = t.Ac } : c).ToList();
                    if (s.Components.Any(c => !selectors.ContainsKey(c.Id))) return null;
                    (s.ScanAt, s.DataAt) = (p, end);
                    return s;
                }
            }
            s.Segments.Add((p, 2 + length));
            p = end;
        }
        return null;
    }

    private sealed class BitReader(byte[] data, int position)
    {
        private uint _bits;
        private int _count;
        public int Position = position;

        private int Bit()
        {
            if (_count == 0)
            {
                byte b = Position < data.Length ? data[Position] : (byte)0xFF;
                if (b == 0xFF)
                {
                    if (Position + 1 < data.Length && data[Position + 1] == 0) Position += 2;
                    else b = 0; // a marker (or the end): zeros are fed until the restart is taken
                }
                else Position++;
                _bits = b;
                _count = 8;
            }
            _count--;
            return (int)(_bits >> _count) & 1;
        }

        public int Receive(int n)
        {
            int v = 0;
            for (int i = 0; i < n; i++) v = v << 1 | Bit();
            return v;
        }

        public int Symbol(Huffman table)
        {
            int code = 0;
            for (int length = 1; length <= 16; length++)
            {
                code = code << 1 | Bit();
                if (table.MaxCode[length] >= 0 && code <= table.MaxCode[length] && code >= table.MinCode[length])
                    return table.Values[table.ValPtr[length] + code - table.MinCode[length]];
            }
            throw new InvalidDataException("Bad Huffman code in a JPEG tile.");
        }

        /// <summary>Drops the bits left in the current byte and steps over a restart marker.</summary>
        public void Restart()
        {
            _count = 0;
            if (Position + 1 < data.Length && data[Position] == 0xFF && data[Position + 1] is >= 0xD0 and <= 0xD7) Position += 2;
        }
    }

    private sealed class BitWriter
    {
        private readonly MemoryStream _out = new();
        private uint _acc;
        private int _count;

        public void Put(int code, int size)
        {
            _acc = _acc << size | (uint)code;
            _count += size;
            while (_count >= 8)
            {
                byte b = (byte)(_acc >> (_count - 8));
                _out.WriteByte(b);
                if (b == 0xFF) _out.WriteByte(0);
                _count -= 8;
            }
            _acc &= (1u << _count) - 1;
        }

        public byte[] Finish()
        {
            if (_count > 0) Put((1 << (8 - _count)) - 1, 8 - _count);
            return _out.ToArray();
        }
    }

    private static int Extend(int v, int size) => size == 0 ? 0 : v < 1 << (size - 1) ? v - (1 << size) + 1 : v;

    private static int SizeOf(int v)
    {
        v = Math.Abs(v);
        int n = 0;
        while (v != 0) { n++; v >>= 1; }
        return n;
    }

    /// <summary>
    /// The joined JPEG, or null when the tiles are not baseline interleaved JPEGs of one kind, are not whole MCUs, or use a symbol their
    /// own tables do not have.
    /// </summary>
    /// <param name="tiles">Each tile as a complete JPEG stream (with the shared tables already in front), row by row.</param>
    /// <param name="across">Tiles per row.</param>
    /// <param name="down">Rows of tiles.</param>
    /// <param name="tileWidth">The width of a tile in pixels (a whole number of MCUs).</param>
    /// <param name="tileHeight">The height of a tile in pixels.</param>
    /// <param name="width">The width of the picture.</param>
    /// <param name="height">The height of the picture.</param>
    public static byte[]? Join(IReadOnlyList<byte[]> tiles, int across, int down, int tileWidth, int tileHeight, int width, int height)
    {
        try { return JoinCore(tiles, across, down, tileWidth, tileHeight, width, height); }
        catch (InvalidDataException) { return null; }
        catch (IndexOutOfRangeException) { return null; }
    }

    private static byte[]? JoinCore(IReadOnlyList<byte[]> tiles, int across, int down, int tileWidth, int tileHeight, int width, int height)
    {
        if (tiles.Count != across * down || width is < 1 or > 65535 || height is < 1 or > 65535) return null;
        var first = Parse(tiles[0]);
        if (first is null) return null;
        var components = first.Components;
        // With one component the MCU is a single block whatever the sampling factors say.
        int maxH = components.Count == 1 ? 1 : components.Max(c => c.H), maxV = components.Count == 1 ? 1 : components.Max(c => c.V);
        int[] hs = components.Select(c => components.Count == 1 ? 1 : c.H).ToArray(), vs = components.Select(c => components.Count == 1 ? 1 : c.V).ToArray();
        int mcuW = 8 * maxH, mcuH = 8 * maxV;
        if (tileWidth % mcuW != 0 || tileHeight % mcuH != 0) return null;
        int tileMcusX = tileWidth / mcuW, tileMcusY = tileHeight / mcuH;
        int mcuCols = (width + mcuW - 1) / mcuW, mcuRows = (height + mcuH - 1) / mcuH;

        // The coefficients of every block of every component, in the grid of the whole image.
        var blocks = new short[components.Count][];
        var blockCols = new int[components.Count];
        for (int c = 0; c < components.Count; c++)
        {
            blockCols[c] = mcuCols * hs[c];
            long total = (long)blockCols[c] * mcuRows * vs[c] * 64;
            if (total > 400_000_000) return null;
            blocks[c] = new short[total];
        }

        for (int t = 0; t < tiles.Count; t++)
        {
            var info = t == 0 ? first : Parse(tiles[t]);
            if (info is null || info.Components.Count != components.Count) return null;
            if (info.Components.Zip(components).Any(pair => pair.First.Id != pair.Second.Id || pair.First.H != pair.Second.H || pair.First.V != pair.Second.V)) return null;
            var tables = t == 0 ? first.Tables : info.Tables;
            var reader = new BitReader(tiles[t], info.DataAt);
            var pred = new int[components.Count];
            int tx = t % across, ty = t / across;
            for (int m = 0; m < tileMcusX * tileMcusY; m++)
            {
                if (info.Interval > 0 && m > 0 && m % info.Interval == 0) { reader.Restart(); Array.Clear(pred); }
                int gx = tx * tileMcusX + m % tileMcusX, gy = ty * tileMcusY + m / tileMcusX;
                bool inside = gx < mcuCols && gy < mcuRows;
                for (int c = 0; c < components.Count; c++)
                {
                    var comp = info.Components[c];
                    if (!tables.TryGetValue(comp.DcTable, out var dc) || !tables.TryGetValue(0x10 | comp.AcTable, out var ac)) return null;
                    for (int v = 0; v < vs[c]; v++)
                        for (int h = 0; h < hs[c]; h++)
                        {
                            var block = new short[64];
                            int s = reader.Symbol(dc);
                            pred[c] += Extend(reader.Receive(s), s);
                            block[0] = (short)pred[c];
                            for (int k = 1; k < 64;)
                            {
                                int rs = reader.Symbol(ac), r = rs >> 4, bits = rs & 15;
                                if (bits == 0) { if (r == 15) { k += 16; continue; } break; }
                                k += r;
                                if (k > 63) throw new InvalidDataException("Bad coefficient index.");
                                block[k++] = (short)Extend(reader.Receive(bits), bits);
                            }
                            if (!inside) continue;
                            int index = ((gy * vs[c] + v) * blockCols[c] + gx * hs[c] + h) * 64;
                            block.CopyTo(blocks[c], index);
                        }
                }
            }
        }

        // The new scan: every MCU of the whole image in raster order, coded with the first tile's tables.
        var writer = new BitWriter();
        var last = new int[components.Count];
        for (int my = 0; my < mcuRows; my++)
            for (int mx = 0; mx < mcuCols; mx++)
                for (int c = 0; c < components.Count; c++)
                {
                    var comp = components[c];
                    var dc = first.Tables[comp.DcTable];
                    var ac = first.Tables[0x10 | comp.AcTable];
                    for (int v = 0; v < vs[c]; v++)
                        for (int h = 0; h < hs[c]; h++)
                        {
                            int at = ((my * vs[c] + v) * blockCols[c] + mx * hs[c] + h) * 64;
                            var b = blocks[c].AsSpan(at, 64);
                            int diff = b[0] - last[c];
                            last[c] = b[0];
                            int size = SizeOf(diff);
                            if (dc.Size[size] == 0) return null;
                            writer.Put(dc.Code[size], dc.Size[size]);
                            if (size > 0) writer.Put(diff < 0 ? diff + (1 << size) - 1 : diff, size);
                            int run = 0;
                            for (int k = 1; k < 64; k++)
                            {
                                if (b[k] == 0) { run++; continue; }
                                while (run > 15)
                                {
                                    if (ac.Size[0xF0] == 0) return null;
                                    writer.Put(ac.Code[0xF0], ac.Size[0xF0]);
                                    run -= 16;
                                }
                                int bits = SizeOf(b[k]), symbol = run << 4 | bits;
                                if (ac.Size[symbol] == 0) return null;
                                writer.Put(ac.Code[symbol], ac.Size[symbol]);
                                writer.Put(b[k] < 0 ? b[k] + (1 << bits) - 1 : b[k], bits);
                                run = 0;
                            }
                            if (run > 0)
                            {
                                if (ac.Size[0] == 0) return null;
                                writer.Put(ac.Code[0], ac.Size[0]);
                            }
                        }
                }
        var scan = writer.Finish();

        var tile0 = tiles[0];
        using var output = new MemoryStream();
        output.Write(tile0, 0, 2);
        foreach (var (at, length) in first.Segments)
        {
            if (tile0[at + 1] == 0xDD) continue; // no restart intervals any more
            int start = (int)output.Length;
            output.Write(tile0, at, length);
            if (at == first.FrameAt)
            {
                var buffer = output.GetBuffer();
                buffer[start + 5] = (byte)(height >> 8); buffer[start + 6] = (byte)height;
                buffer[start + 7] = (byte)(width >> 8); buffer[start + 8] = (byte)width;
            }
        }
        output.Write(tile0, first.ScanAt, first.DataAt - first.ScanAt);
        output.Write(scan);
        output.Write([0xFF, 0xD9]);
        return output.ToArray();
    }
}
