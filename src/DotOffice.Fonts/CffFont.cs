using System.Buffers.Binary;

namespace DotOffice.Fonts;

/// <summary>
/// The PostScript outlines (<c>CFF </c> table) of an OpenType font. Outlines are never interpreted: the table is only parsed far enough to
/// learn how glyph ids map to the character ids a PDF uses, and to cut it down to the glyphs a document draws.
/// </summary>
internal sealed class CffFont
{
    private const int Charset = 15, Encoding = 16, CharStrings = 17, Private = 18, Subrs = 19, Ros = 1230, FdArray = 1236, FdSelect = 1237;

    private readonly ReadOnlyMemory<byte> _memory;
    private readonly int _numGlyphs;
    private readonly int _headerSize;
    private readonly (int Start, int End) _names, _strings;
    private readonly List<DictEntry> _top;
    private readonly ushort[] _charset;
    private readonly int _charsetLength;

    private readonly record struct DictEntry(int Op, double[] Operands, byte[] Raw);

    public CffFont(ReadOnlyMemory<byte> table, int numGlyphs)
    {
        _memory = table;
        _numGlyphs = numGlyphs;
        var d = table.Span;
        if (d.Length < 4 || d[0] != 1) throw new InvalidDataException("Unsupported CFF table.");
        _headerSize = d[2];
        var name = ReadIndex(d, _headerSize);
        var top = ReadIndex(d, name.End);
        var strings = ReadIndex(d, top.End);
                _names = (_headerSize, name.End);
        _strings = (top.End, strings.End);
        if (top.Offsets.Length < 2) throw new InvalidDataException("CFF font without a top DICT.");
        _top = ReadDict(d[top.Offsets[0]..top.Offsets[1]]);
        IsCid = _top.Any(e => e.Op == Ros);

        _charset = new ushort[numGlyphs];
        int offset = (int)Operand(Charset, 0);
        if (offset <= 2) { for (int g = 0; g < numGlyphs; g++) _charset[g] = (ushort)g; }
        else _charsetLength = ReadCharset(d, offset);
    }

    /// <summary>True for a CID-keyed font, where the character id of a glyph is read from the charset instead of being the glyph id.</summary>
    public bool IsCid { get; }

    /// <summary>The character id the PDF uses for a glyph.</summary>
    public ushort Cid(ushort glyph) => IsCid && glyph < _charset.Length ? _charset[glyph] : glyph;

    /// <summary>A copy of the table in which every glyph outside <paramref name="glyphs"/> (and .notdef) is empty.</summary>
    public byte[] Subset(IEnumerable<ushort> glyphs)
    {
        var d = _memory.Span;
        var keep = new HashSet<ushort>(glyphs) { 0 };
        var charStrings = ReadIndex(d, (int)Operand(CharStrings, 0));
        int charset = (int)Operand(Charset, 0), fdSelect = (int)Operand(FdSelect, -1);
        var charsetBytes = charset > 2 ? d.Slice(charset, _charsetLength).ToArray() : [];
        var fdSelectBytes = fdSelect > 0 ? d.Slice(fdSelect, FdSelectLength(d, fdSelect)).ToArray() : [];

        // Private dictionaries (with the local subroutines that follow them) are copied whole, each after the font dictionaries.
        var fonts = new List<List<DictEntry>>();
        var blocks = new List<byte[]>();
        var localSubrs = new List<int[]>();
        if (IsCid)
        {
            var array = ReadIndex(d, (int)Operand(FdArray, 0));
            for (int i = 0; i + 1 < array.Offsets.Length; i++)
            {
                var font = ReadDict(d[array.Offsets[i]..array.Offsets[i + 1]]);
                fonts.Add(font);
                var (block, subrs) = PrivateBlock(d, font);
                blocks.Add(block);
                localSubrs.Add(subrs);
            }
        }
        else
        {
            var (block, subrs) = PrivateBlock(d, _top);
            blocks.Add(block);
            localSubrs.Add(subrs);
        }

        // Subroutines are inlined into the glyphs that are kept and dropped from the font: they would be most of a CJK font's size.
        var globalSubrs = ReadIndex(d, _strings.End).Offsets;
        var strings = new List<byte[]>(_numGlyphs);
        for (int g = 0; g < _numGlyphs && g + 1 < charStrings.Offsets.Length; g++)
            strings.Add(keep.Contains((ushort)g)
                ? Inline(_memory, charStrings.Offsets[g], charStrings.Offsets[g + 1], globalSubrs, localSubrs[Math.Min(FdOf(d, fdSelect, g), localSubrs.Count - 1)])
                : [0x0E]); // endchar

        byte[] Top(int charsetAt, int charStringsAt, int fdArrayAt, int fdSelectAt, int privateAt) =>
            Dict(_top.Where(e => e.Op != Encoding), e => e.Op switch
            {
                Charset when charset > 2 => Fixed(charsetAt, e.Op),
                CharStrings => Fixed(charStringsAt, e.Op),
                FdArray => Fixed(fdArrayAt, e.Op),
                FdSelect => Fixed(fdSelectAt, e.Op),
                Private => Fixed(blocks[0].Length, privateAt, e.Op),
                _ => null,
            });

        byte[] FontDicts(int firstBlockAt)
        {
            var parts = new List<byte[]>();
            int at = firstBlockAt;
            for (int i = 0; i < fonts.Count; i++)
            {
                int blockAt = at;
                parts.Add(Dict(fonts[i], e => e.Op == Private ? Fixed(blocks[i].Length, blockAt, e.Op) : null));
                at += blocks[i].Length;
            }
            return Index(parts);
        }

        var header = d[.._headerSize].ToArray();
        var name = d[_names.Start.._names.End].ToArray();
        // The top dictionary index sits between the names and the strings; copy the strings and global subroutines as they are.
        byte[] stringsAndSubrs = [.. d[_strings.Start.._strings.End], 0, 0]; // strings, then an empty global subroutine index
        int topLength = Index([Top(0, 0, 0, 0, 0)]).Length;
        int at0 = header.Length + name.Length + topLength + stringsAndSubrs.Length;
        int charsetAt = at0, fdSelectAt = charsetAt + charsetBytes.Length, charStringsAt = fdSelectAt + fdSelectBytes.Length;
        var charStringsBytes = Index(strings);
        int fdArrayAt = charStringsAt + charStringsBytes.Length;
        int firstBlockAt = fdArrayAt + (IsCid ? FontDicts(0).Length : 0);

        using var output = new MemoryStream();
        output.Write(header);
        output.Write(name);
        output.Write(Index([Top(charsetAt, charStringsAt, fdArrayAt, fdSelectAt, firstBlockAt)]));
        output.Write(stringsAndSubrs);
        output.Write(charsetBytes);
        output.Write(fdSelectBytes);
        output.Write(charStringsBytes);
        if (IsCid) output.Write(FontDicts(firstBlockAt));
        foreach (var block in blocks) output.Write(block);
        return output.ToArray();
    }

    /// <summary>The length in bytes of a glyph's charstring.</summary>
    internal int CharStringLength(int glyph)
    {
        var index = ReadIndex(_memory.Span, (int)Operand(CharStrings, 0));
        return index.Offsets[glyph + 1] - index.Offsets[glyph];
    }

    private double Operand(int op, double fallback)
    {
        foreach (var e in _top)
            if (e.Op == op && e.Operands.Length > 0) return e.Operands[0];
        return fallback;
    }

    /// <summary>The private dictionary named by <paramref name="dict"/> without its subroutine entry, and the starts of its local subroutines.</summary>
    private static (byte[] Block, int[] Subrs) PrivateBlock(ReadOnlySpan<byte> d, List<DictEntry> dict)
    {
        var entry = dict.FirstOrDefault(e => e.Op == Private);
        if (entry.Operands is not { Length: 2 }) return ([], []);
        int size = (int)entry.Operands[0], start = (int)entry.Operands[1];
        var privateDict = ReadDict(d.Slice(start, size));
        int[] subrs = privateDict.FirstOrDefault(e => e.Op == Subrs).Operands is { Length: 1 } s ? ReadIndex(d, start + (int)s[0]).Offsets : [];
        return (Dict(privateDict.Where(e => e.Op != Subrs), _ => null), subrs);
    }

    private int FdOf(ReadOnlySpan<byte> d, int fdSelect, int glyph)
    {
        if (fdSelect <= 0) return 0;
        if (d[fdSelect] == 0) return d[fdSelect + 1 + glyph];
        int ranges = BinaryPrimitives.ReadUInt16BigEndian(d[(fdSelect + 1)..]);
        for (int i = 0; i < ranges; i++)
        {
            int at = fdSelect + 3 + i * 3, first = BinaryPrimitives.ReadUInt16BigEndian(d[at..]);
            int next = BinaryPrimitives.ReadUInt16BigEndian(d[(at + 3)..]);
            if (glyph >= first && glyph < next) return d[at + 2];
        }
        return 0;
    }

    /// <summary>A Type 2 charstring with its <c>callsubr</c>/<c>callgsubr</c> calls replaced by the subroutines' bodies.</summary>
    private static byte[] Inline(ReadOnlyMemory<byte> all, int start, int end, int[] globals, int[] locals)
    {
        var output = new List<byte>((end - start) * 2);
        int operands = 0, stems = 0, lastStart = -1, lastValue = 0;
        bool done = false;
        void Run(int from, int to, int depth)
        {
            var code = all.Span[from..to];
            for (int i = 0; i < code.Length && !done;)
            {
                int b = code[i];
                if (b >= 32 || b == 28)
                {
                    int length = b == 28 ? 3 : b <= 246 ? 1 : b <= 254 ? 2 : 5;
                    if (i + length > code.Length) { done = true; return; }
                    lastStart = output.Count;
                    lastValue = b == 28 ? BinaryPrimitives.ReadInt16BigEndian(code[(i + 1)..]) : b <= 246 ? b - 139 : b <= 250 ? (b - 247) * 256 + code[i + 1] + 108 : b <= 254 ? -(b - 251) * 256 - code[i + 1] - 108 : 0;
                    for (int k = 0; k < length; k++) output.Add(code[i + k]);
                    operands++;
                    i += length;
                    continue;
                }
                i++;
                switch (b)
                {
                    case 10 or 29:
                        var table = b == 10 ? locals : globals;
                        int count = Math.Max(table.Length - 1, 0), bias = count < 1240 ? 107 : count < 33900 ? 1131 : 32768;
                        int index = lastValue + bias;
                        if (operands == 0 || lastStart < 0 || depth >= 10 || index < 0 || index >= count) { done = true; return; }
                        output.RemoveRange(lastStart, output.Count - lastStart);
                        operands--;
                        lastStart = -1;
                        Run(table[index], table[index + 1], depth + 1);
                        break;
                    case 11:
                        return;
                    case 14:
                        output.Add(14);
                        done = true;
                        return;
                    case 1 or 3 or 18 or 23:
                        stems += operands / 2;
                        operands = 0;
                        output.Add((byte)b);
                        break;
                    case 19 or 20:
                        stems += operands / 2;
                        operands = 0;
                        output.Add((byte)b);
                        for (int k = 0; k < (stems + 7) / 8 && i < code.Length; k++) output.Add(code[i++]);
                        break;
                    case 12:
                        operands = 0;
                        output.Add(12);
                        if (i < code.Length) output.Add(code[i++]);
                        break;
                    default:
                        operands = 0;
                        output.Add((byte)b);
                        break;
                }
            }
        }
        Run(start, end, 0);
        if (output.Count == 0 || output[^1] != 14) output.Add(14);
        return output.ToArray();
    }

    /// <summary>Writes the entries in order, replacing the ones <paramref name="patch"/> gives new bytes for.</summary>
    private static byte[] Dict(IEnumerable<DictEntry> entries, Func<DictEntry, byte[]?> patch)
    {
        using var output = new MemoryStream();
        foreach (var e in entries) output.Write(patch(e) ?? e.Raw);
        return output.ToArray();
    }

    private static byte[] Fixed(int value, int op) => [.. FiveByte(value), .. Operator(op)];

    private static byte[] Fixed(int size, int offset, int op) => [.. FiveByte(size), .. FiveByte(offset), .. Operator(op)];

    private static byte[] FiveByte(int value)
    {
        var b = new byte[5];
        b[0] = 29;
        BinaryPrimitives.WriteInt32BigEndian(b.AsSpan(1), value);
        return b;
    }

    private static byte[] Operator(int op) => op >= 1200 ? [12, (byte)(op - 1200)] : [(byte)op];

    private static byte[] Index(List<byte[]> items)
    {
        using var output = new MemoryStream();
        output.WriteByte((byte)(items.Count >> 8));
        output.WriteByte((byte)items.Count);
        if (items.Count == 0) return output.ToArray();
        output.WriteByte(4);
        uint at = 1;
        Span<byte> word = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(word, at);
        output.Write(word);
        foreach (var item in items)
        {
            at += (uint)item.Length;
            BinaryPrimitives.WriteUInt32BigEndian(word, at);
            output.Write(word);
        }
        foreach (var item in items) output.Write(item);
        return output.ToArray();
    }

    /// <summary>Reads an INDEX at <paramref name="pos"/>: the absolute start of each element plus the end of the last, and where the INDEX ends.</summary>
    private static (int[] Offsets, int End) ReadIndex(ReadOnlySpan<byte> d, int pos)
    {
        if (pos < 0 || pos + 2 > d.Length) throw new InvalidDataException("CFF INDEX out of range.");
        int count = BinaryPrimitives.ReadUInt16BigEndian(d[pos..]);
        if (count == 0) return ([], pos + 2);
        int offSize = d[pos + 2];
        if (offSize is < 1 or > 4) throw new InvalidDataException("Bad CFF offset size.");
        int dataStart = pos + 3 + (count + 1) * offSize - 1;
        var offsets = new int[count + 1];
        for (int i = 0; i <= count; i++)
        {
            int at = pos + 3 + i * offSize, v = 0;
            for (int k = 0; k < offSize; k++) v = (v << 8) | d[at + k];
            offsets[i] = dataStart + v;
            if (offsets[i] > d.Length) throw new InvalidDataException("CFF INDEX out of range.");
        }
        return (offsets, offsets[count]);
    }

    private static List<DictEntry> ReadDict(ReadOnlySpan<byte> d)
    {
        var entries = new List<DictEntry>();
        var operands = new List<double>();
        int entryStart = 0;
        for (int i = 0; i < d.Length;)
        {
            int b = d[i];
            if (b <= 21)
            {
                int op = b;
                i++;
                if (b == 12 && i < d.Length) op = 1200 + d[i++];
                entries.Add(new DictEntry(op, operands.ToArray(), d[entryStart..i].ToArray()));
                operands.Clear();
                entryStart = i;
            }
            else if (b == 28) { operands.Add(BinaryPrimitives.ReadInt16BigEndian(d[(i + 1)..])); i += 3; }
            else if (b == 29) { operands.Add(BinaryPrimitives.ReadInt32BigEndian(d[(i + 1)..])); i += 5; }
            else if (b == 30)
            {
                i++;
                while (i < d.Length && (d[i] & 0x0F) != 0x0F && (d[i] >> 4) != 0x0F) i++;
                i++;
                operands.Add(double.NaN);
            }
            else if (b >= 32 && b <= 246) { operands.Add(b - 139); i++; }
            else if (b >= 247 && b <= 250) { operands.Add((b - 247) * 256 + d[i + 1] + 108); i += 2; }
            else if (b >= 251 && b <= 254) { operands.Add(-(b - 251) * 256 - d[i + 1] - 108); i += 2; }
            else throw new InvalidDataException("Bad CFF DICT byte.");
        }
        return entries;
    }

    /// <summary>Fills the glyph-to-id table from a custom charset at <paramref name="pos"/> and returns the charset's length in bytes.</summary>
    private int ReadCharset(ReadOnlySpan<byte> d, int pos)
    {
        int format = d[pos], at = pos + 1, g = 1;
        while (g < _numGlyphs)
        {
            if (format == 0) { _charset[g++] = BinaryPrimitives.ReadUInt16BigEndian(d[at..]); at += 2; }
            else
            {
                int first = BinaryPrimitives.ReadUInt16BigEndian(d[at..]), left;
                if (format == 1) { left = d[at + 2]; at += 3; }
                else { left = BinaryPrimitives.ReadUInt16BigEndian(d[(at + 2)..]); at += 4; }
                for (int k = 0; k <= left && g < _numGlyphs; k++) _charset[g++] = (ushort)(first + k);
            }
        }
        return at - pos;
    }

    private int FdSelectLength(ReadOnlySpan<byte> d, int pos) =>
        d[pos] == 0 ? 1 + _numGlyphs : 1 + 2 + BinaryPrimitives.ReadUInt16BigEndian(d[(pos + 1)..]) * 3 + 2;
}
