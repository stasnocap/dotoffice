using System.Buffers.Binary;
using System.Text;

namespace DotOffice.Fonts;

/// <summary>
/// A parsed TrueType (glyf-based) font. Holds the raw font bytes and exposes the metrics
/// and tables needed for text measurement, PDF embedding and subsetting.
/// </summary>
public sealed class TrueTypeFont
{
    private readonly byte[] _data;
    private readonly Dictionary<string, (int Offset, int Length)> _tables;
    private readonly Dictionary<int, ushort> _cmap;
    private readonly ushort[] _advanceWidths;
    private readonly int[] _loca;

    public string PostScriptName { get; }
    public string FamilyName { get; }
    public int UnitsPerEm { get; }
    public int NumGlyphs { get; }
    public short Ascender { get; }
    public short Descender { get; }
    public short LineGap { get; }
    public short CapHeight { get; }
    public (short XMin, short YMin, short XMax, short YMax) BoundingBox { get; }
    public double ItalicAngle { get; }
    public bool IsFixedPitch { get; }
    public bool IsBold { get; }
    public bool IsItalic { get; }
    public ushort WeightClass { get; }

    /// <summary>Style name from the name table (e.g. "Bold Italic").</summary>
    public string SubfamilyName { get; }

    /// <summary>Typographic family (name id 16) when present, otherwise <see cref="FamilyName"/>.</summary>
    public string TypographicFamilyName { get; }

    /// <summary>Underline offset from the baseline in font units (negative is below).</summary>
    public short UnderlinePosition { get; }
    public short UnderlineThickness { get; }

    /// <summary>Strikeout offset above the baseline in font units.</summary>
    public short StrikeoutPosition { get; }
    public short StrikeoutSize { get; }

    /// <summary>True when the OS/2 fsType forbids embedding (restricted license).</summary>
    public bool IsEmbeddingRestricted { get; }

    private TrueTypeFont(byte[] data)
    {
        _data = data;
        var sfntVersion = U32(0);
        if (sfntVersion == 0x4F54544F) // 'OTTO'
            throw new NotSupportedException("CFF-based OpenType fonts are not supported yet.");
        if (sfntVersion == 0x74746366) // 'ttcf'
            throw new NotSupportedException("TrueType collections are not supported yet.");
        if (sfntVersion != 0x00010000 && sfntVersion != 0x74727565) // 'true'
            throw new InvalidDataException("Not a TrueType font.");

        int numTables = U16(4);
        _tables = new Dictionary<string, (int, int)>(numTables);
        for (int i = 0; i < numTables; i++)
        {
            int rec = 12 + i * 16;
            var tag = Encoding.ASCII.GetString(data, rec, 4);
            int offset = (int)U32(rec + 8);
            int length = (int)U32(rec + 12);
            if (offset < 0 || length < 0 || (long)offset + length > data.Length)
                throw new InvalidDataException($"Table '{tag}' is out of bounds.");
            _tables[tag] = (offset, length);
        }

        foreach (var required in new[] { "head", "hhea", "maxp", "hmtx", "cmap", "loca", "glyf" })
            if (!_tables.ContainsKey(required))
                throw new InvalidDataException($"Missing required table '{required}'.");

        int head = Table("head");
        UnitsPerEm = U16(head + 18);
        BoundingBox = (S16(head + 36), S16(head + 38), S16(head + 40), S16(head + 42));
        int macStyle = U16(head + 44);
        bool longLoca = S16(head + 50) == 1;

        NumGlyphs = U16(Table("maxp") + 4);

        int hhea = Table("hhea");
        Ascender = S16(hhea + 4);
        Descender = S16(hhea + 6);
        LineGap = S16(hhea + 8);
        int numberOfHMetrics = U16(hhea + 34);

        _advanceWidths = new ushort[NumGlyphs];
        int hmtx = Table("hmtx");
        ushort last = 0;
        for (int g = 0; g < NumGlyphs; g++)
        {
            if (g < numberOfHMetrics) last = (ushort)U16(hmtx + g * 4);
            _advanceWidths[g] = last;
        }

        _loca = new int[NumGlyphs + 1];
        int loca = Table("loca");
        for (int g = 0; g <= NumGlyphs; g++)
            _loca[g] = longLoca ? (int)U32(loca + g * 4) : U16(loca + g * 2) * 2;

        _cmap = ParseCmap(Table("cmap"));

        CapHeight = (short)(Ascender * 0.7);
        StrikeoutPosition = (short)(Ascender * 0.3);
        StrikeoutSize = (short)(UnitsPerEm / 20);
        WeightClass = 400;
        if (_tables.TryGetValue("OS/2", out var os2))
        {
            int o = os2.Offset;
            int version = U16(o);
            WeightClass = (ushort)U16(o + 4);
            int fsType = U16(o + 8);
            IsEmbeddingRestricted = (fsType & 0x000F) == 0x0002;
            if (os2.Length >= 30)
            {
                StrikeoutSize = S16(o + 26);
                StrikeoutPosition = S16(o + 28);
            }
            if (version >= 2 && os2.Length >= 90) CapHeight = S16(o + 88);
        }

        if (_tables.TryGetValue("post", out var post))
        {
            ItalicAngle = (int)U32(post.Offset + 4) / 65536.0;
            IsFixedPitch = U32(post.Offset + 12) != 0;
            UnderlinePosition = S16(post.Offset + 8);
            UnderlineThickness = S16(post.Offset + 10);
        }
        if (UnderlineThickness <= 0)
        {
            UnderlinePosition = (short)(-UnitsPerEm / 10);
            UnderlineThickness = (short)(UnitsPerEm / 20);
        }

        IsBold = (macStyle & 1) != 0 || WeightClass >= 700;
        IsItalic = (macStyle & 2) != 0 || ItalicAngle != 0;

        var names = FontNames.Parse(GetTableData("name"));
        FamilyName = names.GetValueOrDefault(1) ?? "Unknown";
        SubfamilyName = names.GetValueOrDefault(2) ?? "Regular";
        TypographicFamilyName = names.GetValueOrDefault(16) ?? FamilyName;
        PostScriptName = SanitizePostScriptName(names.GetValueOrDefault(6) ?? FamilyName);
    }

    public static TrueTypeFont Load(byte[] data) => new(data);

    public static TrueTypeFont Load(string path) => new(File.ReadAllBytes(path));

    /// <summary>Returns the glyph id for a Unicode code point, or 0 (.notdef) when unmapped.</summary>
    public ushort GetGlyphId(int codePoint) => _cmap.GetValueOrDefault(codePoint);

    public bool HasGlyph(int codePoint) => _cmap.ContainsKey(codePoint);

    /// <summary>Advance width in font units.</summary>
    public int GetAdvanceWidth(ushort glyphId) => glyphId < NumGlyphs ? _advanceWidths[glyphId] : 0;

    /// <summary>Measures text width in points at the given size (no kerning or shaping).</summary>
    public double MeasureText(string text, double fontSize)
    {
        long units = 0;
        foreach (var rune in text.EnumerateRunes())
            units += GetAdvanceWidth(GetGlyphId(rune.Value));
        return units * fontSize / UnitsPerEm;
    }

    internal ReadOnlySpan<byte> GetTableData(string tag) =>
        _tables.TryGetValue(tag, out var t) ? _data.AsSpan(t.Offset, t.Length) : default;

    internal bool HasTable(string tag) => _tables.ContainsKey(tag);

    internal ReadOnlySpan<byte> GetGlyphData(ushort glyphId)
    {
        int start = _loca[glyphId], end = _loca[glyphId + 1];
        if (end <= start) return default;
        return GetTableData("glyf").Slice(start, end - start);
    }

    private int Table(string tag) => _tables[tag].Offset;

    private int U16(int offset) => BinaryPrimitives.ReadUInt16BigEndian(_data.AsSpan(offset));
    private short S16(int offset) => BinaryPrimitives.ReadInt16BigEndian(_data.AsSpan(offset));
    private uint U32(int offset) => BinaryPrimitives.ReadUInt32BigEndian(_data.AsSpan(offset));

    private Dictionary<int, ushort> ParseCmap(int cmap)
    {
        int numSubtables = U16(cmap + 2);
        int best = -1, bestScore = 0;
        for (int i = 0; i < numSubtables; i++)
        {
            int rec = cmap + 4 + i * 8;
            int platform = U16(rec), encoding = U16(rec + 2);
            int sub = cmap + (int)U32(rec + 4);
            int format = U16(sub);
            int score = (platform, encoding, format) switch
            {
                (3, 10, 12) => 4,
                (0, _, 12) => 3,
                (3, 1, 4) => 2,
                (0, _, 4) => 1,
                _ => 0,
            };
            if (score > bestScore) (best, bestScore) = (sub, score);
        }

        var map = new Dictionary<int, ushort>();
        if (best < 0) return map;

        if (U16(best) == 12)
        {
            int groups = (int)U32(best + 12);
            for (int i = 0; i < groups; i++)
            {
                int g = best + 16 + i * 12;
                uint start = U32(g), end = U32(g + 4), glyph = U32(g + 8);
                for (uint c = start; c <= end && c <= 0x10FFFF; c++)
                    map[(int)c] = (ushort)(glyph + (c - start));
            }
            return map;
        }

        int segCount = U16(best + 6) / 2;
        int ends = best + 14;
        int starts = ends + segCount * 2 + 2;
        int deltas = starts + segCount * 2;
        int rangeOffsets = deltas + segCount * 2;
        for (int i = 0; i < segCount; i++)
        {
            int end = U16(ends + i * 2), start = U16(starts + i * 2);
            int delta = S16(deltas + i * 2);
            int rangeOffsetPos = rangeOffsets + i * 2;
            int rangeOffset = U16(rangeOffsetPos);
            for (int c = start; c <= end && c != 0xFFFF; c++)
            {
                int glyph;
                if (rangeOffset == 0)
                {
                    glyph = (c + delta) & 0xFFFF;
                }
                else
                {
                    int addr = rangeOffsetPos + rangeOffset + (c - start) * 2;
                    int raw = U16(addr);
                    glyph = raw == 0 ? 0 : (raw + delta) & 0xFFFF;
                }
                if (glyph != 0) map[c] = (ushort)glyph;
            }
        }
        return map;
    }

    private static string SanitizePostScriptName(string name)
    {
        var sb = new StringBuilder(name.Length);
        foreach (char c in name)
            if (c > 32 && c < 127 && "[](){}<>/%".IndexOf(c) < 0) sb.Append(c);
        return sb.Length > 0 ? sb.ToString() : "Font";
    }
}
