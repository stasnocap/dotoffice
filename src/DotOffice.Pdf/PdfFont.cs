using System.Text;
using DotOffice.Fonts;

namespace DotOffice.Pdf;

/// <summary>
/// A TrueType font embedded as a Type0/CIDFontType2 font with Identity-H encoding.
/// Glyph usage is tracked while drawing so only used glyphs are embedded.
/// </summary>
/// <remarks>
/// Character codes (CIDs) normally equal glyph ids. Shaped text can use one glyph for different characters
/// (e.g. a dotless letter body shared by several Arabic letters) or for no character at all (extra glyphs of
/// a decomposition), so each distinct (glyph, text) pair gets its own CID; a CIDToGIDMap stream maps the
/// extra CIDs back to their glyphs and ToUnicode stays unambiguous.
/// </remarks>
public sealed class PdfFont
{
    private readonly Dictionary<(ushort Glyph, string Text), ushort> _cids = [];
    private readonly SortedDictionary<ushort, (ushort Glyph, string Text)> _usedCids = new();
    private int _nextCid;

    internal PdfFont(TrueTypeFont font, string resourceName, bool embedded = true)
    {
        Font = font;
        ResourceName = resourceName;
        IsEmbedded = embedded;
        _nextCid = font.NumGlyphs;
    }

    public TrueTypeFont Font { get; }

    public bool IsEmbedded { get; }

    internal string ResourceName { get; }

    public double MeasureText(string text, double fontSize) => Font.MeasureText(text, fontSize);

    /// <summary>Maps text to glyphs, records them as used and returns the hex string operand for Tj.</summary>
    internal string Encode(string text)
    {
        var sb = new StringBuilder(text.Length * 4 + 2).Append('<');
        foreach (var rune in text.EnumerateRunes())
            sb.Append(EncodeGlyph(Font.GetGlyphId(rune.Value), rune.ToString()));
        return sb.Append('>').ToString();
    }

    /// <summary>
    /// Records a glyph as used for <paramref name="text"/> (what it stands for in extracted text, possibly
    /// empty) and returns the hex code of its CID.
    /// </summary>
    internal string EncodeGlyph(ushort glyphId, string text) => Cid(glyphId, text).ToString("X4");

    private ushort Cid(ushort glyph, string text)
    {
        if (glyph == 0) return 0;
        if (_cids.TryGetValue((glyph, text), out var cid)) return cid;
        // The glyph's own id is used for its first text; other texts get CIDs past the last glyph.
        if (!_usedCids.ContainsKey(glyph)) cid = glyph;
        else if (_nextCid <= ushort.MaxValue) cid = (ushort)_nextCid++;
        else return glyph; // no CIDs left: share the code (extraction may pick the other text)
        _cids[(glyph, text)] = cid;
        _usedCids[cid] = (glyph, text);
        return cid;
    }

    internal void Write(PdfWriter writer, int type0Id)
    {
        int cidFontId = writer.Allocate();
        int descriptorId = writer.Allocate();
        int fontFileId = IsEmbedded ? writer.Allocate() : 0;
        int toUnicodeId = writer.Allocate();
        bool identity = _usedCids.All(c => c.Key == c.Value.Glyph);
        int cidToGidId = identity ? 0 : writer.Allocate();

        var baseFont = IsEmbedded ? SubsetTag() + "+" + Font.PostScriptName : Font.PostScriptName;
        double scale = 1000.0 / Font.UnitsPerEm;

        writer.WriteObject(type0Id,
            $"<< /Type /Font /Subtype /Type0 /BaseFont /{baseFont} /Encoding /Identity-H " +
            $"/DescendantFonts [{cidFontId} 0 R] /ToUnicode {toUnicodeId} 0 R >>");

        writer.WriteObject(cidFontId,
            $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{baseFont} " +
            "/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> " +
            $"/FontDescriptor {descriptorId} 0 R /DW {PdfFormat.Number(Font.GetAdvanceWidth(0) * scale)} " +
            $"/W {WidthsArray(scale)} /CIDToGIDMap {(cidToGidId != 0 ? $"{cidToGidId} 0 R" : "/Identity")} >>");

        var bbox = Font.BoundingBox;
        int flags = 4 | (Font.IsFixedPitch ? 1 : 0) | (Font.IsItalic ? 64 : 0);
        var fontFile = IsEmbedded ? $" /FontFile2 {fontFileId} 0 R" : "";
        writer.WriteObject(descriptorId,
            $"<< /Type /FontDescriptor /FontName /{baseFont} /Flags {flags} " +
            $"/FontBBox [{PdfFormat.Number(bbox.XMin * scale)} {PdfFormat.Number(bbox.YMin * scale)} " +
            $"{PdfFormat.Number(bbox.XMax * scale)} {PdfFormat.Number(bbox.YMax * scale)}] " +
            $"/ItalicAngle {PdfFormat.Number(Font.ItalicAngle)} /Ascent {PdfFormat.Number(Font.Ascender * scale)} " +
            $"/Descent {PdfFormat.Number(Font.Descender * scale)} /CapHeight {PdfFormat.Number(Font.CapHeight * scale)} " +
            $"/StemV {(Font.IsBold ? 120 : 80)}{fontFile} >>");

        if (!identity)
        {
            var map = new byte[(_usedCids.Keys.Max() + 1) * 2];
            for (int cid = 0; cid * 2 < map.Length; cid++)
            {
                ushort glyph = _usedCids.TryGetValue((ushort)cid, out var used) ? used.Glyph : cid < Font.NumGlyphs ? (ushort)cid : (ushort)0;
                map[cid * 2] = (byte)(glyph >> 8);
                map[cid * 2 + 1] = (byte)glyph;
            }
            writer.WriteStream(cidToGidId, "", map);
        }
        if (IsEmbedded)
        {
            var subset = TrueTypeSubsetter.Subset(Font, _usedCids.Values.Select(v => v.Glyph));
            writer.WriteStream(fontFileId, $"/Length1 {subset.Length}", subset);
        }
        writer.WriteStream(toUnicodeId, "", Encoding.ASCII.GetBytes(ToUnicodeCMap()));
    }

    private string WidthsArray(double scale)
    {
        var sb = new StringBuilder("[");
        int? runStart = null;
        int previous = -2;
        foreach (var (cid, (glyph, _)) in _usedCids)
        {
            if (cid != previous + 1)
            {
                if (runStart is not null) sb.Append("] ");
                sb.Append(cid).Append(" [");
                runStart = cid;
            }
            else sb.Append(' ');
            sb.Append(PdfFormat.Number(Font.GetAdvanceWidth(glyph) * scale));
            previous = cid;
        }
        if (runStart is not null) sb.Append(']');
        return sb.Append(']').ToString();
    }

    private string ToUnicodeCMap()
    {
        var sb = new StringBuilder();
        sb.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n");
        sb.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
        sb.Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n");
        sb.Append("1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n");
        foreach (var chunk in _usedCids.Where(c => c.Value.Text.Length > 0).Chunk(100))
        {
            sb.Append(chunk.Length).Append(" beginbfchar\n");
            foreach (var (cid, (_, text)) in chunk)
                sb.Append('<').Append(cid.ToString("X4")).Append("> <")
                  .Append(Convert.ToHexString(Encoding.BigEndianUnicode.GetBytes(text))).Append(">\n");
            sb.Append("endbfchar\n");
        }
        sb.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
        return sb.ToString();
    }

    private string SubsetTag()
    {
        uint hash = 2166136261;
        foreach (var (cid, (glyph, _)) in _usedCids) hash = (hash ^ cid ^ ((uint)glyph << 16)) * 16777619;
        foreach (char c in Font.PostScriptName) hash = (hash ^ c) * 16777619;
        var tag = new char[6];
        for (int i = 0; i < 6; i++, hash /= 26) tag[i] = (char)('A' + hash % 26);
        return new string(tag);
    }
}
