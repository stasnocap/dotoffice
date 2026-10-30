using System.Text;
using DotOffice.Fonts;

namespace DotOffice.Pdf;

/// <summary>
/// A TrueType font embedded as a Type0/CIDFontType2 font with Identity-H encoding.
/// Glyph usage is tracked while drawing so only used glyphs are embedded.
/// </summary>
public sealed class PdfFont
{
    private readonly SortedDictionary<ushort, string> _usedGlyphs = new();

    internal PdfFont(TrueTypeFont font, string resourceName, bool embedded = true)
    {
        Font = font;
        ResourceName = resourceName;
        IsEmbedded = embedded;
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
        {
            var gid = Font.GetGlyphId(rune.Value);
            if (gid != 0) _usedGlyphs.TryAdd(gid, rune.ToString());
            sb.Append(gid.ToString("X4"));
        }
        return sb.Append('>').ToString();
    }

    internal void Write(PdfWriter writer, int type0Id)
    {
        int cidFontId = writer.Allocate();
        int descriptorId = writer.Allocate();
        int fontFileId = IsEmbedded ? writer.Allocate() : 0;
        int toUnicodeId = writer.Allocate();

        var baseFont = IsEmbedded ? SubsetTag() + "+" + Font.PostScriptName : Font.PostScriptName;
        double scale = 1000.0 / Font.UnitsPerEm;

        writer.WriteObject(type0Id,
            $"<< /Type /Font /Subtype /Type0 /BaseFont /{baseFont} /Encoding /Identity-H " +
            $"/DescendantFonts [{cidFontId} 0 R] /ToUnicode {toUnicodeId} 0 R >>");

        writer.WriteObject(cidFontId,
            $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{baseFont} " +
            "/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> " +
            $"/FontDescriptor {descriptorId} 0 R /DW {PdfFormat.Number(Font.GetAdvanceWidth(0) * scale)} " +
            $"/W {WidthsArray(scale)} /CIDToGIDMap /Identity >>");

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

        if (IsEmbedded)
        {
            var subset = TrueTypeSubsetter.Subset(Font, _usedGlyphs.Keys);
            writer.WriteStream(fontFileId, $"/Length1 {subset.Length}", subset);
        }
        writer.WriteStream(toUnicodeId, "", Encoding.ASCII.GetBytes(ToUnicodeCMap()));
    }

    private string WidthsArray(double scale)
    {
        var sb = new StringBuilder("[");
        int? runStart = null;
        int previous = -2;
        foreach (var gid in _usedGlyphs.Keys)
        {
            if (gid != previous + 1)
            {
                if (runStart is not null) sb.Append("] ");
                sb.Append(gid).Append(" [");
                runStart = gid;
            }
            else sb.Append(' ');
            sb.Append(PdfFormat.Number(Font.GetAdvanceWidth(gid) * scale));
            previous = gid;
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
        foreach (var chunk in _usedGlyphs.Chunk(100))
        {
            sb.Append(chunk.Length).Append(" beginbfchar\n");
            foreach (var (gid, text) in chunk)
                sb.Append('<').Append(gid.ToString("X4")).Append("> <")
                  .Append(Convert.ToHexString(Encoding.BigEndianUnicode.GetBytes(text))).Append(">\n");
            sb.Append("endbfchar\n");
        }
        sb.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
        return sb.ToString();
    }

    private string SubsetTag()
    {
        uint hash = 2166136261;
        foreach (var gid in _usedGlyphs.Keys) hash = (hash ^ gid) * 16777619;
        foreach (char c in Font.PostScriptName) hash = (hash ^ c) * 16777619;
        var tag = new char[6];
        for (int i = 0; i < 6; i++, hash /= 26) tag[i] = (char)('A' + hash % 26);
        return new string(tag);
    }
}
