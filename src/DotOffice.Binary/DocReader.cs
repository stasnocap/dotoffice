using System.Buffers.Binary;
using System.Text;
using DotOffice.Core;
using DotOffice.Docx;

namespace DotOffice.Binary;

/// <summary>Reads Word 97-2003 documents (.doc) into the <see cref="Document"/> model: text, paragraph and character formatting, styles, tables, page breaks and section page setup.</summary>
public static class DocReader
{
    /// <summary>Reads a DOC file. Corrupt, encrypted, unsupported or oversized input throws <see cref="InvalidDataException"/>.</summary>
    public static Document Read(Stream stream, BinaryReaderOptions? options = null)
    {
        options ??= new BinaryReaderOptions();
        var file = new CompoundFile(stream, options);
        if (file.Contains("EncryptedPackage")) throw new InvalidDataException("The document is encrypted.");
        var main = file.ReadStream("WordDocument") ?? throw new InvalidDataException("The file has no WordDocument stream; it is not a Word document.");
        if (main.Length < 0x01A8 || BinaryPrimitives.ReadUInt16LittleEndian(main) != 0xA5EC) throw new InvalidDataException("The WordDocument stream has no valid file information block.");
        int flags = BinaryPrimitives.ReadUInt16LittleEndian(main.AsSpan(0x0A));
        if ((flags & 0x0100) != 0) throw new InvalidDataException("The document is encrypted.");
        if (BinaryPrimitives.ReadUInt16LittleEndian(main.AsSpan(2)) < 0x00C1) throw new InvalidDataException("Word 6.0/95 documents are not supported.");
        var table = file.ReadStream((flags & 0x0200) != 0 ? "1Table" : "0Table") ?? throw new InvalidDataException("The document's table stream is missing.");
        try
        {
            return new DocContext(main, table, file.ReadStream("Data") ?? []).Read();
        }
        catch (Exception e) when (e is IndexOutOfRangeException or ArgumentException or OverflowException)
        {
            throw new InvalidDataException("The document is damaged.", e);
        }
    }
}

internal sealed partial class DocContext(byte[] main, byte[] table, byte[] data)
{
    private const int MaxChars = 50_000_000, MaxBlocks = 2_000_000;

    private sealed record Piece(int CpStart, int CpEnd, long ByteStart, bool Compressed);

    private sealed class Chp
    {
        public bool Bold, Italic, Strike, SmallCaps, Caps, Hidden;
        public int Underline, Size = 20, Font, Color = -1, Script;
        public long PicLocation = -1; // where a picture character's PICF is in the Data stream
        public bool Ole2, FData;
        public Chp Clone() => (Chp)MemberwiseClone();
    }

    private sealed class Pap
    {
        public int Istd, Jc, Left, Right, First, Before, After, Line = 240, LineMult = 1;
        public bool Keep, KeepNext, PageBreakBefore, InTable, Ttp;

        /// <summary>Frame position (sprmPDxaAbs, sprmPDyaAbs, sprmPPc), size and wrapping; <see cref="Framed"/> when any is set.</summary>
        public int DxaAbs, DyaAbs, Pc, FrameWidth, FrameHeight, Wr, DxaFromText, DyaFromText;

        /// <summary>The list override (1-based; 0 for none) and level of a list paragraph.</summary>
        public int Ilfo, Ilvl;
        public bool Framed;
        public List<TabStop> Tabs = [];
        public byte[]? TableDef;

        /// <summary>Paragraph borders: top, left, bottom, right, between.</summary>
        public Border[] Borders = [Border.None, Border.None, Border.None, Border.None, Border.None];
        public Pap Clone() { var p = (Pap)MemberwiseClone(); p.Tabs = [.. Tabs]; p.Borders = [.. Borders]; return p; }
    }

    private sealed class Style
    {
        public int Base = 0x0FFF, Kind;
        public byte[] Papx = [], Chpx = [];
        public Chp? ResolvedChp;
        public Pap? ResolvedPap;
    }

    private readonly List<Piece> _pieces = [];
    private readonly List<string> _fonts = [];
    private readonly Dictionary<string, FontDeclaration> _fontTable = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<int, Style> _styles = [];
    private readonly List<(int Start, int End, byte[] Props)> _chpRuns = [], _papRuns = [];
    private int _blocks;
    private string _text = "";
    private int _textLength;
    private readonly Dictionary<(int, int), HeaderFooter?> _stories = [];

    private int FcLcb(int pair, out int lcb)
    {
        // FibBase is 32 bytes, then csw and its words, then cslw and its longs, then cbRgFcLcb and the (fc, lcb) pairs.
        int csw = BinaryPrimitives.ReadUInt16LittleEndian(main.AsSpan(32));
        int at = 34 + csw * 2;
        int cslw = BinaryPrimitives.ReadUInt16LittleEndian(main.AsSpan(at));
        at += 2 + cslw * 4;
        int count = BinaryPrimitives.ReadUInt16LittleEndian(main.AsSpan(at));
        at += 2;
        lcb = 0;
        if (pair >= count || at + pair * 8 + 8 > main.Length) return 0;
        lcb = (int)BinaryPrimitives.ReadUInt32LittleEndian(main.AsSpan(at + pair * 8 + 4));
        return (int)BinaryPrimitives.ReadUInt32LittleEndian(main.AsSpan(at + pair * 8));
    }

    private int Lw(int index)
    {
        int csw = BinaryPrimitives.ReadUInt16LittleEndian(main.AsSpan(32));
        int at = 34 + csw * 2;
        return (int)BinaryPrimitives.ReadUInt32LittleEndian(main.AsSpan(at + 2 + index * 4));
    }

    private static int U16(ReadOnlySpan<byte> d, int at) => BinaryPrimitives.ReadUInt16LittleEndian(d[at..]);
    private static int I16(ReadOnlySpan<byte> d, int at) => BinaryPrimitives.ReadInt16LittleEndian(d[at..]);
    private static long U32(ReadOnlySpan<byte> d, int at) => BinaryPrimitives.ReadUInt32LittleEndian(d[at..]);

    public Document Read()
    {
        int textLength = Lw(3);
        if (textLength < 0 || textLength > MaxChars) throw new InvalidDataException("The document's text length is not valid.");
        ReadPieces();
        ReadFonts();
        ReadStyles();
        ReadLists();
        ReadRuns(12, _chpRuns, papx: false);
        ReadRuns(13, _papRuns, papx: true);
        // The text stream holds the main story, then the footnotes, the headers and footers, comments, endnotes and text boxes.
        long total = 0;
        for (int k = 3; k <= 10; k++) total += Math.Max(Lw(k), 0);
        _text = ReadText((int)Math.Min(total, MaxChars));
        _textLength = Math.Min(textLength, _text.Length);
        ReadDrawings();
        ReadNotes();
        var document = Build(0, _textLength, story: false);
        // Without a separator story, Word draws its own in an empty paragraph of the Normal style.
        if (_footnotes > 0)
            document.NoteSeparator = Story(-1, 0, 0)?.Blocks ?? [new Paragraph { MarkProperties = Props(StyleChp(0)) }];
        return document;
    }

    // ------------------------------------------------------------------ text

    private void ReadPieces()
    {
        int fc = FcLcb(33, out int lcb);
        if (lcb <= 0 || fc < 0 || (long)fc + lcb > table.Length) throw new InvalidDataException("The document has no piece table.");
        var clx = table.AsSpan(fc, lcb);
        int at = 0;
        while (at < clx.Length && clx[at] == 1) at += 3 + U16(clx, at + 1); // Prc: property modifiers shared by pieces, not used here
        if (at >= clx.Length || clx[at] != 2) throw new InvalidDataException("The piece table is missing.");
        long size = U32(clx, at + 1);
        at += 5;
        if (size < 4 || at + size > clx.Length) throw new InvalidDataException("The piece table is truncated.");
        int count = (int)((size - 4) / 12);
        for (int i = 0; i < count; i++)
        {
            long cpStart = U32(clx, at + 4 * i), cpEnd = U32(clx, at + 4 * (i + 1));
            long fcRaw = U32(clx, at + 4 * (count + 1) + 8 * i + 2);
            bool compressed = (fcRaw & 0x40000000) != 0;
            long start = compressed ? (fcRaw & 0x3FFFFFFF) / 2 : fcRaw & 0x3FFFFFFF;
            if (cpEnd < cpStart || cpEnd > MaxChars) throw new InvalidDataException("The piece table is not valid.");
            _pieces.Add(new Piece((int)cpStart, (int)cpEnd, start, compressed));
        }
    }

    private string ReadText(int length)
    {
        var sb = new StringBuilder(Math.Min(length, 1 << 20));
        foreach (var piece in _pieces)
        {
            if (piece.CpStart >= length) break;
            int chars = Math.Min(piece.CpEnd, length) - piece.CpStart;
            long bytes = piece.Compressed ? chars : 2L * chars;
            if (piece.ByteStart < 0 || piece.ByteStart + bytes > main.Length) throw new InvalidDataException("A piece of the document text is outside the file.");
            var span = main.AsSpan((int)piece.ByteStart, (int)bytes);
            if (piece.Compressed)
            {
                // Compressed text is Windows-1252; a few of its control positions hold typographic characters.
                foreach (byte b in span) sb.Append(Cp1252(b));
            }
            else sb.Append(Encoding.Unicode.GetString(span));
        }
        return sb.ToString();
    }

    private static char Cp1252(byte b) => b switch
    {
        0x82 => '‚', 0x83 => 'ƒ', 0x84 => '„', 0x85 => '…', 0x86 => '†', 0x87 => '‡', 0x88 => 'ˆ', 0x89 => '‰',
        0x8A => 'Š', 0x8B => '‹', 0x8C => 'Œ', 0x8E => 'Ž', 0x91 => '‘', 0x92 => '’', 0x93 => '“', 0x94 => '”',
        0x95 => '•', 0x96 => '–', 0x97 => '—', 0x98 => '˜', 0x99 => '™', 0x9A => 'š', 0x9B => '›', 0x9C => 'œ',
        0x9E => 'ž', 0x9F => 'Ÿ', 0x80 => '€', _ => (char)b,
    };

    private void ReadFonts()
    {
        int fc = FcLcb(15, out int lcb);
        if (lcb < 4 || fc < 0 || (long)fc + lcb > table.Length) return;
        var d = table.AsSpan(fc, lcb);
        int count = U16(d, 0), at = 4;
        for (int i = 0; i < count && at < d.Length; i++)
        {
            int cb = d[at] + 1;
            if (cb < 41 || at + cb > d.Length) break;
            var names = Encoding.Unicode.GetString(d.Slice(at + 40, Math.Max(0, cb - 40)));
            int nul = names.IndexOf('\0');
            var name = nul >= 0 ? names[..nul] : names;
            _fonts.Add(name);
            // FFN: ff (font family, 1 = roman) in bits 4-6 of the second byte; ixchSzAlt (start of the alternative name, 0 when none) at 5
            int alt = d[at + 5];
            string? altName = null;
            if (alt > 0 && alt < names.Length) { altName = names[alt..]; nul = altName.IndexOf('\0'); if (nul >= 0) altName = altName[..nul]; }
            if (name.Length > 0) _fontTable.TryAdd(name, new FontDeclaration(string.IsNullOrEmpty(altName) ? null : altName, (d[at + 1] >> 4 & 7) == 1));
            at += cb;
        }
    }

    private string FontName(int index) => index >= 0 && index < _fonts.Count && _fonts[index].Length > 0 ? _fonts[index] : "Times New Roman";

    // ------------------------------------------------------------------ styles

    private int _defaultFont;

    private void ReadStyles()
    {
        int fc = FcLcb(1, out int lcb);
        if (lcb < 20 || fc < 0 || (long)fc + lcb > table.Length) return;
        var d = table.AsSpan(fc, lcb);
        int cbStshi = U16(d, 0);
        if (2 + cbStshi > d.Length || cbStshi < 18) return;
        int count = U16(d, 2), baseSize = U16(d, 4);
        _defaultFont = U16(d, 14);
        int at = 2 + cbStshi;
        for (int i = 0; i < count && at + 2 <= d.Length && i < 4096; i++)
        {
            int cb = U16(d, at);
            at += 2;
            if (cb > 0 && at + cb <= d.Length && cb >= baseSize + 2)
            {
                var std = d.Slice(at, cb);
                int sgc = U16(std, 2) & 0xF, istdBase = U16(std, 2) >> 4;
                int cupx = U16(std, 4) & 0xF;
                int nameChars = U16(std, baseSize);
                int p = baseSize + 2 + (nameChars + 1) * 2;
                var style = new Style { Base = istdBase, Kind = sgc };
                var upx = new List<byte[]>();
                for (int k = 0; k < cupx && p + 2 <= std.Length; k++)
                {
                    int cbUpx = U16(std, p);
                    p += 2;
                    if (p + cbUpx > std.Length) break;
                    upx.Add(std.Slice(p, cbUpx).ToArray());
                    p += cbUpx + (cbUpx & 1);
                }
                if (sgc == 1) { if (upx.Count > 0) style.Papx = upx[0]; if (upx.Count > 1) style.Chpx = upx[1]; }
                else if (sgc == 2 && upx.Count > 0) style.Chpx = upx[0];
                _styles[i] = style;
            }
            at += cb + (cb & 1 & 0);
        }
    }

    private Chp StyleChp(int istd, int depth = 0)
    {
        if (!_styles.TryGetValue(istd, out var style) || depth > 20) return new Chp { Font = _defaultFont };
        if (style.ResolvedChp is not null) return style.ResolvedChp;
        var chp = style.Base != 0x0FFF && style.Base != istd ? StyleChp(style.Base, depth + 1).Clone() : new Chp { Font = _defaultFont };
        // The character formatting of a paragraph style is its second UPX; of a character style its only one.
        if (style.Kind == 2 || style.Kind == 1) ApplyChp(chp, style.Chpx);
        return style.ResolvedChp = chp;
    }

    /// <summary>Applies a character style's sprms, after those of the character styles it is based on, over the formatting so far.</summary>
    private void ApplyCharacterStyle(Chp chp, int istd, int depth = 0)
    {
        if (!_styles.TryGetValue(istd, out var style) || style.Kind != 2 || depth > 20) return;
        if (style.Base != 0x0FFF && style.Base != istd) ApplyCharacterStyle(chp, style.Base, depth + 1);
        ApplyChp(chp, style.Chpx);
    }

    private Pap StylePap(int istd, int depth = 0)
    {
        if (!_styles.TryGetValue(istd, out var style) || depth > 20) return new Pap();
        if (style.ResolvedPap is not null) return style.ResolvedPap;
        var pap = style.Base != 0x0FFF && style.Base != istd ? StylePap(style.Base, depth + 1).Clone() : new Pap();
        if (style.Kind == 1 && style.Papx.Length >= 2) ApplyPap(pap, style.Papx.AsSpan(2));
        pap.Istd = istd;
        return style.ResolvedPap = pap;
    }

    // ------------------------------------------------------------------ sprms

    private static int OperandSize(int sprm, ReadOnlySpan<byte> d, int at)
    {
        switch (sprm >> 13)
        {
            case 0 or 1: return 1;
            case 2 or 4 or 5: return 2;
            case 3: return 4;
            case 7: return 3;
            default:
                if (sprm == 0xD608) return at + 2 <= d.Length ? 2 + U16(d, at) : d.Length - at;
                return at < d.Length ? 1 + d[at] : 0;
        }
    }

    private static bool Toggle(bool current, int v) => v switch { 0 => false, 1 => true, 0x81 => !current, _ => current };

    private void ApplyChp(Chp chp, ReadOnlySpan<byte> d)
    {
        int at = 0;
        while (at + 2 <= d.Length)
        {
            int sprm = U16(d, at);
            at += 2;
            int size = OperandSize(sprm, d, at);
            if (at + size > d.Length) break;
            var o = d.Slice(at, size);
            switch (sprm)
            {
                case 0x4A30 when size >= 2: ApplyCharacterStyle(chp, U16(o, 0)); break; // its formatting comes first, then what follows in this list
                case 0x0835: chp.Bold = Toggle(chp.Bold, o[0]); break;
                case 0x0836: chp.Italic = Toggle(chp.Italic, o[0]); break;
                case 0x0837: chp.Strike = Toggle(chp.Strike, o[0]); break;
                case 0x083A: chp.SmallCaps = Toggle(chp.SmallCaps, o[0]); break;
                case 0x083B: chp.Caps = Toggle(chp.Caps, o[0]); break;
                case 0x083C: chp.Hidden = Toggle(chp.Hidden, o[0]); break;
                case 0x2A3E: chp.Underline = o[0]; break;
                case 0x4A43 when size >= 2: chp.Size = Math.Max(2, U16(o, 0)); break;
                case 0x4A4F when size >= 2: chp.Font = U16(o, 0); break;
                case 0x2A42: chp.Color = o[0] == 0 ? -1 : o[0]; break;
                case 0x6870 when size >= 4: chp.Color = 0x1000000 | (o[0] << 16) | (o[1] << 8) | o[2]; break;
                case 0x2A48: chp.Script = o[0]; break;
                case 0x6A03 when size >= 4: chp.PicLocation = U32(o, 0); break;
                case 0x080A: chp.Ole2 = o[0] != 0; break;
                case 0x0806: chp.FData = o[0] != 0; break;
            }
            at += size;
        }
    }

    private void ApplyPap(Pap pap, ReadOnlySpan<byte> d)
    {
        int at = 0;
        while (at + 2 <= d.Length)
        {
            int sprm = U16(d, at);
            at += 2;
            int size = OperandSize(sprm, d, at);
            if (at + size > d.Length) break;
            var o = d.Slice(at, size);
            switch (sprm)
            {
                case 0x4600 when size >= 2: pap.Istd = U16(o, 0); break;
                case 0x2403 or 0x2461: pap.Jc = o[0]; break;
                case 0x2405: pap.Keep = o[0] == 1; break;
                case 0x2406: pap.KeepNext = o[0] == 1; break;
                case 0x2407: pap.PageBreakBefore = o[0] == 1; break;
                case 0x2416: pap.InTable = o[0] == 1; break;
                case 0x460B when size >= 2: pap.Ilfo = I16(o, 0); break;
                case 0x260A: pap.Ilvl = o[0]; break;
                case 0x2417: pap.Ttp = o[0] == 1; break;
                case 0x845E or 0x840F when size >= 2: pap.Left = I16(o, 0); break;
                case 0x845D or 0x840E when size >= 2: pap.Right = I16(o, 0); break;
                case 0x8460 or 0x8411 when size >= 2: pap.First = I16(o, 0); break;
                case 0xA413 or 0xA414 when size >= 2:
                    if (sprm == 0xA413) pap.Before = U16(o, 0); else pap.After = U16(o, 0);
                    break;
                case 0x6412 when size >= 4: pap.Line = I16(o, 0); pap.LineMult = U16(o, 2); break;
                case 0xD608: pap.TableDef = o.ToArray(); break;
                case 0x8418 when size >= 2: pap.DxaAbs = I16(o, 0); pap.Framed = true; break;
                case 0x8419 when size >= 2: pap.DyaAbs = I16(o, 0); pap.Framed = true; break;
                case 0x261B:
                    // sprmPPc: pcVert in bits 4-5 and pcHorz in bits 6-7; 3 leaves the current value
                    if ((o[0] >> 4 & 3) != 3) pap.Pc = pap.Pc & ~0x30 | o[0] & 0x30;
                    if ((o[0] >> 6 & 3) != 3) pap.Pc = pap.Pc & ~0xC0 | o[0] & 0xC0;
                    pap.Framed = true;
                    break;
                case 0x841A when size >= 2: pap.FrameWidth = U16(o, 0); pap.Framed = true; break;
                case 0x442B when size >= 2: pap.FrameHeight = U16(o, 0); break;
                case 0x2423: pap.Wr = o[0]; break;
                case 0x842F when size >= 2: pap.DxaFromText = U16(o, 0); break;
                case 0x842E when size >= 2: pap.DyaFromText = U16(o, 0); break;
                case >= 0x6424 and <= 0x6428 when size >= 4: pap.Borders[sprm - 0x6424] = Brc(o, 0) with { Space = o[3] & 0x1F }; break;
                case >= 0xC64E and <= 0xC652 when size >= 9: pap.Borders[sprm - 0xC64E] = BrcNew(o[1..]); break; // cb, then the BRC
                case 0xC615 when size >= 2: ApplyTabs(pap, o, tolerances: true); break;
                case 0xC60D when size >= 2: ApplyTabs(pap, o, tolerances: false); break;
            }
            at += size;
        }
    }

    private static void ApplyTabs(Pap pap, ReadOnlySpan<byte> o, bool tolerances)
    {
        // sprmPChgTabs / sprmPChgTabsPapx: cb, then the deleted positions (count + shorts, and for sprmPChgTabs a
        // close tolerance per deleted tab), then the added ones (count, positions, descriptors).
        int cb = o[0], at = 1;
        if (cb == 0xFF || at >= o.Length) return;
        int deleted = o[at++];
        for (int i = 0; i < deleted && at + 2 * i + 2 <= o.Length; i++)
        {
            double pos = I16(o, at + 2 * i) / 20.0, close = tolerances && at + 2 * (deleted + i) + 2 <= o.Length ? I16(o, at + 2 * (deleted + i)) / 20.0 : 0;
            pap.Tabs.RemoveAll(t => Math.Abs(t.Position - pos) <= Math.Max(0.5, close));
        }
        at += deleted * (tolerances ? 4 : 2);
        if (at >= o.Length) return;
        int added = o[at++];
        if (at + added * 3 > o.Length) return;
        for (int i = 0; i < added; i++)
        {
            double pos = I16(o, at + 2 * i) / 20.0;
            int desc = o[at + 2 * added + i];
            var align = (desc & 7) switch { 1 => TabAlignment.Center, 2 => TabAlignment.Right, 3 => TabAlignment.Decimal, 4 => TabAlignment.Bar, _ => TabAlignment.Left };
            var leader = ((desc >> 3) & 7) switch { 1 => TabLeader.Dot, 2 => TabLeader.Hyphen, 3 => TabLeader.Underscore, 4 => TabLeader.Underscore, 5 => TabLeader.MiddleDot, _ => TabLeader.None };
            pap.Tabs.RemoveAll(t => Math.Abs(t.Position - pos) < 0.5);
            pap.Tabs.Add(new TabStop(pos, align, leader));
        }
        pap.Tabs.Sort((a, b) => a.Position.CompareTo(b.Position));
    }

    // ------------------------------------------------------------------ formatting runs

    /// <summary>Reads the bin table of character (CHPX) or paragraph (PAPX) formatting pages into runs over character positions.</summary>
    private void ReadRuns(int pair, List<(int, int, byte[])> runs, bool papx)
    {
        int fc = FcLcb(pair, out int lcb);
        if (lcb < 8 || fc < 0 || (long)fc + lcb > table.Length) return;
        var plc = table.AsSpan(fc, lcb);
        int n = (lcb - 4) / 8;
        var entries = new List<(long FcStart, long FcEnd, byte[] Props)>();
        for (int i = 0; i < n && i < 100_000; i++)
        {
            long page = U32(plc, 4 * (n + 1) + 4 * i) & 0x3FFFFF;
            long at = page * 512;
            if (at + 512 > main.Length) continue;
            var fkp = main.AsSpan((int)at, 512);
            int crun = fkp[511];
            if (crun == 0 || 4 * (crun + 1) + crun * (papx ? 13 : 1) > 511) continue;
            for (int k = 0; k < crun; k++)
            {
                long start = U32(fkp, 4 * k), end = U32(fkp, 4 * (k + 1));
                int offset = papx ? fkp[4 * (crun + 1) + 13 * k] : fkp[4 * (crun + 1) + k];
                byte[] props = [];
                if (offset != 0)
                {
                    int p = offset * 2;
                    if (papx)
                    {
                        int cb = fkp[p];
                        p++;
                        if (cb == 0) { cb = fkp[p]; p++; cb *= 2; } else cb = cb * 2 - 1;
                        if (p + cb <= 511 && cb >= 0) props = fkp.Slice(p, cb).ToArray();
                    }
                    else
                    {
                        int cb = fkp[p];
                        if (p + 1 + cb <= 511) props = fkp.Slice(p + 1, cb).ToArray();
                    }
                }
                entries.Add((start, end, props));
            }
        }
        entries.Sort((a, b) => a.FcStart.CompareTo(b.FcStart));
        // Map each file range onto the character positions of the pieces it overlaps.
        foreach (var piece in _pieces)
        {
            int bytesPer = piece.Compressed ? 1 : 2;
            long pieceStart = piece.ByteStart, pieceEnd = pieceStart + (long)(piece.CpEnd - piece.CpStart) * bytesPer;
            int lo = 0, hi = entries.Count;
            while (lo < hi) { int mid = (lo + hi) / 2; if (entries[mid].FcEnd <= pieceStart) lo = mid + 1; else hi = mid; }
            for (int e = lo; e < entries.Count && entries[e].FcStart < pieceEnd; e++)
            {
                long a = Math.Max(entries[e].FcStart, pieceStart), b = Math.Min(entries[e].FcEnd, pieceEnd);
                if (b <= a) continue;
                runs.Add((piece.CpStart + (int)((a - pieceStart) / bytesPer), piece.CpStart + (int)((b - pieceStart) / bytesPer), entries[e].Props));
                if (runs.Count > 5_000_000) throw new InvalidDataException("The document has too many formatting runs.");
            }
        }
        runs.Sort((a, b) => a.Item1.CompareTo(b.Item1));
    }

    private static int FindRun(List<(int Start, int End, byte[] Props)> runs, int cp)
    {
        int lo = 0, hi = runs.Count - 1;
        while (lo <= hi)
        {
            int mid = (lo + hi) / 2;
            if (runs[mid].End <= cp) lo = mid + 1;
            else if (runs[mid].Start > cp) hi = mid - 1;
            else return mid;
        }
        return -1;
    }

    // ------------------------------------------------------------------ sections

    private Section MakeSection(int sectionIndex)
    {
        double width = 612, height = 792, left = 90, right = 90, top = 72, bottom = 72;
        bool landscape = false, titlePage = false;
        double headerDistance = 36, footerDistance = 36;
        int fc = FcLcb(6, out int lcb);
        if (lcb >= 16 && fc >= 0 && (long)fc + lcb <= table.Length)
        {
            var plc = table.AsSpan(fc, lcb);
            int n = (lcb - 4) / 16;
            if (sectionIndex < n)
            {
                long sepx = U32(plc, 4 * (n + 1) + 12 * sectionIndex + 2);
                if (sepx is > 0 and < int.MaxValue && sepx + 2 <= main.Length)
                {
                    int cb = U16(main, (int)sepx);
                    var d = main.AsSpan((int)sepx + 2, (int)Math.Min(cb, main.Length - sepx - 2));
                    int at = 0;
                    while (at + 2 <= d.Length)
                    {
                        int sprm = U16(d, at);
                        at += 2;
                        int size = OperandSize(sprm, d, at);
                        if (at + size > d.Length) break;
                        var o = d.Slice(at, size);
                        switch (sprm)
                        {
                            case 0xB01F: width = U16(o, 0) / 20.0; break;
                            case 0xB020: height = U16(o, 0) / 20.0; break;
                            case 0xB021: left = U16(o, 0) / 20.0; break;
                            case 0xB022: right = U16(o, 0) / 20.0; break;
                            case 0x9023: top = I16(o, 0) / 20.0; break;
                            case 0x9024: bottom = I16(o, 0) / 20.0; break;
                            case 0x301D: landscape = o[0] == 2; break;
                            case 0x300A: titlePage = o[0] != 0; break;
                            case 0xB017: headerDistance = U16(o, 0) / 20.0; break;
                            case 0xB018: footerDistance = U16(o, 0) / 20.0; break;
                        }
                        at += size;
                    }
                }
            }
        }
        if (width < 72 || height < 72 || width > 14400 || height > 14400) (width, height) = (612, 792);
        if (left + right > width * 0.8) (left, right) = (width * 0.1, width * 0.1);
        if (top + bottom > height * 0.8) (top, bottom) = (height * 0.1, height * 0.1);
        return new Section
        {
            PageWidth = width, PageHeight = height, Orientation = landscape || width > height ? Orientation.Landscape : Orientation.Portrait,
            Margins = new Thickness(Math.Abs(top), right, Math.Abs(bottom), left),
            HeaderDistance = headerDistance is > 0 and < 1000 ? headerDistance : 36, FooterDistance = footerDistance is > 0 and < 1000 ? footerDistance : 36,
            DifferentFirstPage = titlePage,
            EvenHeader = Story(sectionIndex, 0), DefaultHeader = Story(sectionIndex, 1), EvenFooter = Story(sectionIndex, 2),
            DefaultFooter = Story(sectionIndex, 3), FirstHeader = Story(sectionIndex, 4), FirstFooter = Story(sectionIndex, 5),
        };
    }

    /// <summary>
    /// A header or footer of a section from the header story table (PlcfHdd): six footnote and endnote separator stories, then six per
    /// section (even header, odd header, even footer, odd footer, first header, first footer). An empty story repeats the previous section's.
    /// </summary>
    private HeaderFooter? Story(int sectionIndex, int kind) => Story(sectionIndex, kind, 6 + sectionIndex * 6 + kind);

    /// <summary>The story at <paramref name="index"/> of the header story table; the first is the footnote separator.</summary>
    private HeaderFooter? Story(int sectionIndex, int kind, int index)
    {
        if (_stories.TryGetValue((sectionIndex, kind), out var cached)) return cached;
        HeaderFooter? result = null;
        int fc = FcLcb(11, out int lcb);
        int count = lcb / 4;
        int baseCp = _textLength + Math.Max(Lw(4), 0);
        if (fc >= 0 && lcb >= 8 && (long)fc + lcb <= table.Length)
        {
            var plc = table.AsSpan(fc, lcb);
            if (index + 1 < count)
            {
                long start = baseCp + U32(plc, 4 * index), end = baseCp + U32(plc, 4 * (index + 1));
                if (start >= end) result = sectionIndex > 0 && index >= 6 ? Story(sectionIndex - 1, kind) : null;
                else if (end <= _text.Length)
                {
                    var built = Build((int)start, (int)end, story: true).Sections[0].Blocks;
                    // A story ends with a guard paragraph mark that Word does not show.
                    if (built.Count > 1 && built[^1] is Paragraph { Inlines.Count: 0 }) built.RemoveAt(built.Count - 1);
                    result = new HeaderFooter();
                    result.Blocks.AddRange(built);
                }
            }
        }
        _stories[(sectionIndex, kind)] = result;
        return result;
    }

    /// <summary>The document properties' fFacingPages: different headers and footers on even and odd pages.</summary>
    private bool FacingPages()
    {
        int fc = FcLcb(31, out int lcb);
        return lcb >= 2 && fc >= 0 && (long)fc + 2 <= table.Length && (table[fc] & 1) != 0;
    }

    /// <summary>
    /// Whether Word collapses the space between paragraphs to the larger of the space after and the space before (HTML paragraph auto spacing):
    /// when the Word 2000 copts in the DOP leave fDontUseHTMLParagraphAutoSpacing (byte 512, bit 2) clear. Older DOPs do not have it.
    /// </summary>
    private bool HtmlParagraphSpacing()
    {
        int fc = FcLcb(31, out int lcb);
        return lcb > 512 && fc >= 0 && (long)fc + 513 <= table.Length && (table[fc + 512] & 4) == 0;
    }

    private static FieldKind? PageField(string code)
    {
        var word = code.Trim().Split(' ', 2)[0].ToUpperInvariant();
        return word switch { "PAGE" => FieldKind.Page, "NUMPAGES" => FieldKind.NumPages, "SECTIONPAGES" => FieldKind.SectionPages, _ => null };
    }

    private List<int> SectionEnds()
    {
        var ends = new List<int>();
        int fc = FcLcb(6, out int lcb);
        if (lcb >= 16 && fc >= 0 && (long)fc + lcb <= table.Length)
        {
            var plc = table.AsSpan(fc, lcb);
            int n = (lcb - 4) / 16;
            for (int i = 1; i <= n; i++) ends.Add((int)U32(plc, 4 * i));
        }
        return ends;
    }

    // ------------------------------------------------------------------ building

    private RunProperties Props(Chp c) => new()
    {
        FontFamily = FontName(c.Font), FontSize = Math.Max(1, c.Size / 2.0), Bold = c.Bold, Italic = c.Italic, Strike = c.Strike,
        Underline = c.Underline switch { 0 => UnderlineStyle.None, 3 => UnderlineStyle.Double, 6 => UnderlineStyle.Thick, 4 => UnderlineStyle.Dotted, 7 => UnderlineStyle.Dashed, 11 => UnderlineStyle.Wavy, _ => UnderlineStyle.Single },
        Color = ColorOf(c.Color), SmallCaps = c.SmallCaps, Caps = c.Caps, Hidden = c.Hidden,
        VerticalPosition = c.Script == 1 ? VerticalPosition.Superscript : c.Script == 2 ? VerticalPosition.Subscript : VerticalPosition.Baseline,
    };

    /// <summary>Builds the blocks of the text from <paramref name="from"/> to <paramref name="to"/>: the main story with its sections, or a header or footer story.</summary>
    private Document Build(int from, int to, bool story)
    {
        var text = _text;
        // Word opens .doc files in compatibility mode, where a paragraph after a page break keeps its space before.
        var doc = new Document { Settings = new DocumentSettings { EvenAndOddHeaders = !story && FacingPages(), PageBreakBeforeDropsSpaceBefore = false, CollapseParagraphSpacing = !story && HtmlParagraphSpacing() },
            Fonts = _fontTable,
        };
        var sectionEnds = story ? [] : SectionEnds();
        int sectionIndex = 0;
        var section = story ? new Section() : MakeSection(0);
        doc.Sections.Add(section);
        var blocks = section.Blocks;

        var inlines = new List<Inline>();
        var run = new StringBuilder();
        Chp? runChp = null;
        Pap? tablePap = null;
        Table? currentTable = null;
        List<TableCell>? rowCells = null;
        List<Block>? cellBlocks = null;
        var fieldStack = new List<bool>(); // for each open field: whether its result (after the separator) is being shown
        var fieldCode = new StringBuilder();

        void Flush()
        {
            if (run.Length > 0 && runChp is not null) inlines.Add(new TextRun(run.ToString(), Props(runChp)));
            run.Clear();
        }

        Chp ChpAt(int cp, Pap pap)
        {
            var chp = StyleChp(pap.Istd).Clone();
            int r = FindRun(_chpRuns, cp);
            if (r >= 0) ApplyChp(chp, _chpRuns[r].Props);
            return chp;
        }

        Pap PapAt(int cp)
        {
            int r = FindRun(_papRuns, cp);
            var props = r >= 0 ? _papRuns[r].Props : [];
            int istd = props.Length >= 2 ? U16(props, 0) : 0;
            var pap = StylePap(istd).Clone();
            if (props.Length > 2) ApplyPap(pap, props.AsSpan(2));
            // A list level's indents and tabs come between the style and the paragraph's own formatting.
            if (LevelOf(pap.Ilfo, pap.Ilvl) is { Papx.Length: > 0 } level)
            {
                var (ilfo, ilvl) = (pap.Ilfo, pap.Ilvl);
                pap = StylePap(istd).Clone();
                ApplyPap(pap, level.Papx);
                if (props.Length > 2) ApplyPap(pap, props.AsSpan(2));
                (pap.Ilfo, pap.Ilvl) = (ilfo, ilvl);
            }
            pap.Istd = istd;
            return pap;
        }

        ParagraphProperties ParagraphProps(Pap p) => new()
        {
            Alignment = p.Jc switch { 1 => ParagraphAlignment.Center, 2 => ParagraphAlignment.Right, 3 => ParagraphAlignment.Justify, _ => ParagraphAlignment.Left },
            LeftIndent = p.Left / 20.0, RightIndent = p.Right / 20.0, FirstLineIndent = p.First / 20.0,
            SpaceBefore = p.Before / 20.0, SpaceAfter = p.After / 20.0,
            LineSpacing = p.LineMult == 1 ? new LineSpacing(Math.Max(0.1, p.Line / 240.0), LineSpacingRule.Auto)
                : p.Line < 0 ? new LineSpacing(-p.Line / 20.0, LineSpacingRule.Exact) : p.Line > 0 ? new LineSpacing(p.Line / 20.0, LineSpacingRule.AtLeast) : LineSpacing.Single,
            KeepWithNext = p.KeepNext, KeepLinesTogether = p.Keep, PageBreakBefore = p.PageBreakBefore, TabStops = [.. p.Tabs],
            Frame = p.Framed && !p.InTable ? FrameOf(p) : null,
            Borders = p.Borders.Any(b => b.IsVisible)
                ? new BorderSet { Top = p.Borders[0], Left = p.Borders[1], Bottom = p.Borders[2], Right = p.Borders[3], InsideHorizontal = p.Borders[4] }
                : BorderSet.Empty,
        };

        // The label takes the paragraph mark's formatting with the level's character sprms; symbol-font bullets become Unicode.
        ListLabel? Label(Pap pap, Chp mark)
        {
            if (LevelOf(pap.Ilfo, pap.Ilvl) is not { } level || NextLabel(pap.Ilfo, pap.Ilvl) is not { } text) return null;
            var chp = mark.Clone();
            ApplyChp(chp, level.Chpx);
            var props = Props(chp);
            if (SymbolFonts.IsSymbolFont(props.FontFamily))
            {
                text = SymbolFonts.Map(props.FontFamily, text, unknown: '•');
                props = props with { FontFamily = FontName(mark.Font) };
            }
            return new ListLabel(text, props, SuffixOf(level), LabelAlignment(level));
        }

        void EndParagraph(int markCp, bool cellMark)
        {
            if (++_blocks > MaxBlocks) throw new InvalidDataException("The document has too many paragraphs.");
            var pap = PapAt(markCp);
            var markChp = ChpAt(markCp, pap);
            runChp ??= markChp;
            Flush();
            if (pap.Ttp && cellMark)
            {
                FinishRow(pap);
                inlines.Clear();
                return;
            }
            var paragraph = new Paragraph { Properties = ParagraphProps(pap), MarkProperties = Props(markChp), ListLabel = Label(pap, markChp) };
            paragraph.Inlines.AddRange(inlines);
            inlines.Clear();
            if (pap.InTable)
            {
                currentTable ??= new Table();
                rowCells ??= [];
                cellBlocks ??= [];
                cellBlocks.Add(paragraph);
                if (cellMark)
                {
                    rowCells.Add(new TableCell { Blocks = { } });
                    rowCells[^1].Blocks.AddRange(cellBlocks);
                    cellBlocks = null;
                }
                tablePap = pap;
                return;
            }
            CloseTable();
            blocks.Add(paragraph);
        }

        void FinishRow(Pap ttp)
        {
            if (rowCells is null || currentTable is null) { rowCells = null; cellBlocks = null; return; }
            var row = new TableRow();
            var edges = new List<int>();
            var tcs = new List<(int Flags, BorderSet Borders)>();
            if (ttp.TableDef is { Length: > 3 } def)
            {
                int itcMac = def[2];
                if (3 + (itcMac + 1) * 2 <= def.Length)
                {
                    for (int i = 0; i <= itcMac; i++) edges.Add(I16(def, 3 + 2 * i));
                    int tcAt = 3 + (itcMac + 1) * 2;
                    for (int i = 0; i < itcMac; i++)
                    {
                        if (tcAt + 20 * (i + 1) > def.Length) { tcs.Add((0, BorderSet.Empty)); continue; }
                        int p = tcAt + 20 * i;
                        tcs.Add((U16(def, p), new BorderSet { Top = Brc(def, p + 4), Left = Brc(def, p + 8), Bottom = Brc(def, p + 12), Right = Brc(def, p + 16) }));
                    }
                }
            }
            // Cells merged to the left (fMerged) are covered by the cell that starts the merge.
            for (int i = 0; i < rowCells.Count; i++)
            {
                var cell = rowCells[i];
                var (flags, borders) = i < tcs.Count ? tcs[i] : (0, BorderSet.Empty);
                double width = i + 1 < edges.Count ? (edges[i + 1] - edges[i]) / 20.0 : 0;
                if ((flags & 2) != 0 && row.Cells.Count > 0)
                {
                    var previous = row.Cells[^1];
                    row.Cells[^1] = new TableCell
                    {
                        Width = new TableWidth(previous.Width.Value + width, TableWidthUnit.Points), GridSpan = previous.GridSpan + 1, VerticalMerge = previous.VerticalMerge,
                        Borders = previous.Borders, Shading = previous.Shading,
                    };
                    row.Cells[^1].Blocks.AddRange(previous.Blocks);
                    continue;
                }
                var merge = (flags & 0x40) != 0 ? VerticalMerge.Restart : (flags & 0x20) != 0 ? VerticalMerge.Continue : VerticalMerge.None;
                var built = new TableCell { Width = width > 0 ? new TableWidth(width, TableWidthUnit.Points) : TableWidth.Auto, Borders = borders, VerticalMerge = merge };
                built.Blocks.AddRange(cell.Blocks);
                row.Cells.Add(built);
            }
            currentTable.Rows.Add(row);
            rowCells = null;
            cellBlocks = null;
        }

        void CloseTable()
        {
            if (currentTable is null) return;
            if (rowCells is { Count: > 0 } || cellBlocks is { Count: > 0 })
            {
                // A table cut off before its row-end mark: close the row with what there is.
                rowCells ??= [];
                if (cellBlocks is { Count: > 0 }) { rowCells.Add(new TableCell()); rowCells[^1].Blocks.AddRange(cellBlocks); }
                FinishRow(tablePap ?? new Pap());
            }
            if (currentTable.Rows.Count > 0)
            {
                // The grid is the widths of the row with the most cells.
                var widest = currentTable.Rows.OrderByDescending(r => r.Cells.Sum(c => c.GridSpan)).First();
                foreach (var cell in widest.Cells)
                    for (int s = 0; s < cell.GridSpan; s++) currentTable.Grid.Add(cell.Width.Unit == TableWidthUnit.Points ? cell.Width.Value / cell.GridSpan : 0);
                if (currentTable.Grid.Any(w => w <= 0)) currentTable.Grid.Clear();
                blocks.Add(currentTable);
            }
            currentTable = null;
            tablePap = null;
        }

        int cp = from;
        int fieldDepth = 0;
        int resultHidden = int.MaxValue; // the depth whose result is replaced by a live page-number field
        for (; cp < to; cp++)
        {
            char c = text[cp];
            // Field codes are not shown: only the result (between the separator and the end) is. PAGE, NUMPAGES and SECTIONPAGES
            // become fields evaluated at layout, in place of their stored result.
            if (c == '\x13') { fieldStack.Add(false); fieldDepth++; fieldCode.Clear(); continue; }
            if (c == '\x14' && fieldStack.Count > 0)
            {
                fieldStack[^1] = true;
                if (fieldStack.Count < resultHidden && !fieldStack.Take(fieldStack.Count - 1).Any(shown => !shown) && PageField(fieldCode.ToString()) is { } kind)
                {
                    var chp = ChpAt(cp, PapAt(cp));
                    Flush();
                    inlines.Add(new FieldRun(kind, Props(chp)));
                    resultHidden = fieldStack.Count;
                }
                continue;
            }
            if (c == '\x15' && fieldStack.Count > 0)
            {
                if (fieldStack.Count == resultHidden) resultHidden = int.MaxValue;
                fieldStack.RemoveAt(fieldStack.Count - 1); fieldDepth--; continue;
            }
            if (fieldStack.Count > 0 && !fieldStack[^1]) { if (fieldCode.Length < 256) fieldCode.Append(c); continue; }
            if (fieldStack.Count >= resultHidden || fieldStack.Count > 0 && fieldStack.Any(shown => !shown)) continue;

            switch (c)
            {
                case '\r' or '\a':
                    EndParagraph(cp, cellMark: c == '\a');
                    runChp = null;
                    continue;
                case '\f':
                {
                    Flush();
                    bool sectionBreak = !story && sectionEnds.Contains(cp + 1) && sectionIndex + 1 < sectionEnds.Count;
                    if (sectionBreak)
                    {
                        CloseTable();
                        sectionIndex++;
                        section = MakeSection(sectionIndex);
                        doc.Sections.Add(section);
                        blocks = section.Blocks;
                    }
                    else inlines.Add(new Break(BreakKind.Page, Props(runChp ?? new Chp { Font = _defaultFont })));
                    continue;
                }
                case '\v': Flush(); inlines.Add(new Break(BreakKind.Line, Props(runChp ?? new Chp { Font = _defaultFont }))); continue;
                case '\t':
                {
                    var chp = ChpAt(cp, PapAt(cp));
                    Flush();
                    inlines.Add(new Tab(Props(chp)));
                    continue;
                }
                case '\x08':
                    Flush();
                    AddFloating(cp, Props(ChpAt(cp, PapAt(cp))), inlines);
                    continue;
                case '\x01':
                    if (ChpAt(cp, PapAt(cp)) is { PicLocation: >= 0, Ole2: false, FData: false } pictureChp && InlinePicture(pictureChp.PicLocation, Props(pictureChp)) is { } picture)
                    {
                        Flush();
                        inlines.Add(picture);
                    }
                    continue;
                case '\x02' when !story && NoteAt(cp, doc) is { } note:
                    Flush();
                    inlines.Add(new NoteReference(note, Props(ChpAt(cp, PapAt(cp)))));
                    continue;
                case '\x02' when story && _currentNote is { } owner:
                    Flush();
                    inlines.Add(new NoteMark(owner, Props(ChpAt(cp, PapAt(cp)))));
                    continue;
                case '\x02' or '\x05' or '\x0E' or '\x0F': continue; // anchors of notes and comments
                case '\x03' or '\x04': continue; // the note separator lines, drawn by layout
            }
            if (c == '\u001E') c = '-';
            else if (c == '\u001F') continue;
            var pap = PapAt(cp);
            var current = ChpAt(cp, pap);
            if (runChp is null || !SameFormat(runChp, current)) { Flush(); runChp = current; }
            if (!current.Hidden) run.Append(c);
        }
        if (run.Length > 0 || inlines.Count > 0) EndParagraph(Math.Max(from, to - 1), cellMark: false);
        CloseTable();
        _ = fieldDepth;

        // Remove empty trailing sections created by a final section break.
        for (int s = doc.Sections.Count - 1; s > 0 && doc.Sections[s].Blocks.Count == 0; s--) doc.Sections.RemoveAt(s);
        if (doc.Sections[0].Blocks.Count == 0) doc.Sections[0].Blocks.Add(new Paragraph());
        return doc;
    }

    private static bool SameFormat(Chp a, Chp b) =>
        a.Bold == b.Bold && a.Italic == b.Italic && a.Strike == b.Strike && a.Underline == b.Underline && a.Size == b.Size && a.Font == b.Font
        && a.Color == b.Color && a.Script == b.Script && a.SmallCaps == b.SmallCaps && a.Caps == b.Caps && a.Hidden == b.Hidden;

    private static Border Brc(ReadOnlySpan<byte> d, int at)
    {
        int width = d[at], type = d[at + 1], ico = d[at + 2];
        if (width == 0xFF || type is 0 or 0xFF || width == 0) return Border.None;
        var kind = type switch { 3 => BorderStyle.Double, 6 or 7 => BorderStyle.Dotted, 5 or 8 => BorderStyle.Dashed, 2 => BorderStyle.Thick, _ => BorderStyle.Single };
        return new Border(kind, BrcWidth(width, type), ColorOf(ico == 0 ? -1 : ico) ?? Color.Black);
    }

    /// <summary>
    /// A frame from the frame sprms: horizontal position from the column, margin or page (pcHorz 0-2) at an offset or
    /// aligned (dxaAbs -4 center, -8 right, -12 inside, -16 outside; 0 left); vertical from the margin, page or paragraph
    /// (pcVert 0-2), offset or aligned (dyaAbs -4 top, -8 center, -12 bottom, -16 inside, -20 outside). A height has the
    /// minimum flag in bit 15 (exact when clear).
    /// </summary>
    private static FrameProperties FrameOf(Pap p)
    {
        static (FloatAlignment, double) Pos(int v) => v switch
        {
            -4 => (FloatAlignment.Center, 0), -8 => (FloatAlignment.End, 0), -12 => (FloatAlignment.Inside, 0), -16 => (FloatAlignment.Outside, 0),
            _ => (FloatAlignment.Offset, v / 20.0),
        };
        static (FloatAlignment, double) VPos(int v) => v switch
        {
            -4 => (FloatAlignment.Start, 0), -8 => (FloatAlignment.Center, 0), -12 => (FloatAlignment.End, 0), -16 => (FloatAlignment.Inside, 0), -20 => (FloatAlignment.Outside, 0),
            _ => (FloatAlignment.Offset, v / 20.0),
        };
        var (hAlign, x) = p.DxaAbs == 0 ? (FloatAlignment.Start, 0.0) : Pos(p.DxaAbs);
        var (vAlign, y) = VPos(p.DyaAbs);
        var hRef = (p.Pc >> 6 & 3) switch { 1 => FloatReference.Margin, 2 => FloatReference.Page, _ => FloatReference.Column };
        var vRef = (p.Pc >> 4 & 3) switch { 1 => FloatReference.Page, 2 => FloatReference.Paragraph, _ => FloatReference.Margin };
        int height = p.FrameHeight & 0x7FFF;
        double dx = p.DxaFromText / 20.0, dy = p.DyaFromText / 20.0;
        return new FrameProperties
        {
            Width = p.FrameWidth > 0 ? p.FrameWidth / 20.0 : null,
            Height = height > 0 ? height / 20.0 : null,
            HeightExact = height > 0 && (p.FrameHeight & 0x8000) == 0,
            Horizontal = new FloatPosition(hRef, hAlign, x),
            Vertical = new FloatPosition(vRef, vAlign, y),
            Distance = new Thickness(dy, dx, dy, dx),
            Wrap = p.Wr == 1 ? TextWrap.TopAndBottom : TextWrap.Square,
        };
    }

    /// <summary>A BRC (the 8-byte border of the Word 2000 sprms): COLORREF, width in eighths of a point, type, space in points.</summary>
    private static Border BrcNew(ReadOnlySpan<byte> d)
    {
        int width = d[4], type = d[5];
        if (type is 0 or 0xFF || width == 0) return Border.None;
        var kind = type switch { 3 => BorderStyle.Double, 6 or 7 => BorderStyle.Dotted, 5 or 8 => BorderStyle.Dashed, 2 => BorderStyle.Thick, _ => BorderStyle.Single };
        var color = d[3] == 0xFF ? Color.Black : new Color(d[0], d[1], d[2]);
        return new Border(kind, BrcWidth(width, type), color, d[6] & 0x1F);
    }

    /// <summary>A border's total width in points: a double border (type 3) is two lines of the given width with a gap of the same width.</summary>
    private static double BrcWidth(int eighths, int type) => Math.Max(0.25, eighths / 8.0) * (type == 3 ? 3 : 1);

    private static Color? ColorOf(int c)
    {
        if (c < 0) return null;
        if ((c & 0x1000000) != 0) return new Color((byte)(c >> 16), (byte)(c >> 8), (byte)c);
        return c switch
        {
            1 => Color.Black, 2 => new Color(0, 0, 255), 3 => new Color(0, 255, 255), 4 => new Color(0, 255, 0), 5 => new Color(255, 0, 255), 6 => new Color(255, 0, 0),
            7 => new Color(255, 255, 0), 8 => Color.White, 9 => new Color(0, 0, 128), 10 => new Color(0, 128, 128), 11 => new Color(0, 128, 0), 12 => new Color(128, 0, 128),
            13 => new Color(128, 0, 0), 14 => new Color(128, 128, 0), 15 => new Color(128, 128, 128), 16 => new Color(192, 192, 192), _ => null,
        };
    }
}
