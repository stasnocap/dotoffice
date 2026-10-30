using System.Buffers.Binary;
using System.Text;

namespace DotOffice.Fonts;

/// <summary>
/// Produces a subset TrueType font that keeps the original glyph ids (so a PDF can use an
/// identity CID-to-GID map) but drops outlines of unused glyphs. No cmap is emitted, which is
/// valid for fonts embedded as CIDFontType2.
/// </summary>
public static class TrueTypeSubsetter
{
    private static readonly string[] CopiedTables = ["cvt ", "fpgm", "hhea", "hmtx", "maxp", "prep"];

    public static byte[] Subset(TrueTypeFont font, IEnumerable<ushort> glyphIds)
    {
        var keep = new HashSet<ushort>();
        var pending = new Stack<ushort>(glyphIds.Append((ushort)0));
        while (pending.Count > 0)
        {
            var gid = pending.Pop();
            if (gid >= font.NumGlyphs || !keep.Add(gid)) continue;
            foreach (var component in CompositeComponents(font.GetGlyphData(gid)))
                if (!keep.Contains(component)) pending.Push(component);
        }

        using var glyf = new MemoryStream();
        var loca = new byte[(font.NumGlyphs + 1) * 4];
        for (int g = 0; g < font.NumGlyphs; g++)
        {
            BinaryPrimitives.WriteUInt32BigEndian(loca.AsSpan(g * 4), (uint)glyf.Length);
            if (!keep.Contains((ushort)g)) continue;
            glyf.Write(font.GetGlyphData((ushort)g));
            while (glyf.Length % 4 != 0) glyf.WriteByte(0);
        }
        BinaryPrimitives.WriteUInt32BigEndian(loca.AsSpan(font.NumGlyphs * 4), (uint)glyf.Length);

        var head = font.GetTableData("head").ToArray();
        BinaryPrimitives.WriteUInt32BigEndian(head.AsSpan(8), 0); // checksumAdjustment
        BinaryPrimitives.WriteInt16BigEndian(head.AsSpan(50), 1); // long loca

        var tables = new SortedDictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["head"] = head,
            ["loca"] = loca,
            ["glyf"] = glyf.ToArray(),
        };
        foreach (var tag in CopiedTables)
            if (font.HasTable(tag)) tables[tag] = font.GetTableData(tag).ToArray();

        return WriteFont(tables);
    }

    private static IEnumerable<ushort> CompositeComponents(ReadOnlySpan<byte> glyph)
    {
        var result = new List<ushort>();
        if (glyph.Length < 10 || BinaryPrimitives.ReadInt16BigEndian(glyph) >= 0) return result;

        int pos = 10;
        while (pos + 4 <= glyph.Length)
        {
            int flags = BinaryPrimitives.ReadUInt16BigEndian(glyph[pos..]);
            result.Add(BinaryPrimitives.ReadUInt16BigEndian(glyph[(pos + 2)..]));
            pos += 4;
            pos += (flags & 0x0001) != 0 ? 4 : 2;     // ARG_1_AND_2_ARE_WORDS
            if ((flags & 0x0008) != 0) pos += 2;      // WE_HAVE_A_SCALE
            else if ((flags & 0x0040) != 0) pos += 4; // WE_HAVE_AN_X_AND_Y_SCALE
            else if ((flags & 0x0080) != 0) pos += 8; // WE_HAVE_A_TWO_BY_TWO
            if ((flags & 0x0020) == 0) break;         // MORE_COMPONENTS
        }
        return result;
    }

    private static byte[] WriteFont(SortedDictionary<string, byte[]> tables)
    {
        int numTables = tables.Count;
        int entrySelector = (int)Math.Floor(Math.Log2(numTables));
        int searchRange = (1 << entrySelector) * 16;

        int offset = 12 + numTables * 16;
        int total = offset + tables.Values.Sum(t => (t.Length + 3) & ~3);
        var output = new byte[total];
        var span = output.AsSpan();

        BinaryPrimitives.WriteUInt32BigEndian(span, 0x00010000);
        BinaryPrimitives.WriteUInt16BigEndian(span[4..], (ushort)numTables);
        BinaryPrimitives.WriteUInt16BigEndian(span[6..], (ushort)searchRange);
        BinaryPrimitives.WriteUInt16BigEndian(span[8..], (ushort)entrySelector);
        BinaryPrimitives.WriteUInt16BigEndian(span[10..], (ushort)(numTables * 16 - searchRange));

        int record = 12, headOffset = -1;
        foreach (var (tag, data) in tables)
        {
            Encoding.ASCII.GetBytes(tag, span[record..]);
            BinaryPrimitives.WriteUInt32BigEndian(span[(record + 4)..], Checksum(data));
            BinaryPrimitives.WriteUInt32BigEndian(span[(record + 8)..], (uint)offset);
            BinaryPrimitives.WriteUInt32BigEndian(span[(record + 12)..], (uint)data.Length);
            data.CopyTo(span[offset..]);
            if (tag == "head") headOffset = offset;
            offset += (data.Length + 3) & ~3;
            record += 16;
        }

        BinaryPrimitives.WriteUInt32BigEndian(span[(headOffset + 8)..], 0xB1B0AFBA - Checksum(output));
        return output;
    }

    private static uint Checksum(ReadOnlySpan<byte> data)
    {
        uint sum = 0;
        int i = 0;
        for (; i + 4 <= data.Length; i += 4) sum += BinaryPrimitives.ReadUInt32BigEndian(data[i..]);
        if (i < data.Length)
        {
            Span<byte> tail = stackalloc byte[4];
            data[i..].CopyTo(tail);
            sum += BinaryPrimitives.ReadUInt32BigEndian(tail);
        }
        return sum;
    }
}
