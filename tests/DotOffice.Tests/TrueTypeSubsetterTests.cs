using DotOffice.Fonts;

namespace DotOffice.Tests;

public class TrueTypeSubsetterTests
{
    private static readonly TrueTypeFont Font = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));

    [Fact]
    public void Subset_is_much_smaller_and_keeps_used_outlines()
    {
        var gids = "Hello".Select(c => Font.GetGlyphId(c)).ToArray();
        var subset = TrueTypeSubsetter.Subset(Font, gids);

        Assert.True(subset.Length < Font.GetTableData("glyf").Length / 10);

        var reparsedTables = ParseTableTags(subset);
        Assert.Contains("glyf", reparsedTables);
        Assert.Contains("loca", reparsedTables);
        Assert.DoesNotContain("cmap", reparsedTables);

        // Same glyph ids, identical outline bytes for kept glyphs, empty for dropped ones.
        var h = Font.GetGlyphId('H');
        var x = Font.GetGlyphId('x');
        var parsed = new SubsetReader(subset);
        Assert.Equal(Font.GetGlyphData(h).ToArray(), parsed.Glyph(h));
        Assert.Empty(parsed.Glyph(x));
        Assert.Equal(Font.NumGlyphs, parsed.NumGlyphs);
    }

    [Fact]
    public void Subset_includes_composite_glyph_components()
    {
        // Accented letters in Noto Sans are composites referencing base glyphs.
        var composite = Font.GetGlyphId('é');
        var parsed = new SubsetReader(TrueTypeSubsetter.Subset(Font, [composite]));

        var data = Font.GetGlyphData(composite);
        Assert.True(System.Buffers.Binary.BinaryPrimitives.ReadInt16BigEndian(data) < 0, "expected composite glyph");
        Assert.NotEmpty(parsed.Glyph(Font.GetGlyphId('e')));
    }

    [Fact]
    public void Subset_has_valid_font_checksum()
    {
        var subset = TrueTypeSubsetter.Subset(Font, [Font.GetGlyphId('A')]);
        uint sum = 0;
        for (int i = 0; i < subset.Length; i += 4)
            sum += System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(subset.AsSpan(i));
        Assert.Equal(0xB1B0AFBA, sum);
    }

    private static HashSet<string> ParseTableTags(byte[] font) => new SubsetReader(font).Tables.Keys.ToHashSet();

    private sealed class SubsetReader
    {
        private readonly byte[] _data;
        public Dictionary<string, (int Offset, int Length)> Tables { get; } = [];
        public int NumGlyphs { get; }

        public SubsetReader(byte[] data)
        {
            _data = data;
            int n = U16(4);
            for (int i = 0; i < n; i++)
            {
                int rec = 12 + i * 16;
                Tables[System.Text.Encoding.ASCII.GetString(data, rec, 4)] = ((int)U32(rec + 8), (int)U32(rec + 12));
            }
            NumGlyphs = U16(Tables["maxp"].Offset + 4);
        }

        public byte[] Glyph(int gid)
        {
            int loca = Tables["loca"].Offset, glyf = Tables["glyf"].Offset;
            int start = (int)U32(loca + gid * 4), end = (int)U32(loca + gid * 4 + 4);
            return _data.AsSpan(glyf + start, end - start).ToArray();
        }

        private int U16(int o) => System.Buffers.Binary.BinaryPrimitives.ReadUInt16BigEndian(_data.AsSpan(o));
        private uint U32(int o) => System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(_data.AsSpan(o));
    }
}
