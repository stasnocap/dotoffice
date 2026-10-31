using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;

namespace DotOffice.Tests;

/// <summary>Expected glyphs and positions were produced with HarfBuzz from the same fonts.</summary>
public class TextShaperTests
{
    private static readonly TrueTypeFont Arabic = TrueTypeFont.Load(TestAssets.Path("NotoSansArabic-Regular.ttf"));
    private static readonly TrueTypeFont Hebrew = TrueTypeFont.Load(TestAssets.Path("NotoSansHebrew-Regular.ttf"));
    private static readonly TrueTypeFont Latin = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));
    private static readonly TrueTypeFont Nastaliq = TrueTypeFont.Load(TestAssets.Path("NotoNastaliqUrdu-Regular.ttf"));

    private static (int, int, int, int)[] Positions(ShapedGlyph[] glyphs) =>
        glyphs.Select(g => ((int)g.GlyphId, g.Advance, g.XOffset, g.YOffset)).ToArray();

    [Fact]
    public void Arabic_letters_take_their_joining_forms()
    {
        var glyphs = TextShaper.Shape(Arabic, "سلام", rightToLeft: true, kerning: false);
        Assert.Equal([(37, 784, 0, 0), (71, 220, 0, 0), (11, 379, 0, 0), (74, 484, 0, 0)], Positions(glyphs));
        Assert.Equal([0, 1, 2, 3], glyphs.Select(g => g.Cluster));
        Assert.Equal("سلام", string.Concat(glyphs.Select(g => g.Text)));

        // The same letter differs in isolated, initial, medial and final position.
        int Form(string text, int index) => TextShaper.Shape(Arabic, text, true, false)[index].GlyphId;
        var forms = new[] { Form("ع", 0), Form("عع", 0), Form("ععع", 1), Form("عع", 1) };
        Assert.Equal(4, forms.Distinct().Count());
    }

    [Fact]
    public void Marks_are_positioned_on_their_base_and_decompositions_keep_text_on_the_first_glyph()
    {
        var glyphs = TextShaper.Shape(Arabic, "بِ", rightToLeft: true, kerning: false);
        Assert.Equal([(14, 993, 0, 0), (316, 0, 410, -23), (426, 0, 344, -211)], Positions(glyphs));
        Assert.Equal(["ب", "", "ِ"], glyphs.Select(g => g.Text)); // the dot glyph comes from decomposing the letter
    }

    [Fact]
    public void Cursive_attachment_joins_nastaliq_letters_on_a_slope()
    {
        Assert.Equal([(951, 0, 0, 0), (188, 607, 0, 95), (953, 0, 0, 0), (399, 201, 0, 0), (4, 297, 0, 0), (436, 660, 0, 0)],
            Positions(TextShaper.Shape(Nastaliq, "سلام", rightToLeft: true, kerning: false)));
        Assert.Equal([(3, 330, 0, 0), (143, 406, 0, 0), (126, 300, 0, 0), (551, 422, 0, 0)],
            Positions(TextShaper.Shape(Nastaliq, "اردو", rightToLeft: true, kerning: false)));
        var glyphs = TextShaper.Shape(Nastaliq, "پاکستان", rightToLeft: true, kerning: false);
        Assert.Equal([361, 0, 145, 0], glyphs.Where((_, i) => i is 6 or 7 or 8 or 9).Select(g => g.YOffset)); // the chain climbs
    }

    [Fact]
    public void Hebrew_marks_are_reordered_canonically()
    {
        var glyphs = TextShaper.Shape(Hebrew, "שָׁ", rightToLeft: true, kerning: false);
        Assert.Equal([(96, 730, 0, 0), (100, 0, 539, 0), (79, 0, 227, 0)], Positions(glyphs));
        Assert.Equal(["ש", "ׁ", "ָ"], glyphs.Select(g => g.Text));
    }

    [Fact]
    public void Kerning_is_applied_only_when_requested()
    {
        Assert.Equal([(36, 599, 0, 0), (57, 600, 0, 0)], Positions(TextShaper.Shape(Latin, "AV", false, kerning: true)));
        Assert.Equal([(36, 639, 0, 0), (57, 600, 0, 0)], Positions(TextShaper.Shape(Latin, "AV", false, kerning: false)));
        Assert.False(TextShaper.NeedsShaping("plain text", kerning: false));
        Assert.True(TextShaper.NeedsShaping("plain text", kerning: true));
        Assert.True(TextShaper.NeedsShaping("שלום", kerning: false));
        Assert.True(TextShaper.NeedsShaping("é", kerning: false));
    }

    [Fact]
    public void Shaping_does_not_stop_after_many_runs_in_one_font()
    {
        // The lookup operation limit applies per call; it used to be shared by every call on a font.
        var font = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));
        string text = string.Concat(Enumerable.Repeat("AVTAWAY ", 20));
        var expected = Positions(TextShaper.Shape(font, text, false, kerning: true));
        for (int i = 0; i < 2000; i++) TextShaper.Shape(font, text, false, kerning: true);
        Assert.Equal(expected, Positions(TextShaper.Shape(font, text, false, kerning: true)));
    }

    [Fact]
    public void Fonts_without_layout_tables_map_characters_one_to_one()
    {
        var glyphs = TextShaper.Shape(Latin, "سلام", rightToLeft: true, kerning: false); // no Arabic glyphs: .notdef
        Assert.Equal(4, glyphs.Length);
        Assert.All(glyphs, g => Assert.Equal(0, g.GlyphId));
    }
}
