using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;

namespace DotOffice.Tests;

/// <summary>Expected glyphs and positions were produced with HarfBuzz from the same fonts.</summary>
public class TextShaperTests
{
    private static readonly TrueTypeFont Arabic = TrueTypeFont.Load(TestAssets.Path("NotoSansArabic-Regular.ttf"));
    private static readonly TrueTypeFont Hebrew = TrueTypeFont.Load(TestAssets.Path("NotoSansHebrew-Regular.ttf"));
    private static readonly TrueTypeFont Latin = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));

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
    public void Fonts_without_layout_tables_map_characters_one_to_one()
    {
        var glyphs = TextShaper.Shape(Latin, "سلام", rightToLeft: true, kerning: false); // no Arabic glyphs: .notdef
        Assert.Equal(4, glyphs.Length);
        Assert.All(glyphs, g => Assert.Equal(0, g.GlyphId));
    }
}
