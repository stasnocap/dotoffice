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
    private static readonly TrueTypeFont Devanagari = TrueTypeFont.Load(TestAssets.Path("NotoSansDevanagari-Regular.ttf"));
    private static readonly TrueTypeFont Bengali = TrueTypeFont.Load(TestAssets.Path("NotoSansBengali-Regular.ttf"));
    private static readonly TrueTypeFont Tamil = TrueTypeFont.Load(TestAssets.Path("NotoSansTamil-Regular.ttf"));

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
        Assert.False(TextShaper.NeedsShaping(Latin, "some words", kerning: false));
        Assert.True(TextShaper.NeedsShaping(Latin, "some words", kerning: true));
        Assert.True(TextShaper.NeedsShaping(Latin, "fill", kerning: false)); // the "fi" ligature
        Assert.True(TextShaper.NeedsShaping(Latin, "שלום", kerning: false));
        Assert.True(TextShaper.NeedsShaping(Latin, "é", kerning: false));
    }

    [Fact]
    public void Indic_syllables_are_reordered_and_form_conjuncts()
    {
        // Pre-base matra i moved before its consonant, half forms, reph, split matras and Tamil's shri ligature.
        Assert.Equal([(544, 259, 0, 0), (88, 531, 0, 0), (245, 309, 0, 0), (73, 531, 0, 0), (33, 259, 0, 0)],
            Positions(TextShaper.Shape(Devanagari, "हिन्दी", false, kerning: true)));
        Assert.Equal([(56, 768, 0, 0), (506, 0, -221, 0)], Positions(TextShaper.Shape(Devanagari, "र्क", false, kerning: true)));
        Assert.Equal([(398, 361, 0, 0), (25, 807, 0, 0), (59, 266, 0, 0), (55, 623, 0, 0), (132, 0, -137, 0)],
            Positions(TextShaper.Shape(Bengali, "কোর্স", false, kerning: true)));
        Assert.Equal([(135, 859, 0, 0)], Positions(TextShaper.Shape(Bengali, "ক্ষ", false, kerning: true)));
        Assert.Equal([(46, 901, 0, 0), (18, 825, 0, 0), (41, 640, 0, 0)], Positions(TextShaper.Shape(Tamil, "கொ", false, kerning: true)));
        Assert.Equal([(163, 1514, 0, 0)], Positions(TextShaper.Shape(Tamil, "ஸ்ரீ", false, kerning: true)));
    }

    [Fact]
    public void Joiners_select_half_forms_or_break_conjuncts_and_are_hidden()
    {
        // ZWJ asks for the half form, ZWNJ for the explicit virama; both become an invisible space glyph.
        Assert.Equal([(232, 550, 0, 0), (3, 0, 0, 0), (86, 578, 0, 0)],
            Positions(TextShaper.Shape(Devanagari, "क्\u200dष", false, kerning: true)));
        Assert.Equal([(56, 768, 0, 0), (103, 0, -221, 0), (3, 0, 0, 0), (86, 578, 0, 0)],
            Positions(TextShaper.Shape(Devanagari, "क्\u200cष", false, kerning: true)));
        Assert.Equal([(73, 344, 0, 0), (3, 0, 0, 0), (76, 258, 0, 0)], Positions(TextShaper.Shape(Latin, "f\u200ci", false, kerning: true)));
    }

    [Fact]
    public void Default_ligatures_apply_unless_turned_off()
    {
        var glyphs = TextShaper.Shape(Latin, "office", false, kerning: true);
        Assert.Equal([(82, 605, 0, 0), (1656, 946, 0, 0), (70, 480, 0, 0), (72, 564, 0, 0)], Positions(glyphs));
        Assert.Equal(("ffi", 1), (glyphs[1].Text, glyphs[1].Cluster));
        Assert.Equal(6, TextShaper.Shape(Latin, "office", false, kerning: true, ligatures: false).Length);
    }

    [Fact]
    public void Letters_and_marks_compose_or_decompose_to_the_glyphs_the_font_has()
    {
        var composed = TextShaper.Shape(Latin, "e\u0301", false, kerning: true);
        Assert.Equal([(171, 564, 0, 0)], Positions(composed));
        Assert.Equal("e\u0301", composed[0].Text);
        // No glyph for U+2260: equals sign and combining long solidus, positioned by GPOS.
        var decomposed = TextShaper.Shape(Latin, "\u2260", false, kerning: true);
        Assert.Equal([(32, 572, 0, 0), (2716, 0, -286, 2)], Positions(decomposed));
        Assert.Equal(["\u2260", ""], decomposed.Select(g => g.Text));
        // No precomposed q with marks: canonical order (dot below first), then mark positioning.
        Assert.Equal([(84, 615, 0, 0), (2696, 0, 169, -240), (2667, 0, -309, 0)],
            Positions(TextShaper.Shape(Latin, "q\u0302\u0323", false, kerning: true)));
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
