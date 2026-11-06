using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;

namespace DotOffice.Tests;

/// <summary>
/// The Universal Shaping Engine shaper (Sinhala). The expected glyphs are what DotOffice produces for words whose
/// page rendering matched LibreOffice (which shapes Sinhala with HarfBuzz's USE) pixel for pixel; they are not
/// output of Word.
/// </summary>
public class UseShaperTests
{
    private static readonly TrueTypeFont Sinhala = TrueTypeFont.Load(TestAssets.Path("NotoSansSinhala-Regular.ttf"));
    private static readonly TrueTypeFont Latin = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));

    private static List<GlyphInfo> Shape(string text) => UseShaper.Shape(Sinhala, text, "sinh")!;
    private static int[] Ids(string text) => Shape(text).Select(g => (int)g.Glyph).ToArray();

    [Theory]
    [InlineData(0x0DCA, (byte)UseCategory.HVM)] // al-lakuna is a halant that also acts as a vowel modifier
    [InlineData(0x0D9A, (byte)UseCategory.B)]
    [InlineData(0x0DD9, (byte)UseCategory.VPre)]
    [InlineData(0x0DCF, (byte)UseCategory.VPst)]
    [InlineData(0x0DD2, (byte)UseCategory.VAbv)]
    [InlineData(0x0DD4, (byte)UseCategory.VBlw)]
    [InlineData(0x0D82, (byte)UseCategory.VMPst)]
    [InlineData(0x200D, (byte)UseCategory.CGJ)]
    [InlineData(0x200C, (byte)UseCategory.ZWNJ)]
    [InlineData(0x25CC, (byte)UseCategory.B)]
    [InlineData(0x0F71, (byte)UseCategory.CMBlw)] // Tibetan, from the Microsoft additions
    [InlineData(0x0F90, (byte)UseCategory.SUB)]
    public void Characters_have_the_USE_category_of_the_Unicode_properties(int codePoint, byte expected)
    {
        var starts = UseData.RangeStarts;
        int index = 0;
        while (index + 1 < starts.Length && starts[index + 1] <= codePoint) index++;
        Assert.Equal(expected, UseData.RangeValues[index]);
    }

    [Fact]
    public void A_left_vowel_sign_moves_before_its_consonant()
    {
        // Logical order: KA, E. Visual order: E, KA (the pre-base vowel goes to the start of the syllable).
        var logical = Ids("ක");
        var shaped = Ids("කෙ");
        Assert.Equal(2, shaped.Length);
        Assert.Equal(logical[0], shaped[1]);
        Assert.NotEqual(logical[0], shaped[0]);
    }

    [Fact]
    public void Split_vowel_signs_are_decomposed_around_the_consonant()
    {
        // KA + KOMBUVA HAA AELA-PILLA = E, KA, AELA-PILLA: three glyphs, the consonant in the middle.
        var shaped = Shape("කො");
        Assert.Equal([114, 23, 82], shaped.Select(g => (int)g.Glyph));
        Assert.Equal(Ids("කෙ")[0], shaped[0].Glyph); // the same E sign as in KA+E
    }

    [Fact]
    public void A_consonant_with_al_lakuna_and_zwj_and_ra_takes_the_rakaransaya_mark_with_no_advance()
    {
        // The rakaransaya glyph is a GDEF mark with an advance in hmtx; USE zeroes mark advances (ZERO_WIDTH_MARKS_BY_GDEF_EARLY).
        var shaped = Shape("ක්‍රම");
        Assert.Equal([23, 130, 52], shaped.Select(g => (int)g.Glyph));
        Assert.Equal(0, shaped[1].XAdvance);
        Assert.True(shaped[1].XOffset < 0); // the mark is positioned back over the consonant by GPOS
        Assert.Equal(1007, shaped[0].XAdvance);
    }

    [Fact]
    public void A_ligature_of_the_font_forms_over_ZWJ()
    {
        Assert.Single(Ids("ක්‍ෂ")); // KA, AL, ZWJ, SSA -> one conjunct glyph
    }

    [Fact]
    public void Zwnj_prevents_the_conjunct()
    {
        var joined = Ids("ක්‍ෂ");
        var blocked = Ids("ක‌්ෂ");
        Assert.Single(joined);
        Assert.True(blocked.Length > 1);
    }

    [Fact]
    public void A_vowel_sign_that_would_make_another_vowel_gets_a_dotted_circle()
    {
        // HarfBuzz's vowel constraints: AYANNA + AELA-PILLA looks like AAYANNA, so the sign gets a circle of its own.
        var shaped = Shape("අා");
        Assert.Equal(3, shaped.Count);
        Assert.Equal(Sinhala.GetGlyphId(0x25CC), shaped[1].Glyph);
    }

    [Fact]
    public void A_sign_without_a_base_is_a_broken_cluster_with_a_dotted_circle()
    {
        var shaped = Shape("ා");
        Assert.Equal(2, shaped.Count);
        Assert.Equal(Sinhala.GetGlyphId(0x25CC), shaped[0].Glyph);
    }

    [Fact]
    public void Syllables_follow_the_USE_grammar()
    {
        // KA AL ZWJ RA is one cluster (halant + base), MA starts the next one; ZWJ is invisible to the scanner.
        var shaped = Shape("ක්‍රම");
        Assert.Equal(shaped[0].Syllable, shaped[1].Syllable);
        Assert.NotEqual(shaped[1].Syllable, shaped[2].Syllable);
        // Two consonants without a halant are two syllables.
        var two = Shape("කම");
        Assert.NotEqual(two[0].Syllable, two[1].Syllable);
    }

    [Fact]
    public void Clusters_cover_the_text_of_each_syllable()
    {
        var shaped = TextShaper.Shape(Sinhala, "කොම", false, kerning: true);
        Assert.Equal("කොම", string.Concat(shaped.Select(g => g.Text)));
        Assert.Equal(0, shaped[0].Cluster);
    }

    [Fact]
    public void A_font_without_a_Sinhala_script_gets_the_default_shaper()
    {
        Assert.Null(UseShaper.Shape(Latin, "ක", "sinh"));
        Assert.Null(UseShaper.Shape(Sinhala, "ක", "latn"));
    }

    [Fact]
    public void Sinhala_goes_through_the_text_shaper_and_is_not_one_glyph_per_character()
    {
        var shaped = TextShaper.Shape(Sinhala, "කො", false, kerning: true);
        Assert.Equal(3, shaped.Length);
        Assert.True(TextShaper.NeedsShaping(Sinhala, "කො", kerning: false));
    }
}
