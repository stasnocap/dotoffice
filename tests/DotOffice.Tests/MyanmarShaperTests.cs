using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;

namespace DotOffice.Tests;

/// <summary>
/// The Myanmar shaper. The expectations are relations that follow from the rules of Microsoft's Myanmar script
/// development page as HarfBuzz implements them; whole words were also compared with LibreOffice (which shapes
/// Myanmar with HarfBuzz) at 300 dpi. Neither is output of Word.
/// </summary>
public class MyanmarShaperTests
{
    private static readonly TrueTypeFont Myanmar = TrueTypeFont.Load(TestAssets.Path("NotoSansMyanmar-Regular.ttf"));
    private static readonly TrueTypeFont Latin = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));

    private static List<GlyphInfo> Shape(string text) => MyanmarShaper.Shape(Myanmar, text, "mymr")!;
    private static int[] Ids(string text) => Shape(text).Select(g => (int)g.Glyph).ToArray();
    private static int Nominal(int codePoint) => Myanmar.GetGlyphId(codePoint);

    [Theory]
    [InlineData(0x1000, (byte)MyanmarCategory.C)]
    [InlineData(0x1004, (byte)MyanmarCategory.Ra)] // nga forms the kinzi
    [InlineData(0x1025, (byte)MyanmarCategory.IV)]
    [InlineData(0x1031, (byte)MyanmarCategory.VPre)]
    [InlineData(0x102D, (byte)MyanmarCategory.VAbv)]
    [InlineData(0x102F, (byte)MyanmarCategory.VBlw)]
    [InlineData(0x102C, (byte)MyanmarCategory.VPst)]
    [InlineData(0x1037, (byte)MyanmarCategory.DB)]
    [InlineData(0x1038, (byte)MyanmarCategory.SM)]
    [InlineData(0x1036, (byte)MyanmarCategory.A)]
    [InlineData(0x1039, (byte)MyanmarCategory.H)]
    [InlineData(0x103A, (byte)MyanmarCategory.As)]
    [InlineData(0x103B, (byte)MyanmarCategory.MY)]
    [InlineData(0x103C, (byte)MyanmarCategory.MR)]
    [InlineData(0x103D, (byte)MyanmarCategory.MW)]
    [InlineData(0x103E, (byte)MyanmarCategory.MH)]
    [InlineData(0x1060, (byte)MyanmarCategory.ML)]
    [InlineData(0x1063, (byte)MyanmarCategory.PT)]
    [InlineData(0xFE00, (byte)MyanmarCategory.VS)]
    [InlineData(0x25CC, (byte)MyanmarCategory.DottedCircle)]
    public void Characters_have_the_Myanmar_category_of_the_HarfBuzz_table(int codePoint, byte expected)
    {
        var starts = MyanmarData.RangeStarts;
        int index = 0;
        while (index + 1 < starts.Length && starts[index + 1] <= codePoint) index++;
        Assert.Equal(expected, MyanmarData.RangeValues[index]);
    }

    [Fact]
    public void HarfBuzz_uses_this_shaper_only_for_fonts_with_a_mym2_script()
    {
        Assert.True(MyanmarShaper.Applies(Myanmar.OpenTypeLayout));
        Assert.False(MyanmarShaper.Applies(Latin.OpenTypeLayout)); // a font designed for latn/DFLT gets the default shaper
        Assert.Null(MyanmarShaper.Shape(Latin, "က", "mymr"));
        Assert.Null(MyanmarShaper.Shape(Myanmar, "က", "latn"));
    }

    [Fact]
    public void The_left_vowel_sign_moves_before_its_consonant()
    {
        var shaped = Ids("ကေ");
        Assert.Equal(2, shaped.Length);
        Assert.Equal(Nominal(0x1031), shaped[0]);
        Assert.Equal(Ids("က")[0], shaped[1]);
    }

    [Fact]
    public void Medial_ra_goes_before_the_base_and_the_left_vowel_before_that()
    {
        var withRa = Ids("ကြ");
        Assert.Equal(Ids("က")[0], withRa[^1]);
        Assert.NotEqual(Ids("က")[0], withRa[0]);
        var withE = Ids("ကြေ");
        Assert.Equal(3, withE.Length);
        Assert.Equal(Nominal(0x1031), withE[0]);
        Assert.Equal(withRa, withE[1..]);
    }

    [Fact]
    public void The_vowel_sign_e_stays_in_front_when_other_medials_follow_the_base()
    {
        var shaped = Ids("ကျွေ");
        Assert.Equal(Nominal(0x1031), shaped[0]);
    }

    [Fact]
    public void The_kinzi_is_drawn_after_the_base_it_belongs_to()
    {
        // NGA, ASAT, VIRAMA, KA: the base is KA; the kinzi (a mark glyph) comes after it.
        var shaped = Shape("င်္က");
        Assert.Equal(Ids("က")[0], shaped[0].Glyph);
        Assert.Equal(2, shaped.Count);
    }

    [Fact]
    public void A_broken_cluster_gets_a_dotted_circle_that_the_left_vowel_goes_before()
    {
        var shaped = Ids("ေ");
        Assert.Equal([Nominal(0x1031), Nominal(0x25CC)], shaped);
    }

    [Fact]
    public void Syllables_shape_independently()
    {
        Assert.Equal(Ids("ကေ").Concat(Ids("ကြ")), Ids("ကေကြ"));
    }

    [Fact]
    public void Each_syllable_gets_its_text_on_its_first_glyph()
    {
        var shaped = Shape("ကေ");
        Assert.Equal("ကေ", shaped[0].Text);
    }

    [Fact]
    public void Myanmar_goes_through_the_text_shaper()
    {
        var shaped = TextShaper.Shape(Myanmar, "ကေ", false, kerning: false);
        Assert.Equal(Ids("ကေ"), shaped.Select(g => (int)g.GlyphId));
        Assert.True(TextShaper.NeedsShaping(Myanmar, "က", kerning: false));
    }
}
