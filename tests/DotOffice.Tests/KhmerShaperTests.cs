using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;

namespace DotOffice.Tests;

/// <summary>
/// The Khmer shaper. The expectations are relations that follow from the rules of Microsoft's Khmer script
/// development page as HarfBuzz implements them; whole words were also compared with LibreOffice (which shapes Khmer
/// with HarfBuzz) pixel for pixel. Neither is output of Word.
/// </summary>
public class KhmerShaperTests
{
    private static readonly TrueTypeFont Khmer = TrueTypeFont.Load(TestAssets.Path("NotoSansKhmer-Regular.ttf"));

    private static List<GlyphInfo> Shape(string text) => KhmerShaper.Shape(Khmer, text, "khmr")!;
    private static int[] Ids(string text) => Shape(text).Select(g => (int)g.Glyph).ToArray();
    private static int Nominal(int codePoint) => Khmer.GetGlyphId(codePoint);

    [Theory]
    [InlineData(0x1780, (byte)KhmerCategory.C)]
    [InlineData(0x179A, (byte)KhmerCategory.Ra)]
    [InlineData(0x17A5, (byte)KhmerCategory.V)]
    [InlineData(0x17D2, (byte)KhmerCategory.H)] // coeng
    [InlineData(0x17C1, (byte)KhmerCategory.VPre)]
    [InlineData(0x17B7, (byte)KhmerCategory.VAbv)]
    [InlineData(0x17BB, (byte)KhmerCategory.VBlw)]
    [InlineData(0x17B6, (byte)KhmerCategory.VPst)]
    [InlineData(0x17CC, (byte)KhmerCategory.Robatic)]
    [InlineData(0x17C9, (byte)KhmerCategory.Robatic)] // muusikatoan
    [InlineData(0x17C6, (byte)KhmerCategory.Xgroup)]
    [InlineData(0x17C7, (byte)KhmerCategory.Ygroup)]
    [InlineData(0x200D, (byte)KhmerCategory.ZWJ)]
    [InlineData(0x25CC, (byte)KhmerCategory.DottedCircle)]
    public void Characters_have_the_Khmer_category_of_the_HarfBuzz_table(int codePoint, byte expected)
    {
        var starts = KhmerData.RangeStarts;
        int index = 0;
        while (index + 1 < starts.Length && starts[index + 1] <= codePoint) index++;
        Assert.Equal(expected, KhmerData.RangeValues[index]);
    }

    [Fact]
    public void Only_the_Khmer_script_is_shaped_here()
    {
        Assert.Null(KhmerShaper.Shape(Khmer, "ក", "latn"));
    }

    [Fact]
    public void A_left_vowel_sign_moves_before_its_consonant()
    {
        var ka = Ids("ក");
        var shaped = Ids("កេ");
        Assert.Equal(2, shaped.Length);
        Assert.Equal(ka[0], shaped[1]);
        Assert.Equal(Nominal(0x17C1), shaped[0]);
    }

    [Fact]
    public void A_split_vowel_sign_is_decomposed_into_the_left_vowel_and_the_sign()
    {
        // 17C4 (OO) has no Unicode decomposition; HarfBuzz splits it into 17C1 + 17C4 and the font draws the right part.
        // The font's GSUB then forms the base with the right part (one glyph); the left vowel stays in front.
        var shaped = Ids("កោ");
        Assert.Equal(2, shaped.Length);
        Assert.Equal(Nominal(0x17C1), shaped[0]);
        Assert.NotEqual(Ids("ក")[0], shaped[1]);
        Assert.Equal(shaped[0], Ids("កេ")[0]);
    }

    [Fact]
    public void Coeng_ro_moves_before_the_base()
    {
        // KA, COENG, RO: the subscript ro is drawn first (pref), the base last.
        var shaped = Ids("ក្រ");
        Assert.Equal(Ids("ក")[0], shaped[^1]);
        Assert.NotEqual(Ids("ក")[0], shaped[0]);
        Assert.NotEqual(Nominal(0x179A), shaped[^1]);
    }

    [Fact]
    public void Another_subscript_stays_after_the_base()
    {
        var shaped = Ids("ក្ខ");
        Assert.Equal(Ids("ក")[0], shaped[0]);
    }

    [Fact]
    public void Coeng_ro_goes_before_the_base_and_before_a_left_vowel_sign_stays_first()
    {
        // The left vowel goes to the very start, in front of the pref glyphs, when it comes last in the syllable.
        var shaped = Shape("ក្រេ");
        Assert.Equal(Nominal(0x17C1), shaped[0].Glyph);
        Assert.Equal(Ids("ក")[0], shaped[^1].Glyph);
    }

    [Fact]
    public void A_broken_cluster_gets_a_dotted_circle()
    {
        // A coeng with nothing before it: the circle comes first and the coeng sign after it.
        var shaped = Ids("្ក");
        Assert.Equal(Nominal(0x25CC), shaped[0]);
        Assert.True(shaped.Length >= 2);
    }

    [Fact]
    public void Syllables_shape_independently()
    {
        Assert.Equal(Ids("កេ").Concat(Ids("ក្ខ")), Ids("កេក្ខ"));
    }

    [Fact]
    public void Each_syllable_gets_its_text_on_its_first_glyph()
    {
        var shaped = Shape("កេ");
        Assert.Equal("កេ", shaped[0].Text);
        Assert.Equal(0, shaped[0].Cluster);
    }
}
