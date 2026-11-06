using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;

namespace DotOffice.Tests;

/// <summary>
/// A non-spacing mark with no base before it is Inherited for the script split, so its run has no script and
/// HarfBuzz guesses one from the mark (hb_buffer_guess_segment_properties). The shaped output was also compared with
/// libharfbuzz for these inputs; it is not output of Word.
/// </summary>
public class LoneMarkTests
{
    [Theory]
    [InlineData("NotoSansKhmer-Regular.ttf", 0x17BB, "khmr")]
    [InlineData("NotoSansMyanmar-Regular.ttf", 0x103A, "mymr")]
    [InlineData("NotoSansMyanmar-Regular.ttf", 0x1085, "mymr")]
    [InlineData("NotoSansDevanagari-Regular.ttf", 0x093C, "deva")]
    [InlineData("NotoSansSinhala-Regular.ttf", 0x0DCA, "sinh")]
    public void A_lone_mark_takes_the_script_of_the_mark_and_sits_on_a_dotted_circle(string file, int mark, string tag)
    {
        var font = TrueTypeFont.Load(TestAssets.Path(file));
        var text = char.ConvertFromUtf32(mark);
        Assert.Equal(tag, Assert.Single(ScriptRuns.Split(text)).Tag);
        var glyphs = TextShaper.Shape(font, text, rightToLeft: false, kerning: true);
        Assert.Equal(2, glyphs.Length);
        Assert.Equal(font.GetGlyphId(0x25CC), glyphs[0].GlyphId);
    }

    [Fact]
    public void A_mark_after_a_base_keeps_the_script_of_the_base_and_a_lone_Latin_mark_stays_without_script()
    {
        Assert.Equal("latn", Assert.Single(ScriptRuns.Split("aុ")).Tag);
        Assert.Equal("DFLT", Assert.Single(ScriptRuns.Split("́")).Tag); // Zinh in the UCD too
    }
}
