using DotOffice.Fonts;

namespace DotOffice.Tests;

public class TrueTypeFontTests
{
    private static readonly TrueTypeFont Font = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));

    [Fact]
    public void Reads_basic_metadata()
    {
        Assert.Equal("NotoSans-Regular", Font.PostScriptName);
        Assert.Equal("Noto Sans", Font.FamilyName);
        Assert.Equal(1000, Font.UnitsPerEm);
        Assert.True(Font.NumGlyphs > 1000);
        Assert.True(Font.Ascender > 0);
        Assert.True(Font.Descender < 0);
        Assert.False(Font.IsBold);
        Assert.False(Font.IsItalic);
        Assert.False(Font.IsEmbeddingRestricted);
    }

    [Theory]
    [InlineData('A')]
    [InlineData('z')]
    [InlineData('Ж')]
    [InlineData('Ω')]
    public void Maps_common_characters(char c)
    {
        Assert.True(Font.HasGlyph(c));
        Assert.True(Font.GetAdvanceWidth(Font.GetGlyphId(c)) > 0);
    }

    [Fact]
    public void Unmapped_character_returns_notdef()
    {
        Assert.Equal(0, Font.GetGlyphId(0x1F600)); // emoji not in Noto Sans
    }

    [Fact]
    public void Measures_text_proportionally_to_size()
    {
        var at10 = Font.MeasureText("Hello", 10);
        Assert.True(at10 > 0);
        Assert.Equal(at10 * 2, Font.MeasureText("Hello", 20), 6);
        Assert.True(Font.MeasureText("WWW", 10) > Font.MeasureText("iii", 10));
    }

    [Fact]
    public void Rejects_non_font_data()
    {
        Assert.Throws<InvalidDataException>(() => TrueTypeFont.Load(new byte[64]));
    }
}
