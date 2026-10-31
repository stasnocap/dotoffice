using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Layout;

namespace DotOffice.Tests;

public class ItemizerTests
{
    private static readonly FontResolver Fonts = new(new FontCollection()
        .AddFile(TestAssets.Path("NotoSans-Regular.ttf"))
        .AddFile(TestAssets.Path("NotoSansDevanagari-Regular.ttf")));

    private static List<Atom> Itemize(string text)
    {
        var paragraph = new Paragraph();
        paragraph.Inlines.Add(new TextRun(text, RunProperties.Default with { FontFamily = "Noto Sans", ComplexScriptFontFamily = "Noto Sans" }));
        return new Itemizer(new FontSelector(Fonts, null)).Itemize(paragraph, _ => "");
    }

    [Fact]
    public void Joiners_stay_with_their_letter_in_a_fallback_font()
    {
        // Noto Sans has a ZWJ glyph but no Devanagari: the whole syllable is shaped in the fallback font, so the
        // ZWJ selects the half form of KA (HarfBuzz's glyph 232) instead of drawing a visible virama.
        var atom = Assert.Single(Itemize("क्‍ष"));
        Assert.Equal("Noto Sans Devanagari", atom.Font!.Font.FamilyName);
        Assert.Equal([232, 3, 86], atom.Glyphs!.Select(g => (int)g.GlyphId));
    }
}
