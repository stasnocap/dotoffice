using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;
using DotOffice.Layout;
using DotOffice.Tests.Docx;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Layout;

public class LigatureTests
{
    [Fact]
    public void Docx_runs_have_no_ligatures_unless_w14_ligatures_asks_for_them()
    {
        var doc = DocxReader.Read(new DocxBuilder
        {
            Body = P(R("a") + R("b", "<w14:ligatures w14:val=\"standard\"/>") + R("c", "<w14:ligatures w14:val=\"none\"/>")),
        }.BuildStream());
        var p = (Paragraph)doc.Sections[0].Blocks[0];
        Assert.Equal([false, true, false], p.Inlines.Select(i => i.Properties.Ligatures).ToArray());
    }

    [Fact]
    public void Ligatures_property_decides_whether_fi_forms_a_ligature_and_defaults_on_for_other_formats()
    {
        Assert.True(RunProperties.Default.Ligatures);
        var fonts = new FontResolver(new FontCollection().Add(Noto));
        var font = fonts.Resolve("Noto Sans", false, false);
        var on = Itemizer.Shape(font, "office fi", 12, RunProperties.Default, rtl: false);
        var off = Itemizer.Shape(font, "office fi", 12, RunProperties.Default with { Ligatures = false }, rtl: false);
        // Without ligatures nothing needs shaping, or the glyph count stays the character count.
        int Count(ShapedGlyph[]? g) => g?.Length ?? "office fi".Length;
        Assert.True(Count(on.Glyphs) < Count(off.Glyphs), "the font should ligate fi when allowed");
    }
}
