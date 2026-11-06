using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class MirrorMarginTests
{
    private const string Run = "<w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/></w:rPr>";

    // Left 1 inch, right 0.5 inch, gutter 0.25 inch.
    private const string Margins = "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>" +
                                   "<w:pgMar w:top=\"1440\" w:right=\"720\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"360\"/></w:sectPr>";

    private static LayoutDocument Layout(bool mirror)
    {
        string body = $"<w:p><w:r>{Run}<w:t>one</w:t></w:r></w:p><w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>" +
                      $"<w:p><w:r>{Run}<w:t>two</w:t></w:r></w:p><w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>" +
                      $"<w:p><w:r>{Run}<w:t>three</w:t></w:r></w:p>" + Margins;
        var doc = DocxReader.Read(new DocxBuilder { Body = body, Settings = mirror ? "<w:mirrorMargins/>" : null }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc);
    }

    private static double X(LayoutPage page, string text) => Texts(page).First(t => t.Text == text).X;

    [Fact]
    public void Gutter_widens_the_left_margin()
    {
        var doc = Layout(false);
        Assert.Equal(72 + 18, X(doc.Pages[0], "one"), 0.01);
        Assert.Equal(72 + 18, X(doc.Pages[1], "two"), 0.01);
    }

    [Fact]
    public void Mirrored_margins_swap_sides_on_even_pages()
    {
        var doc = Layout(true);
        Assert.Equal(72 + 18, X(doc.Pages[0], "one"), 0.01);
        Assert.Equal(36, X(doc.Pages[1], "two"), 0.01);
        Assert.Equal(72 + 18, X(doc.Pages[2], "three"), 0.01);
    }
}
