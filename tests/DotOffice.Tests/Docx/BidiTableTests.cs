using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class BidiTableTests
{
    private const string Run = "<w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/></w:rPr>";

    private static string Cell(string text) => $"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r>{Run}<w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static string Table(string tblPr) =>
        $"<w:tbl><w:tblPr><w:tblW w:w=\"4000\" w:type=\"dxa\"/>{tblPr}<w:tblLayout w:type=\"fixed\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr>" + Cell("first") + Cell("second") + "</w:tr></w:tbl>" + P("");

    private static (Document Doc, Dictionary<string, double> X) Layout(string tblPr)
    {
        var doc = DocxReader.Read(new DocxBuilder { Body = Table(tblPr) + Section() }.BuildStream());
        var page = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
        return (doc, Texts(page).ToDictionary(t => t.Text, t => t.X));
    }

    [Fact]
    public void Reads_bidiVisual_with_the_default_on_and_explicit_off()
    {
        Assert.False(((Table)Layout("").Doc.Sections[0].Blocks[0]).Properties.RightToLeft);
        Assert.True(((Table)Layout("<w:bidiVisual/>").Doc.Sections[0].Blocks[0]).Properties.RightToLeft);
        Assert.False(((Table)Layout("<w:bidiVisual w:val=\"0\"/>").Doc.Sections[0].Blocks[0]).Properties.RightToLeft);
    }

    [Fact]
    public void A_right_to_left_table_puts_its_first_cell_at_the_right_edge_of_the_text_area()
    {
        var ltr = Layout("").X;
        var rtl = Layout("<w:bidiVisual/>").X;
        Assert.True(ltr["first"] < ltr["second"]);
        Assert.True(rtl["first"] > rtl["second"]);
        // The 200 pt table is flush with the right margin (72 + 468), cells 100 pt wide, 5.4 pt cell margin.
        Assert.Equal(72 + 468 - 100 + 5.4, rtl["first"], 0.5);
        Assert.Equal(72 + 468 - 200 + 5.4, rtl["second"], 0.5);
    }

    [Fact]
    public void Indent_of_a_right_to_left_table_is_taken_from_its_right_edge()
    {
        var plain = Layout("<w:bidiVisual/>").X["first"];
        var indented = Layout("<w:bidiVisual/><w:tblInd w:w=\"400\" w:type=\"dxa\"/>").X["first"];
        Assert.Equal(-20, indented - plain, 0.01);
    }
}
