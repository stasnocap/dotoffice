using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class TextBoxDirectionTests
{
    private static string Shape(string vert, string anchor = "t") =>
        "<wps:spPr><a:prstGeom prst=\"rect\"/></wps:spPr>" +
        "<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/></w:rPr><w:t>vertical text here</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
        $"<wps:bodyPr lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" anchor=\"{anchor}\"{(vert.Length > 0 ? $" vert=\"{vert}\"" : "")}/>";

    private static LayoutPage Layout(string vert, string anchor = "t")
    {
        var body = P(AnchorShape(60, 200, Position("page", "1270000", "page", "2540000"), Shape(vert, anchor))) + Section();
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
    }

    [Theory]
    [InlineData("vert270", 90)]
    [InlineData("vert", -90)]
    public void Text_is_turned_a_quarter_about_the_center_of_the_body(string vert, double rotation)
    {
        var turn = Layout(vert).Items.OfType<PushItem>().Single(p => p.Rotation != 0);
        Assert.Equal(rotation, turn.Rotation);
        Assert.Equal(100 + 30, turn.CenterX, 0.5);
        Assert.Equal(200 + 100, turn.CenterY, 0.5);
    }

    [Fact]
    public void Lines_are_as_long_as_the_box_is_high()
    {
        // 200 pt long: the text fits on one line before the turn.
        var texts = Texts(Layout("vert270")).ToList();
        Assert.Single(texts.Select(t => Math.Round(t.Baseline, 1)).Distinct());
        Assert.InRange(texts[0].X, 130 - 100 - 0.5, 130 - 100 + 0.5);
    }

    [Fact]
    public void Horizontal_boxes_are_not_turned()
    {
        Assert.DoesNotContain(Layout("").Items, i => i is PushItem { Rotation: not 0 });
    }
}
