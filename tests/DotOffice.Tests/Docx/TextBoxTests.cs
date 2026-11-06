using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using DotOffice.Tests.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class TextBoxTests
{
    private const string NotoRun = "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/>";

    private static Document Read(string body, List<ConversionWarning>? warnings = null)
    {
        var b = new DocxBuilder { Body = body + Section() };
        return DocxReader.Read(b.BuildStream(), new DocxReaderOptions { Warnings = warnings is null ? null : new ListSink(warnings) });
    }

    private sealed class ListSink(List<ConversionWarning> list) : IConversionWarningSink
    {
        public void Warn(ConversionWarning warning) => list.Add(warning);
    }

    private static LayoutPage Layout(Document doc) =>
        new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];

    private const string Box =
        "<wps:spPr><a:xfrm rot=\"0\"><a:off x=\"0\" y=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"/><a:solidFill><a:srgbClr val=\"FFFF00\"/></a:solidFill>" +
        "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln></wps:spPr>" +
        "<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/></w:rPr><w:t>inside</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
        "<wps:bodyPr lIns=\"127000\" tIns=\"254000\" rIns=\"0\" bIns=\"0\" anchor=\"t\"/>";

    [Fact]
    public void Reads_fill_outline_insets_content_and_position_of_a_text_box()
    {
        var doc = Read(P(AnchorShape(200, 100, Position("page", "1270000", "page", "2540000"), Box)));
        var box = Assert.Single(((Paragraph)doc.Sections[0].Blocks[0]).Inlines.OfType<FloatingTextBox>());
        Assert.Equal((200.0, 100.0), (box.Width, box.Height));
        Assert.Equal(new Color(255, 255, 0), box.Fill);
        Assert.Equal(new Color(255, 0, 0), box.Outline);
        Assert.Equal(2.0, box.OutlineWidth);
        Assert.Equal(new Thickness(20, 0, 0, 10), box.Insets);
        Assert.Equal(new FloatPosition(FloatReference.Page, FloatAlignment.Offset, 100), box.Horizontal);
        Assert.Single(box.Blocks);
    }

    [Fact]
    public void Lays_the_content_inside_the_box_below_its_insets_and_clips_it()
    {
        var page = Layout(Read(P(AnchorShape(200, 100, Position("page", "1270000", "page", "2540000"), Box))));
        var fill = Assert.Single(page.Items.OfType<RectItem>());
        Assert.Equal((100.0, 200.0, 200.0, 100.0), (fill.X, fill.Y, fill.Width, fill.Height));
        var text = Assert.Single(Texts(page), t => t.Text.Contains("inside"));
        Assert.Equal(110.0, text.X, 0.01);                       // box x + left inset
        Assert.True(text.Baseline > 220 && text.Baseline < 300); // below the 20 pt top inset, inside the box
        Assert.Equal(4, page.Items.OfType<LineItem>().Count());   // outline
        int push = page.Items.FindIndex(i => i is PushItem { Clip: not null });
        Assert.True(push >= 0 && page.Items.IndexOf(text) > push && page.Items.FindIndex(push, i => i is PopItem) > page.Items.IndexOf(text));
    }

    [Fact]
    public void Square_wrap_moves_body_text_aside_and_none_does_not()
    {
        string body = R(string.Join(' ', Enumerable.Range(1, 60).Select(i => $"w{i}")), NotoRun);
        double Left(string wrap) => Texts(Layout(Read(P(AnchorShape(150, 60, Position("margin", "left", "paragraph", "0"), Box, wrap) + body))))
            .Where(t => !t.Text.Contains("inside") && t.Baseline < 72 + 40).Min(t => t.X);
        Assert.True(Left("<wp:wrapSquare wrapText=\"bothSides\"/>") >= 72 + 150 - 0.01);
        Assert.True(Left("<wp:wrapNone/>") < 72 + 150);
    }

    [Fact]
    public void Unsupported_geometry_keeps_the_text_and_warns_but_a_plain_filled_rectangle_is_drawn()
    {
        var warnings = new List<ConversionWarning>();
        var ellipse = Read(P(AnchorShape(100, 100, Position("page", "0", "page", "0"), Box.Replace("prst=\"rect\"", "prst=\"ellipse\""))), warnings);
        var box = Assert.Single(((Paragraph)ellipse.Sections[0].Blocks[0]).Inlines.OfType<FloatingTextBox>());
        Assert.Null(box.Fill);
        Assert.Contains(warnings, w => w.Code == "docx.shape.geometry");

        var plain = Read(P(AnchorShape(100, 100, Position("page", "0", "page", "0"),
            "<wps:spPr><a:prstGeom prst=\"rect\"/><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>")));
        var rect = Assert.Single(((Paragraph)plain.Sections[0].Blocks[0]).Inlines.OfType<FloatingTextBox>());
        Assert.Equal(new Color(0x44, 0x72, 0xC4), rect.Fill); // the Office theme's accent1 when the document has no theme
        Assert.Null(rect.Outline);
    }
}
