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
        var ellipse = Read(P(AnchorShape(100, 100, Position("page", "0", "page", "0"), Box.Replace("prst=\"rect\"", "prst=\"cloud\""))), warnings);
        var box = Assert.Single(((Paragraph)ellipse.Sections[0].Blocks[0]).Inlines.OfType<FloatingTextBox>());
        Assert.Null(box.Fill);
        Assert.Contains(warnings, w => w.Code == "docx.shape.geometry");

        var plain = Read(P(AnchorShape(100, 100, Position("page", "0", "page", "0"),
            "<wps:spPr><a:prstGeom prst=\"rect\"/><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>")));
        var rect = Assert.Single(((Paragraph)plain.Sections[0].Blocks[0]).Inlines.OfType<FloatingTextBox>());
        Assert.Equal(new Color(0x44, 0x72, 0xC4), rect.Fill); // the Office theme's accent1 when the document has no theme
        Assert.Null(rect.Outline);
    }

    private static string Geometry(string prst, string fill = "<a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill>", string ln = "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill></a:ln>", string text = "") =>
        $"<wps:spPr><a:prstGeom prst=\"{prst}\"/>{fill}{ln}</wps:spPr>{text}<wps:bodyPr lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\"/>";

    [Fact]
    public void An_ellipse_is_filled_and_outlined_as_an_ellipse_not_as_the_frame()
    {
        var page = Layout(Read(P(AnchorShape(100, 60, Position("page", "0", "page", "0"), Geometry("ellipse")))));
        var fill = Assert.Single(page.Items.OfType<EllipseItem>());
        Assert.Equal((100.0, 60.0), (fill.Width, fill.Height));
        Assert.Empty(page.Items.OfType<RectItem>());
        // The outline follows the ellipse: its points lie on the ellipse, so none is at a corner of the frame.
        var lines = page.Items.OfType<LineItem>().Where(l => l.Color == new Color(0, 0, 255)).ToList();
        Assert.True(lines.Count > 20);
        Assert.All(lines, l => Assert.InRange(Math.Pow((l.X1 - 50) / 50, 2) + Math.Pow((l.Y1 - 30) / 30, 2), 0.98, 1.02));
    }

    [Fact]
    public void A_rounded_rectangle_cuts_its_corners_by_the_adjust_value()
    {
        double Corner(string adj)
        {
            var avLst = adj.Length == 0 ? "" : $"<a:avLst><a:gd name=\"adj\" fmla=\"val {adj}\"/></a:avLst>";
            var page = Layout(Read(P(AnchorShape(100, 100, Position("page", "0", "page", "0"), Geometry("roundRect").Replace("<a:prstGeom prst=\"roundRect\"/>", $"<a:prstGeom prst=\"roundRect\">{avLst}</a:prstGeom>")))));
            var polygon = Assert.Single(page.Items.OfType<PolygonItem>());
            return polygon.Points.Min(p => Math.Min(p.X, p.Y) + 0 == p.X ? p.Y : p.X); // distance of the first straight edge from the corner
        }
        Assert.True(Corner("50000") > Corner("10000"));
    }

    [Fact]
    public void A_gradient_fill_is_a_real_gradient_clipped_to_the_outline()
    {
        const string gradient = "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst><a:lin ang=\"5400000\"/></a:gradFill>";
        var rect = Layout(Read(P(AnchorShape(100, 60, Position("page", "0", "page", "0"), Geometry("rect", gradient))))).Items.OfType<GradientItem>().Single();
        Assert.Equal(90, rect.Degree);
        Assert.Equal(2, rect.Stops.Count);
        Assert.Equal(new Color(255, 0, 0), rect.Stops[0].Color);
        Assert.Null(rect.Clip);
        var ellipse = Layout(Read(P(AnchorShape(100, 60, Position("page", "0", "page", "0"), Geometry("ellipse", gradient))))).Items.OfType<GradientItem>().Single();
        Assert.NotNull(ellipse.Clip);
    }

    [Fact]
    public void Gradient_fills_no_longer_warn_but_picture_fills_still_do()
    {
        var warnings = new List<ConversionWarning>();
        Read(P(AnchorShape(100, 60, Position("page", "0", "page", "0"), Geometry("rect", "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst></a:gradFill>"))), warnings);
        Assert.DoesNotContain(warnings, w => w.Code == "docx.shape.fill");
        Read(P(AnchorShape(100, 60, Position("page", "0", "page", "0"), Geometry("rect", "<a:pattFill prst=\"pct5\"/>"))), warnings);
        Assert.Contains(warnings, w => w.Code == "docx.shape.fill");
    }

    [Fact]
    public void Text_in_an_ellipse_is_laid_out_in_its_inscribed_rectangle()
    {
        const string content = "<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/></w:rPr><w:t>inside</w:t></w:r></w:p></w:txbxContent></wps:txbx>";
        double X(string prst) => Texts(Layout(Read(P(AnchorShape(200, 100, Position("page", "0", "page", "0"), Geometry(prst, text: content)))))).Single(t => t.Text == "inside").X;
        Assert.Equal(0, X("rect"), 0.01);
        Assert.Equal(200 * 0.14645, X("ellipse"), 0.01);
        Assert.InRange(X("roundRect"), 0.1, 200 * 0.14645);
    }

    [Fact]
    public void An_inline_shape_takes_its_size_in_the_line_and_is_drawn_at_its_position()
    {
        const string shape = "<w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"1270000\" cy=\"635000\"/><wp:docPr id=\"3\" name=\"I\"/>" +
            "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:spPr><a:prstGeom prst=\"rect\"/><a:solidFill><a:srgbClr val=\"FFFF00\"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
        var doc = Read("<w:p><w:r><w:rPr>" + NotoRun + "</w:rPr><w:t>before</w:t></w:r>" + shape + "<w:r><w:rPr>" + NotoRun + "</w:rPr><w:t>after</w:t></w:r></w:p>");
        Assert.Single(((Paragraph)doc.Sections[0].Blocks[0]).Inlines.OfType<InlineShape>());
        var page = Layout(doc);
        var fill = Assert.Single(page.Items.OfType<RectItem>());
        Assert.Equal((100.0, 50.0), (fill.Width, fill.Height));
        var before = Texts(page).Single(t => t.Text == "before");
        var after = Texts(page).Single(t => t.Text == "after");
        Assert.True(fill.X > before.X && after.X >= fill.X + fill.Width - 0.01);
        // The shape stands on the baseline.
        Assert.Equal(before.Baseline, fill.Y + fill.Height, 0.01);
    }
}
