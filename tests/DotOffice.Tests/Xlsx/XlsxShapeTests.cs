using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxShapeTests
{
    private static readonly byte[] Png = PngTestEncoder.Encode(2, 2, 2, 8, [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]);

    private static (Worksheet Sheet, ConversionWarningList Warnings) Read(string anchors)
    {
        var warnings = new ConversionWarningList();
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: DrawingElement).AddDrawing(1, anchors, Png);
        var wb = XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings });
        return (wb.Sheets[0], warnings);
    }

    private static SheetShape ReadOne(string content, string anchor = "one")
    {
        var (sheet, warnings) = Read(anchor == "one" ? OneCellAnchor(1, 1, 100, 50, content) : TwoCellAnchor(1, 1, 3, 3, content));
        Assert.Empty(warnings.Warnings);
        return Assert.Single(sheet.Shapes);
    }

    private const string SolidRed = "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>";

    [Fact]
    public void Reads_geometry_fill_and_outline_from_the_shape_properties()
    {
        var shape = ReadOne(Sp("<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"ellipse\"/>" + SolidRed +
                               "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln>"));
        Assert.Equal(ShapeGeometry.Ellipse, shape.Geometry);
        Assert.Equal(new Color(255, 0, 0), shape.Fill);
        Assert.Equal(new Color(0, 0, 255), shape.LineColor);
        Assert.Equal((2.0, StrokeStyle.Dashed), (shape.LineWidth, shape.LineStyle));
        Assert.Equal((ImageAnchorKind.OneCell, 100.0, 50.0), (shape.Anchor, shape.Width, shape.Height));
    }

    [Fact]
    public void Style_references_supply_the_theme_fill_and_outline_when_the_shape_states_none()
    {
        var style = "<xdr:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef>" +
                    "<a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>" +
                    "<a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></xdr:style>";
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>", style));
        Assert.Equal(Color.TryParseHex("4472C4"), shape.Fill);
        Assert.Equal(1.0, shape.LineWidth);
        var line = shape.LineColor!.Value;
        Assert.True(line.R < 0x44 && line.B < 0xC4 && line.B > 0x40); // accent1 darkened by the 50 % shade
    }

    [Fact]
    public void Explicit_no_fill_and_no_line_override_the_style()
    {
        var style = "<xdr:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef></xdr:style>";
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/><a:noFill/><a:ln><a:noFill/></a:ln>", style, "<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>hi</a:t></a:r></a:p></xdr:txBody>"));
        Assert.Null(shape.Fill);
        Assert.Null(shape.LineColor);
    }

    [Fact]
    public void Reads_the_text_with_alignment_runs_anchor_and_insets()
    {
        var body = "<xdr:txBody><a:bodyPr anchor=\"ctr\" lIns=\"127000\" wrap=\"none\"/><a:p><a:pPr algn=\"ctr\"/>" +
                   "<a:r><a:rPr sz=\"1400\" b=\"1\"><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill><a:latin typeface=\"Arial\"/></a:rPr><a:t>Hello </a:t></a:r>" +
                   "<a:r><a:rPr i=\"1\"/><a:t>world</a:t></a:r></a:p><a:p><a:endParaRPr sz=\"2000\"/></a:p></xdr:txBody>";
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed, "", body));
        Assert.Equal(ShapeTextAnchor.Middle, shape.TextAnchor);
        Assert.False(shape.WrapText);
        Assert.Equal(10.0, shape.TextInsets.Left);
        Assert.Equal(2, shape.Paragraphs.Count);
        var first = shape.Paragraphs[0];
        Assert.Equal(HorizontalAlignment.Center, first.Alignment);
        Assert.Equal(["Hello ", "world"], first.Runs.Select(r => r.Text));
        Assert.Equal((14.0, true, "Arial", new Color(0, 255, 0)), (first.Runs[0].Font.FontSize, first.Runs[0].Font.Bold, first.Runs[0].Font.FontFamily, first.Runs[0].Font.Color!.Value));
        Assert.True(first.Runs[1].Font.Italic);
        Assert.Equal(11.0, first.Runs[1].Font.FontSize);
        Assert.Equal(20.0, shape.Paragraphs[1].FontSize);
    }

    [Fact]
    public void A_text_box_without_fill_or_line_is_kept_for_its_text()
    {
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/><a:noFill/>", "", "<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>note</a:t></a:r></a:p></xdr:txBody>"));
        Assert.Null(shape.Fill);
        Assert.Equal("note", shape.Paragraphs[0].Runs[0].Text);
    }

    [Fact]
    public void Connectors_and_lines_are_straight_lines_with_flips()
    {
        var connector = "<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"5\" name=\"Connector\"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm flipV=\"1\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1\" cy=\"1\"/></a:xfrm>" +
                        "<a:prstGeom prst=\"straightConnector1\"/><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln></xdr:spPr></xdr:cxnSp>";
        var shape = ReadOne(connector);
        Assert.Equal(ShapeGeometry.Line, shape.Geometry);
        Assert.True(shape.FlipV);
        Assert.Equal(new Color(255, 0, 0), shape.LineColor);
    }

    [Fact]
    public void Rotation_is_counter_clockwise_and_hidden_or_unprinted_shapes_are_skipped()
    {
        var shape = ReadOne(Sp("<a:xfrm rot=\"1800000\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1\" cy=\"1\"/></a:xfrm><a:prstGeom prst=\"rect\"/>" + SolidRed));
        Assert.Equal(-30.0, shape.Rotation);
        var (sheet, _) = Read(OneCellAnchor(0, 0, 10, 10, Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed, nvAttributes: " hidden=\"1\"")) +
                              TwoCellAnchor(0, 0, 2, 2, Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed), clientData: "<xdr:clientData fPrintsWithSheet=\"0\"/>"));
        Assert.Empty(sheet.Shapes);
    }

    [Fact]
    public void Unknown_geometry_is_drawn_as_a_rectangle_with_one_approximation_warning_and_custom_geometry_is_not_drawn()
    {
        var (sheet, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Sp("<a:prstGeom prst=\"star5\"/>" + SolidRed)) + OneCellAnchor(2, 2, 10, 10, Sp("<a:prstGeom prst=\"wave\"/>" + SolidRed)) +
                                     OneCellAnchor(4, 4, 10, 10, CustomShape));
        Assert.Equal([ShapeGeometry.Rectangle, ShapeGeometry.Rectangle], sheet.Shapes.Select(s => s.Geometry));
        Assert.Equal(["xlsx.drawing", "xlsx.drawing-approximated"], warnings.Warnings.Select(w => w.Code).OrderBy(c => c));
        Assert.Contains("star5", warnings.Warnings.Single(w => w.Code == "xlsx.drawing-approximated").Message);
    }

    [Fact]
    public void Group_children_are_placed_inside_the_frame_of_the_group()
    {
        var child = (string fill, int x, int y, int cx, int cy) =>
            Sp($"<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"/><a:solidFill><a:srgbClr val=\"{fill}\"/></a:solidFill>");
        var group = "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"9\" name=\"Group\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>" +
                    "<xdr:grpSpPr><a:xfrm><a:off x=\"1000\" y=\"1000\"/><a:ext cx=\"4000\" cy=\"2000\"/><a:chOff x=\"1000\" y=\"1000\"/><a:chExt cx=\"4000\" cy=\"2000\"/></a:xfrm></xdr:grpSpPr>" +
                    child("FF0000", 1000, 1000, 2000, 1000) + child("00FF00", 3000, 2000, 2000, 1000) + "</xdr:grpSp>";
        var (sheet, _) = Read(OneCellAnchor(1, 1, 200, 100, group));
        Assert.Equal(2, sheet.Shapes.Count);
        Assert.Equal(new FrameFraction(0, 0, 0.5, 0.5), sheet.Shapes[0].Frame);
        Assert.Equal(new FrameFraction(0.5, 0.5, 0.5, 0.5), sheet.Shapes[1].Frame);
        Assert.All(sheet.Shapes, s => Assert.Equal((200.0, 100.0), (s.Width, s.Height)));
    }

    [Fact]
    public void Pictures_and_shapes_share_one_drawing_order()
    {
        var (sheet, _) = Read(OneCellAnchor(0, 0, 10, 10, Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed)) + OneCellAnchor(0, 0, 10, 10, Pic()) + OneCellAnchor(0, 0, 10, 10, Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed)));
        Assert.Equal([0, 2], sheet.Shapes.Select(s => s.ZOrder));
        Assert.Equal(1, sheet.Images[0].ZOrder);
    }
}
