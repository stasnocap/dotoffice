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
        var (sheet, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Sp("<a:prstGeom prst=\"cloud\"/>" + SolidRed)) + OneCellAnchor(2, 2, 10, 10, Sp("<a:prstGeom prst=\"wave\"/>" + SolidRed)) +
                                     OneCellAnchor(4, 4, 10, 10, CustomShape));
        Assert.Equal([ShapeGeometry.Rectangle, ShapeGeometry.Rectangle], sheet.Shapes.Select(s => s.Geometry));
        Assert.Equal(["xlsx.drawing", "xlsx.drawing-approximated"], warnings.Warnings.Select(w => w.Code).OrderBy(c => c));
        Assert.Contains("cloud", warnings.Warnings.Single(w => w.Code == "xlsx.drawing-approximated").Message);
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

    private static string Connector(string preset, string avLst = "", string ln = "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln>") =>
        "<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id=\"5\" name=\"Connector\"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1\" cy=\"1\"/></a:xfrm>" +
        $"<a:prstGeom prst=\"{preset}\"><a:avLst>{avLst}</a:avLst></a:prstGeom>{ln}</xdr:spPr></xdr:cxnSp>";

    [Fact]
    public void A_gradient_fill_keeps_its_stops_angle_and_first_color()
    {
        var linear = ReadOne(Sp("<a:prstGeom prst=\"ellipse\"/><a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst><a:lin ang=\"5400000\"/></a:gradFill>"));
        Assert.Equal(new Color(255, 0, 0), linear.Fill);
        var gradient = linear.FillGradient!;
        Assert.Equal((GradientKind.Linear, 90.0), (gradient.Kind, gradient.Degree));
        Assert.Equal([0.0, 1.0], gradient.Stops.Select(g => g.Position));
        var radial = ReadOne(Sp("<a:prstGeom prst=\"rect\"/><a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FFFFFF\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"000000\"/></a:gs></a:gsLst>" +
                                "<a:path path=\"circle\"><a:fillToRect l=\"25000\" t=\"50000\" r=\"75000\" b=\"50000\"/></a:path></a:gradFill>")).FillGradient!;
        Assert.Equal(GradientKind.Path, radial.Kind);
        Assert.Equal((0.25, 0.25, 0.5, 0.5), (radial.Left, radial.Right, radial.Top, radial.Bottom));
        Assert.Null(ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed)).FillGradient);
    }

    [Fact]
    public void Arrowheads_are_read_from_the_line_ends()
    {
        var ends = "<a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill><a:headEnd type=\"oval\" w=\"lg\" len=\"sm\"/><a:tailEnd type=\"triangle\"/></a:ln>";
        var shape = ReadOne(Connector("straightConnector1", ln: ends));
        Assert.Equal(new LineEnd(LineEndKind.Oval, 5, 2), shape.HeadEnd);
        Assert.Equal(new LineEnd(LineEndKind.Triangle, 3, 3), shape.TailEnd);
        Assert.Equal(LineEndKind.None, ReadOne(Connector("line")).TailEnd.Kind);
    }

    [Fact]
    public void Bent_and_curved_connectors_follow_their_routes_and_adjustments()
    {
        var bent = ReadOne(Connector("bentConnector3", "<a:gd name=\"adj1\" fmla=\"val 25000\"/>"));
        Assert.Equal(ShapeGeometry.Line, bent.Geometry);
        var path = Assert.Single(bent.Paths!);
        Assert.False(path.Closed);
        Assert.Equal([(0.0, 0.0), (0.25, 0.0), (0.25, 1.0), (1.0, 1.0)], path.Points);
        Assert.Equal([(0.0, 0.0), (1.0, 0.0), (1.0, 1.0)], ReadOne(Connector("bentConnector2")).Paths![0].Points);
        Assert.Equal(6, ReadOne(Connector("bentConnector5")).Paths![0].Points.Count);
        var curved = ReadOne(Connector("curvedConnector3")).Paths![0].Points;
        Assert.Equal((0.0, 0.0), curved[0]);
        Assert.Equal((1.0, 1.0), curved[^1]);
        Assert.Contains(curved, p => Math.Abs(p.X - 0.5) < 1e-9 && Math.Abs(p.Y - 0.5) < 1e-9); // the S-curve passes the middle
        Assert.True(curved.Count > 10);
    }

    [Fact]
    public void Curved_connectors_with_four_and_five_segments_follow_the_presets_curves()
    {
        var (sheet, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Connector("curvedConnector4")));
        Assert.Single(sheet.Shapes);
        Assert.Empty(warnings.Warnings);
        // curvedConnector4 (adj1 = adj2 = 50 %): runs down the middle (x = 0.5, from y = 0.25 on), crosses at y = 0.5 and ends at the corner.
        var four = ReadOne(Connector("curvedConnector4")).Paths![0].Points;
        Assert.Equal((0.0, 0.0), four[0]);
        Assert.Equal((1.0, 1.0), four[^1]);
        Assert.Equal(49, four.Count);                                                     // three Bézier segments of 16 steps and the start
        Assert.Contains(four, p => Math.Abs(p.X - 0.5) < 1e-9 && Math.Abs(p.Y - 0.25) < 1e-9);   // the end of the first segment (x2, y1)
        Assert.Contains(four, p => Math.Abs(p.X - 0.75) < 1e-9 && Math.Abs(p.Y - 0.5) < 1e-9);   // the end of the second (x3, y4)
        // Both adjustments move the runs: x2 = 20 %, y4 = 80 %.
        var moved = ReadOne(Connector("curvedConnector4", "<a:gd name=\"adj1\" fmla=\"val 20000\"/><a:gd name=\"adj2\" fmla=\"val 80000\"/>")).Paths![0].Points;
        Assert.Contains(moved, p => Math.Abs(p.X - 0.2) < 1e-9 && Math.Abs(p.Y - 0.4) < 1e-9);   // (x2, y1 = y4 / 2)
        Assert.Contains(moved, p => Math.Abs(p.X - 0.6) < 1e-9 && Math.Abs(p.Y - 0.8) < 1e-9);   // (x3 = (1 + x2) / 2, y4)
        // curvedConnector5: vertical runs at x3 = adj1 and x6 = adj3, joined along y4 = adj2.
        var five = ReadOne(Connector("curvedConnector5", "<a:gd name=\"adj1\" fmla=\"val 20000\"/><a:gd name=\"adj2\" fmla=\"val 60000\"/><a:gd name=\"adj3\" fmla=\"val 80000\"/>")).Paths![0].Points;
        Assert.Equal(65, five.Count);
        Assert.Equal((1.0, 1.0), five[^1]);
        Assert.Contains(five, p => Math.Abs(p.X - 0.2) < 1e-9 && Math.Abs(p.Y - 0.3) < 1e-9);    // (x3, y1)
        Assert.Contains(five, p => Math.Abs(p.X - 0.5) < 1e-9 && Math.Abs(p.Y - 0.6) < 1e-9);    // (x1 = (x3 + x6) / 2, y4)
        Assert.Contains(five, p => Math.Abs(p.X - 0.8) < 1e-9 && Math.Abs(p.Y - 0.8) < 1e-9);    // (x6, y5 = (1 + y4) / 2)
    }

    [Fact]
    public void Custom_geometry_becomes_paths_in_fractions_of_the_frame()
    {
        var geometry = "<a:custGeom><a:avLst/><a:pathLst><a:path w=\"100\" h=\"50\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"100\" y=\"0\"/></a:lnTo>" +
                       "<a:cubicBezTo><a:pt x=\"100\" y=\"25\"/><a:pt x=\"75\" y=\"50\"/><a:pt x=\"50\" y=\"50\"/></a:cubicBezTo><a:close/></a:path></a:pathLst></a:custGeom>";
        var shape = ReadOne(Sp(geometry + SolidRed));
        Assert.Equal(ShapeGeometry.Custom, shape.Geometry);
        var path = Assert.Single(shape.Paths!);
        Assert.True(path.Closed);
        Assert.Equal((0.0, 0.0), path.Points[0]);
        Assert.Equal((1.0, 0.0), path.Points[1]);
        Assert.Equal((0.5, 1.0), path.Points[^1]);
    }

    private static string Xfrm(int cx, int cy) => $"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>";

    [Fact]
    public void Custom_geometry_evaluates_guide_formulas_and_adjust_values()
    {
        var geometry = Xfrm(2000, 1000) + "<a:custGeom><a:avLst><a:gd name=\"adj\" fmla=\"val 25000\"/></a:avLst><a:gdLst>" +
                       "<a:gd name=\"x1\" fmla=\"*/ w adj 100000\"/><a:gd name=\"y1\" fmla=\"+- vc 0 hd4\"/><a:gd name=\"x2\" fmla=\"pin 0 x1 wd2\"/>" +
                       "<a:gd name=\"d\" fmla=\"sin h 5400000\"/><a:gd name=\"m\" fmla=\"max x1 d\"/><a:gd name=\"c\" fmla=\"?: adj x1 w\"/></a:gdLst>" +
                       "<a:pathLst><a:path><a:moveTo><a:pt x=\"l\" y=\"t\"/></a:moveTo><a:lnTo><a:pt x=\"x1\" y=\"y1\"/></a:lnTo><a:lnTo><a:pt x=\"m\" y=\"b\"/></a:lnTo><a:lnTo><a:pt x=\"c\" y=\"vc\"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>";
        var (sheet, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Sp(geometry + SolidRed)));
        Assert.Empty(warnings.Warnings);
        var points = Assert.Single(Assert.Single(sheet.Shapes).Paths!).Points;
        Assert.Equal(4, points.Count);
        Assert.Equal((0.0, 0.0), points[0]);
        Assert.Equal((0.25, 0.25), (Math.Round(points[1].X, 6), Math.Round(points[1].Y, 6))); // x1 = w/4, y1 = vc - hd4
        Assert.Equal((0.5, 1.0), (Math.Round(points[2].X, 6), Math.Round(points[2].Y, 6)));   // max(x1 = 500, d = 1000 sin 90°) = 1000
        Assert.Equal((0.25, 0.5), (Math.Round(points[3].X, 6), Math.Round(points[3].Y, 6)));  // adj > 0 picks x1
    }

    [Fact]
    public void Custom_geometry_arcs_may_use_guides_and_formulas_that_cannot_be_evaluated_are_reported()
    {
        var arc = Xfrm(1000, 1000) + "<a:custGeom><a:avLst/><a:gdLst/><a:pathLst><a:path><a:moveTo><a:pt x=\"l\" y=\"vc\"/></a:moveTo>" +
                  "<a:arcTo wR=\"wd2\" hR=\"hd2\" stAng=\"cd2\" swAng=\"cd2\"/><a:close/></a:path></a:pathLst></a:custGeom>";
        var shape = ReadOne(Sp(arc + SolidRed));
        Assert.Equal(1.0, shape.Paths![0].Points.Max(p => p.X), 6); // a half ellipse from the left edge to the right
        var bad = Xfrm(1000, 1000) + "<a:custGeom><a:avLst/><a:gdLst><a:gd name=\"x\" fmla=\"bogus w\"/></a:gdLst><a:pathLst><a:path><a:moveTo><a:pt x=\"x\" y=\"0\"/></a:moveTo></a:path></a:pathLst></a:custGeom>";
        var (sheet, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Sp(bad + SolidRed)));
        Assert.Empty(sheet.Shapes);
        Assert.Single(warnings.Warnings, w => w.Code == "xlsx.drawing" && w.Message.Contains("custom geometry"));
    }

    [Theory]
    [InlineData("ltUpDiag", HatchKind.UpDiagonal, 0.5)]
    [InlineData("dkHorz", HatchKind.Horizontal, 2)]
    [InlineData("narVert", HatchKind.Vertical, 0.75)]
    [InlineData("wdDnDiag", HatchKind.DownDiagonal, 3)]
    [InlineData("cross", HatchKind.Cross, 1)]
    [InlineData("diagCross", HatchKind.DiagonalCross, 0.75)]
    public void Pattern_fills_become_hatches_of_the_foreground_over_the_background(string preset, HatchKind kind, double weight)
    {
        var fill = $"<a:pattFill prst=\"{preset}\"><a:fgClr><a:srgbClr val=\"0000FF\"/></a:fgClr><a:bgClr><a:srgbClr val=\"FFFF00\"/></a:bgClr></a:pattFill>";
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + fill));
        var pattern = Assert.IsType<ShapePattern>(shape.FillPattern);
        Assert.Equal((kind, weight, new Color(0, 0, 255), new Color(255, 255, 0)), (pattern.Kind, pattern.Weight, pattern.Foreground, pattern.Background));
    }

    [Fact]
    public void Percent_patterns_are_a_tint_and_other_fills_have_no_pattern()
    {
        var fill = "<a:pattFill prst=\"pct25\"><a:fgClr><a:srgbClr val=\"000000\"/></a:fgClr><a:bgClr><a:srgbClr val=\"FFFFFF\"/></a:bgClr></a:pattFill>";
        var pattern = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + fill)).FillPattern!;
        Assert.Equal((HatchKind.Tint, 0.25), (pattern.Kind, pattern.Coverage));
        Assert.Null(ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed)).FillPattern);
    }

    private const string Accent1Style = "<xdr:style><a:lnRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"{0}\"><a:schemeClr val=\"accent1\"/></a:fillRef>" +
                                        "<a:effectRef idx=\"{1}\"><a:schemeClr val=\"accent1\"/></a:effectRef></xdr:style>";

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    public void Gradient_style_references_take_the_theme_gradient_in_the_style_color(int idx)
    {
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>", string.Format(Accent1Style, idx, 0)));
        var gradient = Assert.IsType<GradientFill>(shape.FillGradient);
        Assert.Equal((GradientKind.Linear, 90.0, 3), (gradient.Kind, gradient.Degree, gradient.Stops.Count));
        Assert.Equal(shape.Fill, gradient.Stops[0].Color);
        // The gradient is a variation of accent 1 (4472C4): all stops stay bluish and they differ from each other.
        Assert.All(gradient.Stops, g => Assert.True(g.Color.B > g.Color.R));
        Assert.NotEqual(gradient.Stops[0].Color, gradient.Stops[2].Color);
    }

    [Fact]
    public void A_solid_style_reference_and_an_explicit_fill_do_not_become_gradients()
    {
        Assert.Null(ReadOne(Sp("<a:prstGeom prst=\"rect\"/>", string.Format(Accent1Style, 1, 0))).FillGradient);
        Assert.Null(ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed, string.Format(Accent1Style, 3, 0))).FillGradient);
    }

    [Fact]
    public void Shadows_come_from_the_shape_or_from_the_themes_effect_style()
    {
        var own = "<a:effectLst><a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"2700000\"><a:srgbClr val=\"000000\"><a:alpha val=\"40000\"/></a:srgbClr></a:outerShdw></a:effectLst>";
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + own));
        var shadow = Assert.IsType<ShapeShadow>(shape.Shadow);
        Assert.Equal((2.12132, 2.12132, 0.4, Color.Black), (Math.Round(shadow.Dx, 5), Math.Round(shadow.Dy, 5), shadow.Opacity, shadow.Color)); // 3 pt at 45°
        // Default theme: effect style 3 is a shadow that falls straight down (dir 90°).
        var themed = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed, string.Format(Accent1Style, 1, 3))).Shadow!;
        Assert.Equal((0.0, 1.5, 0.63), (Math.Round(themed.Dx, 6), Math.Round(themed.Dy, 6), themed.Opacity));
        Assert.Null(ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed, string.Format(Accent1Style, 1, 1))).Shadow);
        // An empty effect list on the shape switches the style's shadow off.
        Assert.Null(ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + "<a:effectLst/>", string.Format(Accent1Style, 1, 3))).Shadow);
    }

    [Fact]
    public void A_shadow_keeps_its_blur_and_a_blur_without_distance_still_counts()
    {
        string Shadow(string attributes) => $"<a:effectLst><a:outerShdw {attributes}><a:srgbClr val=\"000000\"/></a:outerShdw></a:effectLst>";
        Assert.Equal(4.0, ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + Shadow("blurRad=\"50800\" dist=\"38100\" dir=\"2700000\""))).Shadow!.Blur);
        var glow = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + Shadow("blurRad=\"127000\" dist=\"0\""))).Shadow!;
        Assert.Equal((10.0, 0.0), (glow.Blur, glow.Dx));
        Assert.Null(ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + Shadow("dist=\"0\""))).Shadow);
    }

    [Fact]
    public void A_picture_fill_loads_its_image_and_a_rect_path_gradient_keeps_its_shape()
    {
        var picture = ReadOne(Sp("<a:prstGeom prst=\"ellipse\"/><a:blipFill><a:blip r:embed=\"rId1\"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>"));
        Assert.Equal(ImageFormat.Png, picture.FillImage!.Format);
        Assert.Null(picture.FillGradient);

        const string Stops = "<a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst>";
        GradientFill Gradient(string path) => ReadOne(Sp($"<a:prstGeom prst=\"rect\"/><a:gradFill>{Stops}<a:path path=\"{path}\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path></a:gradFill>")).FillGradient!;
        Assert.True(Gradient("rect").Rectangular);
        Assert.False(Gradient("circle").Rectangular);
    }

    [Fact]
    public void A_picture_fill_without_its_image_leaves_it_out()
    {
        var (sheet, warnings) = Read(OneCellAnchor(1, 1, 100, 50, Sp("<a:prstGeom prst=\"rect\"/><a:blipFill><a:blip r:embed=\"rId9\"/><a:stretch/></a:blipFill><a:ln><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill></a:ln>")));
        Assert.Null(Assert.Single(sheet.Shapes).FillImage);
        Assert.Contains("picture fill whose image is missing", Assert.Single(warnings.Warnings, w => w.Code == "xlsx.drawing-approximated").Message);
    }

    [Fact]
    public void A_tiled_picture_fill_gets_its_tile_size_offset_alignment_and_opacity()
    {
        // The 2 x 2 pixel picture is 1.5 points square at 96 dots per inch; sx/sy scale it.
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/><a:blipFill><a:blip r:embed=\"rId1\"><a:alphaModFix amt=\"40000\"/></a:blip><a:tile tx=\"127000\" ty=\"254000\" sx=\"400000\" sy=\"200000\" algn=\"ctr\"/></a:blipFill>"));
        Assert.NotNull(shape.FillImage);
        var tile = shape.FillTile!;
        Assert.Equal((6.0, 3.0, 10.0, 20.0, "ctr", "none"), (tile.Width, tile.Height, tile.OffsetX, tile.OffsetY, tile.Alignment, tile.Flip));
        var mirrored = ReadOne(Sp("<a:prstGeom prst=\"rect\"/><a:blipFill><a:blip r:embed=\"rId1\"/><a:tile flip=\"xy\"/></a:blipFill>"));
        Assert.Equal("xy", mirrored.FillTile!.Flip);
        Assert.Equal(0.4, shape.FillImageAlpha);
        var plain = ReadOne(Sp("<a:prstGeom prst=\"rect\"/><a:blipFill><a:blip r:embed=\"rId1\"/><a:stretch/></a:blipFill>"));
        Assert.Null(plain.FillTile);
        Assert.Equal(1, plain.FillImageAlpha);
        // A tile smaller than a point is not worth drawing: the picture is stretched, which is noted.
        var (tiny, warnings) = Read(OneCellAnchor(1, 1, 100, 50, Sp("<a:prstGeom prst=\"rect\"/><a:blipFill><a:blip r:embed=\"rId1\"/><a:tile sx=\"10000\" sy=\"10000\"/></a:blipFill>")));
        Assert.Null(Assert.Single(tiny.Shapes).FillTile);
        Assert.Contains("tiled picture fills of this picture are stretched", Assert.Single(warnings.Warnings, w => w.Code == "xlsx.drawing-approximated").Message);
    }

    [Theory]
    [InlineData("pentagon", ShapeGeometry.Pentagon)]
    [InlineData("octagon", ShapeGeometry.Octagon)]
    [InlineData("plus", ShapeGeometry.Plus)]
    [InlineData("chevron", ShapeGeometry.Chevron)]
    [InlineData("homePlate", ShapeGeometry.HomePlate)]
    [InlineData("star5", ShapeGeometry.Star5)]
    [InlineData("flowChartPreparation", ShapeGeometry.Hexagon)]
    [InlineData("flowChartInputOutput", ShapeGeometry.Parallelogram)]
    [InlineData("flowChartAlternateProcess", ShapeGeometry.RoundRectangle)]
    [InlineData("flowChartTerminator", ShapeGeometry.RoundRectangle)]
    public void More_presets_have_their_own_geometry(string preset, ShapeGeometry expected)
    {
        var (sheet, warnings) = Read(OneCellAnchor(1, 1, 100, 50, Sp($"<a:prstGeom prst=\"{preset}\"/>" + SolidRed)));
        Assert.Equal(expected, Assert.Single(sheet.Shapes).Geometry);
        Assert.DoesNotContain(warnings.Warnings, w => w.Code == "xlsx.drawing");
    }

    [Fact]
    public void Adjust_values_of_the_presets_are_read_as_fractions()
    {
        string Geometry(string preset, string adjust) => $"<a:prstGeom prst=\"{preset}\"><a:avLst>{adjust}</a:avLst></a:prstGeom>" + SolidRed;
        Assert.Equal(0.5, ReadOne(Sp(Geometry("flowChartTerminator", ""))).CornerRadius);
        Assert.Equal(0.3, ReadOne(Sp(Geometry("roundRect", "<a:gd name=\"adj\" fmla=\"val 30000\"/>"))).CornerRadius);
        Assert.Equal(0.5, ReadOne(Sp(Geometry("roundRect", "<a:gd name=\"adj\" fmla=\"val 90000\"/>"))).CornerRadius); // limited to a half
        Assert.Equal(1.0 / 6, ReadOne(Sp(Geometry("roundRect", ""))).CornerRadius);
        var arrow = ReadOne(Sp(Geometry("rightArrow", "<a:gd name=\"adj1\" fmla=\"val 30000\"/><a:gd name=\"adj2\" fmla=\"val 80000\"/>")));
        Assert.Equal((0.3, 0.8), (arrow.Adjust1!.Value, arrow.Adjust2!.Value));
        var chevron = ReadOne(Sp(Geometry("chevron", "<a:gd name=\"adj\" fmla=\"val 20000\"/>")));
        Assert.Equal((0.2, null), (chevron.Adjust1, chevron.Adjust2));
        Assert.Null(ReadOne(Sp(Geometry("octagon", ""))).Adjust1);
    }

    [Fact]
    public void A_reflection_is_read_from_the_effects_and_a_skewed_one_is_noted()
    {
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + "<a:effectLst><a:reflection blurRad=\"0\" stA=\"52000\" endA=\"300\" endPos=\"55000\" dist=\"127000\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/></a:effectLst>"));
        Assert.Equal(new ShapeReflection(0.52, 0.003, 0.55, 10), shape.Reflection);
        var plain = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + "<a:effectLst><a:reflection/></a:effectLst>"));
        Assert.Equal(new ShapeReflection(1, 0, 1, 0), plain.Reflection);
        var (_, warnings) = Read(OneCellAnchor(1, 1, 100, 50, Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + "<a:effectLst><a:reflection blurRad=\"63500\" sy=\"-50000\"/></a:effectLst>")));
        Assert.Contains("reflections are drawn straight below", Assert.Single(warnings.Warnings, w => w.Code == "xlsx.drawing-approximated").Message);
    }

    [Fact]
    public void Glow_and_soft_edges_are_read_from_the_effects()
    {
        var shape = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + "<a:effectLst><a:glow rad=\"127000\"><a:srgbClr val=\"00B0F0\"><a:alpha val=\"40000\"/></a:srgbClr></a:glow><a:softEdge rad=\"63500\"/></a:effectLst>"));
        Assert.Equal(new ShapeGlow(new Color(0, 0xB0, 0xF0), 0.4, 10), shape.Glow);
        Assert.Equal(5, shape.SoftEdge);
        var none = ReadOne(Sp("<a:prstGeom prst=\"rect\"/>" + SolidRed + "<a:effectLst><a:glow rad=\"0\"><a:srgbClr val=\"00B0F0\"/></a:glow></a:effectLst>"));
        Assert.Null(none.Glow);
        Assert.Equal(0, none.SoftEdge);
    }

    [Fact]
    public void A_rotated_group_turns_its_children_and_remembers_the_turn()
    {
        string Child(int x) => Sp($"<a:xfrm rot=\"1800000\"><a:off x=\"{x}\" y=\"0\"/><a:ext cx=\"1000\" cy=\"2000\"/></a:xfrm><a:prstGeom prst=\"rect\"/>" + SolidRed);
        string Group(string rot, string children) => "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"9\" name=\"G\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>" +
            $"<xdr:grpSpPr><a:xfrm rot=\"{rot}\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"4000\" cy=\"2000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"4000\" cy=\"2000\"/></a:xfrm></xdr:grpSpPr>{children}</xdr:grpSp>";
        var (sheet, warnings) = Read(OneCellAnchor(1, 1, 200, 100, Group("5400000", Child(0) + Child(3000))));
        Assert.Empty(warnings.Warnings);                  // no longer reported as drawn unrotated
        Assert.Equal(2, sheet.Shapes.Count);
        // The file turns clockwise; the model counter-clockwise: the child's -30° and the group's -90°.
        Assert.All(sheet.Shapes, s => Assert.Equal(-120, s.Rotation));
        var turn = Assert.Single(sheet.Shapes[0].GroupTurns);
        Assert.Equal((-90.0, new FrameFraction(0, 0, 1, 1)), (turn.Rotation, turn.Frame));
        // A group in a group adds its turn, outermost first; a group that is not turned leaves no trace.
        var inner = Group("2700000", Child(0));
        var (nested, _) = Read(OneCellAnchor(1, 1, 200, 100, Group("5400000", inner)));
        var shape = Assert.Single(nested.Shapes);
        Assert.Equal([-90.0, -45.0], shape.GroupTurns.Select(t => t.Rotation));
        Assert.Equal(-90 - 45 - 30, shape.Rotation);
        var (flat, _) = Read(OneCellAnchor(1, 1, 200, 100, Group("0", Child(0))));
        Assert.Empty(flat.Shapes[0].GroupTurns);
        Assert.Equal(-30, flat.Shapes[0].Rotation);
    }
}
