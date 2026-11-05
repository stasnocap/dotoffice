using DotOffice.Core;
using DotOffice.Layout;

namespace DotOffice.Tests.Layout;

public class SheetShapeLayoutTests
{
    private static readonly Color Red = new(255, 0, 0), Blue = new(0, 0, 255);

    private static SheetShape Shape(ShapeGeometry geometry = ShapeGeometry.Rectangle, Color? fill = null, Color? line = null) => new()
    {
        Anchor = ImageAnchorKind.TwoCell, From = new AnchorPoint(1, 1, 0, 0), To = new AnchorPoint(3, 3, 0, 0),
        Geometry = geometry, Fill = fill, LineColor = line, LineWidth = 2,
    };

    private static LayoutPage Page(SheetShape shape, Action<Worksheet>? configure = null)
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text("a")), (9, 9, SheetLayoutTests.Text("z")));
        sheet.DefaultColumnWidth = 50;
        sheet.DefaultRowHeight = 20;
        sheet.Shapes.Add(shape);
        configure?.Invoke(sheet);
        return SheetLayoutTests.Layout(sheet).Pages[0];
    }

    private static (double Left, double Top) Origin => (50.4, 54);

    [Fact]
    public void A_filled_rectangle_is_a_polygon_over_its_frame_with_four_outline_lines()
    {
        var page = Page(Shape(fill: Red, line: Blue));
        var polygon = page.Items.OfType<PolygonItem>().Single();
        Assert.Equal(Red, polygon.Fill);
        var (left, top) = Origin;
        Assert.Contains((left + 50, top + 20), polygon.Points.Select(p => (Math.Round(p.X, 6), Math.Round(p.Y, 6))));
        Assert.Contains((left + 150, top + 60), polygon.Points.Select(p => (Math.Round(p.X, 6), Math.Round(p.Y, 6))));
        var lines = page.Items.OfType<LineItem>().ToList();
        Assert.Equal(4, lines.Count);
        Assert.All(lines, l => Assert.Equal((2.0, Blue), (l.Width, l.Color)));
    }

    [Fact]
    public void Shapes_without_fill_or_line_draw_nothing_but_their_text()
    {
        var shape = Shape() with { Paragraphs = [new ShapeParagraph(HorizontalAlignment.Left, [new CellTextRun("caption", RunProperties.Default with { FontFamily = "Calibri", FontSize = 10 })], 10)] };
        var page = Page(shape);
        Assert.Empty(page.Items.OfType<PolygonItem>());
        Assert.Contains("caption", string.Concat(page.Items.OfType<TextItem>().Select(t => t.Text)));
    }

    [Fact]
    public void An_ellipse_is_an_ellipse_item()
    {
        var page = Page(Shape(ShapeGeometry.Ellipse, Red));
        var ellipse = page.Items.OfType<EllipseItem>().Single();
        Assert.Equal((100.0, 40.0), (ellipse.Width, ellipse.Height));
    }

    [Theory]
    [InlineData(ShapeGeometry.Triangle, 3)]
    [InlineData(ShapeGeometry.Diamond, 4)]
    [InlineData(ShapeGeometry.RightArrow, 7)]
    [InlineData(ShapeGeometry.Hexagon, 6)]
    public void Polygon_geometries_have_their_corners(ShapeGeometry geometry, int corners)
    {
        Assert.Equal(corners, Page(Shape(geometry, Red)).Items.OfType<PolygonItem>().Single().Points.Count);
    }

    [Fact]
    public void A_line_is_one_line_item_along_the_diagonal_and_flips_choose_the_other_one()
    {
        var page = Page(Shape(ShapeGeometry.Line, line: Blue));
        var line = page.Items.OfType<LineItem>().Single();
        Assert.Equal((50 + Origin.Left, 20 + Origin.Top, 150 + Origin.Left, 60 + Origin.Top), (line.X1, line.Y1, line.X2, line.Y2));
        var flipped = Page(Shape(ShapeGeometry.Line, line: Blue) with { FlipV = true });
        Assert.Contains(flipped.Items.OfType<PushItem>(), p => p.FlipV && !p.FlipH);
    }

    [Fact]
    public void A_horizontal_line_of_zero_height_is_still_drawn()
    {
        var shape = Shape(ShapeGeometry.Line, line: Blue) with { To = new AnchorPoint(3, 1, 0, 0) };
        Assert.Single(Page(shape).Items.OfType<LineItem>());
    }

    [Fact]
    public void Rotation_turns_the_shape_about_its_center()
    {
        var page = Page(Shape(fill: Red) with { Rotation = 30 });
        var turn = page.Items.OfType<PushItem>().Single(p => p.Rotation != 0);
        Assert.Equal((30.0, Origin.Left + 100, Origin.Top + 40), (turn.Rotation, turn.CenterX, turn.CenterY));
    }

    [Fact]
    public void Text_is_drawn_after_the_outline_and_outside_the_mirroring_group()
    {
        var shape = Shape(fill: Red) with
        {
            FlipH = true,
            Paragraphs = [new ShapeParagraph(HorizontalAlignment.Center, [new CellTextRun("mid", RunProperties.Default with { FontFamily = "Calibri", FontSize = 10 })], 10)],
            TextAnchor = ShapeTextAnchor.Middle,
        };
        var items = Page(shape).Items;
        int polygon = items.ToList().FindIndex(i => i is PolygonItem);
        int text = items.ToList().FindIndex(i => i is TextItem t && t.Text.Contains("mid"));
        Assert.True(polygon >= 0 && text > polygon);
        // The flip group is closed before the text starts.
        int flipPush = items.ToList().FindIndex(i => i is PushItem { FlipH: true });
        int depth = 0;
        for (int i = flipPush; i < text; i++)
        {
            if (items[i] is PushItem) depth++;
            else if (items[i] is PopItem) depth--;
        }
        Assert.Equal(0, depth);
        var label = items.OfType<TextItem>().First(t => t.Text.Contains("mid"));
        // Centered: the text sits around the middle of the 100 pt wide frame.
        Assert.InRange(label.X, Origin.Left + 50 + 30, Origin.Left + 50 + 70);
    }

    [Fact]
    public void Shapes_extend_the_used_range_and_follow_the_page_scale()
    {
        var sheet = new Worksheet("S");
        sheet.Shapes.Add(Shape(fill: Red));
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        Assert.Single(page.Items.OfType<PolygonItem>());

        var half = Page(Shape(fill: Red), s => s.PageSetup.Scale = 0.5).Items.OfType<PolygonItem>().Single();
        var xs = half.Points.Select(p => p.X);
        Assert.Equal(50, xs.Max() - xs.Min(), 6);
    }

    [Fact]
    public void Pictures_and_shapes_are_drawn_in_z_order()
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text("a")));
        sheet.DefaultColumnWidth = 50;
        sheet.DefaultRowHeight = 20;
        var png = new ImageData(PngTestEncoder.Encode(2, 2, 2, 8, [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]), "t.png");
        sheet.Shapes.Add(Shape(fill: Red) with { ZOrder = 2 });
        sheet.Images.Add(new SheetImage { Image = png, Anchor = ImageAnchorKind.TwoCell, From = new AnchorPoint(1, 1, 0, 0), To = new AnchorPoint(3, 3, 0, 0), ZOrder = 1 });
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.ToList();
        Assert.True(items.FindIndex(i => i is ImageItem) < items.FindIndex(i => i is PolygonItem));
    }

    [Fact]
    public void A_group_child_covers_its_part_of_the_group_frame()
    {
        var page = Page(Shape(fill: Red) with { Frame = new FrameFraction(0.5, 0, 0.5, 0.5) });
        var xs = page.Items.OfType<PolygonItem>().Single().Points;
        Assert.Equal((Origin.Left + 100, Origin.Left + 150), (xs.Min(p => p.X), xs.Max(p => p.X)));
        Assert.Equal((Origin.Top + 20, Origin.Top + 40), (xs.Min(p => p.Y), xs.Max(p => p.Y)));
    }

    private static readonly IReadOnlyList<GradientStop> RedBlue = [new GradientStop(0, Red), new GradientStop(1, Blue)];

    [Fact]
    public void A_gradient_fill_is_one_gradient_clipped_to_the_outline()
    {
        var fill = new GradientFill { Kind = GradientKind.Linear, Degree = 90, Stops = RedBlue };
        var rectangle = Page(Shape(fill: Red) with { FillGradient = fill }).Items.OfType<GradientItem>().Single();
        Assert.Null(rectangle.Clip);
        Assert.Equal(90, rectangle.Degree);
        var diamond = Page(Shape(ShapeGeometry.Diamond, Red) with { FillGradient = fill }).Items.OfType<GradientItem>().Single();
        Assert.Equal(4, diamond.Clip!.Count);
        var path = Page(Shape(ShapeGeometry.Ellipse, Red) with { FillGradient = new GradientFill { Kind = GradientKind.Path, Left = 0.5, Right = 0.5, Top = 0.5, Bottom = 0.5, Stops = RedBlue } }).Items.OfType<GradientItem>().Single();
        Assert.True(path.Radial);
        Assert.Equal(0.5, path.FocusX);
        Assert.Empty(Page(Shape(fill: Red) with { FillGradient = fill }).Items.OfType<PolygonItem>());
    }

    [Fact]
    public void A_line_with_a_triangle_arrowhead_stops_at_its_base()
    {
        var shape = Shape(ShapeGeometry.Line, line: Blue) with { TailEnd = new LineEnd(LineEndKind.Triangle, 3, 3), LineWidth = 2 };
        var page = Page(shape);
        var head = page.Items.OfType<PolygonItem>().Single();
        Assert.Equal(Blue, head.Fill);
        var line = page.Items.OfType<LineItem>().Single();
        var (left, top) = Origin;
        var tip = (X: left + 150, Y: top + 60);
        Assert.Contains(head.Points, p => Math.Abs(p.X - tip.X) < 1e-6 && Math.Abs(p.Y - tip.Y) < 1e-6);
        double gap = Math.Sqrt(Math.Pow(line.X2 - tip.X, 2) + Math.Pow(line.Y2 - tip.Y, 2));
        Assert.InRange(gap, 5, 6.5); // 90 % of the 6 pt head length
        Assert.Equal((left + 50, top + 20), (Math.Round(line.X1, 6), Math.Round(line.Y1, 6))); // no head at the start
    }

    [Fact]
    public void Every_arrowhead_kind_draws_something_at_the_end()
    {
        foreach (var kind in new[] { LineEndKind.Triangle, LineEndKind.Stealth, LineEndKind.Arrow, LineEndKind.Oval, LineEndKind.Diamond })
        {
            var page = Page(Shape(ShapeGeometry.Line, line: Blue) with { HeadEnd = new LineEnd(kind, 3, 3) });
            Assert.True(page.Items.OfType<PolygonItem>().Any() || page.Items.OfType<EllipseItem>().Any() || page.Items.OfType<LineItem>().Count() > 1, kind.ToString());
        }
    }

    [Fact]
    public void Connector_paths_are_stroked_segment_by_segment_over_the_frame()
    {
        var shape = Shape(ShapeGeometry.Line, line: Blue) with { Paths = [new ShapePath([(0, 0), (0.5, 0), (0.5, 1), (1, 1)], false, Fill: false)] };
        var lines = Page(shape).Items.OfType<LineItem>().ToList();
        var (left, top) = Origin;
        Assert.Equal(3, lines.Count);
        Assert.Equal((left + 50, top + 20), (Math.Round(lines[0].X1, 6), Math.Round(lines[0].Y1, 6)));
        Assert.Equal((left + 100, top + 20), (Math.Round(lines[0].X2, 6), Math.Round(lines[0].Y2, 6)));
        Assert.Equal((left + 150, top + 60), (Math.Round(lines[2].X2, 6), Math.Round(lines[2].Y2, 6)));
    }

    [Fact]
    public void Custom_paths_are_filled_and_outlined()
    {
        var shape = Shape(ShapeGeometry.Custom, Red, Blue) with { Paths = [new ShapePath([(0, 0), (1, 0), (0.5, 1)], true)] };
        var page = Page(shape);
        Assert.Equal(3, page.Items.OfType<PolygonItem>().Single().Points.Count);
        Assert.Equal(3, page.Items.OfType<LineItem>().Count());
        var noFill = Page(shape with { Paths = [new ShapePath([(0, 0), (1, 0), (0.5, 1)], true, Fill: false)] });
        Assert.Empty(noFill.Items.OfType<PolygonItem>());
    }
}
