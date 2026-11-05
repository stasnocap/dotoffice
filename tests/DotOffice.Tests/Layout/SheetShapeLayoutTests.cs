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

    [Fact]
    public void A_pattern_fill_draws_the_background_and_clipped_hatch_lines()
    {
        var pattern = new ShapePattern(HatchKind.Horizontal, Blue, Red, Spacing: 10, Weight: 1);
        var page = Page(Shape() with { FillPattern = pattern, Fill = Blue });
        Assert.Equal(Red, page.Items.OfType<PolygonItem>().Single().Fill);
        var lines = page.Items.OfType<LineItem>().Where(l => l.Color == Blue).ToList();
        Assert.Equal(4, lines.Count); // frame 100 x 40: lines every 10 pt, anchored at the top edge
        Assert.All(lines, l => Assert.Equal(l.Y1, l.Y2, 6));
        Assert.All(lines, l => Assert.Equal(100, Math.Abs(l.X2 - l.X1), 3));
        Assert.Equal(new[] { 0.0, 10, 20, 30 }, lines.Select(l => Math.Round(l.Y1 - 54 - 20)).Order());
    }

    [Fact]
    public void Diagonal_hatches_stay_inside_a_triangle_and_tints_are_one_color()
    {
        var page = Page(Shape(ShapeGeometry.Triangle) with { FillPattern = new ShapePattern(HatchKind.UpDiagonal, Blue, Red, 8, 1), Fill = Blue });
        var (left, top) = (50.4 + 50, 54 + 20);
        foreach (var l in page.Items.OfType<LineItem>().Where(l => l.Color == Blue))
        {
            Assert.InRange(l.X1, left - 0.01, left + 100.01);
            Assert.InRange(l.X2, left - 0.01, left + 100.01);
            Assert.InRange(l.Y1, top - 0.01, top + 40.01);
            Assert.True(l.Y2 < l.Y1 && l.X2 > l.X1); // running up to the right
        }
        Assert.Contains(page.Items, i => i is LineItem l && l.Color == Blue);
        var tint = Page(Shape() with { FillPattern = new ShapePattern(HatchKind.Tint, Color.Black, Color.White, Coverage: 0.5), Fill = Color.Black });
        Assert.Equal(new Color(128, 128, 128), tint.Items.OfType<PolygonItem>().Single().Fill);
        Assert.Empty(tint.Items.OfType<LineItem>());
    }

    [Fact]
    public void A_shadow_is_a_translucent_darker_copy_moved_by_its_offset()
    {
        var page = Page(Shape(fill: Red, line: Blue) with { Shadow = new ShapeShadow(Color.Black, 0.4, 3, 5) });
        var items = page.Items.ToList();
        int push = items.FindIndex(i => i is PushItem { Opacity: 0.4 });
        Assert.True(push > 0);
        var shadowPolygon = Assert.IsType<PolygonItem>(items[push + 1]);
        Assert.Equal(Color.Black, shadowPolygon.Fill);
        var body = items.OfType<PolygonItem>().Single(p => p.Fill == Red);
        Assert.Equal(body.Points[0].X + 3, shadowPolygon.Points[0].X, 6);
        Assert.Equal(body.Points[0].Y + 5, shadowPolygon.Points[0].Y, 6);
        Assert.True(items.IndexOf(body) > push); // the shadow lies underneath
        Assert.DoesNotContain(items.Skip(push).TakeWhile(i => i is not PopItem), i => i is LineItem l && l.Color == Blue);
        Assert.Contains(items.Skip(push).TakeWhile(i => i is not PopItem), i => i is LineItem l && l.Color == Color.Black);
    }

    [Fact]
    public void A_blurred_shadow_is_nine_faint_copies_that_add_up_to_its_opacity()
    {
        var page = Page(Shape(fill: Red) with { Shadow = new ShapeShadow(Color.Black, 0.5, 4, 4, Blur: 6) });
        var copies = page.Items.OfType<PushItem>().Where(p => p.Opacity < 1).ToList();
        Assert.Equal(9, copies.Count);
        double each = copies[0].Opacity;
        Assert.Equal(0.5, 1 - Math.Pow(1 - each, 9), 6);
        Assert.All(copies, c => Assert.Equal(each, c.Opacity, 9));
        // The copies are spread over the blur radius (3 points either side) around the shadow's place.
        var xs = copies.Select(c => c.CenterX).Distinct().Order().ToList();
        Assert.Equal(3, xs.Count);
        Assert.Equal(3 * 2, xs[^1] - xs[0], 6);
        Assert.Single(Page(Shape(fill: Red) with { Shadow = new ShapeShadow(Color.Black, 0.5, 4, 4) }).Items.OfType<PushItem>(), p => p.Opacity < 1);
    }

    [Fact]
    public void A_picture_fill_is_the_image_over_the_frame_cut_at_the_outline()
    {
        var picture = new ImageData(PngTestEncoder.Encode(1, 1, 2, 8, [1, 2, 3]), "fill");
        var rectangle = Page(Shape(fill: Red) with { FillImage = picture }).Items.OfType<ImageItem>().Single();
        Assert.Null(rectangle.Clip);
        var plain = Page(Shape(fill: Red)).Items.OfType<PolygonItem>().Single().Points;
        Assert.Equal((plain.Min(p => p.X), plain.Min(p => p.Y)), (rectangle.X, rectangle.Y));
        var diamond = Page(Shape(ShapeGeometry.Diamond, Red) with { FillImage = picture }).Items.OfType<ImageItem>().Single();
        Assert.Equal(4, diamond.Clip!.Count);
        Assert.DoesNotContain(Page(Shape(fill: Red) with { FillImage = picture }).Items.OfType<PolygonItem>(), p => p.Fill == Red);
        // A gradient wins over a picture; a plain shape has no image.
        Assert.Empty(Page(Shape(fill: Red) with { FillImage = picture, FillGradient = new GradientFill { Kind = GradientKind.Linear, Stops = RedBlue } }).Items.OfType<ImageItem>());
    }

    [Fact]
    public void Rectangular_path_gradients_are_marked_for_concentric_rectangles()
    {
        GradientItem Draw(bool rectangular) => Page(Shape(ShapeGeometry.Ellipse, Red) with
        {
            FillGradient = new GradientFill { Kind = GradientKind.Path, Rectangular = rectangular, Left = 0.5, Right = 0.5, Top = 0.5, Bottom = 0.5, Stops = RedBlue },
        }).Items.OfType<GradientItem>().Single();
        Assert.True(Draw(true).Rectangular);
        Assert.False(Draw(false).Rectangular);
        Assert.False(Page(Shape(fill: Red) with { FillGradient = new GradientFill { Kind = GradientKind.Linear, Stops = RedBlue } }).Items.OfType<GradientItem>().Single().Rectangular);
    }

    [Fact]
    public void A_picture_fill_clip_scales_with_the_sheet()
    {
        var picture = new ImageData(PngTestEncoder.Encode(1, 1, 2, 8, [1, 2, 3]), "fill");
        var item = Page(Shape(ShapeGeometry.Ellipse, Red) with { FillImage = picture }, sheet => sheet.PageSetup.Scale = 0.5).Items.OfType<ImageItem>().Single();
        Assert.Equal(100 * 0.5, item.Width, 6);
        // The clip polygon lies on the (scaled) image frame, not on the unscaled one.
        Assert.InRange(item.Clip!.Min(p => p.X), item.X - 1e-6, item.X + 1);
        Assert.InRange(item.Clip!.Max(p => p.X), item.X + item.Width - 1, item.X + item.Width + 1e-6);
        Assert.InRange(item.Clip!.Max(p => p.Y), item.Y + item.Height - 1, item.Y + item.Height + 1e-6);
    }

    [Fact]
    public void A_tiled_picture_fill_repeats_the_picture_from_the_alignment_corner_and_cuts_each_tile_at_the_outline()
    {
        var picture = new ImageData(PngTestEncoder.Encode(1, 1, 2, 8, [1, 2, 3]), "fill");
        // The frame is 100 x 40 points; 30 x 20 point tiles from the top-left corner: 4 columns by 2 rows.
        var tiles = Page(Shape(fill: Red) with { FillImage = picture, FillTile = new ShapeTile(30, 20, 0, 0) }).Items.OfType<ImageItem>().ToList();
        Assert.Equal(8, tiles.Count);
        var (left, top) = Origin;
        Assert.Equal((left + 50, top + 20), (tiles.Min(t => t.X), tiles.Min(t => t.Y)));
        Assert.All(tiles, t => Assert.Equal((30.0, 20.0), (t.Width, t.Height)));
        Assert.All(tiles, t => Assert.Equal(4, t.Clip!.Count));
        // Mirrored tiles: every other tile across x, y or both is flipped, with a clip that still follows the outline once flipped back.
        List<LayoutItem> Mirrored(string flip) => Page(Shape(fill: Red) with { FillImage = picture, FillTile = new ShapeTile(30, 20, 0, 0, "tl", flip) }).Items;
        Assert.DoesNotContain(Mirrored("none").OfType<PushItem>(), p => p.FlipH || p.FlipV);
        Assert.Equal(4, Mirrored("x").OfType<PushItem>().Count(p => p.FlipH && !p.FlipV));   // columns 1 and 3 of both rows
        Assert.Equal(4, Mirrored("y").OfType<PushItem>().Count(p => p.FlipV && !p.FlipH));   // the second row
        Assert.Equal(6, Mirrored("xy").OfType<PushItem>().Count(p => p.FlipH || p.FlipV));   // every tile of an odd row or column
        var flippedClips = Mirrored("x").OfType<ImageItem>().Where(t => t.Clip is not null).Select(t => t.Clip!.Min(p => p.X)).Distinct().OrderBy(v => v).ToList();
        // The frame starts at left + 50 and is 100 wide: tile column 1 is centered 45 and column 3 105 points in, so its outline is mirrored about those.
        Assert.Contains(flippedClips, v => Math.Abs(v - (left + 40)) < 1e-6);
        Assert.Contains(flippedClips, v => Math.Abs(v - (left + 160)) < 1e-6);
        // From the bottom-right corner (and shifted 5 points right) the grid lines up with that corner instead.
        var right = Page(Shape(fill: Red) with { FillImage = picture, FillTile = new ShapeTile(30, 20, 5, 0, "br") }).Items.OfType<ImageItem>().ToList();
        Assert.Contains(right, t => Math.Abs(t.X + t.Width - (left + 155)) < 1e-6 && Math.Abs(t.Y + t.Height - (top + 60)) < 1e-6);
        // Far too many tiles give up and stretch the picture over the frame.
        var many = Page(Shape(fill: Red) with { FillImage = picture, FillTile = new ShapeTile(1, 1, 0, 0) }).Items.OfType<ImageItem>().ToList();
        Assert.Single(many);
        Assert.Equal(100, many[0].Width);
    }

    [Fact]
    public void A_picture_fill_with_an_alpha_is_painted_translucent()
    {
        var picture = new ImageData(PngTestEncoder.Encode(1, 1, 2, 8, [1, 2, 3]), "fill");
        var items = Page(Shape(fill: Red) with { FillImage = picture, FillImageAlpha = 0.4 }).Items;
        int at = items.ToList().FindIndex(i => i is PushItem { Opacity: 0.4 });
        Assert.True(at >= 0);
        Assert.IsType<ImageItem>(items[at + 1]);
        Assert.IsType<PopItem>(items[at + 2]);
    }

    [Fact]
    public void A_glow_is_four_translucent_strokes_under_the_shape_that_reach_its_radius()
    {
        var glowColor = new Color(0, 176, 240);
        var page = Page(Shape(fill: Red) with { Glow = new ShapeGlow(glowColor, 0.4, 8) });
        var layers = page.Items.OfType<PushItem>().Where(p => p.Opacity < 1).ToList();
        Assert.Equal(4, layers.Count);
        Assert.Equal(0.4, 1 - Math.Pow(1 - layers[0].Opacity, 4), 6);
        var strokes = page.Items.OfType<LineItem>().Where(l => l.Color == glowColor).ToList();
        Assert.Equal([16.0, 12.0, 8.0, 4.0], strokes.Select(l => l.Width).Distinct());   // widest first, up to twice the radius (half of it lies outside)
        // The glow is drawn before the fill so it lies underneath.
        var items = page.Items.ToList();
        Assert.True(items.FindIndex(i => i is LineItem l && l.Color == glowColor) < items.FindIndex(i => i is PolygonItem p && p.Fill == Red));
    }

    [Fact]
    public void Soft_edges_draw_shrinking_copies_that_add_up_to_full_opacity_inside()
    {
        var page = Page(Shape(fill: Red, line: Blue) with { SoftEdge = 6 });
        var fills = page.Items.OfType<PolygonItem>().Where(p => p.Fill == Red).ToList();
        Assert.Equal(6, fills.Count);
        Assert.Equal(fills.Select(f => f.Points.Max(p => p.X) - f.Points.Min(p => p.X)).OrderDescending(), fills.Select(f => f.Points.Max(p => p.X) - f.Points.Min(p => p.X)));
        Assert.True(fills[0].Points.Max(p => p.X) - fills[0].Points.Min(p => p.X) > fills[5].Points.Max(p => p.X) - fills[5].Points.Min(p => p.X) + 9);
        var opacities = page.Items.OfType<PushItem>().Where(p => p.Clip is null).Select(p => p.Opacity).ToList();
        double visible = 1;
        foreach (var o in opacities) visible *= 1 - o;
        Assert.Equal(0, visible, 9);                      // the innermost copy is opaque, so the middle is solid
        Assert.Empty(page.Items.OfType<LineItem>());      // the outline fades with the edges
    }

    [Fact]
    public void Soft_edges_of_a_notched_outline_move_every_edge_inwards_and_the_reflex_corner_out_of_the_way()
    {
        // An L: a bar along the top (the frame is 100 x 40) with a leg on the left. Fractions of the frame: the notch's corner is at (0.4, 0.4).
        var l = Shape(ShapeGeometry.Custom, Red) with
        {
            Paths = [new ShapePath([(0, 0), (1, 0), (1, 0.4), (0.4, 0.4), (0.4, 1), (0, 1)], true)], SoftEdge = 6,
        };
        var fills = Page(l).Items.OfType<PolygonItem>().Where(p => p.Fill == Red).ToList();
        Assert.Equal(6, fills.Count);
        var (left, top) = Origin;
        var (x0, y0) = (left + 50, top + 20);
        // The last copy lies 5 points inside the outline: the reflex corner (40, 16) moves to (35, 11) and the others by 5 along both axes.
        (double X, double Y)[] expected = [(5, 5), (95, 5), (95, 11), (35, 11), (35, 35), (5, 35)];
        Assert.Equal(expected.Length, fills[^1].Points.Count);
        foreach (var (ex, ey) in expected)
            Assert.Contains(fills[^1].Points, q => Math.Abs(q.X - (x0 + ex)) < 1e-6 && Math.Abs(q.Y - (y0 + ey)) < 1e-6);
    }

    [Fact]
    public void Inset_polygons_give_way_when_a_part_is_too_thin_or_the_polygon_degenerates()
    {
        (double, double)[] strip = [(0, 0), (100, 0), (100, 6), (0, 6)];
        Assert.NotNull(SheetLayoutEngine.InsetPolygon(strip, 2));
        Assert.Null(SheetLayoutEngine.InsetPolygon(strip, 4));                      // the strip is only 6 thick: 2 x 4 turns it inside out
        Assert.Null(SheetLayoutEngine.InsetPolygon([(0, 0), (10, 0)], 1));          // not a polygon
        Assert.Null(SheetLayoutEngine.InsetPolygon([(0, 0), (5, 0), (10, 0)], 1));  // no area
        // Either winding gives the same inset.
        var clockwise = SheetLayoutEngine.InsetPolygon([(0, 0), (0, 10), (10, 10), (10, 0)], 2)!;
        var counter = SheetLayoutEngine.InsetPolygon([(0, 0), (10, 0), (10, 10), (0, 10)], 2)!;
        Assert.Equal(counter.OrderBy(q => q.X).ThenBy(q => q.Y), clockwise.OrderBy(q => q.X).ThenBy(q => q.Y));
        Assert.Equal((2, 2), counter.Min(q => q.X) == 2 ? (counter.Min(q => q.X), counter.Min(q => q.Y)) : (0, 0));
    }

    [Fact]
    public void A_reflection_is_the_shape_mirrored_under_its_bottom_edge_in_nested_layers_that_add_up_to_the_fade()
    {
        var page = Page(Shape(fill: Red) with { Reflection = new ShapeReflection(0.5, 0, 1, 6), Paragraphs = [] });
        var mirrors = page.Items.OfType<PushItem>().Where(p => p.FlipV && p.Clip is null).ToList();
        Assert.Equal(16, mirrors.Count);
        var (left, top) = Origin;
        // The frame's bottom is at top + 60; a distance of 6 puts the mirror line 3 points below it.
        Assert.All(mirrors, m => Assert.Equal(top + 60 + 3, m.CenterY, 6));
        // Layer k reaches from the reflection's top (6 points under the shape) to the end of the k-th band of the shape's height (40).
        var clips = page.Items.OfType<PushItem>().Where(p => p.Clip is { } c && c.Y > top + 60).Select(p => p.Clip!.Value).ToList();
        Assert.Equal(16, clips.Count);
        Assert.All(clips, c => Assert.Equal(top + 66, c.Y, 6));
        for (int k = 0; k < 16; k++) Assert.Equal(40.0 * (k + 1) / 16, clips[k].Height, 6);
        // The layers that reach band j (index j or later) together have the opacity of the fade there: 50 % down to 0 over the height.
        for (int j = 0; j < 16; j++)
        {
            double clear = 1;
            for (int k = j; k < 16; k++) clear *= 1 - mirrors[k].Opacity;
            Assert.Equal(0.5 + (0 - 0.5) * (j + 0.5) / 16, 1 - clear, 6);
        }
        // The reflection comes before the shape itself, so that the shape covers a reflection that touches it.
        var items = page.Items.ToList();
        Assert.True(items.FindIndex(i => i is PushItem { FlipV: true }) < items.FindLastIndex(i => i is PolygonItem));
        // A reflection that has faded out by half the height covers only that part; one that starts transparent is not drawn at all.
        var half = Page(Shape(fill: Red) with { Reflection = new ShapeReflection(0.5, 0, 0.5, 0) }).Items.OfType<PushItem>().Where(p => p.Clip is { } c && c.Y >= top + 60 - 1e-9).ToList();
        Assert.Equal(20, half.Max(p => p.Clip!.Value.Height), 6);
        Assert.DoesNotContain(Page(Shape(fill: Red) with { Reflection = new ShapeReflection(0, 0, 1, 0) }).Items.OfType<PushItem>(), p => p.FlipV);
    }

    [Fact]
    public void The_polygon_presets_fill_their_frame_and_follow_their_adjust_values()
    {
        // The default frame is 100 x 40 points from (left + 50, top + 20).
        var (left, top) = Origin;
        List<(double X, double Y)> Points(ShapeGeometry geometry, double? a1 = null, double? a2 = null) =>
            Page(Shape(geometry, Red) with { Adjust1 = a1, Adjust2 = a2 }).Items.OfType<PolygonItem>().Single().Points.Select(p => (Math.Round(p.X - left - 50, 4), Math.Round(p.Y - top - 20, 4))).ToList();
        var pentagon = Points(ShapeGeometry.Pentagon);
        Assert.Equal(5, pentagon.Count);
        Assert.Equal((0.0, 100.0), (pentagon.Min(p => p.X), pentagon.Max(p => p.X)));
        Assert.Equal((0.0, 40.0), (pentagon.Min(p => p.Y), pentagon.Max(p => p.Y)));
        var octagon = Points(ShapeGeometry.Octagon);                                // the corner cut is 0.29289 of the shorter side (40)
        Assert.Contains((11.7156, 0.0), octagon);
        Assert.Contains((0.0, 11.7156), octagon);
        Assert.Contains((20.0, 0.0), Points(ShapeGeometry.Octagon, 0.5));          // limited to half the shorter side
        Assert.Equal(12, Points(ShapeGeometry.Plus).Count);
        Assert.Contains((10.0, 10.0), Points(ShapeGeometry.Plus));                  // an arm a quarter of the shorter side wide
        Assert.Contains((20.0, 20.0), Points(ShapeGeometry.Chevron, 0.5));          // the notch: half the shorter side from the left, in the middle
        Assert.Contains((100.0, 20.0), Points(ShapeGeometry.Chevron, 0.5));         // the point
        Assert.Contains((80.0, 0.0), Points(ShapeGeometry.HomePlate, 0.5));
        var star = Points(ShapeGeometry.Star5);
        Assert.Equal(10, star.Count);
        Assert.Equal((0.0, 100.0), (star.Min(p => p.X), star.Max(p => p.X)));
        // A block arrow's shaft and head follow adj1 and adj2: a shaft 30 % of the height, a head 80 % of the shorter side (32) long.
        var arrow = Points(ShapeGeometry.RightArrow, 0.3, 0.8);
        Assert.Contains((68.0, 14.0), arrow);
        Assert.Contains((0.0, 26.0), arrow);
        Assert.Contains((100.0, 20.0), arrow);
    }

    [Fact]
    public void A_rotated_group_swings_its_children_about_the_groups_center()
    {
        // The group frame is the whole 100 x 40 shape frame; the child is its left half (center 25 left of the group's center).
        var child = Shape(fill: Red) with { Frame = new FrameFraction(0, 0, 0.5, 1), Rotation = 90, GroupTurns = [new GroupTurn(FrameFraction.Whole, 90)] };
        var polygon = Page(child).Items.OfType<PolygonItem>().Single();
        var (left, top) = Origin;
        // The child's own frame is 50 x 40 around (left + 75, top + 40); a quarter turn counter-clockwise about (left + 100, top + 40) puts its center at (left + 100, top + 65).
        double cx = (polygon.Points.Min(p => p.X) + polygon.Points.Max(p => p.X)) / 2, cy = (polygon.Points.Min(p => p.Y) + polygon.Points.Max(p => p.Y)) / 2;
        Assert.Equal((left + 100, top + 65), (Math.Round(cx, 6), Math.Round(cy, 6)));
    }
}
