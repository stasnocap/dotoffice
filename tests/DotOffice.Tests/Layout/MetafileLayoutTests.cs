using System.Buffers.Binary;
using System.Text;
using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Layout;

namespace DotOffice.Tests.Layout;

public class MetafileLayoutTests
{
    /// <summary>Lays out a 400 × 300 slide whose background is the metafile.</summary>
    private static List<LayoutItem> Lay(byte[] metafile, ConversionWarningList? warnings = null)
    {
        var slide = new Slide(400, 300) { BackgroundImage = new ImageData(metafile, "test") };
        var presentation = new Presentation();
        presentation.Slides.Add(slide);
        var engine = new SheetLayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(LayoutTestHelpers.Noto)), Warnings = warnings });
        return engine.Layout(presentation).Pages[0].Items;
    }

    /// <summary>An EMF of 1000 × 750 device pixels at 10 pixels per millimetre, whose frame covers them all, with the given records.</summary>
    private static byte[] Emf(params byte[][] records)
    {
        var header = new byte[108];
        Put(header, 0, 1, 108, 0, 0, 999, 749, 0, 0, 10000, 7500);
        Encoding.ASCII.GetBytes(" EMF").CopyTo(header, 40);
        Put(header, 72, 1000, 750, 100, 75);
        var all = new List<byte[]> { header };
        all.AddRange(records);
        all.Add(Record(14, 0, 0, 0));
        return [.. all.SelectMany(r => r)];
    }

    private static void Put(byte[] d, int offset, params int[] values)
    {
        for (int k = 0; k < values.Length; k++) BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(offset + k * 4), values[k]);
    }

    private static byte[] Record(int type, params int[] values)
    {
        var r = new byte[8 + values.Length * 4];
        Put(r, 0, type, r.Length);
        Put(r, 8, values);
        return r;
    }

    [Fact]
    public void An_emf_rectangle_is_filled_with_the_selected_brush_and_fitted_to_the_picture()
    {
        var items = Lay(Emf(
            Record(39, 1, 0, 0x0000FF, 0),       // CreateBrushIndirect: solid red
            Record(37, 1),                       // SelectObject
            Record(37, unchecked((int)0x80000008)), // NULL_PEN
            Record(43, 0, 0, 500, 375)));        // Rectangle over the top-left quarter
        var path = Assert.Single(items.OfType<PathItem>());
        Assert.Equal(new Color(255, 0, 0), path.Fill);
        Assert.Null(path.Stroke);
        var points = path.Points.Where(p => p.Kind != PathPointKind.Close).ToList();
        Assert.Equal(0, points.Min(p => p.X), 3);
        Assert.Equal(200, points.Max(p => p.X), 3);
        Assert.Equal(150, points.Max(p => p.Y), 3);
        Assert.DoesNotContain(items, i => i is ImageItem);
    }

    [Fact]
    public void An_emf_clip_rectangle_clips_the_shapes_drawn_inside_it_until_the_dc_is_restored()
    {
        var items = Lay(Emf(
            Record(33),                        // SaveDC
            Record(30, 0, 0, 500, 375),        // IntersectClipRect: top-left quarter
            Record(43, 0, 0, 1000, 750),       // Rectangle, clipped
            Record(34, -1),                    // RestoreDC
            Record(43, 0, 0, 100, 100)));      // Rectangle, not clipped
        var paths = items.OfType<PathItem>().ToList();
        Assert.Equal(2, paths.Count);
        int first = items.IndexOf(paths[0]), second = items.IndexOf(paths[1]);
        var clip = Assert.IsType<PushItem>(items[first - 1]);
        Assert.NotNull(clip.ClipOutline);
        Assert.Equal(200, clip.ClipOutline!.Max(p => p.X), 3);
        Assert.Equal(150, clip.ClipOutline!.Max(p => p.Y), 3);
        Assert.IsType<PopItem>(items[first + 1]);
        Assert.Equal(first + 2, second);
        Assert.Equal(items.OfType<PushItem>().Count(), items.OfType<PopItem>().Count());
    }

    [Fact]
    public void Anisotropic_mapping_and_a_world_transform_place_emf_shapes()
    {
        var items = Lay(Emf(
            Record(17, 8),                 // MM_ANISOTROPIC
            Record(9, 100, 100),           // window 100 × 100
            Record(11, 1000, 750),         // onto the whole device
            Record(35, BitConverter.SingleToInt32Bits(1), 0, 0, BitConverter.SingleToInt32Bits(1),
                BitConverter.SingleToInt32Bits(50), 0), // world: 50 units to the right
            Record(42, 0, 0, 50, 100)));   // Ellipse
        var path = Assert.Single(items.OfType<PathItem>());
        var xs = path.Points.Where(p => p.Kind != PathPointKind.Close).Select(p => p.X).ToList();
        Assert.Equal(200, xs.Min(), 1);
        Assert.Equal(400, xs.Max(), 1);
        Assert.Equal(Color.White, path.Fill);
        Assert.Equal(Color.Black, path.Stroke);
    }

    [Fact]
    public void Emf_text_is_drawn_with_the_selected_font_and_color()
    {
        var font = new byte[368];
        Put(font, 0, 82, font.Length, 1, -60);
        Encoding.Unicode.GetBytes("Noto Sans").CopyTo(font, 40);
        var text = new byte[76 + 8];
        Put(text, 0, 84, text.Length);
        Put(text, 36, 100, 200, 2, 76, 0);
        Encoding.Unicode.GetBytes("Hi").CopyTo(text, 76);
        var items = Lay(Emf(font, Record(37, 1), Record(24, 0x00FF00), Record(22, 24), text));
        var run = Assert.Single(items.OfType<TextItem>());
        Assert.Equal("Hi", run.Text);
        Assert.Equal(new Color(0, 255, 0), run.Color);
        Assert.Equal(40, run.X, 3);
        Assert.Equal(80, run.Baseline, 3);
        Assert.Equal(24, run.FontSize, 3);
    }

    [Fact]
    public void A_placeable_wmf_polygon_is_mapped_from_its_bounds()
    {
        var d = new List<byte>();
        void U16(int v) { d.Add((byte)v); d.Add((byte)(v >> 8)); }
        void U32(int v) { U16(v & 0xFFFF); U16(v >> 16 & 0xFFFF); }
        U32(unchecked((int)0x9AC6CDD7)); U16(0); U16(0); U16(0); U16(1000); U16(750); U16(1440); U32(0); U16(0);
        U16(1); U16(9); U16(0x300); U32(0); U16(1); U32(0); U16(0);
        U32(7); U16(0x02FC); U16(0); U32(0xFF0000); U16(0); // CreateBrushIndirect: solid blue
        U32(4); U16(0x012D); U16(0);                        // SelectObject 0
        U32(10); U16(0x0324); U16(3); U16(0); U16(0); U16(1000); U16(0); U16(500); U16(750); // Polygon
        U32(3); U16(0);
        var path = Assert.Single(Lay([.. d]).OfType<PathItem>());
        Assert.Equal(new Color(0, 0, 255), path.Fill);
        Assert.Contains(path.Points, p => Math.Abs(p.X - 400) < 0.01 && Math.Abs(p.Y) < 0.01);
        Assert.Contains(path.Points, p => Math.Abs(p.X - 200) < 0.01 && Math.Abs(p.Y - 300) < 0.01);
    }

    /// <summary>An EMF+ record: type, flags and data.</summary>
    private static byte[] Plus(int type, int flags, params byte[][] data)
    {
        var body = data.SelectMany(b => b).ToArray();
        var r = new byte[12 + body.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(r, (ushort)type);
        BinaryPrimitives.WriteUInt16LittleEndian(r.AsSpan(2), (ushort)flags);
        Put(r, 4, r.Length, body.Length);
        body.CopyTo(r, 12);
        return r;
    }

    private static byte[] I(params int[] values)
    {
        var d = new byte[values.Length * 4];
        Put(d, 0, values);
        return d;
    }

    private static byte[] F(params float[] values) => I([.. values.Select(BitConverter.SingleToInt32Bits)]);

    /// <summary>An EMF comment holding an EMF+ header (254 dpi, so EMF+ pixels are the EMF's device pixels) and the records.</summary>
    private static byte[] PlusComment(bool dual, params byte[][] records)
    {
        var plus = new List<byte[]> { Plus(0x4001, dual ? 1 : 0, I(unchecked((int)0xDBC01002), dual ? 1 : 0, 254, 254)) };
        plus.AddRange(records);
        plus.Add(Plus(0x4002, 0));
        var body = plus.SelectMany(b => b).ToArray();
        var comment = new byte[16 + body.Length];
        Put(comment, 0, 70, comment.Length, body.Length + 4);
        Encoding.ASCII.GetBytes("EMF+").CopyTo(comment, 12);
        body.CopyTo(comment, 16);
        return comment;
    }

    private const int Red = unchecked((int)0xFFFF0000);

    [Fact]
    public void An_emf_plus_rectangle_is_filled_with_its_color_and_fitted_to_the_picture()
    {
        var items = Lay(Emf(PlusComment(false, Plus(0x400A, 0x8000, I(Red, 1), F(0, 0, 500, 375)))));
        var path = Assert.Single(items.OfType<PathItem>());
        Assert.Equal(new Color(255, 0, 0), path.Fill);
        var points = path.Points.Where(p => p.Kind != PathPointKind.Close).ToList();
        Assert.Equal(200, points.Max(p => p.X), 3);
        Assert.Equal(150, points.Max(p => p.Y), 3);
        Assert.DoesNotContain(items, i => i is ImageItem);
    }

    [Fact]
    public void An_emf_plus_pen_strokes_lines_in_page_units_and_a_translucent_brush_fills_in_a_faded_group()
    {
        var pen = Plus(0x4008, 0x0200 | 1, I(unchecked((int)0xDBC01002), 0, 0, 0), F(1), I(unchecked((int)0xDBC01002), 0, unchecked((int)0xFF0000FF)));
        var items = Lay(Emf(PlusComment(false,
            pen,
            Plus(0x4030, 6, F(1)),                             // page unit: millimetres
            Plus(0x400D, 1, I(2), F(0, 10, 50, 10)),           // DrawLines with pen 1
            Plus(0x400E, 0x8000, I(unchecked((int)0x8000FF00)), F(0, 0, 10, 10))))); // FillEllipse, half-transparent green
        var paths = items.OfType<PathItem>().ToList();
        Assert.Equal(2, paths.Count);
        Assert.Equal(new Color(0, 0, 255), paths[0].Stroke);
        Assert.Equal(200, paths[0].Points.Max(p => p.X), 3);
        Assert.Equal(40, paths[0].Points[0].Y, 3);
        Assert.Equal(4, paths[0].LineWidth, 3);
        var fade = Assert.IsType<PushItem>(items[items.IndexOf(paths[1]) - 1]);
        Assert.Equal(128 / 255.0, fade.Opacity, 3);
        Assert.Equal(40, paths[1].Points.Max(p => p.X), 1);
    }

    [Fact]
    public void Emf_plus_strings_are_drawn_with_their_font_color_and_alignment()
    {
        var name = Encoding.Unicode.GetBytes("Noto Sans");
        var font = Plus(0x4008, 0x0600 | 2, I(unchecked((int)0xDBC01002)), F(60), I(2, 0, 0, 9), name);
        var format = Plus(0x4008, 0x0700 | 3, I(unchecked((int)0xDBC01002), 0, 0, 1, 0), new byte[40]);
        var text = Encoding.Unicode.GetBytes("Hi");
        var items = Lay(Emf(PlusComment(false, font, format,
            Plus(0x401C, 0x8000 | 2, I(unchecked((int)0xFF00FF00), 3, 2), F(100, 200, 500, 100), text))));
        var run = Assert.Single(items.OfType<TextItem>());
        Assert.Equal("Hi", run.Text);
        Assert.Equal(new Color(0, 255, 0), run.Color);
        Assert.Equal(24, run.FontSize, 3);
        double width = run.Font.MeasureText("Hi", 24);
        Assert.Equal(40 + (200 - width) / 2, run.X, 3);
        Assert.True(run.Baseline > 80 && run.Baseline < 110);
    }

    [Fact]
    public void An_emf_plus_clip_rectangle_clips_what_is_drawn_after_it()
    {
        var items = Lay(Emf(PlusComment(false,
            Plus(0x4032, 0x0100, F(0, 0, 500, 375)),           // SetClipRect, intersect
            Plus(0x400A, 0x8000, I(Red, 1), F(0, 0, 1000, 750)))));
        var path = Assert.Single(items.OfType<PathItem>());
        var clip = Assert.IsType<PushItem>(items[items.IndexOf(path) - 1]);
        Assert.Equal(200, clip.ClipOutline!.Max(p => p.X), 3);
        Assert.Equal(items.OfType<PushItem>().Count(), items.OfType<PopItem>().Count());
    }

    [Fact]
    public void A_dual_emf_draws_its_gdi_records_and_not_its_emf_plus_ones()
    {
        var items = Lay(Emf(
            PlusComment(true, Plus(0x400A, 0x8000, I(Red, 1), F(0, 0, 500, 375))),
            Record(43, 0, 0, 500, 375)));
        var path = Assert.Single(items.OfType<PathItem>());
        Assert.Equal(Color.White, path.Fill);
    }

    [Fact]
    public void Gdi_records_after_an_emf_plus_get_dc_are_drawn_too()
    {
        var items = Lay(Emf(
            PlusComment(false, Plus(0x400A, 0x8000, I(Red, 1), F(0, 0, 500, 375)), Plus(0x4004, 0)),
            Record(43, 0, 0, 100, 100)));
        var paths = items.OfType<PathItem>().ToList();
        Assert.Equal(2, paths.Count);
        Assert.Equal(new Color(255, 0, 0), paths[0].Fill);
        Assert.Equal(Color.White, paths[1].Fill);
    }

    /// <summary>An EMF+ linear gradient brush object (clamped unless <paramref name="wrap"/> says otherwise) from start to end color over the rectangle.</summary>
    private static byte[] LinearBrush(int id, float[] rect, int start, int end, int wrap = 4, int flags = 0, params byte[][] extra) =>
        Plus(0x4008, 0x100 | id, [I(unchecked((int)0xDBC01002), 4, flags, wrap), F(rect), I(start, end, 0, 0), .. extra]);

    [Fact]
    public void An_emf_plus_linear_gradient_brush_fills_a_gradient_clipped_to_the_shape()
    {
        var items = Lay(Emf(PlusComment(false,
            LinearBrush(1, [0, 0, 500, 375], Red, unchecked((int)0xFF0000FF)),
            Plus(0x400E, 0, I(1), F(0, 0, 500, 375)))));
        var gradient = Assert.Single(items.OfType<GradientItem>());
        Assert.Equal(0, gradient.Degree, 3);
        Assert.Equal(200, gradient.Width, 3);
        Assert.Equal([new GradientStop(0, new Color(255, 0, 0)), new GradientStop(1, new Color(0, 0, 255))], gradient.Stops);
        var clip = items.OfType<PushItem>().Last(p => p.ClipOutline is not null);
        Assert.True(clip.ClipOutline!.Count > 16); // the flattened ellipse
        Assert.Empty(items.OfType<PathItem>());
    }

    [Fact]
    public void A_tiled_emf_plus_linear_gradient_repeats_across_the_shape()
    {
        var items = Lay(Emf(PlusComment(false,
            LinearBrush(1, [0, 0, 250, 375], Red, unchecked((int)0xFF0000FF), wrap: 0),
            Plus(0x400A, 0, I(1, 1), F(0, 0, 500, 375)))));
        var stops = Assert.Single(items.OfType<GradientItem>()).Stops;
        Assert.Equal([0, 0.5, 0.5, 1], stops.Select(s => Math.Round(s.Position, 6)));
        Assert.Equal(new Color(255, 0, 0), stops[2].Color);
    }

    [Fact]
    public void An_emf_plus_gradient_with_preset_colors_uses_them_as_stops()
    {
        var items = Lay(Emf(PlusComment(false,
            LinearBrush(1, [0, 0, 500, 375], 0, 0, 4, 4, I(3), F(0, 0.25f, 1), I(Red, unchecked((int)0xFFFFFFFF), unchecked((int)0xFF0000FF))),
            Plus(0x400A, 0, I(1, 1), F(0, 0, 500, 375)))));
        var stops = Assert.Single(items.OfType<GradientItem>()).Stops;
        Assert.Equal([0, 0.25, 1], stops.Select(s => Math.Round(s.Position, 6)));
        Assert.Equal(Color.White, stops[1].Color);
    }

    [Fact]
    public void An_emf_plus_path_gradient_with_a_rectangular_boundary_runs_in_rectangles_from_its_center()
    {
        var items = Lay(Emf(PlusComment(false,
            Plus(0x4008, 0x100 | 1, I(unchecked((int)0xDBC01002), 3, 0, 0, unchecked((int)0xFFFFFFFF)), F(250, 125), I(1, Red, 4), F(0, 0, 500, 0, 500, 250, 0, 250)),
            Plus(0x400A, 0, I(1, 1), F(0, 0, 500, 250)))));
        var gradient = Assert.Single(items.OfType<GradientItem>());
        Assert.True(gradient.Radial && gradient.Rectangular);
        Assert.Equal(0.5, gradient.FocusX, 3);
        Assert.Equal(Color.White, gradient.Stops[0].Color);
        Assert.Equal(new Color(255, 0, 0), gradient.Stops[^1].Color);
    }

    [Fact]
    public void A_hatched_emf_brush_draws_its_lines_over_the_background_color_inside_the_shape()
    {
        var items = Lay(Emf(
            Record(25, 0x00FFFF),                // SetBkColor: yellow
            Record(39, 1, 2, 0xFF0000, 1),       // CreateBrushIndirect: blue vertical hatch
            Record(37, 1),
            Record(37, unchecked((int)0x80000008)),
            Record(43, 0, 0, 100, 100)));
        var paths = items.OfType<PathItem>().ToList();
        Assert.Equal(new Color(255, 255, 0), paths[0].Fill);
        var lines = paths.Skip(1).ToList();
        Assert.True(lines.Count >= 12); // a line every 8 device pixels across 100
        Assert.All(lines, l => Assert.Equal(new Color(0, 0, 255), l.Stroke));
        Assert.All(lines, l => Assert.Equal(l.Points[0].X, l.Points[1].X, 6));
        Assert.Contains(items, i => i is PushItem { ClipOutline: not null });
    }

    [Fact]
    public void A_transparent_background_leaves_out_the_hatch_background()
    {
        var items = Lay(Emf(Record(18, 1), Record(39, 1, 2, 0xFF0000, 0), Record(37, 1), Record(37, unchecked((int)0x80000008)), Record(43, 0, 0, 100, 100)));
        Assert.All(items.OfType<PathItem>(), p => Assert.Null(p.Fill));
    }

    [Fact]
    public void A_dib_pattern_brush_is_tiled_over_the_shape()
    {
        // A 2 × 2 24-bit DIB (rows padded to 8 bytes) right after the record's fixed part.
        int[] bmi = [40, 2, 2, 1 | 24 << 16, 0, 16, 0, 0, 0, 0];
        int[] bits = [0x00FF0000, 0x00FFFFFF, 0x0000FF00, 0];
        var items = Lay(Emf(Record(94, [1, 0, 32, 40, 72, 16, .. bmi, .. bits]), Record(37, 1), Record(37, unchecked((int)0x80000008)), Record(43, 0, 0, 10, 10)));
        var tiles = items.OfType<ImageItem>().ToList();
        Assert.Equal(25, tiles.Count); // 5 × 5 tiles of 2 pixels cover 10
        Assert.Single(tiles.Select(i => i.Image).Distinct());
    }

    [Fact]
    public void An_emf_gradient_fill_rectangle_is_a_gradient_and_a_triangle_its_mean_color()
    {
        static int[] Vertex(int x, int y, int r, int g, int b) => [x, y, r << 8 | g << 24, b << 8];
        var items = Lay(Emf(
            Record(118, [0, 0, 0, 0, 2, 1, 1, .. Vertex(0, 0, 255, 0, 0), .. Vertex(100, 50, 0, 0, 255), 0, 1]),
            Record(118, [0, 0, 0, 0, 3, 1, 2, .. Vertex(0, 100, 255, 0, 0), .. Vertex(50, 60, 0, 255, 0), .. Vertex(100, 100, 0, 0, 255), 0, 1, 2])));
        var gradient = Assert.Single(items.OfType<GradientItem>());
        Assert.Equal(90, gradient.Degree, 3);
        Assert.Equal(new Color(255, 0, 0), gradient.Stops[0].Color);
        Assert.Equal(new Color(0, 0, 255), gradient.Stops[1].Color);
        Assert.Equal(new Color(85, 85, 85), Assert.Single(items.OfType<PathItem>()).Fill);
    }
}
