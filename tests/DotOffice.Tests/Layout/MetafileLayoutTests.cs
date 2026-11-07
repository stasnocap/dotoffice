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

    [Fact]
    public void An_emf_drawn_only_with_emf_plus_records_keeps_its_placeholder_and_warns()
    {
        var comment = new byte[20];
        Put(comment, 0, 70, 20, 8);
        Encoding.ASCII.GetBytes("EMF+").CopyTo(comment, 12);
        var warnings = new ConversionWarningList();
        var items = Lay(Emf(comment), warnings);
        Assert.Contains(items, i => i is ImageItem);
        Assert.Contains(warnings.Warnings, w => w.Code == "metafile.unsupported");
    }
}
