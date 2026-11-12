using DotOffice.Core;
using DotOffice.Layout;

namespace DotOffice.Tests.Layout;

public class SheetChartLayoutTests
{
    private static readonly Color Red = new(255, 0, 0), Blue = new(0, 0, 255);

    /// <summary>The page position of a chart anchored at A1: the sheet's left and top margins.</summary>
    private const double OriginX = 50.4, OriginY = 54;

    private static SheetChart Chart(ChartKind kind, params ChartSeries[] series) => new()
    {
        Anchor = ImageAnchorKind.Absolute, X = 0, Y = 0, Width = 300, Height = 200, Kind = kind,
        Categories = ["a", "b", "c"], Series = series, Legend = LegendPosition.None, Title = null,
        Ticks = [new(0, "0"), new(10, "10"), new(20, "20")], ValueMin = 0, ValueMax = 20,
    };

    private static ChartSeries Series(Color color, params double?[] values) => new("S", values, color);

    private static LayoutPage Page(SheetChart chart)
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text("x")));
        sheet.Charts.Add(chart);
        return SheetLayoutTests.Layout(sheet).Pages[0];
    }

    private static string Text(LayoutPage page) => string.Concat(page.Items.OfType<TextItem>().Select(t => t.Text));

    private static SheetChart WordChart(ChartKind kind, LegendPosition legend, params ChartSeries[] series) =>
        Chart(kind, series) with { WordLayout = true, FontSize = 10, Legend = legend, MajorGridlines = true, Width = 360, Height = 216 };

    private static (double Left, double Right, double Top, double Bottom) PlotOf(LayoutPage page)
    {
        var rows = page.Items.OfType<LineItem>().Where(l => l.Y1 == l.Y2 && l.X2 - l.X1 > 50).ToList();
        return (rows.Min(l => l.X1), rows.Max(l => l.X2), rows.Min(l => l.Y1), rows.Max(l => l.Y1));
    }

    [Fact]
    public void A_Word_chart_with_outside_ticks_draws_3_pt_marks_at_the_value_ticks_and_slot_boundaries()
    {
        var chart = WordChart(ChartKind.Column, LegendPosition.None, Series(Red, 5, 10, 20)) with { ValueAxisTicks = ChartTickMark.Out, CategoryAxisTicks = ChartTickMark.Out };
        var ticks = Page(chart).Items.OfType<LineItem>().Where(l => Math.Abs(Math.Max(l.X2 - l.X1, l.Y2 - l.Y1) - 3.02) < 1e-6).ToList();
        Assert.Equal(chart.Ticks.Count, ticks.Count(l => l.Y1 == l.Y2));
        Assert.Equal(4, ticks.Count(l => l.X1 == l.X2));
        Assert.DoesNotContain(Page(chart with { ValueAxisTicks = ChartTickMark.None, CategoryAxisTicks = ChartTickMark.None }).Items.OfType<LineItem>(), l => Math.Abs(Math.Max(l.X2 - l.X1, l.Y2 - l.Y1) - 3.02) < 1e-6);
    }

    [Fact]
    public void A_Word_chart_without_title_or_legend_has_Words_margins()
    {
        // Measured in Word 2016: the plot ends 11 pt from the right and starts 11 pt from the top; the bottom labels take 23.7 pt at 10 pt text.
        var (left, right, top, bottom) = PlotOf(Page(WordChart(ChartKind.Column, LegendPosition.None, Series(Red, 5, 10, 20))));
        Assert.Equal(OriginX + 360 - 11, right, 2);
        Assert.Equal(OriginY + 11, top, 2);
        Assert.InRange(OriginY + 216 - bottom, 18, 30);
        Assert.InRange(left - OriginX, 15, 40);
    }

    [Fact]
    public void A_Word_charts_right_legend_has_square_keys_17_pt_apart_and_the_plot_gives_way_by_its_width()
    {
        var chart = WordChart(ChartKind.Column, LegendPosition.Right, Series(Red, 5, 10, 20), Series(Blue, 5, 10, 20));
        var page = Page(chart);
        var keys = page.Items.OfType<RectItem>().Where(r => Math.Abs(r.Width - 5.03) < 1e-6 && Math.Abs(r.Height - 5.03) < 1e-6).OrderBy(r => r.Y).ToList();
        Assert.Equal([Red, Blue], keys.Select(k => k.Fill));
        var text = page.Items.OfType<TextItem>().Where(t => t.Text == "S").OrderBy(t => t.Baseline).ToList();
        Assert.Equal(2, text.Count);
        Assert.InRange(keys[1].Y - keys[0].Y, 17.07, 21); // 5.57 more than the line height, which is 11.5 pt for Arial but larger for the font used here
        Assert.Equal(keys[0].X + 7.23, text[0].X, 2);
        Assert.Equal(keys[0].X - 15.76, PlotOf(page).Right, 1); // the plot ends 15.76 left of the keys: 11 + (text width + 22.04) from the edge
        // The block of keys is centered on the chart's middle.
        Assert.Equal(OriginY + 108, (keys[0].Y + keys[1].Y + 5.03) / 2, 1);
    }

    [Theory]
    [InlineData(8, 4.02, 14.85)]
    [InlineData(14, 7.04, 21.49)]
    public void A_Word_charts_legend_keys_and_entry_pitch_scale_with_the_text_size(double size, double key, double pitch)
    {
        // Measured in Word 2016 at 8 and 14 pt (Arial): square keys 0.503 em, entries 14.85 / 21.49 pt apart.
        var chart = WordChart(ChartKind.Column, LegendPosition.Right, Series(Red, 5, 10, 20), Series(Blue, 5, 10, 20)) with { FontSize = size };
        var keys = Page(chart).Items.OfType<RectItem>().Where(r => Math.Abs(r.Width - key) < 0.01).OrderBy(r => r.Y).ToList();
        Assert.Equal(2, keys.Count);
        Assert.InRange(keys[1].Y - keys[0].Y, pitch - 0.5, pitch + 6); // the test font's line height differs from Arial's
    }

    [Fact]
    public void A_Word_charts_right_legend_lists_stacked_columns_and_clustered_bars_backwards()
    {
        ChartSeries a = new("A", [1, 2, 3], Red), b = new("B", [1, 2, 3], Blue);
        static List<Color> Order(LayoutPage page, bool alongX = false) => page.Items.OfType<RectItem>().Where(r => Math.Abs(r.Width - 5.03) < 1e-6).OrderBy(r => alongX ? r.X : r.Y).Select(r => r.Fill).ToList();
        Assert.Equal([Red, Blue], Order(Page(WordChart(ChartKind.Column, LegendPosition.Right, a, b))));
        Assert.Equal([Blue, Red], Order(Page(WordChart(ChartKind.Column, LegendPosition.Right, a, b) with { Grouping = ChartGrouping.Stacked })));
        Assert.Equal([Blue, Red], Order(Page(WordChart(ChartKind.Bar, LegendPosition.Right, a, b))));
        // Only the vertical legends are reversed.
        Assert.Equal([Red, Blue], Order(Page(WordChart(ChartKind.Column, LegendPosition.Bottom, a, b) with { Grouping = ChartGrouping.Stacked }), true));
    }

    [Fact]
    public void A_Word_charts_bottom_legend_is_one_row_that_takes_23_pt_off_the_plot()
    {
        var plain = PlotOf(Page(WordChart(ChartKind.Column, LegendPosition.None, Series(Red, 5, 10, 20))));
        var page = Page(WordChart(ChartKind.Column, LegendPosition.Bottom, Series(Red, 5, 10, 20), Series(Blue, 5, 10, 20)));
        double row = plain.Bottom - PlotOf(page).Bottom; // the entry pitch (line height + 5.57; 17.07 for Arial 10) + 6
        Assert.InRange(row, 23.07, 27);
        var keys = page.Items.OfType<RectItem>().Where(r => Math.Abs(r.Width - 5.03) < 1e-6).OrderBy(r => r.X).ToList();
        Assert.Equal(2, keys.Count);
        Assert.Equal(keys[0].Y, keys[1].Y, 3);
        Assert.Equal(OriginY + 216 - 6 - (row - 6) / 2 - 2.515, keys[0].Y, 1); // centered in the row, which ends 6 pt above the edge (17.05 above it for Arial 10)
    }

    [Fact]
    public void A_Word_pie_is_the_largest_circle_in_the_space_the_legend_leaves_and_has_no_white_outlines()
    {
        var page = Page(WordChart(ChartKind.Pie, LegendPosition.Right, Series(Red, 5, 10, 20)));
        var slices = page.Items.OfType<PolygonItem>().ToList();
        Assert.NotEmpty(slices);
        Assert.Equal(OriginY + 11, slices.SelectMany(s => s.Points).Min(p => p.Y), 1);
        Assert.Equal((216 - 22) / 2.0, (slices.SelectMany(s => s.Points).Max(p => p.Y) - slices.SelectMany(s => s.Points).Min(p => p.Y)) / 2, 1);
        Assert.DoesNotContain(page.Items.OfType<LineItem>(), l => l.Color == Color.White);
    }

    [Fact]
    public void Without_the_Word_flag_the_shared_layout_is_unchanged()
    {
        var page = Page(Chart(ChartKind.Column, Series(Red, 5, 10, 20)) with { Legend = LegendPosition.Right, MajorGridlines = true });
        Assert.NotEqual(OriginX + 300 - 11, PlotOf(page).Right, 1);
        Assert.DoesNotContain(page.Items.OfType<RectItem>(), r => Math.Abs(r.Width - 5.03) < 1e-6);
    }

    [Fact]
    public void Candlestick_bodies_are_white_when_the_close_rises_and_black_when_it_falls()
    {
        var chart = Chart(ChartKind.Line, Series(Red, 5, 15, 10), Series(Blue, 8, 18, 14), Series(Red, 9, 6, 10)) with { UpDownBars = true };
        var page = Page(chart);
        var black = page.Items.OfType<RectItem>().Where(r => r.Fill == Color.Black).OrderBy(r => r.X).ToList();
        var white = page.Items.OfType<RectItem>().Where(r => r.Fill == Color.White && r.Width < 40 && r.Height < 100).ToList();
        Assert.Equal(3, black.Count(b => b.Height > 0 && b.Width < 40 && b.Width > 5)); // an outline per category (the third is flat)
        Assert.Single(white); // only the first rises (5 -> 9); the second falls, the third is unchanged
        Assert.True(white[0].X > black[0].X && white[0].Width < black[0].Width);
    }

    [Fact]
    public void Columns_are_rectangles_whose_heights_follow_the_values()
    {
        var page = Page(Chart(ChartKind.Column, Series(Red, 5, 10, 20)));
        var bars = page.Items.OfType<RectItem>().Where(r => r.Fill == Red).OrderBy(r => r.X).ToList();
        Assert.Equal(3, bars.Count);
        Assert.Equal(2, bars[1].Height / bars[0].Height, 3);
        Assert.Equal(2, bars[2].Height / bars[1].Height, 3);
        Assert.Equal(bars[0].Y + bars[0].Height, bars[2].Y + bars[2].Height, 3);
        Assert.Contains("10", Text(page));
    }

    [Fact]
    public void Clustered_series_sit_side_by_side_and_stacked_ones_on_top_of_each_other()
    {
        var clustered = Page(Chart(ChartKind.Column, Series(Red, 5, 5, 5), Series(Blue, 5, 5, 5)));
        Assert.True(clustered.Items.OfType<RectItem>().First(r => r.Fill == Blue).X > clustered.Items.OfType<RectItem>().First(r => r.Fill == Red).X);
        var stacked = Page(Chart(ChartKind.Column, Series(Red, 5, 5, 5), Series(Blue, 5, 5, 5)) with { Grouping = ChartGrouping.Stacked });
        var red = stacked.Items.OfType<RectItem>().First(r => r.Fill == Red);
        var blue = stacked.Items.OfType<RectItem>().First(r => r.Fill == Blue);
        Assert.Equal(red.X, blue.X, 3);
        Assert.Equal(red.Y, blue.Y + blue.Height, 3);
    }

    [Fact]
    public void Horizontal_bars_grow_to_the_right()
    {
        var page = Page(Chart(ChartKind.Bar, Series(Red, 5, 10, 20)));
        var bars = page.Items.OfType<RectItem>().Where(r => r.Fill == Red).OrderBy(r => r.Y).ToList();
        Assert.Equal(3, bars.Count);
        Assert.True(bars[0].Width > bars[1].Width && bars[1].Width > bars[2].Width); // top bar is the last category
        Assert.Equal(bars[0].X, bars[2].X, 3);
    }

    [Fact]
    public void A_line_chart_joins_the_points_with_segments_and_marks_them()
    {
        var page = Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = true }));
        Assert.Equal(2, page.Items.OfType<LineItem>().Count(l => l.Color == Red));
        Assert.Equal(3, Filled(page, Red).Count); // the first series' automatic symbol (c:idx 0) is a diamond, filled in its automatic color
    }

    [Fact]
    public void A_marker_is_5_points_across_unless_the_series_gives_a_size()
    {
        var auto = Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = true }));
        Assert.All(Filled(auto, Red), b => Assert.Equal(5, b.Width));
        var big = Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = true, MarkerSize = 9 }));
        Assert.All(Filled(big, Red), b => Assert.Equal(9, b.Width));
    }

    private static (double X, double Y, double Width, double Height) Box(PathItem p)
    {
        var points = p.Points.Where(q => q.Kind != PathPointKind.Close).ToList();
        return (points.Min(q => q.X), points.Min(q => q.Y), points.Max(q => q.X) - points.Min(q => q.X), points.Max(q => q.Y) - points.Min(q => q.Y));
    }

    private static List<(double X, double Y, double Width, double Height)> Filled(LayoutPage page, Color color) =>
        page.Items.OfType<PathItem>().Where(p => p.Fill == color).Select(Box).ToList();

    private static string Signature(LayoutItem item) => item switch
    {
        PathItem p => $"path {p.Fill} {p.Stroke} {p.LineWidth} " + string.Join(' ', p.Points.Select(q => $"{q.Kind}:{q.X:F3},{q.Y:F3}")),
        LineItem l => $"line {l.X1:F3},{l.Y1:F3},{l.X2:F3},{l.Y2:F3} {l.Width} {l.Color}",
        _ => item.ToString()!,
    };

    private static LayoutPage Marked(ChartMarkerSymbol symbol, bool marker = true) =>
        Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = marker, MarkerSize = 9, MarkerSymbol = symbol }));

    [Fact]
    public void A_square_marker_is_a_path_of_the_marker_size_with_no_halo()
    {
        var page = Marked(ChartMarkerSymbol.Square);
        Assert.Equal(3, Filled(page, Red).Count(b => b.Width == 9 && b.Height == 9));
        Assert.DoesNotContain(page.Items.OfType<RectItem>(), r => r.Fill == Color.White && r.Width < 20); // Office 2016 draws no white halo
        Assert.DoesNotContain(page.Items.OfType<EllipseItem>(), e => e.Fill == Red);
    }

    [Fact]
    public void A_diamond_marker_is_a_four_point_polygon_as_wide_and_high_as_the_marker_size()
    {
        var diamonds = Marked(ChartMarkerSymbol.Diamond).Items.OfType<PathItem>().Where(p => p.Fill == Red).ToList();
        Assert.Equal(3, diamonds.Count);
        Assert.All(diamonds, d =>
        {
            Assert.Equal(4, d.Points.Count(p => p.Kind != PathPointKind.Close));
            var box = Box(d);
            Assert.Equal(9, box.Width, 6);
            Assert.Equal(9, box.Height, 6);
        });
    }

    [Fact]
    public void An_x_marker_is_two_crossing_thin_strokes_in_the_automatic_outline_color()
    {
        var page = Marked(ChartMarkerSymbol.X);
        var strokes = page.Items.OfType<LineItem>().Where(l => l.Color == Red && l.Width == 0.75 && Math.Abs(l.X2 - l.X1) == 9 && Math.Abs(l.Y2 - l.Y1) == 9).ToList();
        Assert.Equal(6, strokes.Count);
        Assert.DoesNotContain(page.Items.OfType<EllipseItem>(), e => e.Fill == Red);
    }

    [Fact]
    public void A_none_symbol_draws_no_marker_and_an_automatic_symbol_is_a_diamond_for_c_idx_0()
    {
        var none = Marked(ChartMarkerSymbol.Auto, marker: false);
        Assert.Empty(Filled(none, Red));
        var auto = Marked(ChartMarkerSymbol.Auto);
        Assert.Equal(3, Filled(auto, Red).Count(b => b.Width == 9)); // no white halo either
        Assert.Equal(Marked(ChartMarkerSymbol.Diamond).Items.OfType<PathItem>().Select(Signature), auto.Items.OfType<PathItem>().Select(Signature));
        Assert.NotEqual(Marked(ChartMarkerSymbol.Circle).Items.OfType<PathItem>().Select(Signature), auto.Items.OfType<PathItem>().Select(Signature));
    }

    [Fact]
    public void A_legend_entry_uses_the_series_marker_shape()
    {
        var chart = Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Square }) with { Legend = LegendPosition.Right };
        var page = Page(chart);
        Assert.Equal(3, Filled(page, Red).Count(b => b.Width == 5 && b.Height == 5)); // the three points
        Assert.Single(Filled(page, Red), b => b.Width == 6 && b.Height == 6); // the legend marker is 6 across whatever the plot's size
    }

    [Fact]
    public void A_legend_marker_is_six_across_whatever_the_plot_marker_size()
    {
        var series = Series(Red, 5, 10, 20) with { Marker = true, MarkerSize = 12, MarkerSymbol = ChartMarkerSymbol.Square };
        var page = Page(Chart(ChartKind.Line, series) with { Legend = LegendPosition.Right });
        Assert.Equal(3, Filled(page, Red).Count(b => b.Width == 12));
        Assert.Single(Filled(page, Red), b => b.Width == 6 && b.Height == 6);
    }

    private static LayoutPage Styled(ChartMarkerStyle style, ChartMarkerSymbol symbol = ChartMarkerSymbol.Circle, ChartKind kind = ChartKind.Line) =>
        Page(Chart(kind, Series(Red, 5, 10, 20) with { Marker = true, MarkerSize = 9, MarkerSymbol = symbol, MarkerStyle = style }));

    private static IEnumerable<PathItem> Paths(LayoutPage page) => page.Items.OfType<PathItem>();

    [Fact]
    public void Axis_lines_are_black_at_the_width_the_chart_gives_them()
    {
        var page = Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20)) with { AxisLineWidth = 0.5, MajorGridlines = false, Border = null, Legend = LegendPosition.None });
        Assert.Contains(page.Items, i => i is LineItem { Width: 0.5 } l && l.Color == Color.Black && l.Y1 == l.Y2);
        Assert.Contains(page.Items, i => i is LineItem { Width: 0.5 } l && l.Color == Color.Black && l.X1 == l.X2);
    }

    [Fact]
    public void A_path_scales_with_the_sheet_around_the_page_margin()
    {
        var series = Series(Red, 5, 10, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Circle, MarkerStyle = new ChartMarkerStyle(Fill: Blue, Line: Blue, LineWidth: 2) };
        LayoutPage Scaled(double scale)
        {
            var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text("x")));
            sheet.PageSetup.Scale = scale;
            sheet.Charts.Add(Chart(ChartKind.Line, series));
            return SheetLayoutTests.Layout(sheet).Pages[0];
        }
        var full = Paths(Scaled(1)).ToList();
        var half = Paths(Scaled(0.5)).ToList();
        Assert.NotEmpty(full);
        Assert.Equal(full.Count, half.Count);
        double left = OriginX; // the page margin: scaling keeps it fixed
        for (int i = 0; i < full.Count; i++)
        {
            Assert.Equal(full[i].LineWidth / 2, half[i].LineWidth, 6);
            for (int k = 0; k < full[i].Points.Count; k++)
                Assert.Equal(left + (full[i].Points[k].X - left) / 2, half[i].Points[k].X, 6);
        }
    }

    [Fact]
    public void A_styled_marker_sits_on_the_page_where_an_unstyled_one_does()
    {
        var plain = Series(Red, 5, 10, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Circle };
        var styled = plain with { MarkerStyle = new ChartMarkerStyle(Fill: Blue) };
        var page = Page(Chart(ChartKind.Line, plain, styled));
        var circles = Filled(page, Red);
        var others = Filled(page, Blue);
        Assert.Equal(circles.Count, others.Count);
        for (int i = 0; i < others.Count; i++)
        {
            Assert.Equal(circles[i].X + circles[i].Width / 2, others[i].X + others[i].Width / 2, 1);
            Assert.Equal(circles[i].Y + circles[i].Height / 2, others[i].Y + others[i].Height / 2, 1);
        }
    }

    [Fact]
    public void A_marker_without_a_style_is_filled_in_the_automatic_color_with_a_thin_automatic_outline_and_no_halo()
    {
        var auto = new ChartAutoMarker(0, Blue, Color.Black, 0.5, 5);
        var plain = Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = true, MarkerSize = 9, MarkerSymbol = ChartMarkerSymbol.Circle, Auto = auto }));
        var paths = Paths(plain).ToList();
        Assert.Equal(6, paths.Count);
        Assert.Equal(3, paths.Count(p => p.Fill == Blue && p.Stroke is null)); // the series color (Red) is not used
        Assert.Equal(3, paths.Count(p => p.Fill is null && p.Stroke == Color.Black && p.LineWidth == 0.5));
        Assert.DoesNotContain(plain.Items.OfType<EllipseItem>(), e => e.Fill == Color.White);
    }

    [Fact]
    public void A_hollow_marker_is_an_outline_of_the_stated_width_with_no_fill_and_no_halo()
    {
        var page = Styled(new ChartMarkerStyle(NoFill: true, Line: Blue, LineWidth: 2));
        var rings = Paths(page).ToList();
        Assert.Equal(3, rings.Count);
        Assert.All(rings, p => Assert.Equal((null, (Color?)Blue, 2.0), (p.Fill, p.Stroke, p.LineWidth)));
        Assert.DoesNotContain(page.Items.OfType<EllipseItem>(), e => e.Fill == Color.White && Math.Abs(e.Width - 10.4) < 1e-9);
    }

    [Fact]
    public void A_white_filled_marker_with_an_outline_draws_the_fill_then_the_outline()
    {
        var items = Styled(new ChartMarkerStyle(Fill: Color.White, Line: Red), ChartMarkerSymbol.Square).Items;
        var paths = items.OfType<PathItem>().ToList();
        Assert.Equal(6, paths.Count);
        Assert.Equal((Color?)Color.White, paths[0].Fill);
        Assert.Equal((Color?)Red, paths[1].Stroke);
        Assert.Equal(0.75, paths[1].LineWidth);
        Assert.True(items.IndexOf(paths[0]) < items.IndexOf(paths[1]));
    }

    [Fact]
    public void A_custom_fill_without_an_outline_gets_the_thin_automatic_outline()
    {
        // Excel and PowerPoint 2016 (spec-fill): a marker with a fill and no stated outline is outlined in the automatic color.
        var paths = Paths(Styled(new ChartMarkerStyle(Fill: Blue))).ToList();
        Assert.Equal(6, paths.Count);
        Assert.Equal(3, paths.Count(p => p.Fill == Blue && p.Stroke is null));
        Assert.Equal(3, paths.Count(p => p.Fill is null && p.Stroke == Red && p.LineWidth == 0.75));
    }

    [Fact]
    public void A_no_fill_marker_without_an_outline_is_a_ring_in_the_automatic_outline()
    {
        var rings = Paths(Styled(new ChartMarkerStyle(NoFill: true))).ToList();
        Assert.Equal(3, rings.Count);
        Assert.All(rings, p => Assert.Equal((null, (Color?)Red, 0.75), (p.Fill, p.Stroke, p.LineWidth)));
    }

    [Fact]
    public void A_filled_open_marker_has_an_unoutlined_box_of_the_marker_size_behind_its_strokes()
    {
        // Measured (Office 2016): x, plus and star with a fill are an unoutlined filled box of the marker size under strokes in the outline color.
        foreach (var symbol in new[] { ChartMarkerSymbol.X, ChartMarkerSymbol.Plus, ChartMarkerSymbol.Star })
        {
            var items = Styled(new ChartMarkerStyle(Fill: Blue), symbol).Items;
            var boxes = items.OfType<PathItem>().Where(p => p.Fill == Blue).ToList();
            Assert.Equal(3, boxes.Count);
            Assert.All(boxes, b => Assert.Equal((9.0, 9.0), (Box(b).Width, Box(b).Height)));
            Assert.DoesNotContain(items.OfType<PathItem>(), p => p.Fill is null);
            Assert.Equal(symbol == ChartMarkerSymbol.Star ? 9 : 6, items.OfType<LineItem>().Count(l => l.Color == Red && l.Width == 0.75));
            Assert.True(items.IndexOf(boxes[0]) < items.IndexOf(items.OfType<LineItem>().First(l => l.Color == Red && l.Width == 0.75)));
        }
    }

    [Fact]
    public void A_filled_dot_dash_and_square_keep_their_outline()
    {
        foreach (var symbol in new[] { ChartMarkerSymbol.Dot, ChartMarkerSymbol.Dash, ChartMarkerSymbol.Square })
            Assert.Contains(Styled(new ChartMarkerStyle(Fill: Blue, Line: Red), symbol).Items.OfType<PathItem>(), p => p.Fill is null && p.Stroke == Red);
    }

    [Fact]
    public void A_stock_series_that_states_symbol_none_has_no_legend_marker_but_a_markers_only_series_keeps_it()
    {
        PathItem[] Legend(ChartSeries s) => Page(Chart(ChartKind.Line, s) with { Legend = LegendPosition.Right }).Items.OfType<PathItem>().Where(p => p.Fill == Red).ToArray();
        var hidden = Series(Red, 5, 10, 20) with { Line = false, Marker = false, MarkerNone = true };
        Assert.Empty(Legend(hidden));
        Assert.Equal(4, Legend(Series(Red, 5, 10, 20) with { Line = false, Marker = true }).Length); // three points and the legend entry
        // a point that states its own marker is still drawn in the plot
        var point = hidden with { PointMarkers = new Dictionary<int, ChartPointMarker> { [1] = new(Visible: true, Symbol: ChartMarkerSymbol.Square) } };
        Assert.Single(Legend(point)); // that point only, no legend marker
    }

    [Fact]
    public void Hi_low_lines_use_the_stated_color_and_width_else_the_default()
    {
        var series = new[] { Series(Red, 5, 10, 20), Series(Red, 8, 18, 14) };
        var stated = Page(Chart(ChartKind.Line, series) with { HiLowLines = true, HiLowColor = Blue, HiLowWidth = 0.75 }).Items.OfType<LineItem>().Where(l => l.Color == Blue).ToList();
        Assert.Equal(3, stated.Count);
        Assert.All(stated, l => Assert.Equal(0.75, l.Width));
        var plain = Page(Chart(ChartKind.Line, series) with { HiLowLines = true }).Items.OfType<LineItem>().Where(l => l.Color == new Color(0x40, 0x40, 0x40)).ToList();
        Assert.Equal(3, plain.Count);
        Assert.All(plain, l => Assert.Equal(1, l.Width));
    }

    [Fact]
    public void A_no_fill_open_marker_is_strokes_only_and_an_unfilled_one_has_no_box()
    {
        Assert.Empty(Paths(Styled(new ChartMarkerStyle(NoFill: true), ChartMarkerSymbol.Plus)));
        Assert.Equal(6, Styled(new ChartMarkerStyle(NoFill: true), ChartMarkerSymbol.Plus).Items.OfType<LineItem>().Count(l => l.Color == Red && l.Width == 0.75));
        Assert.Empty(Paths(Marked(ChartMarkerSymbol.Plus)));
    }

    private static LayoutPage Sized(ChartMarkerSymbol symbol, double size) =>
        Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = true, MarkerSize = size, MarkerSymbol = symbol }));

    [Fact]
    public void A_dash_is_a_bar_as_wide_as_the_marker_and_a_fifth_as_high_and_a_dot_is_its_right_half()
    {
        // Excel and PowerPoint 2016 (sizes chart, sizes 5-40): dash width = size, height about 0.2 size, both centered on the point; dot from the point outward.
        foreach (double size in new[] { 5.0, 10, 20, 40 })
        {
            var dash = Filled(Sized(ChartMarkerSymbol.Dash, size), Red);
            var dot = Filled(Sized(ChartMarkerSymbol.Dot, size), Red);
            Assert.Equal(3, dash.Count);
            Assert.Equal((size, size * 0.2), (dash[0].Width, dash[0].Height));
            Assert.Equal((size / 2, size * 0.2), (dot[0].Width, dot[0].Height), (a, b) => Math.Abs(a.Item1 - b.Item1) < 1e-9 && Math.Abs(a.Item2 - b.Item2) < 1e-9);
            Assert.Equal(dash[0].X + size / 2, dot[0].X, 6); // the dot starts at the point, where the dash has its middle
            Assert.Equal(dash[0].Y, dot[0].Y, 6);
        }
    }

    [Fact]
    public void Automatic_symbols_follow_the_series_c_idx_modulo_9()
    {
        // Excel and PowerPoint 2016 (rotation charts, 14 series): diamond, square, triangle, x, star, circle, plus, dot, dash, then again.
        var expected = new[] { ChartMarkerSymbol.Diamond, ChartMarkerSymbol.Square, ChartMarkerSymbol.Triangle, ChartMarkerSymbol.X, ChartMarkerSymbol.Star,
            ChartMarkerSymbol.Circle, ChartMarkerSymbol.Plus, ChartMarkerSymbol.Dot, ChartMarkerSymbol.Dash };
        for (int index = 0; index < 14; index++)
        {
            var auto = new ChartAutoMarker(index, Red, Red, 0.75, 5);
            var drawn = Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = true, Auto = auto })).Items.Where(i => i is PathItem or LineItem).ToList();
            var named = Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = true, MarkerSymbol = expected[index % 9], Auto = auto })).Items.Where(i => i is PathItem or LineItem).ToList();
            Assert.Equal(named.Select(Signature), drawn.Select(Signature));
        }
    }

    [Fact]
    public void An_automatic_marker_uses_the_c_idx_not_the_position_of_the_series()
    {
        // The first series in the list has c:idx 3 (an x: two strokes), the second c:idx 1 (a square).
        var first = Series(Red, 5, 10, 20) with { Marker = true, Auto = new ChartAutoMarker(3, Red, Red, 0.75, 5) };
        var second = Series(Blue, 5, 10, 20) with { Marker = true, Auto = new ChartAutoMarker(1, Blue, Blue, 0.75, 5) };
        var page = Page(Chart(ChartKind.Line, first, second));
        Assert.Equal(6, page.Items.OfType<LineItem>().Count(l => l.Color == Red && l.Width == 0.75));
        Assert.Equal(3, Filled(page, Blue).Count(b => b.Width == 5 && b.Height == 5));
    }

    [Fact]
    public void A_point_marker_without_a_symbol_or_style_keeps_the_series_automatic_symbol_and_color()
    {
        var auto = new ChartAutoMarker(1, Blue, Blue, 0.75, 5);
        var series = Pointed(Series(Red, 5, 10, 20) with { Marker = true, Auto = auto }, (1, new ChartPointMarker(Size: 9)));
        var boxes = Filled(Page(Chart(ChartKind.Line, series)), Blue);
        Assert.Equal([5.0, 9.0, 5.0], boxes.Select(b => b.Width)); // squares for c:idx 1, the middle one larger
    }

    [Fact]
    public void A_point_with_its_own_style_still_takes_its_automatic_parts_from_the_series_c_idx()
    {
        var auto = new ChartAutoMarker(2, Blue, Color.Black, 0.5, 5);
        var series = Pointed(Series(Red, 5, 10, 20) with { Marker = true, Auto = auto }, (0, new ChartPointMarker(Style: new ChartMarkerStyle(Line: Color.White))));
        var paths = Paths(Page(Chart(ChartKind.Line, series))).ToList();
        Assert.Equal(6, paths.Count);
        Assert.Equal(3, paths.Count(p => p.Fill == Blue)); // automatic fill, not the series color
        Assert.Single(paths, p => p.Stroke == Color.White);
        Assert.Equal(2, paths.Count(p => p.Stroke == Color.Black && p.LineWidth == 0.5));
    }

    [Fact]
    public void A_marker_with_no_fill_and_no_line_draws_nothing()
    {
        var page = Styled(new ChartMarkerStyle(NoFill: true, NoLine: true));
        Assert.Empty(Paths(page));
        Assert.DoesNotContain(page.Items.OfType<EllipseItem>(), e => e.Fill == Red);
        Assert.DoesNotContain(Styled(new ChartMarkerStyle(NoFill: true, NoLine: true), ChartMarkerSymbol.X).Items.OfType<LineItem>(), l => l.Width == 1.5);
    }

    [Fact]
    public void An_open_marker_takes_the_outline_color_and_width_and_ignores_a_fill_of_its_own_when_there_is_an_outline()
    {
        var strokes = Styled(new ChartMarkerStyle(Fill: Blue, Line: Color.White, LineWidth: 3), ChartMarkerSymbol.X).Items.OfType<LineItem>().Where(l => l.Color == Color.White && l.Width == 3);
        Assert.Equal(6, strokes.Count());
    }

    [Fact]
    public void A_marker_alpha_wraps_the_shape_in_an_opacity_group()
    {
        var page = Styled(new ChartMarkerStyle(Fill: Blue, FillAlpha: 0.5));
        Assert.Equal(3, page.Items.OfType<PushItem>().Count(p => p.Opacity == 0.5));
    }

    [Fact]
    public void A_legend_entry_uses_the_marker_style()
    {
        var series = Series(Red, 5, 10, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Circle, MarkerStyle = new ChartMarkerStyle(NoFill: true, Line: Blue, LineWidth: 2) };
        var page = Page(Chart(ChartKind.Line, series) with { Legend = LegendPosition.Right });
        Assert.Equal(4, Paths(page).Count(p => p.Fill is null && p.Stroke == Blue && p.LineWidth == 2));
    }

    [Fact]
    public void A_scatter_marker_style_is_drawn_too()
    {
        var series = Series(Red, 5, 10, 20) with { Marker = true, XValues = new double?[] { 1, 2, 3 }, MarkerStyle = new ChartMarkerStyle(NoFill: true, Line: Blue) };
        var chart = Chart(ChartKind.Scatter, series) with { XMin = 0, XMax = 4 };
        Assert.Equal(3, Paths(Page(chart)).Count(p => p.Stroke == Blue));
    }

    [Fact]
    public void A_pie_has_one_slice_per_value_and_a_legend_lists_the_categories()
    {
        var pie = Chart(ChartKind.Pie, new ChartSeries("S", [1, 1, 2], Red, PointColors: [Red, Blue, Color.Black])) with { Legend = LegendPosition.Right, Title = "Share" };
        var page = Page(pie);
        Assert.Equal(3, page.Items.OfType<PolygonItem>().Count());
        var text = Text(page);
        Assert.Contains("Share", text);
        Assert.Contains("abc", text.Replace(" ", ""));
    }

    [Fact]
    public void Title_axis_titles_and_data_labels_are_drawn_and_everything_is_clipped_to_the_frame()
    {
        var chart = Chart(ChartKind.Column, new ChartSeries("S", [5, 10, 20], Red, Labels: ["L5", "L10", "L20"])) with { Title = "Heading", CategoryAxisTitle = "Cats", ValueAxisTitle = "Vals" };
        var page = Page(chart);
        var text = Text(page);
        foreach (var part in new[] { "Heading", "Cats", "Vals", "L5", "L20" }) Assert.Contains(part, text);
        Assert.Contains(page.Items.OfType<PushItem>(), p => p.Clip is { Width: 300, Height: 200 });
    }

    [Fact]
    public void Charts_extend_the_used_range_and_follow_the_z_order()
    {
        var sheet = new Worksheet("S");
        sheet.Charts.Add(Chart(ChartKind.Column, Series(Red, 1, 2, 3)));
        Assert.Single(SheetLayoutTests.Layout(sheet).Pages);
        Assert.NotEmpty(SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<RectItem>());
    }

    [Fact]
    public void An_area_fills_the_polygon_between_its_line_and_the_baseline()
    {
        var chart = Chart(ChartKind.Area, Series(Red, 10, 20, 10)) with { CategoriesBetween = false };
        var polygon = Page(chart).Items.OfType<PolygonItem>().Single(p => p.Fill == Red);
        Assert.Equal(6, polygon.Points.Count);
        double minX = polygon.Points.Min(p => p.X), maxX = polygon.Points.Max(p => p.X);
        Assert.True(maxX - minX > 200); // edge to edge, not centered in bands
        Assert.Equal(polygon.Points.Max(p => p.Y), polygon.Points[^1].Y, 3); // the baseline closes it
    }

    [Fact]
    public void A_doughnut_leaves_a_hole_and_rings_follow_the_series()
    {
        var chart = Chart(ChartKind.Doughnut, new ChartSeries("A", new double?[] { 1, 1 }, Red, PointColors: [Red, Red]), new ChartSeries("B", new double?[] { 1, 3 }, Blue, PointColors: [Blue, Blue])) with { HoleSize = 0.5 };
        var items = Page(chart).Items;
        Assert.Equal(2, items.OfType<PolygonItem>().Count(p => p.Fill == Red));
        Assert.Equal(2, items.OfType<PolygonItem>().Count(p => p.Fill == Blue));
        var all = items.OfType<PolygonItem>().ToList();
        Assert.Equal(4, all.Count);
        double cx = all.SelectMany(p => p.Points).Average(p => p.X), cy = all.SelectMany(p => p.Points).Average(p => p.Y);
        double minRadius = all.SelectMany(p => p.Points).Min(p => Math.Sqrt((p.X - cx) * (p.X - cx) + (p.Y - cy) * (p.Y - cy)));
        Assert.True(minRadius > 10);
    }

    [Fact]
    public void Scatter_points_sit_at_their_x_and_y_values()
    {
        var series = new ChartSeries("S", new double?[] { 0, 20 }, Red, Marker: true, XValues: new double?[] { 0, 10 }, Line: false);
        var chart = Chart(ChartKind.Scatter, series) with { XTicks = [new(0, "0"), new(10, "10")], XMin = 0, XMax = 10 };
        var page = Page(chart);
        var dots = Filled(page, Red).OrderBy(e => e.X).ToList();
        Assert.Equal(2, dots.Count);
        Assert.True(dots[1].X > dots[0].X + 100);
        Assert.True(dots[1].Y < dots[0].Y - 50);
        Assert.DoesNotContain(page.Items.OfType<LineItem>(), l => l.Color == Red);
    }

    [Fact]
    public void A_scatter_group_on_secondary_axes_is_scaled_by_them()
    {
        var first = new ChartSeries("S", new double?[] { 10 }, Red, Marker: true, XValues: new double?[] { 5 }, Line: false);
        var other = new ChartSeries("T", new double?[] { 10 }, Blue, Marker: true, XValues: new double?[] { 5 }, Line: false);
        var chart = Chart(ChartKind.Scatter, first) with
        {
            XTicks = [new(0, "0"), new(10, "10")], XMin = 0, XMax = 10, ValueMin = 0, ValueMax = 20,
            ExtraGroups = [new ChartGroup(ChartKind.Scatter, ChartGrouping.Clustered, [other], true)],
            ShowSecondaryAxis = true, SecondaryTicks = [new(0, "0"), new(100, "100")], SecondaryMin = 0, SecondaryMax = 100, SecondaryXMin = 0, SecondaryXMax = 20,
        };
        var page = Page(chart);
        var a = Filled(page, Red).Single();
        var b = Filled(page, Blue).Single();
        Assert.True(a.X > b.X + 20);   // x = 5 of 20 sits left of x = 5 of 10
        Assert.True(b.Y > a.Y + 40);   // y = 10 of 100 sits below y = 10 of 20
        Assert.Contains("100", Text(page)); // the secondary scale is labelled on the right
    }

    [Fact]
    public void Bubble_area_follows_the_size()
    {
        var series = new ChartSeries("S", new double?[] { 5, 15 }, Red, XValues: new double?[] { 2, 8 }, Sizes: new double?[] { 1, 4 });
        var chart = Chart(ChartKind.Bubble, series) with { XTicks = [new(0, "0"), new(10, "10")], XMin = 0, XMax = 10 };
        var bubbles = Page(chart).Items.OfType<EllipseItem>().Where(e => e.Fill == Red).OrderBy(e => e.Width).ToList();
        Assert.Equal(2, bubbles.Count);
        Assert.Equal(2, bubbles[1].Width / bubbles[0].Width, 3);
    }

    [Fact]
    public void A_radar_series_is_a_closed_polygon_with_a_vertex_per_category()
    {
        var chart = Chart(ChartKind.Radar, Series(Red, 20, 20, 20)) with { Filled = true };
        var polygon = Page(chart).Items.OfType<PolygonItem>().Single();
        Assert.Equal(3, polygon.Points.Count);
        Assert.Equal(3, Page(chart).Items.OfType<LineItem>().Count(l => l.Color == Red));
    }

    [Fact]
    public void Axis_scales_map_values_linearly_logarithmically_and_backwards()
    {
        Assert.Equal(0.25, new AxisScale(0, 20).Fraction(5), 6);
        Assert.Equal(0.75, new AxisScale(0, 20, Reversed: true).Fraction(5), 6);
        Assert.Equal(0.5, new AxisScale(1, 100, LogBase: 10).Fraction(10), 6);
        Assert.Equal(0, new AxisScale(1, 100, LogBase: 10).Fraction(0), 6);
        Assert.Equal(1, new AxisScale(0, 20).Fraction(50), 6);
    }

    [Fact]
    public void A_smoothed_line_is_drawn_as_many_short_segments()
    {
        var straight = Page(Chart(ChartKind.Line, Series(Red, 5, 20, 5))).Items.OfType<LineItem>().Count(l => l.Color == Red && l.Width > 2);
        var smooth = Page(Chart(ChartKind.Line, new ChartSeries("S", new double?[] { 5, 20, 5 }, Red, Smooth: true))).Items.OfType<LineItem>().Count(l => l.Color == Red && l.Width > 2);
        Assert.Equal(2, straight);
        Assert.Equal(24, smooth);
    }

    [Fact]
    public void Reversed_categories_run_from_the_right()
    {
        var normal = Page(Chart(ChartKind.Column, Series(Red, 5, 10, 20))).Items.OfType<RectItem>().Where(r => r.Fill == Red).OrderBy(r => r.X).Select(r => r.Height).ToList();
        var reversed = Page(Chart(ChartKind.Column, Series(Red, 5, 10, 20)) with { ReverseCategories = true }).Items.OfType<RectItem>().Where(r => r.Fill == Red).OrderBy(r => r.X).Select(r => r.Height).ToList();
        Assert.Equal(normal.AsEnumerable().Reverse(), reversed);
    }

    [Fact]
    public void A_reversed_value_axis_hangs_the_columns_from_the_top()
    {
        var reversed = Page(Chart(ChartKind.Column, Series(Red, 5)) with { ReverseValues = true }).Items.OfType<RectItem>().Single(r => r.Fill == Red);
        var normal = Page(Chart(ChartKind.Column, Series(Red, 5))).Items.OfType<RectItem>().Single(r => r.Fill == Red);
        Assert.Equal(normal.Height, reversed.Height, 3);
        Assert.True(reversed.Y < normal.Y);
    }

    [Fact]
    public void A_secondary_axis_draws_its_labels_on_the_right_and_scales_its_group_separately()
    {
        var line = new ChartGroup(ChartKind.Line, ChartGrouping.Clustered, [new ChartSeries("L", new double?[] { 1, 2, 1 }, Blue, Marker: true)], true);
        var chart = Chart(ChartKind.Column, Series(Red, 10, 10, 10)) with
        {
            ExtraGroups = [line], ShowSecondaryAxis = true, SecondaryTicks = [new(0, "0%"), new(2, "200%")], SecondaryMin = 0, SecondaryMax = 2,
        };
        var page = Page(chart);
        var labels = page.Items.OfType<TextItem>().Where(t => t.Text is "0%" or "200%").ToList();
        Assert.Equal(2, labels.Count);
        double plotMiddle = page.Items.OfType<RectItem>().Where(r => r.Fill == Red).Average(r => r.X);
        Assert.All(labels, t => Assert.True(t.X > plotMiddle));
        var dots = Filled(page, Blue);
        Assert.Equal(3, dots.Count);
        var bars = page.Items.OfType<RectItem>().Where(r => r.Fill == Red).ToList();
        Assert.True(dots[1].Y < bars[0].Y); // 2 of 2 is the top, 10 of 20 only the middle
    }

    [Fact]
    public void A_translucent_series_is_wrapped_in_an_opacity_group()
    {
        var page = Page(Chart(ChartKind.Column, Series(Red, 5, 10, 20) with { Alpha = 0.4 }));
        var items = page.Items.ToList();
        int pushes = items.OfType<PushItem>().Count(p => p.Opacity == 0.4);
        Assert.Equal(3, pushes);
        int at = items.FindIndex(i => i is PushItem { Opacity: 0.4 });
        Assert.IsType<RectItem>(items[at + 1]);
        Assert.IsType<PopItem>(items[at + 2]);
    }

    [Fact]
    public void Translucent_bubbles_lose_their_white_halo()
    {
        var series = new ChartSeries("B", new double?[] { 5 }, Red, XValues: new double?[] { 5 }, Sizes: new double?[] { 3 }, Alpha: 0.5);
        var chart = Chart(ChartKind.Bubble, series) with { XMin = 0, XMax = 10, XTicks = [new(0, "0"), new(10, "10")] };
        var page = Page(chart);
        Assert.DoesNotContain(page.Items.OfType<EllipseItem>(), e => e.Fill == Color.White);
        Assert.Contains(page.Items, i => i is PushItem { Opacity: 0.5 });
    }

    [Fact]
    public void A_linear_trendline_is_a_dashed_straight_line_through_the_fit()
    {
        var series = Series(Red, 10, 20, 30) with { Trendline = new ChartTrendline(TrendlineKind.Linear, Blue, 2, true) };
        var lines = Page(Chart(ChartKind.Line, series)).Items.OfType<LineItem>().Where(l => l.Color == Blue).ToList();
        var line = Assert.Single(lines);
        Assert.Equal(LineStyle.Dashed, line.Style);
        Assert.True(line.Y2 < line.Y1); // rising values rise on the page
    }

    [Theory]
    [InlineData(TrendlineKind.Linear, 1, 2, 3, 4)]
    [InlineData(TrendlineKind.Exponential, 2, 4, 8, 16)]
    [InlineData(TrendlineKind.Power, 1, 4, 9, 16)]
    [InlineData(TrendlineKind.Logarithmic, 1, 2, 2.585, 3)]
    [InlineData(TrendlineKind.Polynomial, 1, 4, 9, 16)]
    public void Fits_reproduce_data_that_follows_their_form(TrendlineKind kind, double y1, double y2, double y3, double y4)
    {
        var data = new List<(double X, double Y)> { (1, y1), (2, y2), (3, y3), (4, y4) };
        var curve = ChartTrendlines.Fit(new ChartTrendline(kind, Red, Order: 2), data, 1, 4);
        Assert.NotEmpty(curve);
        foreach (var (x, y) in data)
        {
            var nearest = curve.OrderBy(c => Math.Abs(c.X - x)).First();
            if (Math.Abs(nearest.X - x) < 1e-9) Assert.Equal(y, nearest.Y, 1);
        }
        Assert.Equal(y1, curve[0].Y, 1);
        Assert.Equal(y4, curve[^1].Y, 1);
    }

    [Fact]
    public void Trendlines_extend_forward_and_moving_averages_follow_the_period()
    {
        var data = new List<(double X, double Y)> { (1, 1), (2, 2), (3, 3), (4, 4) };
        var forward = ChartTrendlines.Fit(new ChartTrendline(TrendlineKind.Linear, Red, Forward: 2, Backward: 0.5), data, 1, 4);
        Assert.Equal((0.5, 0.5), (forward[0].X, Math.Round(forward[0].Y, 6)));
        Assert.Equal((6.0, 6.0), (forward[^1].X, Math.Round(forward[^1].Y, 6)));
        var average = ChartTrendlines.Fit(new ChartTrendline(TrendlineKind.MovingAverage, Red, Period: 2), data, 1, 4);
        Assert.Equal(new[] { 1.5, 2.5, 3.5 }, average.Select(p => p.Y));
        Assert.Empty(ChartTrendlines.Fit(new ChartTrendline(TrendlineKind.Linear, Red), [(1, 1)], 1, 1));
        Assert.Empty(ChartTrendlines.Fit(new ChartTrendline(TrendlineKind.Logarithmic, Red), [(-1, 1), (-2, 2)], -2, -1));
    }

    [Fact]
    public void Error_bars_are_a_line_with_end_caps_around_each_value()
    {
        var bars = new ChartErrorBars(false, [2, 2, 2], [2, 2, 2], true, Blue, 1);
        var lines = Page(Chart(ChartKind.Column, Series(Red, 5, 10, 20) with { ErrorBars = [bars] })).Items.OfType<LineItem>().Where(l => l.Color == Blue).ToList();
        Assert.Equal(9, lines.Count); // per point: the bar and two caps
        var vertical = lines.Where(l => l.X1 == l.X2).ToList();
        Assert.Equal(3, vertical.Count);
        Assert.All(vertical, l => Assert.True(Math.Abs(l.Y2 - l.Y1) > 5));
        var noCaps = Page(Chart(ChartKind.Column, Series(Red, 5, 10, 20) with { ErrorBars = [bars with { EndCap = false }] })).Items.OfType<LineItem>().Count(l => l.Color == Blue);
        Assert.Equal(3, noCaps);
    }

    [Fact]
    public void An_inner_manual_plot_layout_fixes_the_plot_rectangle()
    {
        var chart = Chart(ChartKind.Column, Series(Red, 20, 20, 20)) with { PlotArea = new ChartBox(0.2, 0.1, 0.6, 0.5, true), Border = null, Background = null };
        var bars = Page(chart).Items.OfType<RectItem>().Where(r => r.Fill == Red).ToList();
        Assert.Equal(OriginY + 20, bars.Min(b => b.Y), 3); // 10 % of 200
        Assert.Equal(OriginY + 120, bars.Max(b => b.Y + b.Height), 3); // 10 % + 50 % of the height is the zero line
        Assert.InRange(bars.Min(b => b.X), OriginX + 60, OriginX + 120);
        Assert.InRange(bars.Max(b => b.X + b.Width), OriginX + 120, OriginX + 240);
    }

    [Fact]
    public void An_outer_manual_layout_confines_the_axes_and_a_manual_title_sits_where_it_is_put()
    {
        var chart = Chart(ChartKind.Column, Series(Red, 20, 20, 20)) with
        {
            PlotArea = new ChartBox(0.5, 0.5, 0.4, 0.4, false), Title = "Hi", TitlePosition = (0.1, 0.05), Border = null, Background = null,
        };
        var page = Page(chart);
        Assert.All(page.Items.OfType<RectItem>().Where(r => r.Fill == Red), b => Assert.True(b.X >= OriginX + 150 - 0.01 && b.Y >= OriginY + 100 - 0.01));
        var title = page.Items.OfType<TextItem>().First(t => t.Text == "Hi");
        Assert.Equal(OriginX + 30, title.X, 1);
    }

    [Fact]
    public void A_reversed_value_axis_moves_the_category_labels_to_the_top()
    {
        double LabelY(bool reversed) => Page(Chart(ChartKind.Column, Series(Red, 5, 10, 20)) with { ReverseValues = reversed })
            .Items.OfType<TextItem>().First(t => t.Text == "b").Baseline;
        Assert.True(LabelY(false) > OriginY + 150);
        Assert.True(LabelY(true) < OriginY + 40);
    }

    // ------------------------------------------------------------------ trendlines on bars and stacks, labels, legend, text styles

    private static IEnumerable<LineItem> TrendLines(LayoutPage page) => page.Items.OfType<LineItem>().Where(l => l.Color == Blue);

    [Fact]
    public void Trendlines_are_drawn_on_horizontal_bars_along_the_value_axis()
    {
        var series = Series(Red, 5, 10, 20) with { Trendline = new ChartTrendline(TrendlineKind.Linear, Blue, 2, true) };
        var line = Assert.Single(TrendLines(Page(Chart(ChartKind.Bar, series))));
        // The first category is at the bottom, so values grow to the right while the line climbs up the page.
        Assert.True(line.X2 > line.X1);
        Assert.True(line.Y2 < line.Y1);
    }

    [Fact]
    public void Trendlines_on_stacked_groups_follow_the_stacked_totals()
    {
        var first = Series(Red, 5, 5, 5);
        var second = Series(Color.White, 5, 10, 15) with { Trendline = new ChartTrendline(TrendlineKind.Linear, Blue, 2, true) };
        var stacked = Chart(ChartKind.Column, first, second) with { Grouping = ChartGrouping.Stacked };
        var line = Assert.Single(TrendLines(Page(stacked)));
        var alone = Assert.Single(TrendLines(Page(Chart(ChartKind.Column, second))));
        // Totals are 10, 15, 20: the fitted line starts at 10 (the second tick) rather than at 5.
        Assert.True(line.Y1 > line.Y2);
        Assert.True(line.Y1 < alone.Y1);
    }

    [Fact]
    public void The_equation_and_r_squared_of_a_trendline_are_printed_when_asked_for()
    {
        var trend = new ChartTrendline(TrendlineKind.Linear, Blue, 2, true, ShowEquation: true, ShowRSquared: true);
        var withLabel = Text(Page(Chart(ChartKind.Line, Series(Red, 10, 20, 30) with { Trendline = trend })));
        Assert.Contains("y = 10x", withLabel);
        Assert.Contains("R² = 1", withLabel);
        var without = Text(Page(Chart(ChartKind.Line, Series(Red, 10, 20, 30) with { Trendline = trend with { ShowEquation = false, ShowRSquared = false } })));
        Assert.DoesNotContain("R²", without);
        var onlyR = Text(Page(Chart(ChartKind.Line, Series(Red, 10, 20, 30) with { Trendline = trend with { ShowEquation = false } })));
        Assert.DoesNotContain("y =", onlyR);
        Assert.Contains("R²", onlyR);
    }

    [Theory]
    [InlineData(TrendlineKind.Linear, "y = 2x + 1")]
    [InlineData(TrendlineKind.Polynomial, "y = x² + 1")]
    [InlineData(TrendlineKind.Exponential, "y = 2e^0.69315x")]
    [InlineData(TrendlineKind.Power, "y = 3x^2")]
    [InlineData(TrendlineKind.Logarithmic, "y = 2ln(x) + 1")]
    public void Equations_are_fitted_and_written_like_excel(TrendlineKind kind, string expected)
    {
        Func<double, double> f = kind switch
        {
            TrendlineKind.Linear => x => 2 * x + 1,
            TrendlineKind.Polynomial => x => x * x + 1,
            TrendlineKind.Exponential => x => 2 * Math.Exp(Math.Log(2) * x),
            TrendlineKind.Power => x => 3 * x * x,
            _ => x => 2 * Math.Log(x) + 1,
        };
        var data = Enumerable.Range(1, 6).Select(x => ((double)x, f(x))).ToList();
        var (equation, r2) = ChartTrendlines.Describe(new ChartTrendline(kind, Red, Order: 2), data)!.Value;
        Assert.Equal(expected, equation);
        Assert.Equal(1, r2, 6);
    }

    [Fact]
    public void R_squared_measures_the_scatter_around_the_fit_and_a_moving_average_has_no_equation()
    {
        var noisy = new List<(double X, double Y)> { (1, 1), (2, 3), (3, 2), (4, 5), (5, 4) };
        var (equation, r2) = ChartTrendlines.Describe(new ChartTrendline(TrendlineKind.Linear, Red), noisy)!.Value;
        Assert.Equal("y = 0.8x + 0.6", equation);
        Assert.Equal(0.64, r2, 4);
        Assert.Null(ChartTrendlines.Describe(new ChartTrendline(TrendlineKind.MovingAverage, Red), noisy));
        Assert.Null(ChartTrendlines.Describe(new ChartTrendline(TrendlineKind.Linear, Red), [(1, 1)]));
        Assert.Equal("y = -0.5x - 2", ChartTrendlines.Describe(new ChartTrendline(TrendlineKind.Linear, Red), [(0, -2), (2, -3), (4, -4)])!.Value.Equation);
    }

    [Fact]
    public void A_manual_legend_sits_at_its_position_and_the_plot_gives_way_to_its_size_unless_it_overlays()
    {
        var chart = Chart(ChartKind.Column, Series(Red, 5, 10, 20)) with { Legend = LegendPosition.Right, Border = null, Background = null };
        var automatic = Page(chart);
        var box = new ChartBox(0.1, 0.8, 0.3, 0.1, false);
        var manual = Page(chart with { LegendBox = box });
        var swatch = manual.Items.OfType<RectItem>().Where(r => r.Fill == Red && r.Width < 10).Single();
        Assert.Equal(OriginX + 30, swatch.X, 3);
        Assert.InRange(swatch.Y, OriginY + 160, OriginY + 175);
        double Width(LayoutPage page) => page.Items.OfType<RectItem>().Where(r => r.Fill == Red && r.Height > 10).Max(r => r.X + r.Width);
        var overlay = Page(chart with { LegendBox = box, LegendOverlay = true });
        // The manual legend's own width (a third of the chart) is wider than the automatic one, so the plot is narrower; overlaying it keeps the plot whole.
        Assert.True(Width(manual) < Width(automatic));
        Assert.True(Width(overlay) > Width(automatic));
        // A manual legend above the plot pushes the plot down by its height.
        double Top(LayoutPage page) => page.Items.OfType<RectItem>().Where(r => r.Fill == Red && r.Height > 10).Min(r => r.Y);
        var top = chart with { Legend = LegendPosition.Top, LegendBox = new ChartBox(0.1, 0.02, 0.5, 0.2, false) };
        Assert.True(Top(Page(top)) > Top(Page(top with { LegendOverlay = true })));
    }

    [Fact]
    public void Text_styles_change_size_color_and_weight_of_the_title_axes_legend_and_labels()
    {
        var green = new Color(0, 160, 0);
        var chart = Chart(ChartKind.Column, Series(Red, 5, 10, 20) with { Labels = ["a1", "a2", "a3"] }) with
        {
            Title = "Title", Legend = LegendPosition.Right, ValueAxisTitle = "Vertical", CategoryAxisTitle = "Horizontal",
            TextStyles = new ChartTextStyles
            {
                Title = new ChartTextStyle(20, green, Bold: true), ValueAxis = new ChartTextStyle(Color: Blue), CategoryAxis = new ChartTextStyle(Size: 13),
                ValueAxisTitle = new ChartTextStyle(Color: green), Legend = new ChartTextStyle(Italic: true, Size: 12), DataLabels = new ChartTextStyle(Color: Red),
            },
        };
        var texts = Page(chart).Items.OfType<TextItem>().ToList();
        var title = texts.First(t => t.Text == "Title");
        Assert.Equal((20, green), (title.FontSize, title.Color));
        Assert.True(title.SyntheticBold || title.Font.PostScriptName.Contains("Bold", StringComparison.OrdinalIgnoreCase));
        Assert.All(texts.Where(t => t.Text is "0" or "10" or "20"), t => Assert.Equal(Blue, t.Color));
        Assert.Equal(13, texts.First(t => t.Text == "a").FontSize);
        Assert.Equal(green, texts.First(t => t.Text == "Vertical").Color);
        Assert.NotEqual(green, texts.First(t => t.Text == "Horizontal").Color);
        var legend = texts.First(t => t.Text == "S");
        Assert.Equal(12, legend.FontSize);
        Assert.True(legend.SyntheticItalic || legend.Font.PostScriptName.Contains("Italic", StringComparison.OrdinalIgnoreCase));
        Assert.All(texts.Where(t => t.Text is "a1" or "a2" or "a3"), t => Assert.Equal(Red, t.Color));
    }

    [Fact]
    public void A_trendline_number_format_applies_to_the_equation_and_a_manual_position_moves_the_label()
    {
        var noisy = new List<(double X, double Y)> { (1, 1), (2, 3), (3, 2), (4, 5), (5, 4) };
        var format = new ChartTrendline(TrendlineKind.Linear, Red, NumberFormatter: v => v.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal("y = 0.80x + 0.60", ChartTrendlines.Describe(format, noisy)!.Value.Equation);

        SheetChart WithLabel(ChartLabelOffset? at) => Chart(ChartKind.Line, Series(Red, 10, 20, 30) with
        {
            Trendline = new ChartTrendline(TrendlineKind.Linear, Blue, ShowEquation: true, NumberFormatter: v => v.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture), LabelPosition: at),
        });
        TextItem Label(SheetChart c) => Page(c).Items.OfType<TextItem>().Single(t => t.Text == "y"); // the first word of the equation
        var automatic = Label(WithLabel(null));
        Assert.Contains("y = 10.0x", Text(Page(WithLabel(null))));
        var fromEdge = Label(WithLabel(new ChartLabelOffset(0.1, 0.5, true)));
        Assert.Equal(OriginX + 30, fromEdge.X, 3);                       // 10 % of the 300-pt chart from its left edge
        Assert.InRange(fromEdge.Baseline, OriginY + 100, OriginY + 120); // half the 200-pt height down
        var shifted = Label(WithLabel(new ChartLabelOffset(-0.1, 0.1, false)));
        Assert.Equal(automatic.X - 30, shifted.X, 3);
        Assert.Equal(automatic.Baseline + 20, shifted.Baseline, 3);
    }

    [Fact]
    public void Category_labels_turn_by_the_angle_the_axis_asks_for()
    {
        SheetChart WithRotation(double? angle) => Chart(ChartKind.Column, Series(Red, 5, 10, 20)) with
        {
            TextStyles = new ChartTextStyles { CategoryAxis = angle is null ? null : new ChartTextStyle(Rotation: angle) },
        };
        Assert.DoesNotContain(Page(WithRotation(null)).Items.OfType<PushItem>(), p => p.Rotation != 0);
        Assert.DoesNotContain(Page(WithRotation(0)).Items.OfType<PushItem>(), p => p.Rotation != 0);
        var up = Page(WithRotation(90)).Items.OfType<PushItem>().Where(p => p.Rotation != 0).ToList();
        Assert.Equal(3, up.Count);
        Assert.All(up, p => Assert.Equal(90, p.Rotation));
        var down = Page(WithRotation(-45)).Items.OfType<PushItem>().Where(p => p.Rotation != 0).ToList();
        Assert.Equal(3, down.Count);
        Assert.All(down, p => Assert.Equal(-45, p.Rotation));
    }

    [Fact]
    public void Value_axis_labels_and_the_labels_of_horizontal_bars_turn_too()
    {
        static IEnumerable<PushItem> Turned(LayoutPage page) => page.Items.OfType<PushItem>().Where(p => p.Rotation != 0);
        var column = Chart(ChartKind.Column, Series(Red, 5, 10, 20));
        int ticks = column.Ticks.Count;
        var valueTurned = Turned(Page(column with { TextStyles = new ChartTextStyles { ValueAxis = new ChartTextStyle(Rotation: 90) } })).ToList();
        Assert.Equal(ticks, valueTurned.Count);
        Assert.All(valueTurned, p => Assert.Equal(90, p.Rotation));
        // The turned labels are centered on their gridline, so the plot is not narrower than the widest label allows.
        var bar = Chart(ChartKind.Bar, Series(Red, 5, 10, 20));
        var styles = new ChartTextStyles { CategoryAxis = new ChartTextStyle(Rotation: 45), ValueAxis = new ChartTextStyle(Rotation: -45) };
        var turned = Turned(Page(bar with { TextStyles = styles })).ToList();
        Assert.Equal(3 + bar.Ticks.Count, turned.Count);
        Assert.Equal(3, turned.Count(p => p.Rotation == 45));
        Assert.Equal(bar.Ticks.Count, turned.Count(p => p.Rotation == -45));
        Assert.Empty(Turned(Page(bar)));
    }

    [Fact]
    public void Turned_category_labels_of_a_reversed_value_axis_stay_above_the_plot()
    {
        var chart = Chart(ChartKind.Column, Series(Red, 5, 10, 20)) with
        {
            ReverseValues = true,
            TextStyles = new ChartTextStyles { CategoryAxis = new ChartTextStyle(Rotation: -30) },
        };
        var items = Page(chart).Items;
        var turned = items.OfType<PushItem>().Where(p => p.Rotation != 0).ToList();
        Assert.Equal(3, turned.Count);
        Assert.All(turned, p => Assert.Equal(-30, p.Rotation));
        double top = items.OfType<LineItem>().Where(l => l.X1 != l.X2).Min(l => l.Y1);
        Assert.All(turned, p => Assert.True(p.CenterY < top)); // each label is anchored above the first gridline
    }

    [Fact]
    public void A_scatter_chart_turns_its_x_labels_too()
    {
        var series = new ChartSeries("S", new double?[] { 0, 20 }, Red, Marker: true, XValues: new double?[] { 0, 10 }, Line: false);
        var chart = Chart(ChartKind.Scatter, series) with
        {
            XTicks = [new(0, "0"), new(10, "10")], XMin = 0, XMax = 10, TextStyles = new ChartTextStyles { CategoryAxis = new ChartTextStyle(Rotation: 90) },
        };
        Assert.Equal(2, Page(chart).Items.OfType<PushItem>().Count(p => p.Rotation == 90));
    }

    [Fact]
    public void A_manual_legend_size_makes_the_entries_flow_inside_it()
    {
        var a = Series(Red, 1) with { Name = "One" };
        var chart = Chart(ChartKind.Column, a, a with { Name = "Two", Color = Blue }, a with { Name = "Three", Color = Color.Black }, a with { Name = "Four", Color = new Color(0, 160, 0) }) with { Legend = LegendPosition.Right, Border = null, Background = null };
        TextItem Entry(LayoutPage page, string name) => page.Items.OfType<TextItem>().Single(t => t.Text == name);
        // Two entries' height: the four entries need two columns.
        var tall = Page(chart with { LegendBox = new ChartBox(0.5, 0.1, 0.4, 0.8, false) });
        Assert.Equal(Entry(tall, "One").X, Entry(tall, "Four").X, 3);
        var narrow = Page(chart with { LegendBox = new ChartBox(0.5, 0.1, 0.4, 0.2, false) });
        Assert.True(Entry(narrow, "Three").X > Entry(narrow, "One").X + 20);
        Assert.Equal(Entry(narrow, "One").X, Entry(narrow, "Two").X, 3);
        Assert.True(Entry(narrow, "Two").Baseline > Entry(narrow, "One").Baseline);
        // A horizontal legend wraps along rows instead.
        var wide = Page(chart with { Legend = LegendPosition.Bottom, LegendBox = new ChartBox(0.1, 0.5, 0.35, 0.5, false) });
        Assert.True(Entry(wide, "Three").Baseline > Entry(wide, "One").Baseline);
    }

    [Fact]
    public void A_right_legend_that_does_not_fit_starts_at_the_top_and_drops_the_last_entries()
    {
        var entries = Enumerable.Range(0, 60).Select(i => Series(Red, 5, 10, 20) with { Name = "idx" + i }).ToArray();
        var texts = Page(Chart(ChartKind.Line, entries) with { Legend = LegendPosition.Right, Border = null, Background = null }).Items.OfType<TextItem>().ToList();
        var shown = texts.Where(t => t.Text.StartsWith("idx")).ToList();
        Assert.Equal("idx0", shown[0].Text);
        Assert.InRange(shown.Count, 1, 59);
        Assert.Equal(Enumerable.Range(0, shown.Count).Select(i => "idx" + i), shown.Select(t => t.Text));
    }

    [Fact]
    public void A_scatter_series_with_no_line_and_symbol_none_draws_no_markers()
    {
        var none = Series(Red, 5, 10, 20) with { Line = false, MarkerNone = true, XValues = new double?[] { 1, 2, 3 } };
        Assert.DoesNotContain(Page(Chart(ChartKind.Scatter, none) with { XMin = 0, XMax = 4, Legend = LegendPosition.None }).Items, i => i is PathItem or PolygonItem or EllipseItem);
        var markers = none with { MarkerNone = false };
        Assert.Contains(Page(Chart(ChartKind.Scatter, markers) with { XMin = 0, XMax = 4, Legend = LegendPosition.None }).Items, i => i is PathItem or PolygonItem or EllipseItem);
    }

    [Fact]
    public void An_automatic_bottom_legend_that_does_not_fit_one_row_wraps_into_rows_inside_the_chart()
    {
        var entries = Enumerable.Range(0, 14).Select(i => Series(Red, 5, 10, 20) with { Name = "idx" + i }).ToArray();
        var chart = Chart(ChartKind.Line, entries) with { Legend = LegendPosition.Bottom, Border = null, Background = null, Width = 410 };
        var page = Page(chart);
        TextItem Entry(string name) => page.Items.OfType<TextItem>().Single(t => t.Text == name);
        // Two rows of seven, every entry inside the chart's width.
        Assert.Equal(Entry("idx0").Baseline, Entry("idx6").Baseline, 3);
        Assert.Equal(Entry("idx7").Baseline, Entry("idx13").Baseline, 3);
        Assert.True(Entry("idx7").Baseline > Entry("idx0").Baseline);
        Assert.Equal(Entry("idx0").X, Entry("idx7").X, 3);
        Assert.All(page.Items.OfType<TextItem>().Where(t => t.Text.StartsWith("idx")), t => Assert.InRange(t.X, OriginX, OriginX + 410));
        Assert.True(Entry("idx13").X + 25 <= OriginX + 410);
        // The plot gave way by the second row: its category labels sit higher than with a one-row legend.
        var one = Page(Chart(ChartKind.Line, entries[0], entries[1]) with { Legend = LegendPosition.Bottom, Border = null, Background = null });
        Assert.True(one.Items.OfType<TextItem>().Single(t => t.Text == "a").Baseline - page.Items.OfType<TextItem>().Single(t => t.Text == "a").Baseline > 10);
    }

    [Fact]
    public void A_legend_that_fits_one_row_stays_in_one_row()
    {
        var chart = Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Name = "One" }, Series(Blue, 5, 10, 20) with { Name = "Two" }) with { Legend = LegendPosition.Top };
        var page = Page(chart);
        Assert.Equal(page.Items.OfType<TextItem>().Single(t => t.Text == "One").Baseline, page.Items.OfType<TextItem>().Single(t => t.Text == "Two").Baseline, 3);
    }

    [Fact]
    public void A_points_own_label_style_wins_over_the_series_labels()
    {
        var green = new Color(0, 160, 0);
        var series = Series(Red, 5, 10, 20) with { Labels = ["a1", "", "a3"], LabelStyles = [null, null, new ChartTextStyle(Size: 16, Color: green)] };
        var texts = Page(Chart(ChartKind.Column, series)).Items.OfType<TextItem>().ToList();
        Assert.DoesNotContain(texts, t => t.Text == "");
        Assert.Equal(16, texts.Single(t => t.Text == "a3").FontSize);
        Assert.Equal(green, texts.Single(t => t.Text == "a3").Color);
        Assert.NotEqual(16, texts.Single(t => t.Text == "a1").FontSize);
    }

    [Fact]
    public void A_data_label_sits_where_its_position_puts_it()
    {
        double Baseline(ChartSeries series, ChartKind kind = ChartKind.Column) =>
            Page(Chart(kind, series)).Items.OfType<TextItem>().Single(t => t.Text == "lab").Baseline;
        double X(ChartSeries series, ChartKind kind) => Page(Chart(kind, series)).Items.OfType<TextItem>().Single(t => t.Text == "lab").X;
        ChartSeries With(ChartLabelPlacement? place) => new("S", new double?[] { 20 }, Red, false, null, new[] { "lab" }, LabelPlacements: new ChartLabelPlacement?[] { place });
        double outside = Baseline(With(ChartLabelPlacement.OutsideEnd)), center = Baseline(With(ChartLabelPlacement.Center));
        double inEnd = Baseline(With(ChartLabelPlacement.InsideEnd)), inBase = Baseline(With(ChartLabelPlacement.InsideBase));
        Assert.True(outside < inEnd && inEnd < center && center < inBase, $"{outside} {inEnd} {center} {inBase}");
        Assert.Equal(outside, Baseline(With(null)), 3); // no position: the usual outside end of a clustered column
        // Horizontal bars: inside the end lies left of the outside end, the base further left still.
        double barOutside = X(With(ChartLabelPlacement.OutsideEnd), ChartKind.Bar), barInEnd = X(With(ChartLabelPlacement.InsideEnd), ChartKind.Bar);
        double barBase = X(With(ChartLabelPlacement.InsideBase), ChartKind.Bar);
        Assert.True(barBase < barInEnd && barInEnd < barOutside);
    }

    [Fact]
    public void Points_of_a_line_chart_put_their_label_on_the_requested_side()
    {
        ChartSeries With(ChartLabelPlacement? place) => new("S", new double?[] { 10, 20 }, Red, true, null, new[] { "", "lab" }, LabelPlacements: new ChartLabelPlacement?[] { null, place });
        TextItem Label(ChartLabelPlacement? place) => Page(Chart(ChartKind.Line, With(place))).Items.OfType<TextItem>().Single(t => t.Text == "lab");
        var above = Label(null);
        Assert.True(Label(ChartLabelPlacement.Below).Baseline > Label(ChartLabelPlacement.Center).Baseline);
        Assert.True(Label(ChartLabelPlacement.Center).Baseline > above.Baseline);
        Assert.True(Label(ChartLabelPlacement.Left).X < above.X && above.X < Label(ChartLabelPlacement.Right).X);
    }

    [Fact]
    public void A_scatter_chart_titles_its_secondary_horizontal_axis_above_the_plot()
    {
        var primary = new ChartSeries("A", new double?[] { 1, 2 }, Red, Marker: true, XValues: new double?[] { 0, 10 }, Line: false);
        var other = new ChartSeries("B", new double?[] { 1, 2 }, Red, Marker: true, XValues: new double?[] { 0, 5 }, Line: false);
        var chart = Chart(ChartKind.Scatter, primary) with
        {
            XTicks = [new(0, "0"), new(10, "10")], XMin = 0, XMax = 10,
            ExtraGroups = [new ChartGroup(ChartKind.Scatter, ChartGrouping.Clustered, [other], true, 150, 0)],
            ShowSecondaryXAxis = true, SecondaryXTicks = [new(0, "0"), new(5, "5")], SecondaryXMin = 0, SecondaryXMax = 5, SecondaryMin = 0, SecondaryMax = 2,
            SecondaryXAxisTitle = "Top axis",
        };
        double Tick(LayoutPage page) => page.Items.OfType<TextItem>().First(t => t.Text == "5").Baseline;
        var titled = Page(chart);
        var title = titled.Items.OfType<TextItem>().Single(t => t.Text.Contains("Top"));
        Assert.True(title.Baseline < Tick(titled)); // the title sits above the axis' labels
        Assert.True(Tick(titled) > Tick(Page(chart with { SecondaryXAxisTitle = null }))); // and takes room from the plot
    }

    private static double Distance(TextItem t, double width) => Math.Sqrt(Math.Pow(t.X + width / 2 - (OriginX + 150), 2) + Math.Pow(t.Baseline - (OriginY + 100), 2));

    [Theory]
    [InlineData(ChartKind.Column)] [InlineData(ChartKind.Bar)]
    public void Data_labels_on_column_and_bar_fills_are_black_inside_or_out_unless_the_chart_states_a_color(ChartKind kind)
    {
        var series = Series(Red, 5, 10, 20) with { Labels = ["L5", "L10", "L20"] };
        var stacked = Series(Blue, 5, 10, 20) with { Labels = ["M5", "M10", "M20"] };
        var page = Page(Chart(kind, series, stacked) with { Grouping = ChartGrouping.Stacked });
        Assert.All(page.Items.OfType<TextItem>().Where(t => t.Text.StartsWith('L') || t.Text.StartsWith('M')), t => Assert.Equal(Color.Black, t.Color)); // inside, over red and blue
        var outside = Page(Chart(kind, series));
        Assert.All(outside.Items.OfType<TextItem>().Where(t => t.Text.StartsWith('L')), t => Assert.Equal(Color.Black, t.Color));
        var stated = Page(Chart(kind, series, stacked) with { Grouping = ChartGrouping.Stacked, TextColor = Red });
        Assert.All(stated.Items.OfType<TextItem>().Where(t => t.Text.StartsWith('L')), t => Assert.Equal(Red, t.Color));
        var styled = Page(Chart(kind, series, stacked) with { Grouping = ChartGrouping.Stacked, TextStyles = new ChartTextStyles { DataLabels = new ChartTextStyle(Color: Blue) } });
        Assert.All(styled.Items.OfType<TextItem>().Where(t => t.Text.StartsWith('L')), t => Assert.Equal(Blue, t.Color));
    }

    [Fact]
    public void Pie_labels_inside_the_slices_keep_the_contrast_color()
    {
        var pie = new ChartSeries("S", [1, 1, 2], Red, PointColors: [Red, Color.White, Color.Black], Labels: ["A", "B", "C"]);
        var items = Page(Chart(ChartKind.Pie, pie)).Items.OfType<TextItem>().ToList();
        Assert.Equal(Color.White, items.Single(t => t.Text == "A").Color);        // over red
        Assert.Equal(new Color(0x40, 0x40, 0x40), items.Single(t => t.Text == "B").Color); // over white
        Assert.Equal(Color.White, items.Single(t => t.Text == "C").Color);        // over black
    }

    [Fact]
    public void Error_bars_without_a_color_are_drawn_black()
    {
        var bars = new ChartErrorBars(false, [2, 2, 2], [2, 2, 2], true, null, 0.75);
        var lines = Page(Chart(ChartKind.Column, Series(Red, 5, 10, 20) with { ErrorBars = [bars] })).Items.OfType<LineItem>().Where(l => l.Color == Color.Black && l.Width == 0.75).ToList();
        Assert.True(lines.Count >= 9);
    }

    [Fact]
    public void Pie_labels_sit_at_the_asked_position()
    {
        ChartSeries Pie(ChartLabelPlacement? place) => new("S", [1, 1, 2], Red, PointColors: [Red, Blue, Color.Black], Labels: ["A", "B", "C"],
            LabelPlacements: place is { } p ? [p, p, p] : null);
        TextItem Label(ChartLabelPlacement? place, string text) => Page(Chart(ChartKind.Pie, Pie(place))).Items.OfType<TextItem>().Single(t => t.Text == text);
        double cx = OriginX + 150, cy = OriginY + 100;
        // The first slice runs from 12 to 3 o'clock: its middle points up and to the right.
        var inside = Label(null, "A"); var center = Label(ChartLabelPlacement.Center, "A"); var inEnd = Label(ChartLabelPlacement.InsideEnd, "A"); var outside = Label(ChartLabelPlacement.OutsideEnd, "A");
        Assert.True(Distance(center, 5) < Distance(inside, 5) && Distance(inside, 5) < Distance(inEnd, 5));
        Assert.True(outside.X > cx + 30 && outside.Baseline < cy - 30 && Distance(outside, 5) > Distance(inEnd, 5) + 5);  // beyond the rim, on the right
        // On the left the label ends at the rim instead of starting there.
        var left = Label(ChartLabelPlacement.OutsideEnd, "C");   // the last slice (half of the pie) points to the lower left
        Assert.True(left.X < cx - 20);
        // Outside labels are in the plain text color, not the slice's contrast color.
        Assert.NotEqual(Label(null, "C").Color, left.Color);
    }

    [Fact]
    public void A_best_fit_label_stays_inside_a_wide_slice_and_moves_out_of_a_narrow_one()
    {
        var series = new ChartSeries("S", [90, 5, 5], Red, PointColors: [Red, Blue, Color.Black], Labels: ["Big", "TinyLongLabelHere", "T2"],
            LabelPlacements: [ChartLabelPlacement.BestFit, ChartLabelPlacement.BestFit, ChartLabelPlacement.BestFit]);
        var items = Page(Chart(ChartKind.Pie, series)).Items.OfType<TextItem>().ToList();
        double cx = OriginX + 150, cy = OriginY + 100;
        double Reach(string text) { var t = items.Single(i => i.Text == text); return Math.Sqrt(Math.Pow(t.X - cx, 2) + Math.Pow(t.Baseline - cy, 2)); }
        Assert.True(Reach("Big") < 70);          // inside
        Assert.True(Reach("TinyLongLabelHere") > 90 || items.Single(i => i.Text == "TinyLongLabelHere").X + 10 < cx - 70);  // outside the rim
    }

    [Fact]
    public void A_label_with_several_lines_draws_every_line()
    {
        var page = Page(Chart(ChartKind.Column, new ChartSeries("S", [5, 10, 20], Red, Labels: ["Top\nBottom", "x", "y"])));
        var lines = page.Items.OfType<TextItem>().Where(t => t.Text is "Top" or "Bottom").OrderBy(t => t.Baseline).ToList();
        Assert.True(lines.Count == 2, string.Join("|", page.Items.OfType<TextItem>().Select(t => t.Text)));
        Assert.Equal(["Top", "Bottom"], lines.Select(t => t.Text));
        Assert.True(lines[1].Baseline - lines[0].Baseline > 5);
    }

    // ---- per-point markers (c:dPt/c:marker)

    private static ChartSeries Pointed(ChartSeries s, params (int Index, ChartPointMarker Marker)[] points) =>
        s with { PointMarkers = points.ToDictionary(p => p.Index, p => p.Marker) };

    private static double CenterX((double X, double Y, double Width, double Height) b) => b.X + b.Width / 2;

    [Fact]
    public void A_point_marker_changes_only_its_own_point_even_after_a_null_value()
    {
        var series = Pointed(Series(Red, null, 10, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Circle }, (2, new ChartPointMarker(Symbol: ChartMarkerSymbol.Square, Size: 9)));
        var page = Page(Chart(ChartKind.Line, series));
        var boxes = Filled(page, Red);
        Assert.Equal(2, boxes.Count);
        var rect = boxes.Single(b => b.Width == 9);
        var circle = boxes.Single(b => b.Width == 5);
        Assert.True(rect.X + 4.5 > CenterX(circle)); // idx 2 is the third category, to the right of idx 1 // idx 2 is the third category, to the right of idx 1
    }

    [Fact]
    public void A_point_with_no_symbol_draws_no_marker_and_the_others_are_untouched()
    {
        var plain = Page(Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { Marker = true }));
        var hidden = Page(Chart(ChartKind.Line, Pointed(Series(Red, 5, 10, 20) with { Marker = true }, (1, new ChartPointMarker(Visible: false)))));
        Assert.Equal(3, Filled(plain, Red).Count);
        var left = Filled(hidden, Red);
        Assert.Equal(2, left.Count);
        var all = Filled(plain, Red);
        Assert.Equal([all[0], all[2]], left);
    }

    [Fact]
    public void A_visible_point_marker_on_a_series_without_markers_draws_only_that_marker()
    {
        var series = Pointed(Series(Red, 5, 10, 20) with { Marker = false }, (1, new ChartPointMarker(true, ChartMarkerSymbol.Square, 8)));
        var page = Page(Chart(ChartKind.Line, series));
        Assert.Equal(8, Assert.Single(Filled(page, Red)).Width);
    }

    [Fact]
    public void A_point_that_states_only_a_size_keeps_the_series_symbol_and_style()
    {
        var series = Pointed(Series(Red, 5, 10, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Square, MarkerSize = 6, MarkerStyle = new ChartMarkerStyle(Fill: Blue) },
            (1, new ChartPointMarker(Size: 12)));
        var paths = Paths(Page(Chart(ChartKind.Line, series))).Where(p => p.Fill is not null).ToList();
        Assert.Equal(3, paths.Count);
        Assert.All(paths, p => Assert.Equal((Color?)Blue, p.Fill));
        var widths = paths.Select(p => p.Points.Where(q => q.Kind != PathPointKind.Close)).Select(pts => pts.Max(q => q.X) - pts.Min(q => q.X)).ToList();
        Assert.Equal([6.0, 12.0, 6.0], widths.Select(w => Math.Round(w)));
    }

    [Fact]
    public void A_point_style_replaces_the_series_style_as_a_whole()
    {
        // Observed in Word 2016 (one sample): a fill-only point does not keep the series outline.
        var series = Pointed(Series(Red, 5, 10, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Circle, MarkerStyle = new ChartMarkerStyle(Fill: Red, Line: Blue, LineWidth: 2) },
            (0, new ChartPointMarker(Style: new ChartMarkerStyle(Fill: Color.White))));
        var paths = Paths(Page(Chart(ChartKind.Line, series))).ToList();
        Assert.Equal([(Color?)Color.White, Red, Red], paths.Where(p => p.Fill is not null).Select(p => p.Fill));
        // only the two untouched points keep the series' blue 2 pt outline; the fill-only point has the automatic one (Excel and PowerPoint 2016)
        var rings = paths.Where(p => p.Stroke is not null).ToList();
        Assert.Equal(3, rings.Count);
        Assert.Equal(2, rings.Count(p => p.Stroke == Blue && p.LineWidth == 2.0));
        Assert.Equal(1, rings.Count(p => p.Stroke == Red));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_point_stating_a_symbol_without_a_size_has_size_5(bool withFill)
    {
        // Word, Excel and PowerPoint 2016, one retest chart: the series size is not kept, with or without the point's own fill.
        var series = Pointed(Series(Red, 5, 10, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Circle, MarkerSize = 12, MarkerStyle = new ChartMarkerStyle(Fill: Blue) },
            (0, new ChartPointMarker(Symbol: ChartMarkerSymbol.Square, Style: withFill ? new ChartMarkerStyle(Fill: Color.White) : null)));
        var square = Paths(Page(Chart(ChartKind.Line, series))).First(p => p.Fill is not null);
        Assert.Equal(5, square.Points.Where(p => p.Kind != PathPointKind.Close).Max(p => p.X) - square.Points.Where(p => p.Kind != PathPointKind.Close).Min(p => p.X), 3);
    }

    [Fact]
    public void A_point_outline_alone_does_not_keep_the_series_fill()
    {
        var series = Pointed(Series(Red, 5, 10, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Circle, MarkerStyle = new ChartMarkerStyle(Fill: Blue, Line: Blue, LineWidth: 2) },
            (0, new ChartPointMarker(Style: new ChartMarkerStyle(Line: Color.White, LineWidth: 1))));
        var paths = Paths(Page(Chart(ChartKind.Line, series))).ToList();
        // the automatic color, not the series' explicit fill (Word 2016, one chart); the fill is drawn before the outline
        Assert.Equal((Color?)Red, paths[0].Fill);
        Assert.Equal((Color?)Color.White, paths[1].Stroke);
    }

    [Fact]
    public void A_scatter_point_marker_follows_the_point_index_across_a_null_value()
    {
        var series = Pointed(Series(Red, 5, null, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Circle, XValues = new double?[] { 1, 2, 3 } }, (2, new ChartPointMarker(Symbol: ChartMarkerSymbol.Square, Size: 9)));
        var page = Page(Chart(ChartKind.Scatter, series) with { XMin = 0, XMax = 4 });
        var boxes = Filled(page, Red);
        var rect = Assert.Single(boxes, b => b.Width == 9);
        var circle = Assert.Single(boxes, b => b.Width == 5);
        Assert.True(rect.X + 4.5 > CenterX(circle));
    }

    [Fact]
    public void A_radar_point_marker_belongs_to_the_category_with_the_same_index()
    {
        var series = Pointed(Series(Red, 20, 20, 20) with { Marker = true, MarkerSymbol = ChartMarkerSymbol.Circle }, (1, new ChartPointMarker(Symbol: ChartMarkerSymbol.Square, Size: 9)));
        var page = Page(Chart(ChartKind.Radar, series));
        var boxes = Filled(page, Red);
        var rect = Assert.Single(boxes, b => b.Width == 9);
        var circles = boxes.Where(b => b.Width == 5).ToList();
        Assert.Equal(2, circles.Count);
        // category 0 is at the top and 2 at the lower left, so category 1 (lower right) is right of both
        Assert.All(circles, c => Assert.True(rect.X + 4.5 > CenterX(c)));
    }

    [Fact]
    public void Gridlines_use_the_chart_color_and_width_and_axis_text_is_black_unless_stated()
    {
        var chart = Chart(ChartKind.Column, Series(Red, 5, 10, 20));
        var plain = Page(chart);
        Assert.Contains(plain.Items, i => i is LineItem l && l.Color == Color.Black && l.Width == 0.75); // the model's default (Excel)
        Assert.All(plain.Items.OfType<TextItem>().Where(t => t.Text is "0" or "10" or "20" or "a"), t => Assert.Equal(Color.Black, t.Color));
        var stated = Page(chart with { GridlineColor = Blue, GridlineWidth = 0.5 });
        Assert.Contains(stated.Items, i => i is LineItem l && l.Color == Blue && l.Width == 0.5);
        var colored = Page(chart with { TextStyles = new ChartTextStyles { ValueAxis = new ChartTextStyle(Color: Blue) } });
        Assert.All(colored.Items.OfType<TextItem>().Where(t => t.Text is "0" or "10" or "20"), t => Assert.Equal(Blue, t.Color));
    }

    [Fact]
    public void Line_scatter_and_radar_lines_are_drawn_in_the_stated_width_and_2_25_otherwise()
    {
        foreach (var kind in new[] { ChartKind.Line, ChartKind.Scatter, ChartKind.Radar })
        {
            ChartSeries Own(double? width) => Series(Red, 5, 10, 20) with { LineWidth = width, XValues = new double?[] { 1, 2, 3 } };
            SheetChart Make(ChartSeries s) => Chart(kind, s) with { XMin = 0, XMax = 4 };
            var plain = Page(Make(Own(null)));
            Assert.Contains(plain.Items, i => i is LineItem l && l.Color == Red && l.Width == 2.25);
            var wide = Page(Make(Own(4)));
            Assert.Contains(wide.Items, i => i is LineItem l && l.Color == Red && l.Width == 4);
            Assert.DoesNotContain(wide.Items, i => i is LineItem l && l.Color == Red && l.Width == 2.25);
        }
    }

    [Fact]
    public void A_line_and_its_legend_key_are_drawn_in_the_line_color_and_width()
    {
        var chart = Chart(ChartKind.Line, Series(Red, 5, 10, 20) with { LineColor = Blue, LineWidth = 1.5 }) with { Legend = LegendPosition.Bottom };
        var lines = Page(chart).Items.OfType<LineItem>().Where(l => l.Color == Blue).ToList();
        Assert.All(lines, l => Assert.Equal(1.5, l.Width));
        Assert.Contains(lines, l => l.Y1 == l.Y2); // legend key
        Assert.DoesNotContain(Page(chart).Items, i => i is LineItem l && l.Color == Red && l.Width >= 1.5);
    }
}
