using DotOffice.Core;
using DotOffice.Layout;

namespace DotOffice.Tests.Layout;

public class SheetChartLayoutTests
{
    private static readonly Color Red = new(255, 0, 0), Blue = new(0, 0, 255);

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
        Assert.Equal(3, page.Items.OfType<EllipseItem>().Count(e => e.Fill == Red));
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
        var dots = page.Items.OfType<EllipseItem>().Where(e => e.Fill == Red).OrderBy(e => e.X).ToList();
        Assert.Equal(2, dots.Count);
        Assert.True(dots[1].X > dots[0].X + 100);
        Assert.True(dots[1].Y < dots[0].Y - 50);
        Assert.DoesNotContain(page.Items.OfType<LineItem>(), l => l.Color == Red);
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
}
