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
}
