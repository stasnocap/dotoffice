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
        var dots = page.Items.OfType<EllipseItem>().Where(e => e.Fill == Blue).ToList();
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
    public void A_manual_legend_sits_at_its_position_and_leaves_the_plot_alone()
    {
        var chart = Chart(ChartKind.Column, Series(Red, 5, 10, 20)) with { Legend = LegendPosition.Right, Border = null, Background = null };
        var automatic = Page(chart);
        var manual = Page(chart with { LegendBox = new ChartBox(0.1, 0.8, 0.3, 0.1, false) });
        var swatch = manual.Items.OfType<RectItem>().Where(r => r.Fill == Red && r.Width < 10).Single();
        Assert.Equal(OriginX + 30, swatch.X, 3);
        Assert.InRange(swatch.Y, OriginY + 160, OriginY + 175);
        // The automatic legend takes room on the right; the manual one does not, so the columns are wider.
        double Width(LayoutPage page) => page.Items.OfType<RectItem>().Where(r => r.Fill == Red && r.Height > 10).Max(r => r.X + r.Width);
        Assert.True(Width(manual) > Width(automatic));
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
}
