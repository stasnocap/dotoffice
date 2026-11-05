using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxChartTests
{
    private const string Ns = "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";

    private static string Space(string plot, string extra = "") =>
        $"<c:chartSpace {Ns}><c:chart>{extra}<c:plotArea><c:layout/>{plot}</c:plotArea></c:chart></c:chartSpace>";

    private static string Ser(int i, string? name, string? cats, string vals, string extra = "") =>
        $"<c:ser><c:idx val=\"{i}\"/><c:order val=\"{i}\"/>{(name is null ? "" : $"<c:tx><c:strRef><c:f>{name}</c:f></c:strRef></c:tx>")}{extra}" +
        $"{(cats is null ? "" : $"<c:cat><c:strRef><c:f>{cats}</c:f></c:strRef></c:cat>")}<c:val><c:numRef><c:f>{vals}</c:f></c:numRef></c:val></c:ser>";

    private static string Bars(string dir, string grouping, string sers, string extra = "") =>
        $"<c:barChart><c:barDir val=\"{dir}\"/><c:grouping val=\"{grouping}\"/>{sers}{extra}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
        "<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/><c:majorGridlines/></c:valAx>";

    private static (Worksheet Sheet, ConversionWarningList Warnings) Read(string chart)
    {
        var warnings = new ConversionWarningList();
        var rows = Row(1, C("A1", "x") + C("B1", "North")) + Row(2, Inline("A2", "Q1") + C("B2", "10")) + Row(3, Inline("A3", "Q2") + C("B3", "25")) + Row(4, Inline("A4", "Q3") + C("B4", "40"));
        var builder = new XlsxBuilder().AddSheet("S", rows, tail: DrawingElement).AddChartDrawing(1, TwoCellAnchor(1, 5, 5, 15, ChartRef()), chart);
        return (XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings }).Sheets[0], warnings);
    }

    [Fact]
    public void Reads_a_column_chart_from_the_cells_it_refers_to()
    {
        var (sheet, warnings) = Read(Space(Bars("col", "clustered", Ser(0, "S!$B$1", "S!$A$2:$A$4", "S!$B$2:$B$4")), "<c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>") );
        Assert.Empty(warnings.Warnings);
        var chart = Assert.Single(sheet.Charts);
        Assert.Equal((ChartKind.Column, ChartGrouping.Clustered, "Sales"), (chart.Kind, chart.Grouping, chart.Title));
        Assert.Equal(["Q1", "Q2", "Q3"], chart.Categories);
        var series = Assert.Single(chart.Series);
        Assert.Equal("North", series.Name);
        Assert.Equal(new double?[] { 10, 25, 40 }, series.Values.ToArray());
        Assert.Equal(Color.TryParseHex("4472C4"), series.Color);
        Assert.Equal(ImageAnchorKind.TwoCell, chart.Anchor);
    }

    [Fact]
    public void Cached_values_win_over_the_cells_and_explicit_colors_are_kept()
    {
        var ser = "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx><c:v>Lit</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></c:spPr>" +
                  "<c:cat><c:strLit><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>a</c:v></c:pt><c:pt idx=\"1\"><c:v>b</c:v></c:pt></c:strLit></c:cat>" +
                  "<c:val><c:numRef><c:f>S!$B$2:$B$4</c:f><c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>1</c:v></c:pt><c:pt idx=\"1\"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>";
        var chart = Read(Space(Bars("bar", "clustered", ser))).Sheet.Charts.Single();
        Assert.Equal(ChartKind.Bar, chart.Kind);
        Assert.Equal(["a", "b"], chart.Categories);
        Assert.Equal(new double?[] { 1, 3 }, chart.Series[0].Values.ToArray());
        Assert.Equal(new Color(255, 0, 0), chart.Series[0].Color);
        Assert.Equal("0.5", chart.Ticks[1].Label);
        Assert.Equal("1.0", chart.Ticks[2].Label);
    }

    [Theory]
    [InlineData(0, 100, 0, 120, 20)]
    [InlineData(0, 8, 0, 9, 1)]
    [InlineData(0, 1000, 0, 1200, 200)]
    [InlineData(-40, 100, -60, 120, 20)]
    [InlineData(0, 40, 0, 45, 5)]
    public void The_automatic_axis_follows_excels_scale(double lo, double hi, double min, double max, double step)
    {
        var ser = "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:val><c:numLit><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>" + lo.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                  "</c:v></c:pt><c:pt idx=\"1\"><c:v>" + hi.ToString(System.Globalization.CultureInfo.InvariantCulture) + "</c:v></c:pt></c:numLit></c:val></c:ser>";
        var chart = Read(Space(Bars("col", "clustered", ser))).Sheet.Charts.Single();
        Assert.Equal((min, max, step), (chart.ValueMin, chart.ValueMax, chart.Ticks[1].Value - chart.Ticks[0].Value));
    }

    [Fact]
    public void Fixed_axis_limits_and_units_are_honored_and_stacked_axes_use_the_sums()
    {
        var axes = "<c:valAx><c:axId val=\"2\"/><c:scaling><c:min val=\"0\"/><c:max val=\"50\"/></c:scaling><c:majorUnit val=\"10\"/></c:valAx>";
        var chart = Read(Space(Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4")).Replace("<c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/><c:majorGridlines/></c:valAx>", axes))).Sheet.Charts.Single();
        Assert.Equal((0.0, 50.0), (chart.ValueMin, chart.ValueMax));
        Assert.Equal(6, chart.Ticks.Count);
        var stacked = Read(Space(Bars("col", "stacked", Ser(0, null, null, "S!$B$2:$B$4") + Ser(1, null, null, "S!$B$2:$B$4")))).Sheet.Charts.Single();
        Assert.True(stacked.ValueMax >= 80);
        var percent = Read(Space(Bars("col", "percentStacked", Ser(0, null, null, "S!$B$2:$B$4")))).Sheet.Charts.Single();
        Assert.Equal(("0%", "100%"), (percent.Ticks[0].Label, percent.Ticks[^1].Label));
    }

    [Fact]
    public void Reads_line_markers_pie_colors_labels_and_legend()
    {
        var line = "<c:lineChart><c:grouping val=\"standard\"/>" + Ser(0, null, null, "S!$B$2:$B$4", "<c:marker><c:symbol val=\"none\"/></c:marker>") + Ser(1, null, null, "S!$B$2:$B$4") + "</c:lineChart>";
        var lineChart = Read(Space(line)).Sheet.Charts.Single();
        Assert.Equal([false, true], lineChart.Series.Select(s => s.Marker));
        var pie = "<c:pieChart><c:varyColors val=\"1\"/>" + Ser(0, "S!$B$1", "S!$A$2:$A$4", "S!$B$2:$B$4", "<c:dPt><c:idx val=\"1\"/><c:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></c:spPr></c:dPt>").Replace("</c:ser>", "<c:dLbls><c:showVal val=\"0\"/><c:showPercent val=\"1\"/></c:dLbls></c:ser>") + "</c:pieChart>";
        var pieChart = Read(Space(pie, "<c:autoTitleDeleted val=\"0\"/>") .Replace("</c:plotArea>", "</c:plotArea><c:legend><c:legendPos val=\"b\"/></c:legend>")).Sheet.Charts.Single();
        Assert.Equal((ChartKind.Pie, LegendPosition.Bottom, "North"), (pieChart.Kind, pieChart.Legend, pieChart.Title));
        Assert.Equal(new Color(0, 255, 0), pieChart.Series[0].PointColors![1]);
        Assert.Equal("13%", pieChart.Series[0].Labels![0]);
        Assert.Equal("53%", pieChart.Series[0].Labels![2]);
    }

    [Fact]
    public void Unsupported_chart_types_are_named_in_the_warning_and_not_drawn()
    {
        var (sheet, warnings) = Read(Space("<c:stockChart/>"));
        Assert.Empty(sheet.Charts);
        var warning = Assert.Single(warnings.Warnings);
        Assert.Equal("xlsx.drawing", warning.Code);
        Assert.Contains("stock", warning.Message);
    }

    private const string SecondaryAxes =
        "<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/><c:majorGridlines/></c:valAx>" +
        "<c:catAx><c:axId val=\"3\"/><c:scaling/><c:delete val=\"1\"/></c:catAx><c:valAx><c:axId val=\"4\"/><c:scaling><c:min val=\"0\"/><c:max val=\"2\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/>" +
        "<c:title><c:tx><c:rich><a:p><a:r><a:t>Share</a:t></a:r></a:p></c:rich></c:tx></c:title><c:crosses val=\"max\"/></c:valAx>";

    [Fact]
    public void A_combination_chart_keeps_every_bar_line_and_area_group()
    {
        var plot = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" + Ser(0, null, null, "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
                   "<c:lineChart><c:grouping val=\"standard\"/>" + Ser(1, null, null, "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>" +
                   "<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/></c:valAx>";
        var (sheet, warnings) = Read(Space(plot));
        Assert.Empty(warnings.Warnings);
        var chart = Assert.Single(sheet.Charts);
        Assert.Equal(ChartKind.Column, chart.Kind);
        var extra = Assert.Single(chart.ExtraGroups);
        Assert.Equal((ChartKind.Line, false), (extra.Kind, extra.Secondary));
        Assert.Equal("Series2", extra.Series[0].Name);
        Assert.Equal(new Color(0xED, 0x7D, 0x31), extra.Series[0].Color); // the second series takes the second accent
    }

    [Fact]
    public void A_group_on_its_own_value_axis_gets_the_secondary_scale()
    {
        var plot = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" + Ser(0, null, null, "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
                   "<c:lineChart><c:grouping val=\"standard\"/>" + Ser(1, null, null, "S!$B$2:$B$4") + "<c:axId val=\"3\"/><c:axId val=\"4\"/></c:lineChart>" + SecondaryAxes;
        var (sheet, warnings) = Read(Space(plot));
        Assert.Empty(warnings.Warnings);
        var chart = sheet.Charts.Single();
        Assert.True(chart.ExtraGroups[0].Secondary);
        Assert.True(chart.ShowSecondaryAxis);
        Assert.Equal((0.0, 2.0, "Share"), (chart.SecondaryMin, chart.SecondaryMax, chart.SecondaryAxisTitle));
        Assert.True(chart.ValueMax >= 40); // the primary scale comes from the columns only
    }

    [Fact]
    public void A_combination_with_an_unsupported_group_names_what_is_left_out()
    {
        var plot = Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4")) + "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/></c:scatterChart>";
        var (sheet, warnings) = Read(Space(plot));
        Assert.Single(sheet.Charts);
        Assert.Single(warnings.Warnings, w => w.Code == "xlsx.drawing-approximated" && w.Message.Contains("combination") && w.Message.Contains("scatter"));
    }

    [Fact]
    public void Reads_reversed_and_logarithmic_axes_and_smoothed_lines()
    {
        var axes = "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"maxMin\"/></c:scaling><c:delete val=\"0\"/></c:catAx>" +
                   "<c:valAx><c:axId val=\"2\"/><c:scaling><c:logBase val=\"10\"/><c:orientation val=\"maxMin\"/></c:scaling><c:delete val=\"0\"/></c:valAx>";
        var plot = "<c:lineChart><c:grouping val=\"standard\"/>" + Ser(0, null, null, "S!$B$2:$B$4", "<c:smooth val=\"1\"/>").Replace("<c:smooth val=\"1\"/>", "") .Replace("</c:ser>", "<c:smooth val=\"1\"/></c:ser>") +
                   "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>" + axes;
        var chart = Read(Space(plot)).Sheet.Charts.Single();
        Assert.True(chart.ReverseCategories);
        Assert.True(chart.ReverseValues);
        Assert.Equal(10.0, chart.LogBase);
        Assert.Equal((10.0, 100.0), (chart.ValueMin, chart.ValueMax)); // data 10..40
        Assert.Equal([10.0, 100.0], chart.Ticks.Select(t => t.Value));
        Assert.True(chart.Series[0].Smooth);
        Assert.False(Read(Space(plot.Replace("<c:smooth val=\"1\"/>", ""))).Sheet.Charts.Single().Series[0].Smooth);
    }

    private static string XyAxes(string extraY = "") =>
        "<c:valAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:majorGridlines/></c:valAx>" +
        $"<c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/>{extraY}</c:valAx>";

    private static string XySer(string extra, string x, string y, string size = "") =>
        $"<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>{extra}<c:xVal><c:numRef><c:f>{x}</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>{y}</c:f></c:numRef></c:yVal>" +
        $"{(size.Length == 0 ? "" : $"<c:bubbleSize><c:numRef><c:f>{size}</c:f></c:numRef></c:bubbleSize>")}</c:ser>";

    [Fact]
    public void Reads_a_scatter_chart_with_its_own_x_scale_and_marker_only_series()
    {
        var markersOnly = "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val=\"circle\"/><c:spPr><a:solidFill><a:srgbClr val=\"00AA00\"/></a:solidFill></c:spPr></c:marker>";
        var plot = "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/>" + XySer(markersOnly, "S!$B$2:$B$4", "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>" + XyAxes();
        var (sheet, warnings) = Read(Space(plot));
        Assert.Empty(warnings.Warnings);
        var chart = Assert.Single(sheet.Charts);
        var series = Assert.Single(chart.Series);
        Assert.Equal(ChartKind.Scatter, chart.Kind);
        Assert.Equal(new double?[] { 10, 25, 40 }, series.XValues!.ToArray());
        Assert.False(series.Line);
        Assert.True(series.Marker);
        Assert.Equal(new Color(0, 0xAA, 0), series.Color);
        Assert.Equal((0.0, 45.0), (chart.XMin, chart.XMax));
        Assert.True(chart.XGridlines);
        Assert.Equal(chart.Ticks.Count, chart.XTicks.Count);
    }

    [Fact]
    public void A_scatter_group_on_its_own_pair_of_value_axes_gets_secondary_scales()
    {
        var secondary = "<c:valAx><c:axId val=\"3\"/><c:scaling><c:min val=\"0\"/><c:max val=\"100\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"t\"/></c:valAx>" +
                        "<c:valAx><c:axId val=\"4\"/><c:scaling><c:min val=\"0\"/><c:max val=\"500\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Right</a:t></a:r></a:p></c:rich></c:tx></c:title><c:crosses val=\"max\"/></c:valAx>";
        var plot = "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/>" + XySer("", "S!$B$2:$B$4", "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>" +
                   "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/>" + XySer("", "S!$B$2:$B$4", "S!$B$2:$B$4") + "<c:axId val=\"3\"/><c:axId val=\"4\"/></c:scatterChart>" + XyAxes() + secondary;
        var (sheet, warnings) = Read(Space(plot));
        Assert.Empty(warnings.Warnings);
        var chart = Assert.Single(sheet.Charts);
        var group = Assert.Single(chart.ExtraGroups);
        Assert.True(group.Secondary);
        Assert.Equal(ChartKind.Scatter, group.Kind);
        Assert.Equal((0.0, 500.0, "Right"), (chart.SecondaryMin, chart.SecondaryMax, chart.SecondaryAxisTitle));
        Assert.Equal((0.0, 100.0), (chart.SecondaryXMin, chart.SecondaryXMax));
        Assert.True(chart.ShowSecondaryAxis && chart.ShowSecondaryXAxis);
        Assert.Equal((0.0, 45.0), (chart.XMin, chart.XMax)); // the primary range comes from the first group alone
    }

    [Fact]
    public void A_second_scatter_group_on_the_same_axes_is_drawn_on_them()
    {
        var plot = "<c:scatterChart>" + XySer("", "S!$B$2:$B$4", "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>" +
                   "<c:scatterChart>" + XySer("", "S!$B$2:$B$4", "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>" + XyAxes();
        var chart = Read(Space(plot)).Sheet.Charts.Single();
        Assert.False(Assert.Single(chart.ExtraGroups).Secondary);
        Assert.False(chart.ShowSecondaryAxis);
    }

    [Fact]
    public void Reads_a_bubble_chart_with_sizes()
    {
        var plot = "<c:bubbleChart>" + XySer("", "S!$B$2:$B$4", "S!$B$2:$B$4", "S!$B$2:$B$4") + "<c:bubbleScale val=\"50\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:bubbleChart>" + XyAxes();
        var chart = Read(Space(plot)).Sheet.Charts.Single();
        Assert.Equal(ChartKind.Bubble, chart.Kind);
        Assert.Equal(new double?[] { 10, 25, 40 }, chart.Series[0].Sizes!.ToArray());
        Assert.Equal(50, chart.BubbleScale);
    }

    [Fact]
    public void Reads_area_charts_with_stacking_and_the_edge_crossing()
    {
        var plot = "<c:areaChart><c:grouping val=\"stacked\"/>" + Ser(0, null, "S!$A$2:$A$4", "S!$B$2:$B$4") + Ser(1, null, null, "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:areaChart>" +
                   "<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/><c:crossBetween val=\"midCat\"/></c:valAx>";
        var chart = Read(Space(plot)).Sheet.Charts.Single();
        Assert.Equal((ChartKind.Area, ChartGrouping.Stacked, false), (chart.Kind, chart.Grouping, chart.CategoriesBetween));
        Assert.True(chart.ValueMax >= 80);
        Assert.True(Read(Space(plot.Replace("midCat", "between"))).Sheet.Charts.Single().CategoriesBetween);
    }

    [Fact]
    public void Reads_doughnut_and_radar_charts()
    {
        var doughnut = "<c:doughnutChart><c:varyColors val=\"1\"/>" + Ser(0, "S!$B$1", "S!$A$2:$A$4", "S!$B$2:$B$4").Replace("</c:ser>", "<c:dLbls><c:showVal val=\"0\"/><c:showPercent val=\"1\"/></c:dLbls></c:ser>") +
                       "<c:firstSliceAng val=\"90\"/><c:holeSize val=\"60\"/></c:doughnutChart>";
        var ring = Read(Space(doughnut)).Sheet.Charts.Single();
        Assert.Equal((ChartKind.Doughnut, 0.6, 90.0), (ring.Kind, ring.HoleSize, ring.FirstSliceAngle));
        Assert.Equal("13%", ring.Series[0].Labels![0]);
        Assert.Equal(3, ring.Series[0].PointColors!.Count);
        var radar = "<c:radarChart><c:radarStyle val=\"filled\"/>" + Ser(0, null, "S!$A$2:$A$4", "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:radarChart>" +
                    "<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/><c:majorGridlines/></c:valAx>";
        var web = Read(Space(radar)).Sheet.Charts.Single();
        Assert.Equal((ChartKind.Radar, true), (web.Kind, web.Filled));
        Assert.Equal(["Q1", "Q2", "Q3"], web.Categories);
    }

    private static string Line(string sers, string extra = "") =>
        $"<c:lineChart><c:grouping val=\"standard\"/>{sers}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>" +
        "<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/></c:valAx>";

    [Fact]
    public void Reads_the_opacity_of_a_fill()
    {
        var ser = Ser(0, null, null, "S!$B$2:$B$4", "<c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"><a:alpha val=\"40000\"/></a:srgbClr></a:solidFill></c:spPr>");
        var chart = Read(Space(Bars("col", "clustered", ser))).Sheet.Charts.Single();
        Assert.Equal(0.4, chart.Series[0].Alpha, 6);
        Assert.Equal(1, Read(Space(Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4")))).Sheet.Charts.Single().Series[0].Alpha);
    }

    [Fact]
    public void Reads_a_trendline_with_its_type_look_and_extent()
    {
        var trend = "<c:trendline><c:spPr><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill><a:prstDash val=\"sysDash\"/></a:ln></c:spPr>" +
                    "<c:trendlineType val=\"poly\"/><c:order val=\"3\"/><c:forward val=\"1.5\"/><c:backward val=\"0.5\"/></c:trendline>";
        var series = Read(Space(Line(Ser(0, null, null, "S!$B$2:$B$4", trend)))).Sheet.Charts.Single().Series[0];
        var t = Assert.IsType<ChartTrendline>(series.Trendline);
        Assert.Equal((TrendlineKind.Polynomial, 3, 1.5, 0.5, 1.0, true, new Color(0, 255, 0)), (t.Kind, t.Order, t.Forward, t.Backward, t.Width, t.Dashed, t.Color));
        var plain = Read(Space(Line(Ser(0, null, null, "S!$B$2:$B$4", "<c:trendline><c:trendlineType val=\"linear\"/></c:trendline>")))).Sheet.Charts.Single().Series[0].Trendline!;
        Assert.Equal((TrendlineKind.Linear, false, new Color(0x44, 0x72, 0xC4)), (plain.Kind, plain.Dashed, plain.Color));
    }

    [Theory]
    [InlineData("<c:errValType val=\"fixedVal\"/><c:val val=\"3\"/>", 3, 3)]
    [InlineData("<c:errValType val=\"percentage\"/><c:val val=\"10\"/>", 1, 1)]
    [InlineData("<c:errBarType val=\"plus\"/><c:errValType val=\"fixedVal\"/><c:val val=\"3\"/>", 3, 0)]
    [InlineData("<c:errBarType val=\"minus\"/><c:errValType val=\"fixedVal\"/><c:val val=\"3\"/>", 0, 3)]
    [InlineData("<c:errValType val=\"cust\"/><c:plus><c:numLit><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>4</c:v></c:pt></c:numLit></c:plus><c:minus><c:numLit><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:minus>", 4, 2)]
    public void Reads_error_bars_of_every_value_type(string body, double plus, double minus)
    {
        var bars = "<c:errBars><c:errDir val=\"y\"/>" + (body.Contains("errBarType") ? "" : "<c:errBarType val=\"both\"/>") + body + "<c:noEndCap val=\"1\"/></c:errBars>";
        var series = Read(Space(Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4", bars)))).Sheet.Charts.Single().Series[0];
        var read = Assert.Single(series.ErrorBars!);
        Assert.Equal((plus, minus, false), (read.Plus[0], read.Minus[0], read.EndCap));
    }

    [Fact]
    public void Error_bars_widen_the_automatic_scale()
    {
        var plain = Read(Space(Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4")))).Sheet.Charts.Single();
        var bars = "<c:errBars><c:errBarType val=\"plus\"/><c:errValType val=\"fixedVal\"/><c:val val=\"30\"/></c:errBars>";
        var wide = Read(Space(Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4", bars)))).Sheet.Charts.Single();
        Assert.True(wide.ValueMax > plain.ValueMax);
    }

    [Fact]
    public void Reads_manual_layouts_given_from_the_edge_and_ignores_factor_mode()
    {
        string Layout(string mode) => "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/>" + mode + "<c:x val=\"0.1\"/><c:y val=\"0.2\"/><c:w val=\"0.7\"/><c:h val=\"0.6\"/></c:manualLayout></c:layout>";
        string Chart(string layout) => $"<c:chartSpace {Ns}><c:chart><c:title>{(layout == "" ? "" : "<c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"0.3\"/><c:y val=\"0.02\"/></c:manualLayout></c:layout>")}</c:title><c:plotArea>{layout}" +
                                       Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4")) + "</c:plotArea></c:chart></c:chartSpace>";
        var chart = Read(Chart(Layout("<c:xMode val=\"edge\"/><c:yMode val=\"edge\"/>"))).Sheet.Charts.Single();
        Assert.Equal(new ChartBox(0.1, 0.2, 0.7, 0.6, true), chart.PlotArea);
        Assert.Equal((0.3, 0.02), chart.TitlePosition);
        Assert.Null(Read(Chart(Layout(""))).Sheet.Charts.Single().PlotArea);
        Assert.Null(Read(Chart("<c:layout/>")).Sheet.Charts.Single().PlotArea);
    }

    [Fact]
    public void Reads_trendline_labels_the_legend_position_and_text_formatting()
    {
        var trend = "<c:trendline><c:trendlineType val=\"linear\"/><c:dispRSqr val=\"1\"/><c:dispEq val=\"1\"/>" +
                    "<c:trendlineLbl><c:txPr><a:p><a:pPr><a:defRPr sz=\"1100\" b=\"1\"/></a:pPr></a:p></c:txPr></c:trendlineLbl></c:trendline>";
        string Rich(string props, string text, string inner = "") => $"<c:tx><c:rich><a:p><a:pPr><a:defRPr sz=\"1800\"/></a:pPr><a:r><a:rPr lang=\"en-US\" {props}>{inner}</a:rPr><a:t>{text}</a:t></a:r></a:p></c:rich></c:tx>";
        string Tx(string props, string inner = "") => $"<c:txPr><a:p><a:pPr><a:defRPr {props}>{inner}</a:defRPr></a:pPr></a:p></c:txPr>";
        string Fill(string hex) => $"<a:solidFill><a:srgbClr val=\"{hex}\"/></a:solidFill>";
        var plot = $"<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>{Ser(0, null, null, "S!$B$2:$B$4", trend)}<c:dLbls>{Tx("sz=\"800\"", Fill("FF0000"))}<c:showVal val=\"1\"/></c:dLbls><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
                   $"<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/><c:title>{Rich("sz=\"1200\" i=\"1\"", "Cat")}</c:title>{Tx("sz=\"1300\"")}</c:catAx>" +
                   $"<c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/>{Tx("b=\"1\"", Fill("00FF00") + "<a:latin typeface=\"Arial\"/>")}</c:valAx>";
        var legend = "<c:legend><c:legendPos val=\"r\"/><c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"0.7\"/><c:y val=\"0.1\"/><c:w val=\"0.2\"/><c:h val=\"0.3\"/></c:manualLayout></c:layout>" +
                     Tx("sz=\"900\" i=\"1\"") + "</c:legend>";
        var title = $"<c:title>{Rich("sz=\"2000\" b=\"1\"", "Big", Fill("0000FF"))}</c:title>";
        var chart = Read(Space(plot, title + legend)).Sheet.Charts.Single();
        var t = chart.Series[0].Trendline!;
        Assert.True(t.ShowEquation && t.ShowRSquared);
        Assert.Equal(new ChartTextStyle(11, null, true, null, null), t.LabelStyle);
        Assert.Equal(new ChartBox(0.7, 0.1, 0.2, 0.3, false), chart.LegendBox);
        Assert.Equal(new ChartTextStyle(9, null, null, true, null), chart.TextStyles.Legend);
        Assert.Equal(new ChartTextStyle(13), chart.TextStyles.CategoryAxis);
        Assert.Equal(new ChartTextStyle(12, null, null, true, null), chart.TextStyles.CategoryAxisTitle); // the run overrides the paragraph default
        Assert.Equal(new ChartTextStyle(null, new Color(0, 255, 0), true, null, "Arial"), chart.TextStyles.ValueAxis);
        Assert.Equal(new ChartTextStyle(8, new Color(255, 0, 0)), chart.TextStyles.DataLabels);
        Assert.Equal(20, chart.TextStyles.Title!.Size);
        Assert.Equal(new Color(0, 0, 255), chart.TextStyles.Title.Color);
        Assert.Null(Read(Space(Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4")))).Sheet.Charts.Single().TextStyles.Legend);
    }

    [Fact]
    public void The_charts_base_text_style_colors_every_part_but_not_the_size()
    {
        var plot = Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4"));
        var space = $"<c:chartSpace {Ns}><c:chart><c:legend><c:legendPos val=\"r\"/></c:legend><c:plotArea><c:layout/>{plot}</c:plotArea></c:chart>" +
                    "<c:txPr><a:p><a:pPr><a:defRPr sz=\"1200\"><a:solidFill><a:srgbClr val=\"336699\"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>";
        var chart = Read(space).Sheet.Charts.Single();
        Assert.Equal(12, chart.FontSize);
        Assert.Equal(new ChartTextStyle(Color: new Color(0x33, 0x66, 0x99)), chart.TextStyles.Legend);
    }

    [Fact]
    public void Area_and_line_groups_on_the_secondary_axis_keep_their_transparency()
    {
        var area = Ser(1, null, null, "S!$B$2:$B$4", "<c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"><a:alpha val=\"30000\"/></a:srgbClr></a:solidFill></c:spPr>");
        var plot = "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" + Ser(0, null, null, "S!$B$2:$B$4") + "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
                   "<c:areaChart><c:grouping val=\"standard\"/>" + area + "<c:axId val=\"3\"/><c:axId val=\"4\"/></c:areaChart>" + SecondaryAxes;
        var group = Read(Space(plot)).Sheet.Charts.Single().ExtraGroups.Single();
        Assert.True(group.Secondary);
        Assert.Equal(0.3, group.Series[0].Alpha, 6);
    }

    [Fact]
    public void Reads_a_trendline_labels_number_format_and_position()
    {
        string Trend(string label) => $"<c:trendline><c:trendlineType val=\"linear\"/><c:dispEq val=\"1\"/><c:trendlineLbl>{label}</c:trendlineLbl></c:trendline>";
        var factor = "<c:layout><c:manualLayout><c:x val=\"-0.1\"/><c:y val=\"0.2\"/></c:manualLayout></c:layout><c:numFmt formatCode=\"0.00\" sourceLinked=\"0\"/>";
        var edge = "<c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"0.3\"/><c:y val=\"0.4\"/></c:manualLayout></c:layout><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>";
        var first = Read(Space(Line(Ser(0, null, null, "S!$B$2:$B$4", Trend(factor))))).Sheet.Charts.Single().Series[0].Trendline!;
        Assert.Equal(new ChartLabelOffset(-0.1, 0.2, false), first.LabelPosition);
        Assert.Equal("1.50", first.NumberFormatter!(1.5));
        var second = Read(Space(Line(Ser(0, null, null, "S!$B$2:$B$4", Trend(edge))))).Sheet.Charts.Single().Series[0].Trendline!;
        Assert.Equal(new ChartLabelOffset(0.3, 0.4, true), second.LabelPosition);
        Assert.Null(second.NumberFormatter);
    }

    [Fact]
    public void Reads_the_rotation_of_axis_text_and_the_size_of_a_manual_legend()
    {
        string Axis(string body) => $"<c:txPr><a:bodyPr {body}/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr>";
        string Plot(string categoryText) => $"<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>{Ser(0, null, null, "S!$B$2:$B$4")}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
            $"<c:catAx><c:axId val=\"1\"/><c:scaling/><c:delete val=\"0\"/>{categoryText}</c:catAx><c:valAx><c:axId val=\"2\"/><c:scaling/><c:delete val=\"0\"/></c:valAx>";
        double? Rotation(string body) => Read(Space(Plot(Axis(body)))).Sheet.Charts.Single().TextStyles.CategoryAxis?.Rotation;
        Assert.Equal(90, Rotation("rot=\"-5400000\" vert=\"horz\""));  // clockwise in the file, counter-clockwise on the page
        Assert.Equal(-45, Rotation("rot=\"2700000\""));
        Assert.Equal(0, Rotation("rot=\"0\""));
        Assert.Null(Rotation("rot=\"-60000000\""));                      // automatic
        Assert.Equal(90, Rotation("vert=\"vert270\""));
        var legend = "<c:legend><c:legendPos val=\"r\"/><c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"0.6\"/><c:y val=\"0.1\"/><c:w val=\"0.3\"/><c:h val=\"0.2\"/></c:manualLayout></c:layout></c:legend>";
        var box = Read(Space(Plot(""), legend)).Sheet.Charts.Single().LegendBox!.Value;
        Assert.Equal((0.3, 0.2), (box.Width, box.Height));
    }

    [Fact]
    public void A_points_own_data_label_can_show_hide_format_and_style_it()
    {
        var labels = "<c:dLbls>" +
                     "<c:dLbl><c:idx val=\"0\"/><c:delete val=\"1\"/></c:dLbl>" +
                     "<c:dLbl><c:idx val=\"1\"/><c:numFmt formatCode=\"0.0&quot; u&quot;\" sourceLinked=\"0\"/><c:txPr><a:p><a:pPr><a:defRPr sz=\"1400\" b=\"1\"/></a:pPr></a:p></c:txPr><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbl>" +
                     "<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbls>";
        var series = Read(Space(Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4", labels)))).Sheet.Charts.Single().Series[0];
        Assert.Equal(["", "25.0 u", "40"], series.Labels);
        Assert.Equal(new ChartTextStyle(14, null, true), series.LabelStyles![1]);
        Assert.Null(series.LabelStyles[2]);
        // Only one point labelled while the series itself shows none.
        var only = "<c:dLbls><c:dLbl><c:idx val=\"2\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/></c:dLbl><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/></c:dLbls>";
        var single = Read(Space(Bars("col", "clustered", Ser(0, null, null, "S!$B$2:$B$4", only)))).Sheet.Charts.Single().Series[0];
        Assert.Equal(["", "", "40"], single.Labels);
    }
}
