using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class ChartTests
{
    private const string ChartType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    private const string ThemeColors =
        "<a:clrScheme name=\"T\"><a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>" +
        "<a:accent1><a:srgbClr val=\"112233\"/></a:accent1><a:accent2><a:srgbClr val=\"445566\"/></a:accent2><a:accent3><a:srgbClr val=\"778899\"/></a:accent3>" +
        "<a:accent4><a:srgbClr val=\"AA1122\"/></a:accent4><a:accent5><a:srgbClr val=\"22AA33\"/></a:accent5><a:accent6><a:srgbClr val=\"3344AA\"/></a:accent6>" +
        "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme>";

    private sealed class ListSink(List<ConversionWarning> list) : IConversionWarningSink
    {
        public void Warn(ConversionWarning warning) => list.Add(warning);
    }

    private static string Values(string tag, params double[] values) =>
        $"<c:{tag}><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"{values.Length}\"/>" +
        string.Concat(values.Select((v, i) => $"<c:pt idx=\"{i}\"><c:v>{v}</c:v></c:pt>")) + $"</c:numCache></c:numRef></c:{tag}>";

    private static string ChartXml(bool cached = true) =>
        "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:chart>" +
        "<c:autoTitleDeleted val=\"1\"/><c:plotArea><c:layout/><c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" +
        "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx><c:v>S0</c:v></c:tx>" +
        (cached ? "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val=\"3\"/><c:pt idx=\"0\"><c:v>a</c:v></c:pt><c:pt idx=\"1\"><c:v>b</c:v></c:pt><c:pt idx=\"2\"><c:v>c</c:v></c:pt></c:strCache></c:strRef></c:cat>" +
                   Values("val", 5, 10, 20)
                 : "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val>") +
        "</c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" +
        "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>" +
        "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>" +
        "</c:plotArea><c:plotVisOnly val=\"1\"/></c:chart></c:chartSpace>";

    private const string ChartGraphic =
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"{0}\"/></a:graphicData></a:graphic>";

    private static string Inline(string rel, double w, double h) =>
        $"<w:r><w:drawing><wp:inline><wp:extent cx=\"{(long)(w * 12700)}\" cy=\"{(long)(h * 12700)}\"/><wp:docPr id=\"1\" name=\"Chart\"/>{string.Format(ChartGraphic, rel)}</wp:inline></w:drawing></w:r>";

    private static string Anchor(string rel, double w, double h, string position) =>
        "<w:r><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" relativeHeight=\"1\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">" +
        $"<wp:simplePos x=\"0\" y=\"0\"/>{position}<wp:extent cx=\"{(long)(w * 12700)}\" cy=\"{(long)(h * 12700)}\"/><wp:wrapNone/><wp:docPr id=\"2\" name=\"Chart\"/>{string.Format(ChartGraphic, rel)}</wp:anchor></w:drawing></w:r>";

    private static (LayoutPage Page, List<ConversionWarning> Warnings, Document Doc) Run(Func<DocxBuilder, string> body, string? theme = null)
    {
        var b = new DocxBuilder { Theme = theme };
        b.Body = body(b) + Section();
        var warnings = new List<ConversionWarning>();
        var doc = DocxReader.Read(b.BuildStream(), new DocxReaderOptions { Warnings = new ListSink(warnings) });
        var page = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)), Warnings = new ListSink(warnings) }).Layout(doc).Pages[0];
        return (page, warnings, doc);
    }

    private static List<RectItem> Bars(LayoutPage page, Color color) => page.Items.OfType<RectItem>().Where(r => r.Fill == color && r.Width > 5 && r.Height > 5).OrderBy(r => r.X).ToList();

    private static readonly Color Accent1 = Color.TryParseHex("5B9BD5")!.Value;

    [Fact]
    public void Inline_chart_draws_its_series_inside_the_extent_and_takes_line_height()
    {
        var (page, warnings, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, ChartXml()), 300, 200)) + P(R("after", "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/>")));
        Assert.DoesNotContain(warnings, w => w.Code.StartsWith("docx.chart"));
        var bars = Bars(page, Accent1);
        Assert.Equal(3, bars.Count);
        Assert.Equal(2, bars[1].Height / bars[0].Height, 2);
        Assert.All(bars, r => Assert.True(r.X >= 72 && r.X + r.Width <= 72 + 300 && r.Y >= 72 && r.Y + r.Height <= 72 + 200 + 0.01));
        // The paragraph after the chart starts below its 200 pt.
        var after = page.Items.OfType<TextItem>().First(t => t.Text.Contains("after"));
        Assert.True(after.Baseline > 72 + 200);
        Assert.DoesNotContain(page.Items, i => i is ChartItem);
    }

    [Fact]
    public void A_chart_has_Words_default_frame()
    {
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, ChartXml()), 300, 200)));
        // Stroked 0.5 pt, centred on the frame's edge (measured in Word 2021).
        var frame = Assert.Single(page.Items.OfType<PathItem>(), p => p.Stroke == new Color(0x89, 0x89, 0x89) && Math.Abs(p.LineWidth - 0.5) < 1e-6);
        Assert.Equal(4, frame.Points.Count(p => p.Kind != PathPointKind.Close));
        // Word fills a plot area without a stated fill white.
        Assert.Contains(page.Items, i => i is RectItem r && r.Fill == Color.White && r.Width < 300 && r.Height < 200 && r.X > 0);
    }

    [Theory]
    [InlineData(216, "0,1,2,3,4,5,6")]
    [InlineData(115, "0,1,2,3,4,5,6")]
    [InlineData(113, "0,2,4,6")]
    public void A_Word_charts_value_axis_takes_the_finest_step_whose_label_lines_fit_the_plot(double height, string expected)
    {
        // Data maximum 5: 0 to 6. The plot is the height - 11 - 23.71; six intervals need 6 x 11.5 + 10.25 = 79.25 pt of it.
        // Measured in Word 2016: a 78.99 pt plot has step 2, a 79.99 pt plot step 1.
        var xml = ChartXml().Replace("<c:v>20</c:v>", "<c:v>5</c:v>").Replace("<c:v>10</c:v>", "<c:v>3</c:v>");
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), 360, height)));
        var labels = page.Items.OfType<TextItem>().Where(t => t.Text.Length == 1 && char.IsDigit(t.Text[0]) && t.X < 72 + 30).OrderBy(t => t.Text).Select(t => t.Text);
        Assert.Equal(expected, string.Join(",", labels));
    }

    private static List<LineItem> ChartLines(string xml) => Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), 360, 216))).Page.Items.OfType<LineItem>().ToList();

    private static List<LineItem> OfLength(IEnumerable<LineItem> lines, double length) => lines.Where(l => Math.Abs(Math.Max(l.X2 - l.X1, l.Y2 - l.Y1) - length) < 1e-6).ToList();

    private static string WithTicks(string major, string? minor = null)
    {
        string Marks(string axis) => (major.Length > 0 ? $"<c:majorTickMark val=\"{major}\"/>" : "") + (minor is null ? "" : $"<c:minorTickMark val=\"{minor}\"/>") + axis;
        return ChartXml().Replace("<c:crossAx val=\"2\"/>", Marks("<c:crossAx val=\"2\"/>")).Replace("<c:crossAx val=\"1\"/>", Marks("<c:crossAx val=\"1\"/>"));
    }

    [Fact]
    public void A_Word_chart_draws_tick_marks_as_the_file_asks_for_them()
    {
        // Measured in Word 2016: out and in are 3.02 pt, cross (and a missing majorTickMark) 6.04 pt centered on the axis, none draws nothing.
        // 3 categories: 4 slot boundaries; value axis 0..25 by 5: 6 ticks.
        var outside = OfLength(ChartLines(WithTicks("out")), 3.02);
        Assert.Equal(4 + 6, outside.Count);
        var inside = OfLength(ChartLines(WithTicks("in")), 3.02);
        Assert.Equal(4 + 6, inside.Count);
        var axisY = outside.Where(l => l.X1 == l.X2).Select(l => l.Y1).Min();
        Assert.All(inside.Where(l => l.X1 == l.X2), l => Assert.Equal(axisY, l.Y2, 3)); // from the plot down to the axis line
        Assert.All(outside.Where(l => l.X1 == l.X2), l => Assert.True(l.Y2 > l.Y1 && l.Y1 == axisY)); // out: down from the axis line
        Assert.Equal(4 + 6, OfLength(ChartLines(WithTicks("cross")), 6.04).Count);
        Assert.Equal(4 + 6, OfLength(ChartLines(ChartXml()), 6.04).Count);
        Assert.Empty(OfLength(ChartLines(WithTicks("none")), 3.02));
        Assert.Empty(OfLength(ChartLines(WithTicks("none")), 6.04));
    }

    [Fact]
    public void A_Word_chart_draws_outside_minor_tick_marks_a_fifth_of_a_step_apart()
    {
        var lines = ChartLines(WithTicks("out", "out"));
        Assert.Equal(5 * 5 + 1, OfLength(lines, 2.31).Count); // the value axis: five per step, and one at the end
        Assert.Equal(3 + 1, OfLength(lines, 2.23).Count); // the category axis: slot centers and the last boundary
        Assert.Empty(OfLength(ChartLines(WithTicks("out", "in")), 2.31));
    }

    private static readonly Color Grey = new(0x89, 0x89, 0x89);

    private static string WithC14Style(string xml) => xml.Replace("<c:chart>", "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:c14=\"http://schemas.microsoft.com/office/drawing/2007/8/2/chart\"><mc:Choice Requires=\"c14\"><c14:style val=\"102\"/></mc:Choice><mc:Fallback><c:style val=\"2\"/></mc:Fallback></mc:AlternateContent><c:chart>");

    private static LineItem LongLine(IEnumerable<LineItem> lines, bool horizontal) => lines.Where(l => (l.Y1 == l.Y2) == horizontal && Math.Abs(Math.Max(l.X2 - l.X1, l.Y2 - l.Y1)) > 100).OrderByDescending(l => horizontal ? l.Y1 : l.X2).First();

    [Fact]
    public void A_Word_chart_draws_axis_lines_and_ticks_black_unless_it_has_a_c14_style_or_states_the_line()
    {
        var plain = ChartLines(WithTicks("out"));
        Assert.Equal(Color.Black, LongLine(plain, true).Color);
        Assert.All(OfLength(plain, 3.02), l => Assert.Equal(Color.Black, l.Color));
        var c14 = ChartLines(WithC14Style(WithTicks("out")));
        Assert.Equal(Grey, LongLine(c14, true).Color); // measured: 898989 for c14:style 102 and 118
        Assert.Equal(Grey, LongLine(c14, false).Color);
        Assert.All(OfLength(c14, 3.02), l => Assert.Equal(Grey, l.Color));
        var stated = ChartXml().Replace("<c:crossAx val=\"2\"/>", "<c:spPr><a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln></c:spPr><c:crossAx val=\"2\"/>");
        var lines = ChartLines(stated.Replace("<c:crossAx val=\"2\"/>", "<c:majorTickMark val=\"out\"/><c:crossAx val=\"2\"/>").Replace("<c:crossAx val=\"1\"/>", "<c:majorTickMark val=\"out\"/><c:crossAx val=\"1\"/>"));
        var categoryAxis = LongLine(lines, true);
        Assert.Equal((new Color(255, 0, 0), 2.0), (categoryAxis.Color, categoryAxis.Width)); // the stated category axis line
        Assert.Equal(Color.Black, LongLine(lines, false).Color); // the value axis states none
        Assert.Equal(4, OfLength(lines, 3.02).Count(l => l.X1 == l.X2 && l.Color == new Color(255, 0, 0) && l.Width == 2.0)); // its ticks follow it
        Assert.All(OfLength(lines, 3.02).Where(l => l.Y1 == l.Y2), l => Assert.Equal(Color.Black, l.Color));
    }

    [Fact]
    public void A_Word_chart_without_chart_text_properties_has_an_18_pt_regular_title_as_measured_in_Word_2021()
    {
        var xml = ChartXml().Replace("<c:autoTitleDeleted val=\"1\"/>", "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:autoTitleDeleted val=\"0\"/>");
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), 360, 216)));
        var title = page.Items.OfType<TextItem>().First(t => t.Text == "Sales");
        Assert.Equal(18, title.FontSize, 3);
        Assert.False(title.SyntheticBold);
    }

    [Fact]
    public void A_Word_chart_title_stays_18_pt_regular_when_the_chart_text_properties_state_a_typeface()
    {
        var xml = ChartXml().Replace("<c:autoTitleDeleted val=\"1\"/>", "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:autoTitleDeleted val=\"0\"/>")
            .Replace("</c:chartSpace>", "<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"1000\"><a:latin typeface=\"Arial\"/></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>");
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), 360, 216)));
        var title = page.Items.OfType<TextItem>().First(t => t.Text == "Sales");
        Assert.Equal(18, title.FontSize, 3);
        Assert.False(title.SyntheticBold);
    }

    private static string TopRightLegendXml(bool title)
    {
        string ser(int i, double a, double b, double c) =>
            $"<c:ser><c:idx val=\"{i}\"/><c:order val=\"{i}\"/><c:tx><c:v>Ser{(char)(65 + i)}</c:v></c:tx>" + Values("val", a, b, c) + "</c:ser>";
        var xml = ChartXml().Replace("<c:ser>", "<c:SER>").Replace("</c:ser>", "</c:SER>");
        int s0 = xml.IndexOf("<c:SER>"), s1 = xml.IndexOf("</c:SER>") + "</c:SER>".Length;
        xml = xml[..s0] + ser(0, 3, 5, 4) + ser(1, 4, 4.5, 4.3) + ser(2, 5, 4, 4.6) + xml[s1..];
        xml = xml.Replace("<c:plotVisOnly", "<c:legend><c:legendPos val=\"tr\"/></c:legend><c:plotVisOnly");
        return title ? xml.Replace("<c:autoTitleDeleted val=\"1\"/>", "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title><c:autoTitleDeleted val=\"0\"/>") : xml;
    }

    private static List<string> ValueAxisLabels(string xml, double height, int count)
    {
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), 360, height)));
        return page.Items.OfType<TextItem>().Where(t => t.Text.Length == 1 && char.IsDigit(t.Text[0])).Select(t => t.Text).Take(count).ToList(); // the legend entries read "1", "2", "3" after the axis labels
    }

    [Fact]
    public void A_Word_value_axis_weighs_its_step_against_the_plot_the_top_right_legend_leaves_as_measured_in_Word_2021()
    {
        // 144 pt: the three entries do not fit above the plot, so it keeps 103.29 pt and Word shows steps of 1 (7 labels) without a title; with a title it keeps 73.6 pt and shows steps of 2.
        Assert.Equal(["0", "1", "2", "3", "4", "5", "6"], ValueAxisLabels(TopRightLegendXml(false), 144, 7));
        Assert.Equal(["0", "2", "4", "6"], ValueAxisLabels(TopRightLegendXml(true), 144, 4));
    }

    private static LayoutPage CategoryLabelPage(string label, int n, double width, int size = 10)
    {
        var xml = ChartXml()
            .Replace("<c:ptCount val=\"3\"/><c:pt idx=\"0\"><c:v>a</c:v></c:pt><c:pt idx=\"1\"><c:v>b</c:v></c:pt><c:pt idx=\"2\"><c:v>c</c:v></c:pt>",
                $"<c:ptCount val=\"{n}\"/>" + string.Concat(Enumerable.Range(0, n).Select(i => $"<c:pt idx=\"{i}\"><c:v>{label}</c:v></c:pt>")))
            .Replace("</c:chartSpace>", $"<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"{size * 100}\"><a:latin typeface=\"Arial\"/></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>");
        return Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), width, 216))).Page;
    }

    /// <summary>How far the plot's bottom stands above the chart's bottom (the plot's top is 11 below the chart's top).</summary>
    private static double PlotBottomOffset(LayoutPage page)
    {
        var axis = page.Items.OfType<LineItem>().Where(l => l.X1 == l.X2 && Math.Abs(l.Y2 - l.Y1) > 100).OrderBy(l => l.X1).First();
        return Math.Min(axis.Y1, axis.Y2) + 205 - Math.Max(axis.Y1, axis.Y2);
    }

    private static List<TextItem> Labels(LayoutPage page, string text) => page.Items.OfType<TextItem>().Where(t => t.Text == text).ToList();

    [Fact]
    public void Word_category_labels_stay_upright_until_the_widest_no_longer_fits_its_slot_as_measured_in_Word_2021_at_10_pt()
    {
        var wide = CategoryLabelPage("North", 10, 400);
        Assert.Equal(10, Labels(wide, "North").Count);
        Assert.DoesNotContain(wide.Items, i => i is PushItem { Rotation: not 0 });
        var narrow = CategoryLabelPage("North", 10, 200); // slot about 16 pt, label about 26: Word turns them rather than letting them touch
        Assert.All(narrow.Items.OfType<PushItem>().Where(p => p.Rotation != 0), p => Assert.Equal(45, p.Rotation));
        Assert.Equal(10, narrow.Items.OfType<PushItem>().Count(p => p.Rotation == 45));
    }

    [Fact]
    public void Word_wraps_category_labels_at_spaces_when_the_longest_word_fits_the_slot_and_keeps_the_last_baseline()
    {
        var upright = CategoryLabelPage("Northern", 3, 200);
        var page = CategoryLabelPage("Northern Region", 3, 200); // slot 53 pt: "Northern" fits, the whole label does not
        var first = Labels(page, "Northern"); var second = Labels(page, "Region");
        Assert.Equal(3, first.Count);
        Assert.Equal(3, second.Count);
        Assert.Equal(11.5, second[0].Baseline - first[0].Baseline, 3); // lines 1.15 em apart
        Assert.Equal(Labels(upright, "Northern")[0].Baseline, second[0].Baseline, 1); // the last baseline stays
        Assert.Equal(PlotBottomOffset(upright) + 11.5, PlotBottomOffset(page), 1); // the plot gives way by one line
        Assert.DoesNotContain(page.Items, i => i is PushItem { Rotation: not 0 });
    }

    [Fact]
    public void Word_turns_category_labels_by_45_degrees_when_the_slot_is_at_least_15_7_pt_and_sizes_the_plot_by_their_width()
    {
        var page = CategoryLabelPage("NorthernRegion", 6, 200); // slot 28 pt: a single word too wide to fit, wide enough to turn by 45
        var labels = page.Items.OfType<TextItem>().Where(t => t.Text == "NorthernRegion").ToList();
        Assert.Equal(6, labels.Count);
        Assert.Equal(6, page.Items.OfType<PushItem>().Count(p => p.Rotation == 45));
        double w = labels[0].Font.MeasureText("NorthernRegion", labels[0].FontSize);
        Assert.Equal(20.5 + 0.7071 * w, PlotBottomOffset(page), 0); // measured: 20.5 + 0.7071 W above the chart's bottom
    }

    [Fact]
    public void Word_turns_category_labels_by_45_degrees_at_14_pt_with_a_bottom_of_6_5_plus_1_4_em_plus_the_width()
    {
        var page = CategoryLabelPage("NorthernRegion", 6, 300, 14);
        var labels = page.Items.OfType<TextItem>().Where(t => t.Text == "NorthernRegion").ToList();
        Assert.Equal(6, page.Items.OfType<PushItem>().Count(p => p.Rotation == 45));
        double w = labels[0].Font.MeasureText("NorthernRegion", labels[0].FontSize);
        Assert.Equal(6.5 + 1.4 * 14 + 0.7071 * w + (0.555 * 14 + 5.05 - 11), PlotBottomOffset(page), 0); // PlotBottomOffset takes the plot top as 11; at 14 pt it is 12.82 // Word 2021 measured the gap at 8, 10 and 14 pt
    }

    [Theory]
    [InlineData(8, 18, true)]   // the test font sets 18 n at 89.0 pt at 8 pt: it fits the 90.8 pt a right legend allows (W / 3 less one line height, 9.2); a fixed 11.5 pt would break it
    [InlineData(14, 10, false)] // 86.5 pt at 14 pt does not fit the 83.9 pt (it would fit a fixed 88.5)
    public void A_Word_right_legend_wraps_a_name_at_a_third_of_the_chart_less_one_line_as_measured_in_Word_2021(int size, int count, bool fits)
    {
        string word = new('n', count);
        var xml = ChartXml().Replace("<c:v>S0</c:v>", $"<c:v>x {word}</c:v>").Replace("</c:plotArea>", "</c:plotArea><c:legend><c:legendPos val=\"r\"/></c:legend>")
            .Replace("</c:chartSpace>", $"<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"{size * 100}\"><a:latin typeface=\"Arial\"/></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>");
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), 300, 216)));
        Assert.Equal(fits, page.Items.OfType<TextItem>().Any(t => t.Text == word));
    }

    [Fact]
    public void Word_turns_category_labels_by_90_degrees_when_the_slot_is_between_11_1_and_15_7_pt()
    {
        var page = CategoryLabelPage("NorthernRegion", 24, 360); // slot 13.5 pt
        var labels = page.Items.OfType<TextItem>().Where(t => t.Text == "NorthernRegion").ToList();
        Assert.Equal(24, labels.Count);
        Assert.Equal(24, page.Items.OfType<PushItem>().Count(p => p.Rotation == 90));
        double w = labels[0].Font.MeasureText("NorthernRegion", labels[0].FontSize);
        Assert.Equal(12.6 + w, PlotBottomOffset(page), 0); // measured: 12.6 + W above the chart's bottom
    }

    [Theory]
    [InlineData("-5400000", 90)]
    [InlineData("5400000", -90)]
    public void Word_wraps_category_labels_asked_to_turn_by_90_degrees_at_half_the_height_less_12_as_measured_in_Word_2021(string rot, double angle)
    {
        // H 216: lines of at most 96 pt, so the 6 words wrap; the plot's bottom stands 12.51 + the widest line above the chart's bottom.
        var xml = ChartXml().Replace("<c:v>a</c:v>", "<c:v>Alpha beta gamma delta epsilon zeta</c:v>")
            .Replace("</c:catAx>", $"<c:txPr><a:bodyPr rot=\"{rot}\" vert=\"horz\"/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr></c:catAx>")
            .Replace("</c:chartSpace>", "<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"1000\"><a:latin typeface=\"Arial\"/></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>");
        var page = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), 360, 216))).Page;
        // Drawn a word at a time: the two lines are told apart by their baselines, 1.15 em apart.
        var words = page.Items.OfType<TextItem>().Where(t => "Alpha beta gamma delta epsilon zeta".Split(' ').Contains(t.Text)).ToList();
        var lines = words.GroupBy(t => Math.Round(t.Baseline, 1)).Select(g => string.Join(" ", g.OrderBy(t => t.X).Select(t => t.Text))).ToList();
        Assert.Equal(["Alpha beta gamma", "delta epsilon zeta"], lines.Order().ToList());
        Assert.Equal(4, page.Items.OfType<PushItem>().Count(p => p.Rotation == angle)); // a turn per line: 2 + "b" + "c"
        var font = words[0];
        double widest = lines.Max(l => font.Font.MeasureText(l, font.FontSize));
        Assert.Equal(12.51 + widest, PlotBottomOffset(page), 0);
    }

    [Fact]
    public void Word_shows_every_other_category_label_upright_when_that_shows_at_least_as_many_as_turning_would()
    {
        var page = CategoryLabelPage("A", 80, 360); // slot 4 pt: every 2nd label upright (40) beats every 3rd turned (27)
        Assert.Equal(40, Labels(page, "A").Count);
        Assert.DoesNotContain(page.Items, i => i is PushItem { Rotation: not 0 });
    }

    [Fact]
    public void Anchored_chart_is_placed_at_its_position()
    {
        var (page, warnings, _) = Run(b => P(Anchor(b.AddPart("charts/chart1.xml", ChartType, ChartXml()), 200, 150, Position("page", "254000", "page", "508000"))));
        Assert.DoesNotContain(warnings, w => w.Code.StartsWith("docx.chart"));
        var bars = Bars(page, Accent1);
        Assert.Equal(3, bars.Count);
        Assert.All(bars, r => Assert.True(r.X >= 20 && r.X + r.Width <= 220 && r.Y >= 40 && r.Y + r.Height <= 190.01));
    }

    [Fact]
    public void Chart_colors_come_from_the_documents_theme_part()
    {
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, ChartXml()), 300, 200)), $"{ThemeColors}");
        Assert.Equal(3, Bars(page, Color.TryParseHex("112233")!.Value).Count);
        Assert.Empty(Bars(page, Accent1));
    }

    [Fact]
    public void A_Word_chart_of_six_series_shades_the_automatic_colors_as_measured_in_Word_2016()
    {
        // Six series make two levels, the first shaded; Word 2016 draws accent 1 as 5089BC (not 5089BD) and accent 2 as D26E2A (not D36E2A).
        var one = ChartXml();
        var ser = one[one.IndexOf("<c:ser>", StringComparison.Ordinal)..(one.IndexOf("</c:ser>", StringComparison.Ordinal) + 8)];
        var six = one.Replace(ser, string.Concat(Enumerable.Range(0, 6).Select(i => ser.Replace("<c:idx val=\"0\"/><c:order val=\"0\"/>", $"<c:idx val=\"{i}\"/><c:order val=\"{i}\"/>"))));
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, six), 400, 250)));
        var fills = page.Items.OfType<RectItem>().Where(r => r.Width > 5 && r.Height > 5).Select(r => r.Fill).ToList();
        Assert.Contains(Color.TryParseHex("5089BC")!.Value, fills);
        Assert.Contains(Color.TryParseHex("D26E2A")!.Value, fills);
        Assert.DoesNotContain(Color.TryParseHex("5089BD")!.Value, fills);
        Assert.DoesNotContain(Color.TryParseHex("D36E2A")!.Value, fills);
    }

    [Fact]
    public void A_Word_chart_of_eleven_series_tints_the_second_level_by_0_23_as_measured_in_Word_2021()
    {
        var one = ChartXml();
        var ser = one[one.IndexOf("<c:ser>", StringComparison.Ordinal)..(one.IndexOf("</c:ser>", StringComparison.Ordinal) + 8)];
        var xml = one.Replace(ser, string.Concat(Enumerable.Range(0, 11).Select(i => ser.Replace("<c:idx val=\"0\"/><c:order val=\"0\"/>", $"<c:idx val=\"{i}\"/><c:order val=\"{i}\"/>"))));
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), 500, 250)));
        var fills = page.Items.OfType<RectItem>().Where(r => r.Width > 5 && r.Height > 5).Select(r => r.Fill).ToList();
        foreach (var hex in new[] { "5089BC", "D26E2A", "97B9E0", "F1A78A", "BFBFBF", "FFD184", "8FA2D4" }) Assert.Contains(Color.TryParseHex(hex)!.Value, fills);
    }

    [Fact]
    public void Without_a_theme_part_the_office_2013_accents_are_used()
    {
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, ChartXml()), 300, 200)));
        Assert.Equal(3, Bars(page, Accent1).Count);
    }

    [Fact]
    public void A_missing_chart_part_warns_and_leaves_the_extent_empty()
    {
        var (page, warnings, _) = Run(b => P(Inline(b.AddDanglingRelationship("charts/none.xml", ChartType), 300, 200)) + P(R("after", "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/>")));
        Assert.Contains(warnings, w => w.Code == "docx.chart-missing");
        Assert.Empty(page.Items.OfType<RectItem>());
        Assert.True(page.Items.OfType<TextItem>().First(t => t.Text.Contains("after")).Baseline > 72 + 200);
    }

    [Fact]
    public void A_missing_relationship_and_broken_xml_warn_without_throwing()
    {
        var (page, warnings, _) = Run(b => P(Inline("rId999", 100, 50)) + P(Inline(b.AddPart("charts/bad.xml", ChartType, "<c:chartSpace"), 100, 50)));
        Assert.Contains(warnings, w => w.Code == "docx.chart-missing");
        Assert.Contains(warnings, w => w.Code == "docx.chart-unreadable");
        Assert.Empty(page.Items.OfType<RectItem>());
    }

    [Fact]
    public void A_chart_without_cached_values_warns_and_draws_nothing()
    {
        var (page, warnings, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, ChartXml(cached: false)), 300, 200)));
        Assert.Contains(warnings, w => w.Code == "docx.chart-no-cache");
        Assert.Empty(page.Items.OfType<RectItem>());
    }
}
