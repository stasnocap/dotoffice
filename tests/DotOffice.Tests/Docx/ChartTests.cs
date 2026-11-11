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
        Assert.Contains(page.Items, i => i is RectItem r && r.Fill == new Color(0x89, 0x89, 0x89) || i is LineItem l && l.Color == new Color(0x89, 0x89, 0x89) && Math.Abs(l.Width - 0.5) < 1e-6);
    }

    [Theory]
    [InlineData(216, "0,1,2,3,4,5,6")]
    [InlineData(118, "0,2,4,6")]
    public void A_Word_charts_value_axis_takes_the_finest_step_whose_ticks_are_14_pt_apart(double height, string expected)
    {
        // Data maximum 5: 0 to 6; the plot is 216 - 11 - 23.7 = 181 pt high (step 1 at 30 pt) or 83 pt (step 1 would put the ticks 13.9 pt apart).
        var xml = ChartXml().Replace("<c:v>20</c:v>", "<c:v>5</c:v>").Replace("<c:v>10</c:v>", "<c:v>3</c:v>");
        var (page, _, _) = Run(b => P(Inline(b.AddPart("charts/chart1.xml", ChartType, xml), 360, height)));
        var labels = page.Items.OfType<TextItem>().Where(t => t.Text.Length == 1 && char.IsDigit(t.Text[0]) && t.X < 72 + 30).OrderBy(t => t.Text).Select(t => t.Text);
        Assert.Equal(expected, string.Join(",", labels));
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
