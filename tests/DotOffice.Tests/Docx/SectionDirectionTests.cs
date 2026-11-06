using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class SectionDirectionTests
{
    private const string Font = "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/><w:sz w:val=\"28\"/>";
    private static readonly string Lorem = string.Join(' ', Enumerable.Repeat("lorem ipsum dolor sit amet", 60));

    private static LayoutDocument Layout(string direction, int paragraphs = 1, string? sectionExtra = null)
    {
        var body = string.Concat(Enumerable.Repeat(P(R(Lorem, Font)), paragraphs)) +
                   Section(extra: (direction.Length > 0 ? $"<w:textDirection w:val=\"{direction}\"/>" : "") + sectionExtra);
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc);
    }

    /// <summary>Where a point of the page content, laid out turned, ends up on the physical page.</summary>
    private static (double X, double Y) Physical(LayoutPage page, double x, double y)
    {
        var turn = page.Items.OfType<PushItem>().First();
        double u = x - turn.CenterX, v = y - turn.CenterY;
        double radians = turn.Rotation * Math.PI / 180;
        // The group turns counter-clockwise on the screen, where y grows downwards.
        return (turn.CenterX + u * Math.Cos(radians) + v * Math.Sin(radians), turn.CenterY - u * Math.Sin(radians) + v * Math.Cos(radians));
    }

    [Fact]
    public void Horizontal_section_is_not_turned()
    {
        var page = Layout("").Pages[0];
        Assert.DoesNotContain(page.Items, i => i is PushItem);
    }

    [Fact]
    public void Section_running_down_keeps_the_page_size_and_starts_at_the_top_of_the_right_margin()
    {
        var page = Layout("tbRl").Pages[0];
        Assert.Equal(612, page.Width, 0.5);
        Assert.Equal(792, page.Height, 0.5);
        Assert.Equal(-90, page.Items.OfType<PushItem>().First().Rotation);
        var first = Texts(page).First();
        var (x, y) = Physical(page, first.X, first.Baseline);
        Assert.Equal(72, y, 0.5);
        // The baseline of the first line is inside the one-inch right margin boundary, one line in from the edge.
        Assert.InRange(x, 612 - 72 - 20, 612 - 72);
    }

    [Fact]
    public void Lines_of_a_section_running_down_stack_from_right_to_left_and_stay_inside_the_margins()
    {
        var page = Layout("tbRl").Pages[0];
        var lines = Texts(page).Select(t => Physical(page, t.X, t.Baseline)).GroupBy(p => Math.Round(p.X, 1)).OrderByDescending(g => g.Key).ToList();
        Assert.True(lines.Count > 3);
        Assert.All(lines, g => Assert.All(g, p => Assert.InRange(p.Y, 72 - 0.5, 792 - 72 + 0.5)));
        Assert.All(lines, g => Assert.InRange(g.Key, 72 - 0.5, 612 - 72 + 0.5));
        // The lines are as long as the page is high between its margins, so they wrap there.
        Assert.True(lines[0].Max(p => p.Y) > 792 - 72 - 150);
    }

    [Fact]
    public void Section_running_up_starts_at_the_bottom_of_the_left_margin_and_stacks_from_left_to_right()
    {
        var page = Layout("btLr").Pages[0];
        Assert.Equal(90, page.Items.OfType<PushItem>().First().Rotation);
        var points = Texts(page).Select(t => Physical(page, t.X, t.Baseline)).ToList();
        Assert.Equal(792 - 72, points[0].Y, 0.5);
        Assert.InRange(points[0].X, 72, 72 + 20);
        var columns = points.Select(p => Math.Round(p.X, 1)).Distinct().ToList();
        Assert.True(columns.Count > 3);
        Assert.Equal(columns.OrderBy(c => c), columns);
    }

    [Fact]
    public void Lines_other_than_tbRl_and_btLr_stay_horizontal()
    {
        foreach (var direction in new[] { "lrTb", "lrTbV", "tbLrV" })
            Assert.DoesNotContain(Layout(direction).Pages[0].Items, i => i is PushItem);
    }

    [Fact]
    public void Text_that_overflows_the_lines_that_fit_across_the_page_continues_on_the_next_page()
    {
        var doc = Layout("tbRl", paragraphs: 12);
        Assert.True(doc.Pages.Count > 1);
        Assert.All(doc.Pages, p => Assert.Equal((612, 792), (p.Width, p.Height)));
    }

    [Fact]
    public void Section_after_a_vertical_one_is_horizontal_on_a_new_page()
    {
        var first = $"{P(R("vertical", Font), $"<w:sectPr><w:textDirection w:val=\"tbRl\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:type w:val=\"continuous\"/></w:sectPr>")}";
        var body = first + P(R("horizontal", Font)) + Section();
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        var pages = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages;
        Assert.Equal(2, pages.Count);
        Assert.Contains(pages[0].Items, i => i is PushItem);
        Assert.DoesNotContain(pages[1].Items, i => i is PushItem);
    }

    [Fact]
    public void Page_borders_of_a_section_running_down_stay_on_their_physical_edges()
    {
        // Margins: left 50, right 100, top and bottom 72. Distinct widths tell the edges apart; LibreOffice puts the box around the text, 10pt out.
        string borders = "<w:pgBorders w:offsetFrom=\"text\"><w:top w:val=\"single\" w:sz=\"8\" w:space=\"10\"/><w:left w:val=\"single\" w:sz=\"16\" w:space=\"10\"/>" +
                         "<w:bottom w:val=\"single\" w:sz=\"24\" w:space=\"10\"/><w:right w:val=\"single\" w:sz=\"32\" w:space=\"10\"/></w:pgBorders>";
        var body = P(R(Lorem, Font)) + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"2000\" w:bottom=\"1440\" w:left=\"1000\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>" +
                   borders + "<w:textDirection w:val=\"tbRl\"/></w:sectPr>";
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        var page = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
        double Centre(double width, bool x)
        {
            var line = page.Items.OfType<LineItem>().Single(l => Math.Abs(l.Width - width) < 0.01);
            var a = Physical(page, line.X1, line.Y1);
            return x ? a.X : a.Y;
        }
        Assert.Equal(72 - 10 - 0.5, Centre(1, false), 0.5);
        Assert.Equal(50 - 10 - 1, Centre(2, true), 0.5);
        Assert.Equal(792 - 72 + 10 + 1.5, Centre(3, false), 0.5);
        Assert.Equal(612 - 100 + 10 + 2, Centre(4, true), 0.5);
    }

    private const string TextBox =
        "<wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"/><a:solidFill><a:srgbClr val=\"FFFF00\"/></a:solidFill></wps:spPr>" +
        "<wps:txbx><w:txbxContent><w:p><w:r><w:t>BOX</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\"/>";

    private static LayoutPage FloatPage(string direction)
    {
        // A 100 x 60 pt box at offset 200 / 300 pt of the page, with text wrapped around it. The offsets are in the turned frame, as in LibreOffice.
        var body = P(AnchorShape(100, 60, Position("page", "2540000", "page", "3810000"), TextBox, "<wp:wrapSquare wrapText=\"bothSides\"/>") + R(Lorem, Font)) +
                   Section(extra: $"<w:textDirection w:val=\"{direction}\"/>");
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
    }

    // The physical box (x0, x1, y0, y1): 100 wide and 60 high, upright. tbRl: the turned rectangle hangs from the top right corner (312, 200), as in LibreOffice.
    private static (double X0, double X1, double Y0, double Y1) PhysicalBox(string direction) =>
        direction == "tbRl" ? (212, 312, 200, 260) : (300, 400, 532, 592);

    [Theory]
    [InlineData("tbRl")]
    [InlineData("btLr")]
    public void Floating_box_in_a_vertical_section_stays_upright(string direction)
    {
        var page = FloatPage(direction);
        var fill = Assert.Single(page.Items.OfType<RectItem>());
        Assert.Equal((100.0, 60.0), (fill.Width, fill.Height));
        var box = PhysicalBox(direction);
        var (x, y) = Physical(page, fill.X + fill.Width / 2, fill.Y + fill.Height / 2);
        Assert.Equal((box.X0 + box.X1) / 2, x, 0.5);
        Assert.Equal((box.Y0 + box.Y1) / 2, y, 0.5);
        // The box is drawn against the turn of the page.
        Assert.Equal(-page.Items.OfType<PushItem>().First().Rotation, page.Items.OfType<PushItem>().Skip(1).First().Rotation);
    }

    [Theory]
    [InlineData("tbRl")]
    [InlineData("btLr")]
    public void Text_of_a_vertical_section_wraps_around_a_floating_box(string direction)
    {
        var page = FloatPage(direction);
        var box = PhysicalBox(direction);
        var words = Texts(page).Where(t => t.Text.Trim().Length > 0 && !t.Text.Contains("BOX"))
            .Select(t => Physical(page, t.X, t.Baseline)).ToList();
        Assert.True(words.Count > 100);
        Assert.DoesNotContain(words, w => w.X > box.X0 + 1 && w.X < box.X1 - 1 && w.Y > box.Y0 + 1 && w.Y < box.Y1 - 1);
        // Text goes on the wider side of the box, as in a horizontal section; lines through its band start right at its edge.
        double edge = direction == "tbRl" ? box.Y1 : box.Y0;
        Assert.Contains(words, w => w.X > box.X0 + 1 && w.X < box.X1 - 1 && Math.Abs(w.Y - edge) < 1);
    }

    private static List<LayoutItem> Mixed(string textDirection)
    {
        var body = P(R("ab日本語cd", Font)) + Section(extra: $"<w:textDirection w:val=\"{textDirection}\"/>");
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0].Items;
    }

    [Fact]
    public void East_Asian_characters_in_text_running_down_are_turned_back_upright_and_Latin_is_not()
    {
        var items = Mixed("tbRl");
        var texts = items.OfType<TextItem>().ToList();
        Assert.Contains(texts, t => t.Text == "日");
        Assert.Contains(texts, t => t.Text == "語");
        foreach (var t in texts.Where(t => t.Text is "日" or "本" or "語"))
        {
            int at = items.IndexOf(t);
            var push = Assert.IsType<PushItem>(items[at - 1]);
            Assert.Equal(90, push.Rotation);
            Assert.Equal(t.X + t.Font.MeasureText(t.Text, t.FontSize) / 2, push.CenterX, 0.01);
            Assert.IsType<PopItem>(items[at + 1]);
        }
        Assert.All(texts.Where(t => t.Text.Contains('a') || t.Text.Contains('d')), t => Assert.False(items[items.IndexOf(t) - 1] is PushItem { Rotation: 90 }));
    }

    [Fact]
    public void Text_running_up_keeps_East_Asian_characters_turned_with_the_line()
    {
        Assert.DoesNotContain(Mixed("btLr"), i => i is PushItem { Rotation: -90 or 90 } p && p.CenterY != 396);
    }
}
