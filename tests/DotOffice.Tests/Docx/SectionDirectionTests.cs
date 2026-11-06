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
}
