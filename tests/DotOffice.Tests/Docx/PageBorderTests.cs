using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class PageBorderTests
{
    private static string Edge(string name, string attrs = "w:val=\"single\" w:sz=\"8\" w:space=\"24\" w:color=\"FF0000\"") => $"<w:{name} {attrs}/>";

    private static string Borders(string attrs, params string[] edges) => $"<w:pgBorders {attrs}>{string.Concat(edges)}</w:pgBorders>";

    private static (Document Doc, List<LayoutPage> Pages, List<ConversionWarning> Warnings) Layout(string pgBorders, string body = "")
    {
        var warnings = new List<ConversionWarning>();
        var doc = DocxReader.Read(new DocxBuilder { Body = P("one") + body + Section(extra: pgBorders) }.BuildStream(), new DocxReaderOptions { Warnings = new ListSink(warnings) });
        var pages = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages.ToList();
        return (doc, pages, warnings);
    }

    private sealed class ListSink(List<ConversionWarning> list) : IConversionWarningSink
    {
        public void Warn(ConversionWarning warning) => list.Add(warning);
    }

    private static List<LineItem> Lines(LayoutPage page) => page.Items.OfType<LineItem>().Where(l => l.Color == Color.TryParseHex("FF0000")).ToList();

    [Fact]
    public void Page_offset_places_every_edge_at_its_space_from_the_page_edge()
    {
        var page = Layout(Borders("w:offsetFrom=\"page\"", Edge("top"), Edge("left"), Edge("bottom"), Edge("right"))).Pages[0];
        var lines = Lines(page);
        Assert.Equal(4, lines.Count);
        var horizontal = lines.Where(l => l.Y1 == l.Y2).OrderBy(l => l.Y1).ToList();
        var vertical = lines.Where(l => l.X1 == l.X2).OrderBy(l => l.X1).ToList();
        Assert.Equal(24 + 0.5, horizontal[0].Y1, 0.01);
        Assert.Equal(page.Height - 24 - 0.5, horizontal[1].Y1, 0.01);
        Assert.Equal(24 + 0.5, vertical[0].X1, 0.01);
        Assert.Equal(page.Width - 24 - 0.5, vertical[1].X1, 0.01);
        // Corners meet: the horizontals reach the outer edge of the verticals.
        Assert.True(horizontal[0].X1 <= vertical[0].X1 && horizontal[0].X2 >= vertical[1].X1);
    }

    [Fact]
    public void Text_offset_surrounds_the_text_area_at_the_space_outside_the_margins()
    {
        var page = Layout(Borders("w:offsetFrom=\"text\"", Edge("top"), Edge("left"), Edge("bottom"), Edge("right"))).Pages[0];
        var lines = Lines(page);
        Assert.Equal(72 - 24 - 0.5, lines.First(l => l.Y1 == l.Y2).Y1, 0.01);
        Assert.Equal(72 - 24 - 0.5, lines.Where(l => l.X1 == l.X2).Min(l => l.X1), 0.01);
        Assert.Equal(page.Width - 72 + 24 + 0.5, lines.Where(l => l.X1 == l.X2).Max(l => l.X1), 0.01);
    }

    [Fact]
    public void Only_the_named_edges_are_drawn()
    {
        var lines = Lines(Layout(Borders("", Edge("bottom"))).Pages[0]);
        Assert.Single(lines);
        Assert.Equal(lines[0].Y1, lines[0].Y2);
    }

    [Fact]
    public void Display_first_page_and_not_first_page_select_pages()
    {
        string body = "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>" + P("two");
        var first = Layout(Borders("w:display=\"firstPage\"", Edge("top")), body).Pages;
        Assert.True(first.Count >= 2);
        Assert.NotEmpty(Lines(first[0]));
        Assert.Empty(Lines(first[1]));
        var rest = Layout(Borders("w:display=\"notFirstPage\"", Edge("top")), body).Pages;
        Assert.Empty(Lines(rest[0]));
        Assert.NotEmpty(Lines(rest[1]));
    }

    [Fact]
    public void Z_order_back_paints_before_text_and_front_after()
    {
        var back = Layout(Borders("w:zOrder=\"back\"", Edge("top"))).Pages[0];
        var front = Layout(Borders("w:zOrder=\"front\"", Edge("top"))).Pages[0];
        Assert.IsType<LineItem>(back.Items[0]);
        Assert.IsType<LineItem>(front.Items[^1]);
    }

    [Fact]
    public void An_art_border_is_a_single_line_with_a_warning()
    {
        var (_, pages, warnings) = Layout(Borders("", Edge("top", "w:val=\"apples\" w:sz=\"24\" w:space=\"24\" w:color=\"FF0000\"")));
        Assert.Single(Lines(pages[0]));
        Assert.Contains(warnings, w => w.Code == "docx.page-border-art");
    }

    [Fact]
    public void No_pgBorders_means_no_page_borders()
    {
        Assert.Null(Layout("").Doc.Sections[0].PageBorders);
    }
}
