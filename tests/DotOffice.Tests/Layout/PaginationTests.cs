using DotOffice.Core;
using DotOffice.Layout;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Layout;

public class PaginationTests
{
    /// <summary>A section whose body holds exactly <paramref name="lines"/> single-spaced lines of 10pt Noto Sans.</summary>
    private static Section LinesPerPage(int lines)
    {
        double lineHeight = Math.Max(Noto.WinAscent + Noto.WinDescent, Noto.Ascender - Noto.Descender + Noto.LineGap) * 10.0 / Noto.UnitsPerEm;
        double body = lineHeight * lines + 0.5;
        return new Section { PageHeight = 144 + body, Margins = new Thickness(72, 72, 72, 72) };
    }

    private static Paragraph Lines(int count, string prefix, ParagraphProperties? props = null)
    {
        var p = new Paragraph { Properties = props ?? ParagraphProperties.Default, MarkProperties = Text };
        for (int i = 1; i <= count; i++)
        {
            p.Inlines.Add(new TextRun($"{prefix}{i}", Text));
            if (i < count) p.Inlines.Add(new Break(BreakKind.Line, Text));
        }
        return p;
    }

    private static List<string> PageLines(LayoutPage page) =>
        LayoutTestHelpers.Lines(page).Select(l => string.Concat(l.Items.Select(i => i.Text))).ToList();

    [Fact]
    public void Long_text_flows_onto_new_pages()
    {
        var layout = Render(Doc(LinesPerPage(10), Lines(25, "L")));
        Assert.Equal([10, 10, 5], layout.Pages.Select(p => PageLines(p).Count));
        Assert.Equal("L11", PageLines(layout.Pages[1])[0]);
    }

    [Fact]
    public void Widow_control_moves_a_line_so_no_page_starts_with_a_lone_last_line()
    {
        var layout = Render(Doc(LinesPerPage(10), Lines(5, "A"), Lines(6, "B")));
        // 5 A-lines + 5 B-lines would fit, leaving B6 alone on page 2: B5 must move with it.
        Assert.Equal(["B5", "B6"], PageLines(layout.Pages[1]));
    }

    [Fact]
    public void Orphan_control_keeps_two_lines_at_the_bottom_or_moves_the_paragraph()
    {
        var layout = Render(Doc(LinesPerPage(10), Lines(9, "A"), Lines(4, "B")));
        Assert.Equal(9, PageLines(layout.Pages[0]).Count); // B1 alone would be an orphan
        Assert.Equal(["B1", "B2", "B3", "B4"], PageLines(layout.Pages[1]));

        var off = ParagraphProperties.Default with { WidowControl = false };
        var noControl = Render(Doc(LinesPerPage(10), Lines(9, "A", off), Lines(4, "B", off)));
        Assert.Equal("B1", PageLines(noControl.Pages[0])[^1]);
    }

    [Fact]
    public void Keep_with_next_moves_a_heading_to_the_next_page()
    {
        var heading = Para("Heading", ParagraphProperties.Default with { KeepWithNext = true });
        var layout = Render(Doc(LinesPerPage(10), Lines(9, "A"), heading, Lines(3, "B")));
        Assert.Equal(9, PageLines(layout.Pages[0]).Count);
        Assert.Equal("Heading", PageLines(layout.Pages[1])[0]);
    }

    [Fact]
    public void Keep_lines_together_and_page_break_before()
    {
        var keep = Render(Doc(LinesPerPage(10), Lines(7, "A"), Lines(5, "B", ParagraphProperties.Default with { KeepLinesTogether = true })));
        Assert.Equal(7, PageLines(keep.Pages[0]).Count);

        var breakBefore = Render(Doc(LinesPerPage(10), Lines(2, "A"), Lines(2, "B", ParagraphProperties.Default with { PageBreakBefore = true })));
        Assert.Equal(2, breakBefore.Pages.Count);
        Assert.Equal("B1", PageLines(breakBefore.Pages[1])[0]);
    }

    [Fact]
    public void Inline_page_break_continues_the_paragraph_on_a_new_page()
    {
        var p = Para(ParagraphProperties.Default, new TextRun("before", Text), new Break(BreakKind.Page, Text), new TextRun("after", Text));
        var layout = Render(Doc(p));
        Assert.Equal(2, layout.Pages.Count);
        Assert.Equal(["before"], PageLines(layout.Pages[0]));
        Assert.Equal(["after"], PageLines(layout.Pages[1]));
    }

    [Fact]
    public void Sections_set_page_size_and_restart_numbering()
    {
        var doc = new Document();
        var first = new Section();
        first.Blocks.Add(Para("portrait"));
        var footer = new HeaderFooter();
        footer.Blocks.Add(Para(ParagraphProperties.Default, new TextRun("p", Text), new FieldRun(FieldKind.Page, Text),
            new TextRun("/", Text), new FieldRun(FieldKind.NumPages, Text)));
        var landscape = new Section
        {
            PageWidth = 792, PageHeight = 612, Orientation = Orientation.Landscape, DefaultFooter = footer,
            PageNumberStart = 10, PageNumberFormat = PageNumberFormat.LowerRoman,
        };
        landscape.Blocks.Add(Para("landscape"));
        doc.Sections.Add(first);
        doc.Sections.Add(landscape);

        var layout = Render(doc);
        Assert.Equal((612.0, 792.0), (layout.Pages[0].Width, layout.Pages[0].Height));
        Assert.Equal((792.0, 612.0), (layout.Pages[1].Width, layout.Pages[1].Height));
        Assert.Contains("px/2", PageText(layout.Pages[1]));
    }

    [Fact]
    public void Headers_and_footers_get_page_numbers_and_first_page_variants()
    {
        var header = new HeaderFooter();
        header.Blocks.Add(Para("default header"));
        var firstHeader = new HeaderFooter();
        firstHeader.Blocks.Add(Para("first header"));
        var footer = new HeaderFooter();
        footer.Blocks.Add(Para(ParagraphProperties.Default with { Alignment = ParagraphAlignment.Right },
            new FieldRun(FieldKind.Page, Text), new TextRun(" of ", Text), new FieldRun(FieldKind.NumPages, Text)));
        var section = LinesPerPage(10);
        section = new Section
        {
            PageHeight = section.PageHeight, Margins = section.Margins, DifferentFirstPage = true,
            DefaultHeader = header, FirstHeader = firstHeader, DefaultFooter = footer, FirstFooter = footer,
        };
        var layout = Render(Doc(section, Lines(25, "L")));

        Assert.Equal(3, layout.Pages.Count);
        Assert.StartsWith("first header", PageText(layout.Pages[0]));
        Assert.StartsWith("default header", PageText(layout.Pages[1]));
        for (int i = 0; i < 3; i++)
        {
            var footerLine = LayoutTestHelpers.Lines(layout.Pages[i]).Last().Items;
            Assert.Equal($"{i + 1} of 3", string.Concat(footerLine.Select(t => t.Text)));
            Assert.Equal(72 + 468, Right(footerLine[^1]), 2); // right-aligned with the real number
        }
    }

    [Fact]
    public void Tall_header_pushes_the_body_down()
    {
        var header = new HeaderFooter();
        header.Blocks.Add(Lines(8, "H"));
        var layout = Render(Doc(new Section { DefaultHeader = header }, Para("body")));
        var body = LayoutTestHelpers.Lines(layout.Pages[0]).Single(l => l.Items[0].Text == "body");
        var lastHeader = LayoutTestHelpers.Lines(layout.Pages[0]).Single(l => l.Items[0].Text == "H8");
        Assert.True(body.Baseline > lastHeader.Baseline);
    }

    [Theory]
    [InlineData(SectionStart.EvenPage, 2)]  // page 2 is already even: no blank page
    [InlineData(SectionStart.OddPage, 3)]   // page 2 would be even: blank page 2, content on page 3
    public void Even_and_odd_page_sections(SectionStart start, int expectedPage)
    {
        var doc = new Document();
        var a = new Section();
        a.Blocks.Add(Para("one"));
        var b = new Section { Start = start };
        b.Blocks.Add(Para("two"));
        doc.Sections.Add(a);
        doc.Sections.Add(b);
        var layout = Render(doc);
        Assert.Equal(expectedPage, layout.Pages.Count);
        Assert.Equal("two", PageText(layout.Pages[expectedPage - 1]));
        if (expectedPage == 3) Assert.Empty(layout.Pages[1].Items);
    }

    [Fact]
    public void Continuous_section_stays_on_the_same_page()
    {
        var doc = new Document();
        var a = new Section();
        a.Blocks.Add(Para("one"));
        var b = new Section { Start = SectionStart.Continuous, Margins = new Thickness(72, 144, 72, 144) };
        b.Blocks.Add(Para("two"));
        doc.Sections.Add(a);
        doc.Sections.Add(b);
        var page = Render(doc).Pages.Single();
        Assert.Equal(144, Texts(page).Single(t => t.Text == "two").X, 2);
    }
}
