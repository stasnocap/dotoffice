using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Layout;
using DotOffice.Tests.Docx;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Layout;

public class ColumnTests
{
    private static readonly double LineHeight =
        Math.Max(Noto.WinAscent + Noto.WinDescent, Noto.Ascender - Noto.Descender + Noto.LineGap) * 10.0 / Noto.UnitsPerEm;

    /// <summary>A section whose columns hold exactly <paramref name="lines"/> lines; content area is 468pt wide.</summary>
    private static Section Section(int lines, SectionColumns columns, SectionStart start = SectionStart.NextPage) => new()
    {
        PageHeight = 144 + LineHeight * lines + 0.5, Margins = new Thickness(72, 72, 72, 72), Columns = columns, Start = start,
    };

    private static Paragraph Lines(int count, string prefix)
    {
        var p = new Paragraph { Properties = ParagraphProperties.Default with { WidowControl = false }, MarkProperties = Text };
        for (int i = 1; i <= count; i++)
        {
            p.Inlines.Add(new TextRun($"{prefix}{i}", Text));
            if (i < count) p.Inlines.Add(new Break(BreakKind.Line, Text));
        }
        return p;
    }

    /// <summary>Text items by column x, top to bottom.</summary>
    private static Dictionary<double, List<string>> Columns(LayoutPage page) =>
        Texts(page).GroupBy(t => Math.Round(t.X)).OrderBy(g => g.Key)
                   .ToDictionary(g => g.Key, g => g.OrderBy(t => t.Baseline).Select(t => t.Text).ToList());

    private static SectionColumns Two(bool separator = false) => new() { Count = 2, Spacing = 36, Separator = separator };

    [Fact]
    public void Text_fills_columns_left_to_right_then_moves_to_the_next_page()
    {
        var layout = Render(Doc(Section(5, Two()), Lines(13, "L")));
        Assert.Equal(2, layout.Pages.Count);
        var first = Columns(layout.Pages[0]);
        Assert.Equal([72.0, 72 + 216 + 36], first.Keys);
        Assert.Equal(["L1", "L2", "L3", "L4", "L5"], first[72]);
        Assert.Equal(["L6", "L7", "L8", "L9", "L10"], first[324]);
        Assert.Equal(["L11", "L12", "L13"], Columns(layout.Pages[1])[72]);
    }

    [Fact]
    public void Footnotes_go_below_the_column_of_their_reference()
    {
        var note = new Note(NoteKind.Footnote) { Mark = "1", Blocks = { Para(Words(30)) } };
        var p = new Paragraph { MarkProperties = Text };
        p.Inlines.AddRange([new TextRun("b", Text), new NoteReference(note, Text)]);
        var page = Assert.Single(Render(Doc(Section(20, Two()), Lines(25, "L"), p)).Pages);

        var noteWords = Texts(page).Where(t => t.Text.StartsWith('w')).ToList();
        Assert.All(noteWords, t => Assert.True(t.X >= 324 - 0.01 && Right(t) <= 540 + 0.01, $"{t.Text} at {t.X}"));
        var separator = Assert.Single(page.Items.OfType<LineItem>());
        Assert.Equal(324, separator.X1, 1);
        // The first column is not shortened by the second column's footnote.
        double noteTop = noteWords.Min(t => t.Baseline);
        Assert.Contains(Texts(page), t => t.X < 300 && t.Baseline > noteTop);
    }

    [Fact]
    public void Paragraphs_wrap_at_the_column_width()
    {
        var layout = Render(Doc(Section(40, Two()), Para(Words(60))));
        Assert.All(Texts(layout.Pages[0]), t => Assert.True(Right(t) <= 72 + 216 + 0.01, $"{t.Text} ends at {Right(t)}"));
    }

    [Fact]
    public void Column_breaks_start_the_next_column_or_page()
    {
        var p = new Paragraph { MarkProperties = Text };
        p.Inlines.AddRange([new TextRun("a", Text), new Break(BreakKind.Column, Text), new TextRun("b", Text)]);
        var layout = Render(Doc(Section(5, Two()), p));
        Assert.Equal(["a"], Columns(layout.Pages[0])[72]);
        Assert.Equal(["b"], Columns(layout.Pages[0])[324]);

        var single = Render(Doc(Section(5, SectionColumns.Single), p));
        Assert.Equal(2, single.Pages.Count); // a column break in a one-column section acts as a page break
    }

    [Fact]
    public void Continuous_section_break_balances_columns_and_continues_below()
    {
        var doc = new Document();
        var columns = Section(20, Two());
        columns.Blocks.Add(Lines(7, "L"));
        var after = Section(20, SectionColumns.Single, SectionStart.Continuous);
        after.Blocks.Add(Para("after"));
        doc.Sections.Add(columns);
        doc.Sections.Add(after);

        var page = Assert.Single(Render(doc).Pages);
        var cols = Columns(page);
        Assert.Equal(["L1", "L2", "L3", "L4", "after"], cols[72]);
        Assert.Equal(["L5", "L6", "L7"], cols[324]);
        var l4 = Texts(page).Single(t => t.Text == "L4");
        var text = Texts(page).Single(t => t.Text == "after");
        Assert.Equal(l4.Baseline + LineHeight, text.Baseline, 1);
    }

    [Fact]
    public void Explicit_column_widths_and_separator_line()
    {
        var columns = new SectionColumns
        {
            Count = 2, Separator = true,
            Definitions = [new ColumnDefinition(100, 20), new ColumnDefinition(348, 0)],
        };
        var layout = Render(Doc(Section(3, columns), Lines(3, "A"), Para(Words(30))));
        var page = layout.Pages[0];
        Assert.Equal(["A1", "A2", "A3"], Columns(page)[72]);
        var wide = Texts(page).Where(t => t.X >= 192).ToList();
        Assert.NotEmpty(wide);
        Assert.Contains(wide, t => Right(t) > 72 + 120 + 200); // second paragraph uses the wide column
        var separator = Assert.Single(page.Items.OfType<LineItem>());
        Assert.Equal(72 + 110, separator.X1, 2);
    }

    [Fact]
    public void A_paragraph_continuing_in_a_column_of_another_width_is_broken_for_that_width()
    {
        var columns = new SectionColumns { Count = 2, Definitions = [new ColumnDefinition(100, 20), new ColumnDefinition(348, 0)] };
        var page = Assert.Single(Render(Doc(Section(4, columns), Para(Words(25)))).Pages);
        var narrow = Texts(page).Where(t => t.Text.StartsWith("word") && t.X < 192).ToList();
        var wide = Texts(page).Where(t => t.Text.StartsWith("word") && t.X >= 192).ToList();
        Assert.All(narrow, t => Assert.True(Right(t) <= 172 + 0.01));
        Assert.Contains(wide, t => Right(t) > 192 + 200); // wide lines, not narrow ones moved across
        Assert.Equal(25, narrow.Count + wide.Count); // no word lost or repeated
        Assert.True(wide.Select(t => t.Baseline).Distinct().Count() < 4);
    }

    [Fact]
    public void A_paragraph_moved_on_by_widow_control_is_broken_for_the_new_column()
    {
        // Three lines fill the narrow column but one: widow control moves the paragraph's first line on with the second.
        var columns = new SectionColumns { Count = 2, Definitions = [new ColumnDefinition(100, 20), new ColumnDefinition(348, 0)] };
        var p = Para(Words(25));
        p = new Paragraph { Properties = p.Properties with { WidowControl = true }, MarkProperties = Text, Inlines = { p.Inlines[0] } };
        var page = Assert.Single(Render(Doc(Section(4, columns), Lines(3, "A"), p)).Pages);
        var words = Texts(page).Where(t => t.Text.StartsWith("word")).ToList();
        Assert.All(words, t => Assert.True(t.X >= 192 - 0.01, $"{t.Text} at {t.X}"));
        Assert.Contains(words, t => Right(t) > 192 + 200); // laid out at the wide column's width
    }

    [Fact]
    public void Reader_parses_column_settings()
    {
        var doc = DocxReader.Read(new DocxBuilder
        {
            Body = DocxBuilder.P(DocxBuilder.R("x"), DocxBuilder.Section(extra: "<w:cols w:num=\"3\" w:space=\"360\" w:sep=\"1\"/>"))
                   + DocxBuilder.P(DocxBuilder.R("y"))
                   + DocxBuilder.Section(extra: "<w:type w:val=\"continuous\"/><w:cols w:num=\"2\" w:equalWidth=\"0\"><w:col w:w=\"2000\" w:space=\"500\"/><w:col w:w=\"6000\"/></w:cols>"),
        }.BuildStream());
        Assert.Equal((3, 18.0, true, 0), (doc.Sections[0].Columns.Count, doc.Sections[0].Columns.Spacing, doc.Sections[0].Columns.Separator,
            doc.Sections[0].Columns.Definitions.Count));
        Assert.Equal([new ColumnDefinition(100, 25), new ColumnDefinition(300, 0)], doc.Sections[1].Columns.Definitions);
        Assert.Equal(SectionStart.Continuous, doc.Sections[1].Start);
    }
}
