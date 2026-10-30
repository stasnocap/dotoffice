using DotOffice.Core;
using DotOffice.Layout;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Layout;

public class ParagraphLayoutTests
{
    private const double Left = 72, Width = 468; // Letter page with 1" margins

    [Fact]
    public void Wraps_text_within_the_column()
    {
        var page = Render(Doc(Para(Words(200)))).Pages.Single();
        var lines = Lines(page);
        Assert.True(lines.Count > 5);
        foreach (var (_, items) in lines)
        {
            Assert.True(items[0].X >= Left - 0.01);
            Assert.True(Right(items.Last(i => i.Text != " ")) <= Left + Width + 0.01);
        }
        // Words are never split across lines and stay in order (trailing spaces at line ends are not drawn).
        Assert.Equal(Words(200), string.Join(' ', Texts(page).Where(t => t.Text != " ").Select(t => t.Text)));
    }

    [Fact]
    public void Justified_lines_fill_the_column_except_the_last()
    {
        var page = Render(Doc(Para(Words(120), ParagraphProperties.Default with { Alignment = ParagraphAlignment.Justify }))).Pages.Single();
        var lines = Lines(page);
        foreach (var (_, items) in lines[..^1])
            Assert.Equal(Left + Width, Right(items.Last(i => i.Text != " ")), 2);
        Assert.True(Right(lines[^1].Items.Last(i => i.Text != " ")) < Left + Width - 20);
    }

    [Theory]
    [InlineData(ParagraphAlignment.Left, 0)]
    [InlineData(ParagraphAlignment.Center, 0.5)]
    [InlineData(ParagraphAlignment.Right, 1)]
    public void Aligns_lines(ParagraphAlignment alignment, double factor)
    {
        var page = Render(Doc(Para("Hello", ParagraphProperties.Default with { Alignment = alignment }))).Pages.Single();
        var t = Texts(page).Single();
        double w = Noto.MeasureText("Hello", 10);
        Assert.Equal(Left + (Width - w) * factor, t.X, 2);
    }

    [Fact]
    public void Applies_indents_and_first_line_offsets()
    {
        var props = ParagraphProperties.Default with { LeftIndent = 36, RightIndent = 72, FirstLineIndent = 18 };
        var lines = Lines(Render(Doc(Para(Words(80), props))).Pages.Single());
        Assert.Equal(Left + 54, lines[0].Items[0].X, 2);
        Assert.Equal(Left + 36, lines[1].Items[0].X, 2);
        Assert.All(lines, l => Assert.True(Right(l.Items.Last(i => i.Text != " ")) <= Left + Width - 72 + 0.01));

        var hanging = Lines(Render(Doc(Para(Words(80), props with { FirstLineIndent = -18 }))).Pages.Single());
        Assert.Equal(Left + 18, hanging[0].Items[0].X, 2);
    }

    [Fact]
    public void Default_and_custom_tab_stops()
    {
        var tabs = new[] { new TabStop(200, TabAlignment.Right), new TabStop(300, TabAlignment.Center), new TabStop(400, TabAlignment.Decimal) };
        var p = Para(ParagraphProperties.Default with { TabStops = tabs },
            new TextRun("a", Text), new Tab(Text), new TextRun("right", Text), new Tab(Text), new TextRun("mid", Text),
            new Tab(Text), new TextRun("12.5", Text));
        var items = Texts(Render(Doc(Para(ParagraphProperties.Default, new TextRun("x", Text), new Tab(Text), new TextRun("y", Text)), p)).Pages.Single()).ToList();

        Assert.Equal(Left + 36, items.Single(t => t.Text == "y").X, 2); // default stop every 36pt
        Assert.Equal(Left + 200, Right(items.Single(t => t.Text == "right")), 2);
        var mid = items.Single(t => t.Text == "mid");
        Assert.Equal(Left + 300, (mid.X + Right(mid)) / 2, 2);
        var dec = items.Single(t => t.Text == "12.5");
        Assert.Equal(Left + 400, dec.X + Noto.MeasureText("12", 10), 2);
    }

    [Fact]
    public void List_label_tabs_to_the_hanging_indent()
    {
        var p = new Paragraph
        {
            Properties = ParagraphProperties.Default with { LeftIndent = 36, FirstLineIndent = -18 },
            MarkProperties = Text,
            ListLabel = new ListLabel("1.", Text),
        };
        p.Inlines.Add(new TextRun(Words(60), Text));
        var lines = Lines(Render(Doc(p)).Pages.Single());
        Assert.Equal("1.", lines[0].Items[0].Text);
        Assert.Equal(Left + 18, lines[0].Items[0].X, 2);
        Assert.Equal(Left + 36, lines[0].Items[1].X, 2);  // text after the label starts at the indent
        Assert.Equal(Left + 36, lines[1].Items[0].X, 2);  // wrapped lines align with it
    }

    [Theory]
    [InlineData(LineSpacingRule.Exact, 30, 30)]
    [InlineData(LineSpacingRule.AtLeast, 30, 30)]
    [InlineData(LineSpacingRule.AtLeast, 2, -1)]
    [InlineData(LineSpacingRule.Auto, 2, -2)]
    public void Line_spacing_rules(LineSpacingRule rule, double value, double expected)
    {
        double natural = FontMetricsHeight();
        if (expected == -1) expected = natural;
        if (expected == -2) expected = natural * 2;
        var lines = Lines(Render(Doc(Para(Words(60), ParagraphProperties.Default with { LineSpacing = new LineSpacing(value, rule) }))).Pages.Single());
        Assert.Equal(expected, lines[1].Baseline - lines[0].Baseline, 2);
    }

    private static double FontMetricsHeight()
    {
        double scale = 10.0 / Noto.UnitsPerEm;
        double win = Noto.WinAscent + Noto.WinDescent, hhea = Noto.Ascender - Noto.Descender + Noto.LineGap;
        return Math.Max(win, hhea) * scale;
    }

    [Fact]
    public void Spacing_before_and_after_add_up_and_contextual_spacing_removes_them()
    {
        var props = ParagraphProperties.Default with { SpaceBefore = 10, SpaceAfter = 20, StyleId = "S" };
        var plain = Lines(Render(Doc(Para("a", props), Para("b", props))).Pages.Single());
        double gap = plain[1].Baseline - plain[0].Baseline;
        Assert.Equal(FontMetricsHeight() + 30, gap, 2);

        var contextual = props with { ContextualSpacing = true };
        var lines = Lines(Render(Doc(Para("a", contextual), Para("b", contextual))).Pages.Single());
        Assert.Equal(FontMetricsHeight(), lines[1].Baseline - lines[0].Baseline, 2);
    }

    [Fact]
    public void Breaks_overlong_words_instead_of_overflowing()
    {
        var page = Render(Doc(Para(new string('W', 400)))).Pages.Single();
        var lines = Lines(page);
        Assert.True(lines.Count > 1);
        Assert.All(lines, l => Assert.True(Right(l.Items[^1]) <= Left + Width + 0.01));
        Assert.Equal(400, PageText(page).Length);
    }

    [Fact]
    public void Draws_decorations_highlight_and_images()
    {
        var run = Text with { Underline = UnderlineStyle.Single, Strike = true, Highlight = new Color(255, 255, 0), Color = new Color(200, 0, 0) };
        var image = new ImageData([0x89, (byte)'P', (byte)'N', (byte)'G']);
        var page = Render(Doc(Para(ParagraphProperties.Default, new TextRun("decorated", run), new ImageRun(image, 50, 40, Text)))).Pages.Single();
        var text = Texts(page).Single();
        Assert.Equal(new Color(200, 0, 0), text.Color);
        Assert.Equal(2, page.Items.OfType<LineItem>().Count()); // underline + strike
        var highlight = page.Items.OfType<RectItem>().Single();
        Assert.True(page.Items.IndexOf(highlight) < page.Items.IndexOf(text)); // background painted first
        var img = page.Items.OfType<ImageItem>().Single();
        Assert.Equal((50.0, 40.0), (img.Width, img.Height));
        Assert.Equal(text.Baseline, img.Y + img.Height, 2); // sits on the baseline
    }

    [Fact]
    public void Synthesizes_missing_bold_and_italic()
    {
        var page = Render(Doc(Para("x", run: Text with { Bold = true, Italic = true }))).Pages.Single();
        var t = Texts(page).Single();
        Assert.True(t.SyntheticBold && t.SyntheticItalic);
    }

    [Fact]
    public void Superscript_is_smaller_and_raised()
    {
        var page = Render(Doc(Para(ParagraphProperties.Default, new TextRun("x", Text),
            new TextRun("2", Text with { VerticalPosition = VerticalPosition.Superscript })))).Pages.Single();
        var (x, sup) = (Texts(page).First(), Texts(page).Last());
        Assert.True(sup.FontSize < x.FontSize);
        Assert.True(sup.Baseline < x.Baseline);
    }

    [Fact]
    public void Missing_glyphs_and_unknown_fonts_produce_warnings()
    {
        var warnings = new ConversionWarningList();
        Render(Doc(Para("emoji \U0001F600", run: Text with { FontFamily = "Calibri" })), warnings);
        Assert.Contains(warnings.Warnings, w => w.Code == "font.substituted");
        Assert.Contains(warnings.Warnings, w => w.Code == "font.missing-glyph");
    }

    [Fact]
    public void Paragraph_borders_and_shading_are_drawn_around_groups()
    {
        var border = new Border(BorderStyle.Single, 1, Color.Black, 2);
        var props = ParagraphProperties.Default with
        {
            Borders = new BorderSet { Top = border, Bottom = border, Left = border, Right = border },
            Shading = new Color(230, 230, 230),
        };
        var page = Render(Doc(Para("one", props), Para("two", props))).Pages.Single();
        var horizontal = page.Items.OfType<LineItem>().Where(l => Math.Abs(l.Y1 - l.Y2) < 0.01).ToList();
        Assert.Equal(2, horizontal.Count); // grouped: one top border, one bottom border
        Assert.Equal(2, page.Items.OfType<RectItem>().Count());
    }
}