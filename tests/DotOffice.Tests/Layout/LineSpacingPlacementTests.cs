using DotOffice.Core;
using DotOffice.Layout;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Layout;

/// <summary>Where line spacing puts the baseline, from Word 2016 PDFs (Arial 12: single line 13.8 pt, top margin 72).</summary>
public class LineSpacingPlacementTests
{
    // Arial 12: ascent 10.85 and descent 2.95 (natural 13.8), as the layout model computes them.
    private const double Ascent = 10.85, Descent = 2.95, Natural = Ascent + Descent;

    [Theory]
    [InlineData(1.15, 15.87)]
    [InlineData(1.5, 20.7)]
    [InlineData(2, 27.6)]
    public void Auto_spacing_above_single_puts_the_extra_below_and_keeps_the_pitch(double factor, double pitch)
    {
        var (height, baseline) = ParagraphLayouter.LineHeight(new LineSpacing(factor, LineSpacingRule.Auto), Ascent, Descent);
        Assert.Equal(pitch, height, 0.05);
        Assert.Equal(Ascent, baseline, 0.001);          // measured first baseline 83.18 for all three factors (72 + ~11.2)
        Assert.Equal(83.18, 72 + baseline, 0.5);
    }

    [Theory]
    [InlineData(1.15)]
    [InlineData(1.5)]
    [InlineData(2)]
    public void Auto_spacing_above_single_first_baseline_is_the_single_one_on_the_page(double factor)
    {
        double single = FirstBaseline(new LineSpacing(1, LineSpacingRule.Auto));
        Assert.Equal(single, FirstBaseline(new LineSpacing(factor, LineSpacingRule.Auto)), 0.001);
    }

    [Fact]
    public void Mid_page_paragraph_at_1_5_starts_a_natural_line_below_a_single_one_and_has_pitch_1_5()
    {
        var single = ParagraphProperties.Default;
        var spaced = ParagraphProperties.Default with { LineSpacing = new LineSpacing(1.5, LineSpacingRule.Auto) };
        var page = Render(Doc(Para("one", single), Para(Words(60), spaced))).Pages.Single();
        var lines = Lines(page);
        double natural = NotoLineHeight();
        Assert.Equal(natural, lines[1].Baseline - lines[0].Baseline, 0.01);          // measured +13.8 for Arial 12
        Assert.Equal(natural * 1.5, lines[2].Baseline - lines[1].Baseline, 0.01);    // measured pitch 20.7
    }

    [Theory]
    [InlineData(20, 4.92)]
    [InlineData(10, -3.12)]
    public void Exact_spacing_moves_the_baseline_by_the_ascent_share_of_the_height_change(double exact, double measured)
    {
        var (height, baseline) = ParagraphLayouter.LineHeight(new LineSpacing(exact, LineSpacingRule.Exact), Ascent, Descent);
        Assert.Equal(exact, height, 0.001);
        Assert.Equal(measured, baseline - Ascent, 0.5);
        Assert.Equal(Ascent / Natural * (exact - Natural), baseline - Ascent, 0.001);
    }

    [Fact]
    public void Auto_spacing_below_single_moves_the_baseline_by_the_ascent_share()
    {
        var (height, baseline) = ParagraphLayouter.LineHeight(new LineSpacing(0.8, LineSpacingRule.Auto), Ascent, Descent);
        Assert.Equal(Natural * 0.8, height, 0.001);
        Assert.Equal(-2.28, baseline - Ascent, 0.5);
    }

    [Theory]
    [InlineData(20)]
    [InlineData(10)]
    public void At_least_spacing_still_puts_all_the_extra_above(double value)
    {
        var (height, baseline) = ParagraphLayouter.LineHeight(new LineSpacing(value, LineSpacingRule.AtLeast), Ascent, Descent);
        double expected = Math.Max(value, Natural);
        Assert.Equal(expected, height, 0.001);
        Assert.Equal(expected - Descent, baseline, 0.001);
        if (value == 20) Assert.Equal(6.24, baseline - Ascent, 0.5);   // measured +6.24
        else Assert.Equal(Ascent, baseline, 0.001);                    // measured no change
    }

    private static double FirstBaseline(LineSpacing spacing) =>
        Lines(Render(Doc(Para("x", ParagraphProperties.Default with { LineSpacing = spacing }))).Pages.Single())[0].Baseline;

    private static IEnumerable<string> PageLines(LayoutPage page) => Lines(page).Select(l => string.Concat(l.Items.Select(i => i.Text)));

    [Fact]
    public void Paragraph_holding_only_a_page_break_leaves_no_empty_line_on_the_next_page()
    {
        var breaker = Para(ParagraphProperties.Default, new Break(BreakKind.Page, Text));
        var withBreak = Render(Doc(Para("first"), breaker, Para("after"))).Pages;
        var before = Render(Doc(Para("first"), Para("after", ParagraphProperties.Default with { PageBreakBefore = true }))).Pages;
        Assert.Equal(2, withBreak.Count);
        Assert.Equal(["first"], PageLines(withBreak[0]));
        Assert.Equal(Lines(before[1])[0].Baseline, Lines(withBreak[1])[0].Baseline, 0.01);
    }

    [Fact]
    public void Page_break_paragraph_takes_its_line_on_the_first_page_and_a_final_one_adds_no_page()
    {
        // A page of exactly N lines: the break paragraph is the last one that fits, so the break happens after it, not before.
        var breaker = Para(ParagraphProperties.Default, new Break(BreakKind.Page, Text));
        var section = new Section { PageHeight = 72 + 72 + 3 * NotoLineHeight() + 0.5 };
        var layout = Render(Doc(section, Para("a"), Para("b"), breaker, Para("c")));
        Assert.Equal(2, layout.Pages.Count);
        Assert.Equal(["a", "b"], PageLines(layout.Pages[0]));
        Assert.Equal(["c"], PageLines(layout.Pages[1]));
        Assert.Equal(2, Render(Doc(Para("first"), breaker)).Pages.Count);   // the final break makes a blank last page
    }

    [Fact]
    public void Page_break_followed_by_text_in_the_same_paragraph_still_continues_on_the_next_page()
    {
        var p = Para(ParagraphProperties.Default, new Break(BreakKind.Page, Text), new TextRun("after", Text));
        var layout = Render(Doc(p));
        Assert.Equal(2, layout.Pages.Count);
        Assert.Equal(["after"], PageLines(layout.Pages[1]));
        Assert.Equal(72 + NotoLineHeight() - NotoDescent(), Lines(layout.Pages[1])[0].Baseline, 0.01);
    }

    [Fact]
    public void Page_break_after_text_leaves_no_empty_line_on_the_next_page()
    {
        var p = Para(ParagraphProperties.Default, new TextRun("before", Text), new Break(BreakKind.Page, Text));
        var layout = Render(Doc(p, Para("after")));
        Assert.Equal(["after"], PageLines(layout.Pages[1]));
        Assert.Equal(72 + NotoLineHeight() - NotoDescent(), Lines(layout.Pages[1])[0].Baseline, 0.01);   // Word 2016: first line of the page
    }

    [Fact]
    public void Two_page_breaks_in_one_paragraph_leave_one_blank_page()
    {
        var p = Para(ParagraphProperties.Default, new Break(BreakKind.Page, Text), new Break(BreakKind.Page, Text));
        var layout = Render(Doc(Para("a"), p, Para("b")));
        Assert.Equal(3, layout.Pages.Count);
        Assert.Equal(["b"], PageLines(layout.Pages[2]));
    }

    [Fact]
    public void Split_page_break_and_paragraph_mark_has_no_effect()
    {
        var breaker = Para(ParagraphProperties.Default, new Break(BreakKind.Page, Text));
        var doc = new Document();
        doc.Sections.Add(new Section());
        doc.Sections[0].Blocks.AddRange([Para("first"), breaker, Para("after")]);
        var layout = Render(doc);
        Assert.Equal(2, layout.Pages.Count);
        Assert.Equal(72 + NotoLineHeight() - NotoDescent(), Lines(layout.Pages[1])[0].Baseline, 0.01);
    }

    [Fact]
    public void Document_ending_in_a_page_break_paragraph_has_a_blank_last_page()
    {
        var breaker = Para(ParagraphProperties.Default, new Break(BreakKind.Page, Text));
        var layout = Render(Doc(Para("e1"), Para("e2"), breaker));
        Assert.Equal(2, layout.Pages.Count);
        Assert.Empty(Lines(layout.Pages[1]));
        var textBreak = Para(ParagraphProperties.Default, new TextRun("e2", Text), new Break(BreakKind.Page, Text));
        Assert.Equal(2, Render(Doc(Para("e1"), textBreak)).Pages.Count);
    }

    [Fact]
    public void Auto_spacing_above_single_fits_the_last_line_by_its_natural_height()
    {
        // Body of 3.5 natural lines: the second line at factor 2 starts at 2N, so it ends at 3N naturally but 4N with its extra space.
        var props = ParagraphProperties.Default with { LineSpacing = new LineSpacing(2, LineSpacingRule.Auto) };
        Section Small() => new() { PageHeight = 72 + 72 + 3.5 * NotoLineHeight() };
        Assert.Single(Render(Doc(Small(), Para("a", props), Para("b", props))).Pages);
        var three = Render(Doc(Small(), Para("a", props), Para("b", props), Para("c", props))).Pages;
        Assert.Equal(["a b", "c"], three.Select(p => string.Join(" ", PageLines(p))));
    }

    [Fact]
    public void Exact_and_at_least_spacing_need_their_full_height_at_the_page_bottom()
    {
        foreach (var rule in new[] { LineSpacingRule.Exact, LineSpacingRule.AtLeast })
        {
            var props = ParagraphProperties.Default with { LineSpacing = new LineSpacing(2 * NotoLineHeight(), rule) };
            var layout = Render(Doc(new Section { PageHeight = 72 + 72 + 3.5 * NotoLineHeight() }, Para("a", props), Para("b", props)));
            Assert.Equal(["a"], PageLines(layout.Pages[0]));
        }
    }
}
