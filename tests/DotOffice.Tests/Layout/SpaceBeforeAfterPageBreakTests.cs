using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Tests.Docx;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Layout;

/// <summary>
/// Space before on the paragraph that starts a page, from Word 2016 PDFs (24 pt before, Arial 12, top margin 72): dropped after a paragraph
/// ending in a page break (compatibility modes 15, 14, 12 and no settings part) and, in mode 15 only, with w:pageBreakBefore; kept for the
/// first paragraph of the document, after a next-page section break and after a column break.
/// </summary>
public class SpaceBeforeAfterPageBreakTests
{
    private const double Before = 24;
    private static readonly ParagraphProperties Spaced = ParagraphProperties.Default with { SpaceBefore = Before };

    /// <summary>Baseline of a first line at the top of a page with no space before.</summary>
    private static double Top => 72 + NotoLineHeight() - NotoDescent();

    private static Paragraph Breaker(string text = "")
    {
        var p = Para(text);
        p.Inlines.Add(new Break(BreakKind.Page, Text));
        return p;
    }

    private static double TargetBaseline(Document doc)
    {
        var layout = Render(doc);
        var line = layout.Pages.SelectMany(Lines).Single(l => l.Items.Any(i => i.Text == "target"));
        return line.Baseline;
    }

    private static Document WithSettings(DocumentSettings settings, params Block[] blocks)
    {
        var doc = new Document { Settings = settings };
        var section = new Section();
        section.Blocks.AddRange(blocks);
        doc.Sections.Add(section);
        return doc;
    }

    [Theory]
    [InlineData("text", true)]
    [InlineData("", true)]
    [InlineData("text", false)]
    [InlineData("", false)]
    public void Dropped_after_a_paragraph_ending_in_a_page_break_in_every_compatibility_mode(string text, bool mode15)
    {
        var settings = new DocumentSettings { PageBreakBeforeDropsSpaceBefore = mode15, TableEdgeAtIndent = mode15 };
        var doc = WithSettings(settings, Para("first"), Breaker(text), Para("target", Spaced));
        Assert.Equal(Top, TargetBaseline(doc), 0.001);   // measured: same position as with no space before
    }

    [Fact]
    public void Dropped_after_a_page_break_when_the_target_has_double_line_spacing()
    {
        var props = Spaced with { LineSpacing = new LineSpacing(2, LineSpacingRule.Auto) };
        Assert.Equal(Top, TargetBaseline(Doc(Breaker("text"), Para("target", props))), 0.001);
    }

    [Fact]
    public void Text_after_a_page_break_inside_the_paragraph_is_at_the_top_of_the_page()
    {
        var p = Para(Spaced, new TextRun("a", Text), new Break(BreakKind.Page, Text), new TextRun("target", Text));
        Assert.Equal(Top, TargetBaseline(Doc(p)), 0.001);
    }

    [Fact]
    public void Page_break_before_drops_it_in_compatibility_mode_15()
    {
        var doc = Doc(Para("first"), Para("target", Spaced with { PageBreakBefore = true }));
        Assert.True(doc.Settings.PageBreakBeforeDropsSpaceBefore);
        Assert.Equal(2, Render(doc).Pages.Count);
        Assert.Equal(Top, TargetBaseline(doc), 0.001);
    }

    [Fact]
    public void Page_break_before_keeps_it_in_older_compatibility_modes()
    {
        var doc = WithSettings(new DocumentSettings { PageBreakBeforeDropsSpaceBefore = false },
            Para("first"), Para("target", Spaced with { PageBreakBefore = true }));
        Assert.Equal(Top + Before, TargetBaseline(doc), 0.001);   // measured in modes 14 and 12: 24 pt lower
    }

    [Fact]
    public void Kept_for_the_first_paragraph_of_the_document()
    {
        Assert.Equal(Top + Before, TargetBaseline(Doc(Para("target", Spaced))), 0.001);
    }

    [Fact]
    public void Kept_after_a_next_page_section_break()
    {
        var doc = Doc(Para("first"));
        var second = new Section { Start = SectionStart.NextPage };
        second.Blocks.Add(Para("target", Spaced));
        doc.Sections.Add(second);
        Assert.Equal(2, Render(doc).Pages.Count);
        Assert.Equal(Top + Before, TargetBaseline(doc), 0.001);
    }

    [Fact]
    public void Kept_after_a_column_break()
    {
        var breaker = Para("first");
        breaker.Inlines.Add(new Break(BreakKind.Column, Text));
        var doc = Doc(breaker, Para("target", Spaced));
        Assert.Equal(2, Render(doc).Pages.Count);   // a column break in a one-column section starts a page
        Assert.True(TargetBaseline(doc) >= Top + Before - 0.001);
    }

    [Fact]
    public void Still_dropped_after_a_natural_page_break()
    {
        var section = new Section { PageHeight = 72 + 72 + 2 * NotoLineHeight() + 0.5 };
        var doc = Doc(section, Para("a"), Para("b"), Para("target", Spaced));
        Assert.Equal(2, Render(doc).Pages.Count);
        Assert.Equal(Top, TargetBaseline(doc), 0.001);
    }

    /// <summary>Not measured in Word, so left as before: the space stays.</summary>
    [Fact]
    public void Unmeasured_targets_keep_their_space_after_a_page_break()
    {
        var shaded = Spaced with { Shading = new Color(200, 200, 200) };
        var auto = Spaced with { SpaceBeforeAuto = true };
        var bordered = Spaced with { Borders = BorderSet.Empty with { Top = new Border(BorderStyle.Single, 0.5, Color.Black, 0) } };
        Assert.Equal(Top + Before, TargetBaseline(Doc(Breaker("text"), Para("target", shaded))), 0.001);
        Assert.Equal(Top + Before, TargetBaseline(Doc(Breaker("text"), Para("target", auto))), 0.001);
        Assert.True(TargetBaseline(Doc(Breaker("text"), Para("target", bordered))) >= Top + Before - 0.001);
        Assert.Equal(Top + Before, TargetBaseline(Doc(Para("first"), Para("target", shaded with { PageBreakBefore = true }))), 0.001);
    }

    [Theory]
    [InlineData("<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat>", true)]
    [InlineData("<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"14\"/></w:compat>", false)]
    [InlineData("<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"12\"/></w:compat>", false)]
    [InlineData(null, true)]   // no settings part: read as mode 15, where Word itself keeps the space (known difference)
    public void Docx_compatibility_mode_decides_the_page_break_before_rule(string? settings, bool drops)
    {
        var doc = DocxReader.Read(new DocxBuilder { Body = DocxBuilder.P(DocxBuilder.R("x")) + DocxBuilder.Section(), Settings = settings }.BuildStream());
        Assert.Equal(drops, doc.Settings.PageBreakBeforeDropsSpaceBefore);
    }
}
