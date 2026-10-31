using System.Text;
using DotOffice.Converter;
using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using DotOffice.Tests.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;

namespace DotOffice.Tests.Docx;

public class FieldTests
{
    private static Document Read(DocxBuilder builder, IConversionWarningSink? warnings = null) =>
        DocxReader.Read(builder.BuildStream(), new DocxReaderOptions { Warnings = warnings });

    private static Paragraph[] Paragraphs(Document doc) => doc.Sections.SelectMany(s => s.Blocks).OfType<Paragraph>().ToArray();

    private static string Field(string instruction, string result, string rPr = "") =>
        $"<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> {instruction} </w:instrText></w:r>" +
        $"<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>{R(result, rPr)}<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    private static string Bookmark(string name, string content, int id = 1) =>
        $"<w:bookmarkStart w:id=\"{id}\" w:name=\"{name}\"/>{content}<w:bookmarkEnd w:id=\"{id}\"/>";

    private static LayoutDocument Layout(Document doc) =>
        new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(LayoutTestHelpers.Noto)) }).Layout(doc);

    [Theory]
    [InlineData("PAGE \\* roman \\* MERGEFORMAT", FieldKind.Page, PageNumberFormat.LowerRoman)]
    [InlineData("PAGE \\* ROMAN", FieldKind.Page, PageNumberFormat.UpperRoman)]
    [InlineData("NUMPAGES \\* alphabetic", FieldKind.NumPages, PageNumberFormat.LowerLetter)]
    [InlineData("SECTIONPAGES \\* Arabic", FieldKind.SectionPages, PageNumberFormat.Decimal)]
    [InlineData("PAGE \\* MERGEFORMAT", FieldKind.Page, null)]
    public void Parses_page_fields_and_number_format_switches(string instruction, FieldKind kind, PageNumberFormat? format)
    {
        var p = Paragraphs(Read(new DocxBuilder { Body = P(Field(instruction, "7")) }))[0];
        var field = Assert.IsType<FieldRun>(Assert.Single(p.Inlines));
        Assert.Equal((kind, format, "7"), (field.Kind, field.NumberFormat, field.CachedResult));
    }

    [Fact]
    public void Pageref_and_hyperlinks_carry_their_targets()
    {
        var b = new DocxBuilder();
        var url = b.AddHyperlink("https://example.com/a?b=1");
        b.Body = P(Bookmark("Target", R("Target heading")))
                 + P(Field("PAGEREF Target \\h", "3"))
                 + P($"<w:hyperlink r:id=\"{url}\">{R("external")}</w:hyperlink><w:hyperlink w:anchor=\"Target\">{R("internal")}</w:hyperlink>")
                 + P(Field("HYPERLINK \"https://example.org\" \\l \"part\"", "field link"));
        var ps = Paragraphs(Read(b));

        Assert.Equal("Target", Assert.IsType<BookmarkStart>(ps[0].Inlines[0]).Name);
        var pageRef = Assert.IsType<FieldRun>(Assert.Single(ps[1].Inlines));
        Assert.Equal((FieldKind.PageRef, "Target"), (pageRef.Kind, pageRef.Bookmark));
        Assert.Equal(new Hyperlink(null, "Target"), pageRef.Link);

        Assert.Equal(new Hyperlink("https://example.com/a?b=1"), ps[2].Inlines[0].Link);
        Assert.Equal(new Hyperlink(null, "Target"), ps[2].Inlines[1].Link);
        var fieldLink = Assert.IsType<TextRun>(Assert.Single(ps[3].Inlines));
        Assert.Equal(("field link", new Hyperlink("https://example.org", "part")), (fieldLink.Text, fieldLink.Link));
    }

    [Fact]
    public void Ref_fields_show_the_bookmarked_text_including_forward_references()
    {
        var warnings = new ConversionWarningList();
        var doc = Read(new DocxBuilder
        {
            Body = P(Field("REF Later \\h", "stale", "<w:b/>")) + P(Field("REF Missing", "cached text")) + P(Field("REF Nowhere", ""))
                   + P(R("See ") + Bookmark("Later", R("the conclusion") + R(" section", "<w:i/>"))),
        }, warnings);
        var ps = Paragraphs(doc);

        var resolved = Assert.IsType<TextRun>(Assert.Single(ps[0].Inlines));
        Assert.Equal(("the conclusion section", true), (resolved.Text, resolved.Properties.Bold));
        Assert.Equal(new Hyperlink(null, "Later"), resolved.Link);
        Assert.Equal("cached text", Assert.IsType<TextRun>(Assert.Single(ps[1].Inlines)).Text);
        Assert.StartsWith("Error!", Assert.IsType<TextRun>(Assert.Single(ps[2].Inlines)).Text);
        Assert.Contains(warnings.Warnings, w => w.Code == "docx.ref-missing");
    }

    [Fact]
    public void Unknown_fields_and_dates_show_their_cached_result()
    {
        var p = Paragraphs(Read(new DocxBuilder
        {
            Body = P(Field("DATE \\@ \"d MMMM yyyy\"", "1 January 2020") + Field("AUTHOR", "Someone") + Field("MERGEFIELD Name", "«Name»")),
        }))[0];
        Assert.Equal("1 January 2020Someone«Name»", string.Concat(p.Inlines.OfType<TextRun>().Select(t => t.Text)));
    }

    [Fact]
    public void Date_without_cached_result_is_evaluated_with_its_picture()
    {
        var p = Paragraphs(Read(new DocxBuilder
        {
            Body = P("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>DATE \\@ \"yyyy\"</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"),
        }))[0];
        Assert.Equal(DateTime.Now.Year.ToString(), Assert.IsType<TextRun>(Assert.Single(p.Inlines)).Text);
        Assert.Equal("Mon, 05 Feb 24 3:07 PM", FieldInstruction.FormatDate(new DateTime(2024, 2, 5, 15, 7, 0), "ddd, dd MMM yy h:mm AM/PM"));
    }

    [Fact]
    public void Toc_keeps_its_cached_entries_with_links_and_live_page_numbers()
    {
        // The shape Word writes: TOC field spanning several paragraphs, entries linking to _Toc bookmarks.
        string Entry(string text, string bookmark) =>
            $"<w:hyperlink w:anchor=\"{bookmark}\">{R(text)}<w:r><w:tab/></w:r>{Field($"PAGEREF {bookmark} \\h", "9")}</w:hyperlink>";
        var tabs = "<w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9350\"/></w:tabs>";
        var body = new StringBuilder();
        body.Append(P("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h \\z \\u </w:instrText></w:r>" +
                      "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + Entry("Chapter one", "_Toc1"), tabs));
        body.Append(P(Entry("Chapter two", "_Toc2") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>", tabs));
        body.Append(P(Bookmark("_Toc1", R("Chapter one")), "<w:pageBreakBefore/>"));
        body.Append(P(R("text")));
        body.Append(P(Bookmark("_Toc2", R("Chapter two"), 2), "<w:pageBreakBefore/>"));
        var doc = Read(new DocxBuilder { Body = body.ToString() });

        var toc = Paragraphs(doc)[..2];
        Assert.All(toc, p => Assert.Contains(p.Inlines, i => i is FieldRun { Kind: FieldKind.PageRef }));
        Assert.All(toc.SelectMany(p => p.Inlines), i => Assert.NotNull(i.Link));

        var layout = Layout(doc);
        Assert.Equal(3, layout.Pages.Count);
        var lines = LayoutTestHelpers.Lines(layout.Pages[0]);
        Assert.Equal("2", lines[0].Items[^1].Text);
        Assert.Equal("3", lines[1].Items[^1].Text);
        Assert.Equal(72 + 467.5, LayoutTestHelpers.Right(lines[1].Items[^1]), 1); // right tab aligned with the final number
        Assert.Equal(2, layout.Pages[0].Items.OfType<LinkItem>().Count());
        Assert.Equal("_Toc2", Assert.Single(layout.Pages[2].Items.OfType<AnchorItem>()).Name);
    }

    [Fact]
    public void Section_pages_and_formatted_page_numbers_are_evaluated_per_page()
    {
        var footer = new HeaderFooter();
        var text = LayoutTestHelpers.Text;
        footer.Blocks.Add(LayoutTestHelpers.Para(ParagraphProperties.Default, new FieldRun(FieldKind.Page, text) { NumberFormat = PageNumberFormat.UpperRoman },
            new TextRun("/", text), new FieldRun(FieldKind.SectionPages, text), new TextRun("/", text), new FieldRun(FieldKind.NumPages, text)));
        var doc = new Document();
        var first = new Section { DefaultFooter = footer };
        first.Blocks.Add(LayoutTestHelpers.Para("a"));
        first.Blocks.Add(new Paragraph { Properties = ParagraphProperties.Default with { PageBreakBefore = true } });
        var second = new Section { DefaultFooter = footer };
        second.Blocks.Add(LayoutTestHelpers.Para("b"));
        doc.Sections.Add(first);
        doc.Sections.Add(second);

        var pages = Layout(doc).Pages;
        Assert.Equal(["I/2/3", "II/2/3", "III/1/3"], pages.Select(p => string.Concat(LayoutTestHelpers.Lines(p).Last().Items.Select(t => t.Text))));
    }

    [Fact]
    public async Task Pdf_contains_uri_and_internal_link_annotations()
    {
        var b = new DocxBuilder();
        var url = b.AddHyperlink("https://example.com/");
        b.Body = P($"<w:hyperlink r:id=\"{url}\">{R("web")}</w:hyperlink> " + $"<w:hyperlink w:anchor=\"There\">{R("jump")}</w:hyperlink>")
                 + P(Bookmark("There", R("destination")), "<w:pageBreakBefore/>");
        using var output = new MemoryStream();
        await DocumentConverter.ConvertAsync(b.BuildStream(), output, DocumentFormat.Docx, DocumentFormat.Pdf,
            new ConvertOptions { Fonts = new FontCollection().Add(LayoutTestHelpers.Noto), UseSystemFonts = false });
        var pdf = Encoding.Latin1.GetString(output.ToArray());

        Assert.Contains("/Subtype /Link", pdf);
        Assert.Contains("/URI (https://example.com/)", pdf);
        Assert.Matches(@"/Dest \[\d+ 0 R /XYZ null [\d.]+ null\]", pdf);
    }
}
