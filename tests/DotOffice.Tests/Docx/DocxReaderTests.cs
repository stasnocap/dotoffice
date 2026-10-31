using System.IO.Compression;
using DotOffice.Core;
using DotOffice.Docx;
using static DotOffice.Tests.Docx.DocxBuilder;

namespace DotOffice.Tests.Docx;

public class DocxReaderTests
{
    private const string Theme =
        "<a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/></a:majorFont>" +
        "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/></a:minorFont></a:fontScheme>";

    private const string Styles =
        "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault>" +
        "<w:pPrDefault><w:pPr><w:spacing w:after=\"160\" w:line=\"259\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults>" +
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:basedOn w:val=\"Normal\"/><w:pPr><w:keepNext/><w:spacing w:before=\"240\" w:after=\"0\"/></w:pPr>" +
        "<w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:hAnsiTheme=\"majorHAnsi\"/><w:b/><w:color w:val=\"2F5496\"/><w:sz w:val=\"32\"/></w:rPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:basedOn w:val=\"Heading1\"/><w:rPr><w:i/><w:sz w:val=\"26\"/></w:rPr></w:style>" +
        "<w:style w:type=\"character\" w:styleId=\"Strong\"><w:rPr><w:b/></w:rPr></w:style>" +
        "<w:style w:type=\"character\" w:styleId=\"Emph\"><w:basedOn w:val=\"Strong\"/><w:rPr><w:i/></w:rPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Loop1\"><w:basedOn w:val=\"Loop2\"/></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Loop2\"><w:basedOn w:val=\"Loop1\"/><w:rPr><w:sz w:val=\"40\"/></w:rPr></w:style>";

    private static Document Read(DocxBuilder builder, IConversionWarningSink? warnings = null) =>
        DocxReader.Read(builder.BuildStream(), new DocxReaderOptions { Warnings = warnings });

    private static Paragraph[] Paragraphs(Document doc) => doc.Sections.SelectMany(s => s.Blocks).OfType<Paragraph>().ToArray();

    private static TextRun Run(Paragraph p, int index = 0) => p.Inlines.OfType<TextRun>().ElementAt(index);

    [Fact]
    public void Resolves_styles_through_basedOn_chain_and_theme_fonts()
    {
        var doc = Read(new DocxBuilder
        {
            Theme = Theme, Styles = Styles,
            Body = P(R("Body")) + P(R("Title"), "<w:pStyle w:val=\"Heading1\"/>") + P(R("Sub"), "<w:pStyle w:val=\"Heading2\"/>"),
        });
        var ps = Paragraphs(doc);

        var body = Run(ps[0]).Properties;
        Assert.Equal(("Calibri", 11.0, false), (body.FontFamily, body.FontSize, body.Bold));
        Assert.Equal(8, ps[0].Properties.SpaceAfter);
        Assert.Equal(259 / 240.0, ps[0].Properties.LineSpacing.Value, 6);

        var h1 = Run(ps[1]).Properties;
        Assert.Equal(("Calibri Light", 16.0, true, false), (h1.FontFamily, h1.FontSize, h1.Bold, h1.Italic));
        Assert.Equal(new Color(0x2F, 0x54, 0x96), h1.Color);
        Assert.True(ps[1].Properties.KeepWithNext);
        Assert.Equal((12.0, 0.0), (ps[1].Properties.SpaceBefore, ps[1].Properties.SpaceAfter));

        var h2 = Run(ps[2]).Properties; // inherits bold/color/font from Heading1
        Assert.Equal((13.0, true, true), (h2.FontSize, h2.Bold, h2.Italic));
        Assert.Equal("Heading2", ps[2].Properties.StyleId);
    }

    [Fact]
    public void Toggle_properties_flip_across_style_levels_and_are_absolute_when_direct()
    {
        var doc = Read(new DocxBuilder
        {
            Styles = Styles,
            Body = P(R("a", "<w:rStyle w:val=\"Strong\"/>") + R("b", "<w:rStyle w:val=\"Strong\"/><w:b/>") + R("c", "<w:b w:val=\"0\"/>")
                     + R("d", "<w:rStyle w:val=\"Emph\"/>"), "<w:pStyle w:val=\"Heading1\"/>")
                   + P(R("e", "<w:rStyle w:val=\"Strong\"/>") + R("f", "<w:b w:val=\"false\"/><w:i/>")),
        });
        var ps = Paragraphs(doc);
        // Bold paragraph style + bold character style toggles bold off.
        Assert.False(Run(ps[0], 0).Properties.Bold);
        Assert.True(Run(ps[0], 1).Properties.Bold);   // direct formatting is absolute
        Assert.False(Run(ps[0], 2).Properties.Bold);
        Assert.False(Run(ps[0], 3).Properties.Bold);  // Emph inherits bold from Strong (basedOn merges, no toggle)
        Assert.True(Run(ps[0], 3).Properties.Italic);
        Assert.True(Run(ps[1], 0).Properties.Bold);
        Assert.Equal((false, true), (Run(ps[1], 1).Properties.Bold, Run(ps[1], 1).Properties.Italic));
    }

    [Fact]
    public void BasedOn_cycles_do_not_hang()
    {
        var doc = Read(new DocxBuilder { Styles = Styles, Body = P(R("x"), "<w:pStyle w:val=\"Loop1\"/>") });
        Assert.Equal(20, Run(Paragraphs(doc)[0]).Properties.FontSize);
    }

    [Fact]
    public void Direct_paragraph_formatting_and_tabs_merge_over_styles()
    {
        var styles = Styles + "<w:style w:type=\"paragraph\" w:styleId=\"Tabbed\"><w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"1440\"/>" +
                     "<w:tab w:val=\"right\" w:pos=\"9000\" w:leader=\"dot\"/></w:tabs><w:ind w:left=\"720\" w:firstLine=\"360\"/></w:pPr></w:style>";
        var doc = Read(new DocxBuilder
        {
            Styles = styles,
            Body = P(R("x"), "<w:pStyle w:val=\"Tabbed\"/><w:tabs><w:tab w:val=\"clear\" w:pos=\"1440\"/><w:tab w:val=\"center\" w:pos=\"4320\"/></w:tabs>" +
                             "<w:ind w:hanging=\"360\"/><w:jc w:val=\"both\"/><w:spacing w:line=\"300\" w:lineRule=\"exact\"/>" +
                             "<w:pBdr><w:top w:val=\"single\" w:sz=\"8\" w:space=\"1\" w:color=\"FF0000\"/></w:pBdr><w:shd w:val=\"clear\" w:fill=\"EEEEEE\"/>"),
        });
        var p = Paragraphs(doc)[0].Properties;
        Assert.Equal([new TabStop(216, TabAlignment.Center), new TabStop(450, TabAlignment.Right, TabLeader.Dot)], p.TabStops);
        Assert.Equal((36.0, -18.0), (p.LeftIndent, p.FirstLineIndent));
        Assert.Equal(ParagraphAlignment.Justify, p.Alignment);
        Assert.Equal(new LineSpacing(15, LineSpacingRule.Exact), p.LineSpacing);
        Assert.Equal(new Border(BorderStyle.Single, 1, new Color(255, 0, 0), 1), p.Borders.Top);
        Assert.Equal(new Color(0xEE, 0xEE, 0xEE), p.Shading);
    }

    private const string ListNumbering =
        "<w:abstractNum w:abstractNumId=\"0\">" +
        "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl>" +
        "<w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerLetter\"/><w:lvlText w:val=\"%1.%2)\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"360\"/></w:pPr></w:lvl>" +
        "<w:lvl w:ilvl=\"2\"><w:start w:val=\"1\"/><w:numFmt w:val=\"upperRoman\"/><w:lvlText w:val=\"%3\"/><w:lvlRestart w:val=\"1\"/></w:lvl>" +
        "</w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/>" +
        "<w:rPr><w:rFonts w:ascii=\"Symbol\" w:hAnsi=\"Symbol\" w:hint=\"default\"/></w:rPr></w:lvl>" +
        "<w:lvl w:ilvl=\"1\"><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"o\"/><w:suff w:val=\"space\"/><w:rPr><w:rFonts w:ascii=\"Courier New\"/></w:rPr></w:lvl></w:abstractNum>" +
        "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>" +
        "<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>" +
        "<w:num w:numId=\"3\"><w:abstractNumId w:val=\"0\"/><w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"5\"/></w:lvlOverride></w:num>" +
        "<w:num w:numId=\"4\"><w:abstractNumId w:val=\"0\"/></w:num>";

    private static string Item(int numId, int ilvl, string text) =>
        P(R(text), $"<w:numPr><w:ilvl w:val=\"{ilvl}\"/><w:numId w:val=\"{numId}\"/></w:numPr>");

    [Fact]
    public void Numbers_nested_lists_with_restart_semantics()
    {
        var doc = Read(new DocxBuilder
        {
            Numbering = ListNumbering,
            Body = Item(1, 0, "a") + Item(1, 1, "b") + Item(1, 1, "c") + Item(1, 2, "d") + Item(1, 1, "e") + Item(1, 2, "f")
                   + Item(1, 0, "g") + Item(1, 1, "h") + Item(1, 2, "i") + P(R("plain")) + Item(4, 0, "j") + Item(3, 0, "k") + Item(3, 0, "l"),
        });
        var labels = Paragraphs(doc).Select(p => p.ListLabel?.Text).ToArray();
        // Level 2 has lvlRestart=1: using level 1 ("e") does not restart it, using level 0 ("g") does.
        // num 4 shares abstractNum 0 with num 1 and continues it; num 3 overrides the start.
        Assert.Equal(new[] { "1.", "1.a)", "1.b)", "I", "1.c)", "II", "2.", "2.a)", "I", null, "3.", "5.", "6." }, labels);

        var first = Paragraphs(doc)[0].Properties;
        Assert.Equal((36.0, -18.0), (first.LeftIndent, first.FirstLineIndent));
        Assert.Equal(ListLabelSuffix.Tab, Paragraphs(doc)[0].ListLabel!.Suffix);
    }

    [Fact]
    public void Symbol_font_bullets_are_mapped_to_unicode()
    {
        var doc = Read(new DocxBuilder { Numbering = ListNumbering, Body = Item(2, 0, "x") + Item(2, 1, "y") });
        var ps = Paragraphs(doc);
        Assert.Equal("•", ps[0].ListLabel!.Text);
        Assert.NotEqual("Symbol", ps[0].ListLabel!.Properties.FontFamily);
        Assert.Equal(("o", "Courier New", ListLabelSuffix.Space),
            (ps[1].ListLabel!.Text, ps[1].ListLabel!.Properties.FontFamily, ps[1].ListLabel!.Suffix));
    }

    [Fact]
    public void Numbering_attached_through_paragraph_style()
    {
        var styles = Styles + "<w:style w:type=\"paragraph\" w:styleId=\"ListNum\"><w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr><w:ind w:left=\"360\"/></w:pPr></w:style>";
        var doc = Read(new DocxBuilder
        {
            Styles = styles, Numbering = ListNumbering,
            Body = P(R("a"), "<w:pStyle w:val=\"ListNum\"/>") + P(R("b"), "<w:pStyle w:val=\"ListNum\"/>"),
        });
        var ps = Paragraphs(doc);
        Assert.Equal(["1.", "2."], ps.Select(p => p.ListLabel?.Text));
        Assert.Equal(18, ps[0].Properties.LeftIndent); // style indentation wins over numbering from the style
    }

    [Theory]
    [InlineData(1, "decimal", "1")]
    [InlineData(4, "upperRoman", "IV")]
    [InlineData(1994, "lowerRoman", "mcmxciv")]
    [InlineData(28, "upperLetter", "BB")]
    [InlineData(3, "lowerLetter", "c")]
    [InlineData(7, "decimalZero", "07")]
    [InlineData(22, "ordinal", "22nd")]
    [InlineData(5, "none", "")]
    public void Formats_numbers(int value, string format, string expected) =>
        Assert.Equal(expected, Numbering.FormatNumber(value, format));

    [Fact]
    public void Reads_sections_headers_footers_and_page_fields()
    {
        var b = new DocxBuilder();
        var header = b.AddHeader(P(R("Header text")));
        var firstHeader = b.AddHeader(P(R("First page header")));
        var footer = b.AddFooter(P(R("Page ") + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> PAGE \\* MERGEFORMAT </w:instrText></w:r>"
                                     + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
                                     + R(" of ") + "<w:fldSimple w:instr=\"NUMPAGES\"><w:r><w:t>9</w:t></w:r></w:fldSimple>"
                                     + "<w:fldSimple w:instr=\"DATE\"><w:r><w:t>2020-01-01</w:t></w:r></w:fldSimple>"));
        b.Body = P(R("Portrait"), Section(extra: $"<w:headerReference w:type=\"default\" r:id=\"{header}\"/><w:headerReference w:type=\"first\" r:id=\"{firstHeader}\"/>" +
                                                 $"<w:footerReference w:type=\"default\" r:id=\"{footer}\"/><w:titlePg/>"))
                 + P(R("Landscape")) + Section(15840, 12240, "<w:type w:val=\"nextPage\"/>");
        var doc = Read(b);

        Assert.Equal(2, doc.Sections.Count);
        var (s1, s2) = (doc.Sections[0], doc.Sections[1]);
        Assert.Equal((612.0, 792.0, Orientation.Portrait, true), (s1.PageWidth, s1.PageHeight, s1.Orientation, s1.DifferentFirstPage));
        Assert.Equal((792.0, 612.0, Orientation.Landscape), (s2.PageWidth, s2.PageHeight, s2.Orientation));
        Assert.Equal(new Thickness(72, 72, 72, 72), s1.Margins);
        Assert.Equal(36, s1.HeaderDistance);
        Assert.Single(s1.Blocks);
        Assert.Same(s1.DefaultHeader, s2.DefaultHeader); // inherited
        Assert.Same(s1.DefaultFooter, s2.DefaultFooter);
        Assert.Equal("First page header", Run((Paragraph)s1.FirstHeader!.Blocks[0]).Text);

        var footerInlines = ((Paragraph)s1.DefaultFooter!.Blocks[0]).Inlines;
        Assert.Collection(footerInlines,
            i => Assert.Equal("Page ", ((TextRun)i).Text),
            i => Assert.True(i is FieldRun { Kind: FieldKind.Page } f && f.Properties.Bold), // takes the cached result's formatting
            i => Assert.Equal(" of ", ((TextRun)i).Text),
            i => Assert.Equal(FieldKind.NumPages, ((FieldRun)i).Kind),
            i => Assert.Equal("2020-01-01", ((TextRun)i).Text)); // unsupported field: cached result shown
    }

    [Fact]
    public void Reads_inline_images_and_breaks()
    {
        var b = new DocxBuilder();
        var png = PngTestEncoder.Encode(2, 2, 2, 8, new int[12]);
        var rel = b.AddImage("image1.png", png);
        b.Body = P(R("before") + InlineImage(rel, 144, 72) + "<w:r><w:br/><w:t>line</w:t><w:br w:type=\"page\"/><w:tab/><w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/></w:r>");
        var p = Paragraphs(Read(b))[0];
        var image = Assert.IsType<ImageRun>(p.Inlines[1]);
        Assert.Equal((144.0, 72.0, ImageFormat.Png), (image.Width, image.Height, image.Image.Format));
        Assert.Equal(png, image.Image.Data);
        Assert.Equal(BreakKind.Line, Assert.IsType<Break>(p.Inlines[2]).Kind);
        Assert.Equal(BreakKind.Page, Assert.IsType<Break>(p.Inlines[4]).Kind);
        Assert.IsType<Tab>(p.Inlines[5]);
        Assert.Equal("✔", Assert.IsType<TextRun>(p.Inlines[6]).Text);
    }

    [Fact]
    public void Reads_tables_with_spans_merges_borders_and_style_conditions()
    {
        var styles = Styles +
            "<w:style w:type=\"table\" w:default=\"1\" w:styleId=\"TableNormal\"><w:tblPr><w:tblCellMar><w:left w:w=\"108\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style>" +
            "<w:style w:type=\"table\" w:styleId=\"Grid\"><w:basedOn w:val=\"TableNormal\"/><w:tblPr><w:tblBorders>" +
            "<w:top w:val=\"single\" w:sz=\"4\"/><w:left w:val=\"single\" w:sz=\"4\"/><w:bottom w:val=\"single\" w:sz=\"4\"/><w:right w:val=\"single\" w:sz=\"4\"/>" +
            "<w:insideH w:val=\"dashed\" w:sz=\"4\"/><w:insideV w:val=\"single\" w:sz=\"4\"/></w:tblBorders></w:tblPr>" +
            "<w:tblStylePr w:type=\"firstRow\"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val=\"clear\" w:fill=\"4472C4\"/></w:tcPr></w:tblStylePr>" +
            "<w:tblStylePr w:type=\"lastRow\"><w:rPr><w:i/></w:rPr></w:tblStylePr></w:style>";
        string Cell(string text, string tcPr = "") => $"<w:tc><w:tcPr>{tcPr}</w:tcPr>{P(R(text))}</w:tc>";
        var table =
            "<w:tbl><w:tblPr><w:tblStyle w:val=\"Grid\"/><w:tblW w:w=\"5000\" w:type=\"pct\"/><w:jc w:val=\"center\"/>" +
            "<w:tblLook w:val=\"0060\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"4000\"/></w:tblGrid>" +
            $"<w:tr><w:trPr><w:tblHeader/></w:trPr>{Cell("H1", "<w:gridSpan w:val=\"2\"/>")}{Cell("H2")}</w:tr>" +
            $"<w:tr><w:trPr><w:trHeight w:val=\"400\" w:hRule=\"exact\"/><w:cantSplit/></w:trPr>{Cell("A", "<w:vMerge w:val=\"restart\"/><w:vAlign w:val=\"center\"/>")}{Cell("B")}" +
            $"{Cell("C", "<w:tcBorders><w:right w:val=\"nil\"/></w:tcBorders><w:tcMar><w:left w:w=\"0\" w:type=\"dxa\"/></w:tcMar>")}</w:tr>" +
            $"<w:tr>{Cell("", "<w:vMerge/>")}{Cell("E")}{Cell("F", "<w:shd w:val=\"clear\" w:fill=\"FFFF00\"/>")}</w:tr></w:tbl>";
        var doc = Read(new DocxBuilder { Styles = styles, Body = table + P("") });

        var t = Assert.IsType<Table>(doc.Sections[0].Blocks[0]);
        Assert.Equal([100.0, 150.0, 200.0], t.Grid);
        Assert.Equal((new TableWidth(100, TableWidthUnit.Percent), TableAlignment.Center), (t.Properties.Width, t.Properties.Alignment));
        Assert.Equal(3, t.Rows.Count);
        Assert.True(t.Rows[0].IsHeader);
        Assert.Equal((20.0, RowHeightRule.Exact, true), (t.Rows[1].Height, t.Rows[1].HeightRule, t.Rows[1].CantSplit));

        var h1 = t.Rows[0].Cells[0];
        Assert.Equal(2, h1.GridSpan);
        Assert.Equal(new Color(0x44, 0x72, 0xC4), h1.Shading);
        Assert.True(Run((Paragraph)h1.Blocks[0]).Properties.Bold);        // firstRow condition
        Assert.False(Run((Paragraph)t.Rows[1].Cells[1].Blocks[0]).Properties.Bold);
        Assert.True(Run((Paragraph)t.Rows[2].Cells[1].Blocks[0]).Properties.Italic); // lastRow condition

        var a = t.Rows[1].Cells[0];
        Assert.Equal((VerticalMerge.Restart, CellVerticalAlignment.Center), (a.VerticalMerge, a.VerticalAlignment));
        Assert.Equal(VerticalMerge.Continue, t.Rows[2].Cells[0].VerticalMerge);
        Assert.Equal(new Thickness(0, 5.4, 0, 5.4), a.Margins);

        // Outer edges use table borders, inner edges insideH/insideV, direct tcBorders win.
        Assert.Equal(BorderStyle.Single, a.Borders.Left.Style);
        Assert.Equal(BorderStyle.Dashed, a.Borders.Top.Style);
        var c = t.Rows[1].Cells[2];
        Assert.Equal(BorderStyle.None, c.Borders.Right.Style);
        Assert.Equal(0, c.Margins.Left);
        Assert.Equal(BorderStyle.Single, t.Rows[2].Cells[2].Borders.Bottom.Style);
        Assert.Equal(new Color(255, 255, 0), t.Rows[2].Cells[2].Shading);
    }

    [Fact]
    public void Nested_tables_are_read_into_cells()
    {
        var inner = "<w:tbl><w:tblGrid><w:gridCol w:w=\"1000\"/></w:tblGrid><w:tr><w:tc>" + P(R("inner")) + "</w:tc></w:tr></w:tbl>";
        var doc = Read(new DocxBuilder { Body = $"<w:tbl><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr><w:tc>{inner}{P("")}</w:tc></w:tr></w:tbl>" });
        var outer = Assert.IsType<Table>(doc.Sections[0].Blocks[0]);
        var nested = Assert.IsType<Table>(outer.Rows[0].Cells[0].Blocks[0]);
        Assert.Equal("inner", Run((Paragraph)nested.Rows[0].Cells[0].Blocks[0]).Text);
    }

    [Fact]
    public void Skips_unknown_and_unsupported_content_with_warnings()
    {
        var warnings = new ConversionWarningList();
        var body =
            "<w:p><w:customThing><w:r><w:t>hidden</w:t></w:r></w:customThing>" +
            "<w:hyperlink r:id=\"x\"><w:r><w:t>link </w:t></w:r></w:hyperlink>" +
            "<w:ins><w:r><w:t>inserted </w:t></w:r></w:ins><w:del><w:r><w:delText>deleted</w:delText></w:r></w:del>" +
            "<w:sdt><w:sdtContent><w:r><w:t>sdt </w:t></w:r></w:sdtContent></w:sdt>" +
            "<mc:AlternateContent><mc:Choice Requires=\"w14\"><w:r><w:t>choice</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback></mc:AlternateContent>" +
            "<w:r><w:pict/><w:footnoteReference w:id=\"1\"/></w:r>" +
            "<w:r><w:drawing><wp:anchor/></w:drawing></w:r></w:p>" +
            "<w:unknownBlock><w:p/></w:unknownBlock>";
        var doc = Read(new DocxBuilder { Body = body }, warnings);
        var p = Paragraphs(doc).Single();
        Assert.Equal("link inserted sdt choice", string.Concat(p.Inlines.OfType<TextRun>().Select(t => t.Text)));
        Assert.Contains(warnings.Warnings, w => w.Code == "docx.vml");
        Assert.Contains(warnings.Warnings, w => w.Code == "docx.note-missing"); // no footnotes part
        Assert.DoesNotContain(warnings.Warnings, w => w.Code == "docx.floating");
    }

    [Fact]
    public void Missing_optional_parts_use_defaults()
    {
        var doc = Read(new DocxBuilder { Body = P(R("x")) });
        var run = Run(Paragraphs(doc)[0]).Properties;
        Assert.Equal(("Times New Roman", 10.0), (run.FontFamily, run.FontSize));
        var s = doc.Sections.Single();
        Assert.Equal((612.0, 792.0), (s.PageWidth, s.PageHeight));
    }

    [Fact]
    public void Reads_document_title()
    {
        Assert.Equal("My title", Read(new DocxBuilder { Title = "My title", Body = P("") }).Title);
    }

    [Fact]
    public void Rejects_non_packages()
    {
        Assert.Throws<InvalidDataException>(() => DocxReader.Read(new MemoryStream([1, 2, 3, 4])));
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) zip.CreateEntry("hello.txt");
        ms.Position = 0;
        Assert.Throws<InvalidDataException>(() => DocxReader.Read(ms));
    }

    [Fact]
    public void Enforces_decompression_limits()
    {
        var builder = new DocxBuilder { Body = P(R(new string('x', 200_000))) };
        Assert.Throws<InvalidDataException>(() =>
            DocxReader.Read(builder.BuildStream(), new DocxReaderOptions { MaxPartSize = 50_000 }));
        Assert.Throws<InvalidDataException>(() =>
            DocxReader.Read(builder.BuildStream(), new DocxReaderOptions { MaxTotalUncompressedSize = 100_000 }));
    }

    [Fact]
    public void Survives_deep_nesting_and_huge_dimensions()
    {
        var warnings = new ConversionWarningList();
        var deep = string.Concat(Enumerable.Repeat("<w:sdt><w:sdtContent>", 400)) + P(R("deep")) + string.Concat(Enumerable.Repeat("</w:sdtContent></w:sdt>", 400));
        var doc = Read(new DocxBuilder { Body = deep + P(R("ok")) + Section(99_999_999, 99_999_999) }, warnings);
        Assert.Contains(warnings.Warnings, w => w.Code == "docx.depth");
        Assert.Equal("ok", Run(Paragraphs(doc).Last()).Text);
        Assert.Equal(14_400, doc.Sections[0].PageWidth);
    }

    [Fact]
    public void Resolves_relationship_targets()
    {
        Assert.Equal("word/media/a.png", OpcPackage.Resolve("word/document.xml", "media/a.png"));
        Assert.Equal("media/a.png", OpcPackage.Resolve("word/document.xml", "../media/a.png"));
        Assert.Equal("word/x.xml", OpcPackage.Resolve("word/document.xml", "/word/x.xml"));
        Assert.Equal("word/a b.png", OpcPackage.Resolve("word/document.xml", "a%20b.png"));
    }
}
public class DocxFuzzTests
{
    /// <summary>Corrupted packages must either convert or fail with InvalidDataException — never anything else.</summary>
    [Fact]
    public void Mutated_documents_fail_cleanly()
    {
        var builder = new DocxBuilder
        {
            Body = string.Concat(Enumerable.Range(0, 20).Select(i =>
                DocxBuilder.P(DocxBuilder.R($"Paragraph {i}") + "<w:r><w:tab/><w:t>x</w:t></w:r>", "<w:jc w:val=\"both\"/>")))
                + "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc>" + DocxBuilder.P(DocxBuilder.R("cell")) + "</w:tc></w:tr></w:tbl>",
        };
        var original = builder.Build();
        var rng = new Random(1234);
        for (int i = 0; i < 60; i++)
        {
            var bytes = (byte[])original.Clone();
            for (int k = 0; k < 1 + i % 6; k++) bytes[rng.Next(bytes.Length)] ^= (byte)rng.Next(1, 256);
            try
            {
                DocxReader.Read(new MemoryStream(bytes));
            }
            catch (InvalidDataException)
            {
            }
        }
    }
}
