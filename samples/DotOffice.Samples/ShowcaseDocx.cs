using System.Buffers.Binary;
using System.IO.Compression;
using System.Security;
using System.Text;

namespace DotOffice.Samples;

/// <summary>
/// Generates a self-contained showcase DOCX (raw WordprocessingML) exercising the features of the
/// converter: styles, lists, tables, headers/footers, sections, images, pagination, a table of contents,
/// fields and links, columns, footnotes and endnotes, floating images and right-to-left text.
/// Everything in it is generated, so it can be redistributed freely.
/// </summary>
internal static class ShowcaseDocx
{
    private const string Ns =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"";

    private const string Lorem =
        "DotOffice lays out paragraphs with the Unicode line breaking algorithm, measures text with the " +
        "fonts' own advance widths and paginates the result with widow and orphan control. Justified text " +
        "stretches the spaces between words so that both edges of the column line up, while the last line " +
        "of every paragraph stays left aligned. Tables, lists, headers and footers are handled by the same engine.";

    private static readonly string[] Headings =
    [
        "Character formatting", "Paragraph formatting", "Lists", "Images", "Tables", "Pagination", "Landscape section",
        "Multiple columns", "Footnotes and endnotes", "Floating images", "Right-to-left text", "Fields and links",
    ];

    public static void Write(string path, byte[]? jpeg)
    {
        var body = new StringBuilder();
        body.Append(P(R("DotOffice showcase"), "<w:pStyle w:val=\"Title\"/>"));
        body.Append(P(R("A generated document converted to PDF without LibreOffice.", "<w:i/>"), "<w:pStyle w:val=\"Subtitle\"/>"));
        body.Append(TableOfContents());

        body.Append(Heading(1));
        body.Append(P(R("Plain, ") + R("bold", "<w:b/>") + R(", ") + R("italic", "<w:i/>") + R(", ") + R("bold italic", "<w:b/><w:i/>")
                      + R(", ") + R("underlined", "<w:u w:val=\"single\"/>") + R(", ") + R("double underline", "<w:u w:val=\"double\"/>")
                      + R(", ") + R("struck", "<w:strike/>") + R(", ") + R("colored", "<w:color w:val=\"C00000\"/>") + R(", ")
                      + R("highlighted", "<w:highlight w:val=\"yellow\"/>") + R(", ") + R("Small Caps", "<w:smallCaps/>") + R(", ")
                      + R("all caps", "<w:caps/>") + R(", E = mc") + R("2", "<w:vertAlign w:val=\"superscript\"/>") + R(", H")
                      + R("2", "<w:vertAlign w:val=\"subscript\"/>") + R("O, ") + R("spaced", "<w:spacing w:val=\"40\"/>") + R(".")));
        body.Append(P(R("Toggle semantics: this paragraph uses the bold \u201CEmphasis Heading\u201D style; ") + R("a bold character style inside it turns bold off", "<w:rStyle w:val=\"Strong\"/>")
                      + R(", while ") + R("direct bold stays bold", "<w:b/>") + R("."), "<w:pStyle w:val=\"EmphasisPara\"/>"));
        body.Append(P(R("Unicode: Привет, мир! Γειά σου κόσμε! Ångström, naïve café — “quotes” … ✓")));

        body.Append(Heading(2));
        body.Append(P(R(Lorem), "<w:jc w:val=\"both\"/>"));
        body.Append(P(R("Centered paragraph with 1.5 line spacing. " + Lorem[..160]), "<w:jc w:val=\"center\"/><w:spacing w:line=\"360\" w:lineRule=\"auto\"/>"));
        body.Append(P(R("Right aligned text."), "<w:jc w:val=\"right\"/>"));
        body.Append(P(R("First-line indent of half an inch. " + Lorem[..200]), "<w:ind w:firstLine=\"720\"/>"));
        body.Append(P(R("Hanging indent: " + Lorem[..220]), "<w:ind w:left=\"1440\" w:hanging=\"720\"/>"));
        body.Append(P(R("Bordered and shaded paragraph with some padding between the border and the text."),
            "<w:pBdr><w:top w:val=\"single\" w:sz=\"8\" w:space=\"4\" w:color=\"2F5496\"/><w:left w:val=\"single\" w:sz=\"8\" w:space=\"4\" w:color=\"2F5496\"/>" +
            "<w:bottom w:val=\"single\" w:sz=\"8\" w:space=\"4\" w:color=\"2F5496\"/><w:right w:val=\"single\" w:sz=\"8\" w:space=\"4\" w:color=\"2F5496\"/></w:pBdr>" +
            "<w:shd w:val=\"clear\" w:fill=\"DEEAF6\"/><w:ind w:left=\"360\" w:right=\"360\"/>"));
        body.Append(P(R("Tab stops:") + Tab() + R("left") + Tab() + R("center") + Tab() + R("right") ,
            "<w:tabs><w:tab w:val=\"left\" w:pos=\"2160\"/><w:tab w:val=\"center\" w:pos=\"5040\"/><w:tab w:val=\"right\" w:pos=\"9360\"/></w:tabs>"));
        body.Append(P(R("Chapter one") + Tab() + R("12"), "<w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9360\"/></w:tabs>"));
        body.Append(P(R("Decimal tab:") + Tab() + R("1234.56"), "<w:tabs><w:tab w:val=\"decimal\" w:pos=\"4320\"/></w:tabs>"));
        body.Append(P(R("Decimal tab:") + Tab() + R("7.5"), "<w:tabs><w:tab w:val=\"decimal\" w:pos=\"4320\"/></w:tabs>"));

        body.Append(Heading(3));
        foreach (var (level, text) in new[] { (0, "First item"), (1, "Nested item a"), (1, "Nested item b"), (2, "Third level"), (0, "Second item"), (1, "Nested again (restarts at a)"), (0, "Third item with a long text that wraps onto the next line to show the hanging indentation of list items.") })
            body.Append(P(R(text), $"<w:pStyle w:val=\"ListParagraph\"/><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"1\"/></w:numPr>"));
        body.Append(P(R("Bullets, mixed with numbers:")));
        foreach (var (level, text) in new[] { (0, "Bullet one"), (1, "Sub-bullet"), (2, "Square bullet"), (0, "Bullet two") })
            body.Append(P(R(text), $"<w:pStyle w:val=\"ListParagraph\"/><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"2\"/></w:numPr>"));
        body.Append(P(R("Roman numerals starting at 4:")));
        for (int i = 0; i < 3; i++)
            body.Append(P(R($"Item {i + 1}"), "<w:pStyle w:val=\"ListParagraph\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"3\"/></w:numPr>"));

        var parts = new List<(string Name, byte[] Data, string RelId)>();
        body.Append(Heading(4));
        parts.Add(("image1.png", GradientPng(160, 100), "rIdImg1"));
        body.Append(P(R("PNG with alpha over a shaded paragraph: ") + Image("rIdImg1", 160, 100, 1),
            "<w:shd w:val=\"clear\" w:fill=\"FFF2CC\"/>"));
        if (jpeg is not null)
        {
            parts.Add(("image2.jpeg", jpeg, "rIdImg2"));
            body.Append(P(R("JPEG: ") + Image("rIdImg2", 120, 75, 2) + R(" inline with text.")));
        }

        body.Append(Heading(5));
        body.Append(SimpleTable());
        body.Append(P(R("")));
        body.Append(LongTable());

        body.Append(Heading(6, "<w:pageBreakBefore/>"));
        for (int i = 1; i <= 12; i++)
            body.Append(P(R($"Paragraph {i}. " + Lorem + " " + Lorem), "<w:jc w:val=\"both\"/>"));

        // Landscape section.
        body.Append(P(R("End of the portrait section."), "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>" +
            "<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/><w:headerReference w:type=\"first\" r:id=\"rIdFirstHeader\"/><w:titlePg/>" +
            "<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>"));
        body.Append(Heading(7));
        body.Append(P(R("This section uses a landscape page. Headers and footers are inherited from the previous section. " + Lorem), "<w:jc w:val=\"both\"/>"
            + "<w:sectPr><w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"/>" +
              "<w:pgMar w:top=\"1080\" w:right=\"1080\" w:bottom=\"1080\" w:left=\"1080\" w:header=\"540\" w:footer=\"540\" w:gutter=\"0\"/></w:sectPr>"));

        // 8. Columns: a one-column heading, two columns with a separator and a column break, then three
        // columns balanced by a continuous section break, and one column again.
        body.Append(Heading(8));
        body.Append(P(R("The next paragraphs flow through two columns separated by a line. A column break ends the first column early; "
                        + "the three columns after it are balanced because a continuous section break follows them."),
            Portrait("<w:type w:val=\"nextPage\"/>")));
        body.Append(P(R("Column one. " + Lorem), "<w:jc w:val=\"both\"/>"));
        body.Append(P(R("Still column one, then a column break.") + "<w:r><w:br w:type=\"column\"/></w:r>" + R("Column two starts here. " + Lorem),
            "<w:jc w:val=\"both\"/>"));
        body.Append(P(R("Two columns end here."), Portrait("<w:type w:val=\"continuous\"/><w:cols w:num=\"2\" w:space=\"720\" w:sep=\"1\"/>")));
        for (int i = 1; i <= 6; i++)
            body.Append(P(R($"Balanced column text {i}. " + Lorem[..(120 + i * 25)]), "<w:jc w:val=\"both\"/>"));
        body.Append(P(R("Three balanced columns end here."), Portrait("<w:type w:val=\"continuous\"/><w:cols w:num=\"3\" w:space=\"360\"/>")));
        body.Append(P(R("Back to a single column below the balanced columns.")));

        // 9. Footnotes and endnotes.
        body.Append(Heading(9));
        body.Append(P(R("Footnotes are numbered automatically") + NoteRef("footnote", 1) + R(" and placed at the bottom of the page that "
                      + "references them, below a short separator line") + NoteRef("footnote", 2)
                      + R(". Endnotes are collected after the last section") + NoteRef("endnote", 1) + R(".")));

        // 10. Floating images.
        body.Append(Heading(10));
        body.Append(P(Anchor("rIdImg1", 150, 94, 3, "<wp:positionH relativeFrom=\"margin\"><wp:align>right</wp:align></wp:positionH>"
                                                + "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV>",
                             "<wp:wrapSquare wrapText=\"bothSides\"/>", 114300)
                      + R("Square wrapping: this paragraph flows around a floating image anchored to the right margin. " + Lorem + " " + Lorem),
            "<w:jc w:val=\"both\"/>"));
        body.Append(P(Anchor("rIdImg1", 80, 50, 4, "<wp:positionH relativeFrom=\"column\"><wp:align>center</wp:align></wp:positionH>"
                                               + "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>127000</wp:posOffset></wp:positionV>",
                             "<wp:wrapTopAndBottom/>", 63500)
                      + R("Top-and-bottom wrapping keeps lines above and below this centered image. " + Lorem), "<w:jc w:val=\"both\"/>"));

        // 11. Right-to-left text: Hebrew and Arabic paragraphs with numbers, brackets and embedded English.
        body.Append(Heading(11));
        const string hebrew = "<w:rFonts w:cs=\"Noto Sans Hebrew\"/><w:rtl/>";
        const string arabic = "<w:rFonts w:cs=\"Noto Sans Arabic\"/><w:rtl/><w:szCs w:val=\"24\"/>";
        body.Append(P(R("זהו טקסט בעברית עם מספר 2026 (בסוגריים) ומילה ", hebrew) + R("English") + R(" באמצע.", hebrew), "<w:bidi/>"));
        body.Append(P(R("שָׁלוֹם עוֹלָם — טקסט מנוקד.", hebrew), "<w:bidi/>"));
        body.Append(P(R("مرحبا بالعالم، هذا نص عربي مع الأرقام ١٢٣ و456 وكلمة ", arabic) + R("PDF") + R(" في الوسط.", arabic), "<w:bidi/><w:jc w:val=\"both\"/>"));
        body.Append(P(R("بِسْمِ اللَّهِ الرَّحْمَٰنِ الرَّحِيمِ", arabic), "<w:bidi/><w:jc w:val=\"center\"/>"));
        body.Append(P(R("A left-to-right paragraph can contain ") + R("كلمات عربية", arabic) + R(" and ") + R("מילים בעברית", hebrew) + R(" in their own direction.")));
        body.Append(P(R("Kerning (w:kern): AVATAR WAVE Toyota", "<w:kern w:val=\"1\"/><w:sz w:val=\"28\"/>") + R("  without: ") +
                      R("AVATAR WAVE Toyota", "<w:sz w:val=\"28\"/>")));

        // 12. Fields and links.
        body.Append(Heading(12));
        body.Append(P(R("This paragraph is on page ") + Field("PAGE", "1") + R(" (in roman numerals: ") + Field("PAGE \\* roman", "i")
                      + R("). The footnote section starts on page ") + Field("PAGEREF _Toc9 \\h", "1") + R(" and is titled \u201C")
                      + Field("REF _Toc9 \\h", "stale cached title") + R("\u201D. Document date: ") + Field("DATE \\@ \"d MMMM yyyy\"", "31 October 2026")
                      + R(". Unknown fields keep their cached result: ") + Field("MERGEFIELD Customer", "\u00ABCustomer\u00BB") + R(".")));
        body.Append(P(R("Links: ") + $"<w:hyperlink r:id=\"rIdLink\">{R("the DotOffice repository", "<w:rStyle w:val=\"Hyperlink\"/>")}</w:hyperlink>"
                      + R(", an internal link to ") + $"<w:hyperlink w:anchor=\"_Toc8\">{R("the columns section", "<w:rStyle w:val=\"Hyperlink\"/>")}</w:hyperlink>"
                      + R(" and a HYPERLINK field to ") + Field("HYPERLINK \"https://www.unicode.org/reports/tr9/\"", "UAX #9", "<w:rStyle w:val=\"Hyperlink\"/>") + R(".")));
        body.Append(P(R("The end of the document. The endnote follows.")));
        body.Append(Portrait("<w:type w:val=\"continuous\"/>"));

        using var zip = new ZipArchive(File.Create(path), ZipArchiveMode.Create);
        Entry(zip, "[Content_Types].xml",
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"png\" ContentType=\"image/png\"/>" +
            "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>" +
            "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
            "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" +
            "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>" +
            "<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>" +
            "<Override PartName=\"/word/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" +
            "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>" +
            "<Override PartName=\"/word/header2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>" +
            "<Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>" +
            "<Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/>" +
            "<Override PartName=\"/word/endnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml\"/>" +
            "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/></Types>");
        Entry(zip, "_rels/.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/></Relationships>");
        Entry(zip, "docProps/core.xml",
            "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
            "<dc:title>DotOffice showcase</dc:title><dc:creator>DotOffice samples</dc:creator></cp:coreProperties>");
        var rels = new StringBuilder("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        void Rel(string id, string type, string target) =>
            rels.Append($"<Relationship Id=\"{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{type}\" Target=\"{target}\"/>");
        Rel("rIdStyles", "styles", "styles.xml");
        Rel("rIdNumbering", "numbering", "numbering.xml");
        Rel("rIdSettings", "settings", "settings.xml");
        Rel("rIdTheme", "theme", "theme/theme1.xml");
        Rel("rIdHeader", "header", "header1.xml");
        Rel("rIdFirstHeader", "header", "header2.xml");
        Rel("rIdFooter", "footer", "footer1.xml");
        Rel("rIdFootnotes", "footnotes", "footnotes.xml");
        Rel("rIdEndnotes", "endnotes", "endnotes.xml");
        rels.Append("<Relationship Id=\"rIdLink\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" " +
                    "Target=\"https://github.com/stasnocap/dotoffice\" TargetMode=\"External\"/>");
        foreach (var (name, data, relId) in parts)
        {
            Rel(relId, "image", "media/" + name);
            using var s = zip.CreateEntry("word/media/" + name).Open();
            s.Write(data);
        }
        Entry(zip, "word/_rels/document.xml.rels", rels.Append("</Relationships>").ToString());
        Entry(zip, "word/document.xml", $"<w:document {Ns}><w:body>{body}</w:body></w:document>");
        Entry(zip, "word/styles.xml", $"<w:styles {Ns}>{Styles}</w:styles>");
        Entry(zip, "word/numbering.xml", $"<w:numbering {Ns}>{Numbering}</w:numbering>");
        Entry(zip, "word/footnotes.xml", $"<w:footnotes {Ns}>" + Separators("footnote")
            + Note("footnote", 1, "Footnote text uses the FootnoteText style. Its number is the w:footnoteRef mark.")
            + Note("footnote", 2, "A second footnote on the same page.") + "</w:footnotes>");
        Entry(zip, "word/endnotes.xml", $"<w:endnotes {Ns}>" + Separators("endnote")
            + Note("endnote", 1, "Endnotes are numbered with lowercase roman numerals by default.") + "</w:endnotes>");
        Entry(zip, "word/settings.xml", $"<w:settings {Ns}><w:defaultTabStop w:val=\"720\"/><w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat></w:settings>");
        Entry(zip, "word/theme/theme1.xml",
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\"><a:themeElements><a:fontScheme name=\"Office\">" +
            "<a:majorFont><a:latin typeface=\"Cambria\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
            "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>");
        Entry(zip, "word/header1.xml", $"<w:hdr {Ns}>" + P(R("DotOffice showcase") + Tab() + Tab() + R("Phase 3", "<w:i/>"), "<w:pStyle w:val=\"Header\"/><w:pBdr><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"808080\"/></w:pBdr>") + "</w:hdr>");
        Entry(zip, "word/header2.xml", $"<w:hdr {Ns}>" + P(R("First page header — different from the rest", "<w:color w:val=\"808080\"/>"), "<w:jc w:val=\"center\"/>") + "</w:hdr>");
        Entry(zip, "word/footer1.xml", $"<w:ftr {Ns}>" + P(R("Page ") + "<w:fldSimple w:instr=\" PAGE \"><w:r><w:rPr><w:b/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>" + R(" of ")
                                                      + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> NUMPAGES </w:instrText></w:r>"
                                                      + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
                                                      "<w:jc w:val=\"center\"/>") + "</w:ftr>");
    }

    private static string Heading(int number, string pPr = "") =>
        P($"<w:bookmarkStart w:id=\"{number}\" w:name=\"_Toc{number}\"/>{R($"{number}. {Headings[number - 1]}")}<w:bookmarkEnd w:id=\"{number}\"/>",
          "<w:pStyle w:val=\"Heading1\"/>" + pPr);

    /// <summary>A TOC field as Word saves it: cached entries linking to the headings, with PAGEREF page numbers (stale here).</summary>
    private static string TableOfContents()
    {
        var sb = new StringBuilder(P(R("Contents"), "<w:pStyle w:val=\"TOCHeading\"/>"));
        for (int i = 1; i <= Headings.Length; i++)
        {
            var entry = $"<w:hyperlink w:anchor=\"_Toc{i}\">{R($"{i}. {Headings[i - 1]}")}{Tab()}{Field($"PAGEREF _Toc{i} \\h", "0")}</w:hyperlink>";
            if (i == 1) entry = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-1\" \\h \\z \\u </w:instrText></w:r>"
                                + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + entry;
            if (i == Headings.Length) entry += "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";
            sb.Append(P(entry, "<w:pStyle w:val=\"TOC1\"/>"));
        }
        return sb.ToString();
    }

    private static string Field(string instruction, string cached, string rPr = "") =>
        $"<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> {SecurityElement.Escape(instruction)} </w:instrText></w:r>" +
        $"<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>{R(cached, rPr)}<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    private static string NoteRef(string kind, int id) =>
        $"<w:r><w:rPr><w:rStyle w:val=\"{(kind == "footnote" ? "FootnoteReference" : "EndnoteReference")}\"/></w:rPr><w:{kind}Reference w:id=\"{id}\"/></w:r>";

    private static string Note(string kind, int id, string text) =>
        $"<w:{kind} w:id=\"{id}\"><w:p><w:pPr><w:pStyle w:val=\"FootnoteText\"/></w:pPr>" +
        $"<w:r><w:rPr><w:rStyle w:val=\"{(kind == "footnote" ? "FootnoteReference" : "EndnoteReference")}\"/></w:rPr><w:{kind}Ref/></w:r>{R(" " + text)}</w:p></w:{kind}>";

    private static string Separators(string kind) =>
        $"<w:{kind} w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:{kind}>" +
        $"<w:{kind} w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{kind}>";

    /// <summary>A portrait letter-size w:sectPr with the given extra settings (section type, columns).</summary>
    private static string Portrait(string extra) =>
        $"<w:sectPr>{extra}<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" " +
        "w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>";

    /// <summary>A floating picture (wp:anchor); <paramref name="distance"/> is the wrap distance in EMU.</summary>
    private static string Anchor(string relId, double widthPt, double heightPt, int id, string position, string wrap, long distance) =>
        $"<w:r><w:drawing><wp:anchor distT=\"{distance}\" distB=\"{distance}\" distL=\"{distance}\" distR=\"{distance}\" simplePos=\"0\" relativeHeight=\"{id}\" " +
        $"behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/>{position}" +
        $"<wp:extent cx=\"{(long)(widthPt * 12700)}\" cy=\"{(long)(heightPt * 12700)}\"/>{wrap}<wp:docPr id=\"{id}\" name=\"Floating {id}\"/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        $"<pic:pic><pic:nvPicPr><pic:cNvPr id=\"{id}\" name=\"float{id}\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{relId}\"/>" +
        $"<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{(long)(widthPt * 12700)}\" cy=\"{(long)(heightPt * 12700)}\"/>" +
        "</a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>";

    private static string SimpleTable()
    {
        static string Cell(string content, string tcPr = "", string pPr = "") => $"<w:tc><w:tcPr>{tcPr}</w:tcPr>{P(content, pPr)}</w:tc>";
        return "<w:tbl><w:tblPr><w:tblStyle w:val=\"GridTable\"/><w:tblW w:w=\"5000\" w:type=\"pct\"/><w:tblLook w:val=\"04A0\"/></w:tblPr>" +
               "<w:tblGrid><w:gridCol w:w=\"2340\"/><w:gridCol w:w=\"2340\"/><w:gridCol w:w=\"2340\"/><w:gridCol w:w=\"2340\"/></w:tblGrid>" +
               "<w:tr>" + Cell(R("Spans two columns"), "<w:gridSpan w:val=\"2\"/>", "<w:jc w:val=\"center\"/>") + Cell(R("Header C")) + Cell(R("Header D")) + "</w:tr>" +
               "<w:tr>" + Cell(R("Vertically merged, centered"), "<w:vMerge w:val=\"restart\"/><w:vAlign w:val=\"center\"/>") + Cell(R("b1")) + Cell(R("c1")) +
               Cell(R("Shaded cell"), "<w:shd w:val=\"clear\" w:fill=\"FCE4D6\"/>") + "</w:tr>" +
               "<w:tr>" + Cell("", "<w:vMerge/>") + Cell(R("b2 with more text that wraps inside its cell")) + Cell(R("c2")) +
               Cell(R("Right"), "", "<w:jc w:val=\"right\"/>") + "</w:tr>" +
               "<w:tr>" + Cell(R("a3")) + Cell(R("b3")) + Cell(R("Thick red border"), "<w:tcBorders><w:top w:val=\"single\" w:sz=\"18\" w:color=\"C00000\"/>" +
               "<w:bottom w:val=\"single\" w:sz=\"18\" w:color=\"C00000\"/><w:left w:val=\"single\" w:sz=\"18\" w:color=\"C00000\"/><w:right w:val=\"single\" w:sz=\"18\" w:color=\"C00000\"/></w:tcBorders>") +
               Cell(R("d3")) + "</w:tr></w:tbl>";
    }

    private static string LongTable()
    {
        var sb = new StringBuilder("<w:tbl><w:tblPr><w:tblStyle w:val=\"GridTable\"/><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblLook w:val=\"04A0\"/></w:tblPr>" +
                                   "<w:tblGrid><w:gridCol w:w=\"1200\"/><w:gridCol w:w=\"5160\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>");
        sb.Append("<w:tr><w:trPr><w:tblHeader/></w:trPr>");
        foreach (var h in new[] { "#", "Description (header row repeats on every page)", "Amount" })
            sb.Append($"<w:tc>{P(R(h))}</w:tc>");
        sb.Append("</w:tr>");
        for (int i = 1; i <= 40; i++)
        {
            sb.Append("<w:tr>");
            sb.Append($"<w:tc>{P(R(i.ToString()))}</w:tc>");
            sb.Append($"<w:tc>{P(R(i % 7 == 0 ? "A longer description that needs to wrap over several lines in this column, to make rows of different heights." : $"Row {i} of a table that is long enough to break across pages."))}</w:tc>");
            sb.Append($"<w:tc>{P(R($"{i * 123.45:F2}"), "<w:jc w:val=\"right\"/>")}</w:tc>");
            sb.Append("</w:tr>");
        }
        return sb.Append("</w:tbl>").ToString();
    }

    private const string Styles =
        "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"minorHAnsi\" w:cstheme=\"minorBidi\"/>" +
        "<w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/><w:lang w:val=\"en-US\"/></w:rPr></w:rPrDefault>" +
        "<w:pPrDefault><w:pPr><w:spacing w:after=\"160\" w:line=\"259\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults>" +
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
        "<w:style w:type=\"character\" w:default=\"1\" w:styleId=\"DefaultParagraphFont\"><w:name w:val=\"Default Paragraph Font\"/></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Title\"><w:name w:val=\"Title\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/><w:contextualSpacing/></w:pPr>" +
        "<w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:hAnsiTheme=\"majorHAnsi\"/><w:spacing w:val=\"-10\"/><w:sz w:val=\"56\"/></w:rPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Subtitle\"><w:name w:val=\"Subtitle\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:color w:val=\"5A5A5A\"/><w:sz w:val=\"24\"/></w:rPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:next w:val=\"Normal\"/>" +
        "<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before=\"240\" w:after=\"80\"/></w:pPr>" +
        "<w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:hAnsiTheme=\"majorHAnsi\"/><w:b/><w:color w:val=\"2F5496\"/><w:sz w:val=\"32\"/></w:rPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"EmphasisPara\"><w:name w:val=\"Emphasis Heading\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/></w:rPr></w:style>" +
        "<w:style w:type=\"character\" w:styleId=\"Strong\"><w:name w:val=\"Strong\"/><w:rPr><w:b/><w:color w:val=\"7030A0\"/></w:rPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"ListParagraph\"><w:name w:val=\"List Paragraph\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:spacing w:after=\"40\"/><w:ind w:left=\"720\"/><w:contextualSpacing/></w:pPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Header\"><w:name w:val=\"header\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:tabs><w:tab w:val=\"center\" w:pos=\"4680\"/><w:tab w:val=\"right\" w:pos=\"9360\"/></w:tabs><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"TOCHeading\"><w:name w:val=\"TOC Heading\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:spacing w:before=\"240\" w:after=\"80\"/></w:pPr>" +
        "<w:rPr><w:rFonts w:asciiTheme=\"majorHAnsi\" w:hAnsiTheme=\"majorHAnsi\"/><w:color w:val=\"2F5496\"/><w:sz w:val=\"28\"/></w:rPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"TOC1\"><w:name w:val=\"toc 1\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9350\"/></w:tabs>" +
        "<w:spacing w:after=\"40\"/></w:pPr></w:style>" +
        "<w:style w:type=\"character\" w:styleId=\"Hyperlink\"><w:name w:val=\"Hyperlink\"/><w:rPr><w:color w:val=\"0563C1\"/><w:u w:val=\"single\"/></w:rPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"FootnoteText\"><w:name w:val=\"footnote text\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr>" +
        "<w:rPr><w:sz w:val=\"18\"/><w:szCs w:val=\"18\"/></w:rPr></w:style>" +
        "<w:style w:type=\"character\" w:styleId=\"FootnoteReference\"><w:name w:val=\"footnote reference\"/><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr></w:style>" +
        "<w:style w:type=\"character\" w:styleId=\"EndnoteReference\"><w:name w:val=\"endnote reference\"/><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr></w:style>" +
        "<w:style w:type=\"table\" w:default=\"1\" w:styleId=\"TableNormal\"><w:name w:val=\"Normal Table\"/><w:tblPr><w:tblInd w:w=\"0\" w:type=\"dxa\"/>" +
        "<w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style>" +
        "<w:style w:type=\"table\" w:styleId=\"GridTable\"><w:name w:val=\"Grid Table 4 Accent 1\"/><w:basedOn w:val=\"TableNormal\"/>" +
        "<w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr>" +
        "<w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:color=\"8EAADB\"/><w:left w:val=\"single\" w:sz=\"4\" w:color=\"8EAADB\"/>" +
        "<w:bottom w:val=\"single\" w:sz=\"4\" w:color=\"8EAADB\"/><w:right w:val=\"single\" w:sz=\"4\" w:color=\"8EAADB\"/>" +
        "<w:insideH w:val=\"single\" w:sz=\"4\" w:color=\"8EAADB\"/><w:insideV w:val=\"single\" w:sz=\"4\" w:color=\"8EAADB\"/></w:tblBorders>" +
        "<w:tblCellMar><w:top w:w=\"40\" w:type=\"dxa\"/><w:bottom w:w=\"40\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr>" +
        "<w:tblStylePr w:type=\"firstRow\"><w:rPr><w:b/><w:color w:val=\"FFFFFF\"/></w:rPr><w:tcPr><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"4472C4\"/></w:tcPr></w:tblStylePr>" +
        "<w:tblStylePr w:type=\"firstCol\"><w:rPr><w:b/></w:rPr></w:tblStylePr>" +
        "<w:tblStylePr w:type=\"band1Horz\"><w:tcPr><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"D9E2F3\"/></w:tcPr></w:tblStylePr></w:style>";

    private const string Numbering =
        "<w:abstractNum w:abstractNumId=\"0\"><w:multiLevelType w:val=\"multilevel\"/>" +
        "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl>" +
        "<w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerLetter\"/><w:lvlText w:val=\"%2)\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"360\"/></w:pPr></w:lvl>" +
        "<w:lvl w:ilvl=\"2\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:lvlText w:val=\"%1.%2.%3\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2160\" w:hanging=\"720\"/></w:pPr></w:lvl></w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"1\"><w:multiLevelType w:val=\"hybridMultilevel\"/>" +
        "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Symbol\" w:hAnsi=\"Symbol\" w:hint=\"default\"/></w:rPr></w:lvl>" +
        "<w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"o\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"1440\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Courier New\" w:hAnsi=\"Courier New\" w:cs=\"Courier New\" w:hint=\"default\"/></w:rPr></w:lvl>" +
        "<w:lvl w:ilvl=\"2\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"2160\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Wingdings\" w:hAnsi=\"Wingdings\" w:hint=\"default\"/></w:rPr></w:lvl></w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"2\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"4\"/><w:numFmt w:val=\"upperRoman\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl></w:abstractNum>" +
        "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num><w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num><w:num w:numId=\"3\"><w:abstractNumId w:val=\"2\"/></w:num>";

    private static string P(string runs, string pPr = "") => $"<w:p>{(pPr.Length > 0 ? $"<w:pPr>{pPr}</w:pPr>" : "")}{runs}</w:p>";

    private static string R(string text, string rPr = "") =>
        $"<w:r>{(rPr.Length > 0 ? $"<w:rPr>{rPr}</w:rPr>" : "")}<w:t xml:space=\"preserve\">{SecurityElement.Escape(text)}</w:t></w:r>";

    private static string Tab() => "<w:r><w:tab/></w:r>";

    private static string Image(string relId, double widthPt, double heightPt, int id) =>
        $"<w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{(long)(widthPt * 12700)}\" cy=\"{(long)(heightPt * 12700)}\"/>" +
        $"<wp:docPr id=\"{id}\" name=\"Picture {id}\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        $"<pic:pic><pic:nvPicPr><pic:cNvPr id=\"{id}\" name=\"img{id}\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{relId}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        $"<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{(long)(widthPt * 12700)}\" cy=\"{(long)(heightPt * 12700)}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>" +
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

    private static void Entry(ZipArchive zip, string name, string content)
    {
        using var s = zip.CreateEntry(name).Open();
        s.Write(Encoding.UTF8.GetBytes(content));
    }

    /// <summary>RGBA PNG: a color gradient with a transparent radial hole, so alpha is visible.</summary>
    private static byte[] GradientPng(int width, int height)
    {
        var raw = new byte[height * (1 + width * 4)];
        for (int y = 0; y < height; y++)
        {
            int row = y * (1 + width * 4);
            for (int x = 0; x < width; x++)
            {
                int i = row + 1 + x * 4;
                double dx = x - width / 2.0, dy = y - height / 2.0, d = Math.Sqrt(dx * dx + dy * dy) / (height / 2.0);
                raw[i] = (byte)(255 * x / width);
                raw[i + 1] = (byte)(120 + 100 * y / height);
                raw[i + 2] = (byte)(255 - 255 * x / width);
                raw[i + 3] = (byte)Math.Clamp(d * 255, 0, 255);
            }
        }
        using var png = new MemoryStream();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
        var ihdr = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(ihdr, width);
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(4), height);
        ihdr[8] = 8;
        ihdr[9] = 6;
        Chunk(png, "IHDR", ihdr);
        using var z = new MemoryStream();
        using (var zlib = new ZLibStream(z, CompressionLevel.Optimal, leaveOpen: true)) zlib.Write(raw);
        Chunk(png, "IDAT", z.ToArray());
        Chunk(png, "IEND", []);
        return png.ToArray();
    }

    private static void Chunk(Stream s, string type, byte[] data)
    {
        Span<byte> buf = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(buf, data.Length);
        s.Write(buf);
        var typeBytes = Encoding.ASCII.GetBytes(type);
        s.Write(typeBytes);
        s.Write(data);
        uint crc = 0xFFFFFFFF;
        foreach (byte b in typeBytes.Concat(data))
        {
            crc ^= b;
            for (int k = 0; k < 8; k++) crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
        }
        BinaryPrimitives.WriteUInt32BigEndian(buf, ~crc);
        s.Write(buf);
    }
}