using System.IO.Compression;
using System.Text;
using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using DotOffice.Pdf;
using DotOffice.Tests.Docx;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Layout;

public class ComplexScriptTests
{
    private static readonly TrueTypeFont Arabic = TrueTypeFont.Load(TestAssets.Path("NotoSansArabic-Regular.ttf"));
    private static readonly TrueTypeFont Hebrew = TrueTypeFont.Load(TestAssets.Path("NotoSansHebrew-Regular.ttf"));

    private static readonly RunProperties Rtl = Text with { ComplexScriptFontFamily = "Noto Sans Arabic", RightToLeft = true };

    private static LayoutPage Layout(params Block[] blocks) =>
        new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto).Add(Arabic).Add(Hebrew)) })
            .Layout(Doc(blocks)).Pages[0];

    private static ParagraphProperties RtlParagraph => ParagraphProperties.Default with { RightToLeft = true };

    [Fact]
    public void Reads_complex_script_run_and_paragraph_properties()
    {
        var doc = DocxReader.Read(new DocxBuilder
        {
            Theme = "<a:fontScheme name=\"T\"><a:majorFont><a:latin typeface=\"Cambria\"/><a:cs typeface=\"Times New Roman\"/></a:majorFont>" +
                    "<a:minorFont><a:latin typeface=\"Calibri\"/><a:cs typeface=\"Arial\"/></a:minorFont></a:fontScheme>",
            Body = P(R("x", "<w:rFonts w:cs=\"Noto Sans Arabic\"/><w:bCs/><w:iCs/><w:szCs w:val=\"28\"/><w:sz w:val=\"20\"/><w:rtl/><w:kern w:val=\"16\"/>")
                     + R("y", "<w:rFonts w:cstheme=\"minorBidi\"/><w:cs/>"), "<w:bidi/><w:jc w:val=\"left\"/>"),
        }.BuildStream());
        var p = (Paragraph)doc.Sections[0].Blocks[0];
        Assert.True(p.Properties.RightToLeft);
        var x = p.Inlines[0].Properties;
        Assert.Equal(("Noto Sans Arabic", 14.0, true, true, true, 8.0),
            (x.ComplexScriptFontFamily, x.ComplexScriptFontSize, x.ComplexScriptBold, x.ComplexScriptItalic, x.RightToLeft, x.KerningMinSize));
        Assert.Equal(10, x.FontSize);
        var y = p.Inlines[1].Properties;
        Assert.Equal(("Arial", true, false, (bool?)null), (y.ComplexScriptFontFamily, y.ComplexScript, y.RightToLeft, y.ComplexScriptBold));
    }

    [Fact]
    public void Rtl_paragraph_starts_at_the_right_and_reads_right_to_left()
    {
        var page = Layout(Para("אחת שתיים", RtlParagraph, Text with { ComplexScriptFontFamily = "Noto Sans Hebrew" }));
        var first = Texts(page).Single(t => t.Text == "אחת");
        var second = Texts(page).Single(t => t.Text == "שתיים");
        Assert.True(first.RightToLeft && first.Glyphs is not null);
        Assert.Equal(72 + 468, Right(first), 1);
        Assert.True(second.X < first.X);
    }

    [Fact]
    public void Rtl_words_inside_ltr_paragraphs_are_reversed_in_place()
    {
        var p = new Paragraph { MarkProperties = Text };
        p.Inlines.AddRange([new TextRun("left ", Text), new TextRun("كلمة أخرى", Rtl), new TextRun(" right", Text)]);
        var items = Texts(Layout(p)).ToList();
        double X(string text) => items.Single(t => t.Text == text).X;
        Assert.True(X("left") < X("أخرى") && X("أخرى") < X("كلمة") && X("كلمة") < X("right"));
        Assert.Equal(72, X("left"), 2);
    }

    [Fact]
    public void Latin_words_in_an_rtl_run_keep_their_order()
    {
        var p = new Paragraph { Properties = RtlParagraph, MarkProperties = Text };
        p.Inlines.AddRange([new TextRun("عربي ", Rtl), new TextRun("English words.", Rtl)]);
        var items = Texts(Layout(p)).ToList();
        double X(string text) => items.Single(t => t.Text.Trim() == text).X;
        // The space between the English words stays left-to-right; the full stop still follows the RTL run.
        Assert.True(X("English") < X("words") && X("words") < X("عربي"));
        Assert.True(X(".") < X("English"));
    }

    [Fact]
    public void Numbers_and_mirrored_brackets_in_rtl_text()
    {
        var page = Layout(Para("(שלום) 123", RtlParagraph, Text with { ComplexScriptFontFamily = "Noto Sans Hebrew", RightToLeft = true }));
        var items = Texts(page).OrderBy(t => t.X).ToList();
        // Visual order, left to right; the Hebrew font has no brackets, so they come from Noto Sans.
        Assert.Equal(["123", " ", ")", "שלום", "("], items.Select(t => t.Text));
        // The closing bracket ends up on the left, drawn with the mirrored glyph but keeping ")" as its text.
        var closing = items[2];
        Assert.True(closing.RightToLeft);
        Assert.Equal((Noto.GetGlyphId('('), ")"), (closing.Glyphs![0].GlyphId, closing.Glyphs[0].Text));
    }

    [Fact]
    public void Arabic_text_is_shaped_with_the_complex_script_font_and_size()
    {
        var run = Rtl with { ComplexScriptFontSize = 16 };
        var page = Layout(Para("سلام", RtlParagraph, run));
        var item = Texts(page).Single();
        Assert.Same(Arabic, item.Font);
        Assert.Equal(16, item.FontSize);
        Assert.Equal([37, 71, 11, 74], item.Glyphs!.Select(g => (int)g.GlyphId)); // joining forms
        Assert.Equal((784 + 220 + 379 + 484) * 16.0 / Arabic.UnitsPerEm, 72 + 468 - item.X, 3);
    }

    [Fact]
    public void Kerning_applies_from_the_requested_size()
    {
        double Width(double? kern) => Texts(Layout(Para("AV", run: Text with { FontSize = 20, KerningMinSize = kern }))).Sum(t =>
            t.Glyphs is { } g ? g.Sum(x => x.Advance) * t.FontSize / Noto.UnitsPerEm : t.Font.MeasureText(t.Text, t.FontSize));
        Assert.True(Width(10) < Width(null) - 0.5);
        Assert.Equal(Width(null), Width(24), 6); // below the threshold: no kerning
    }

    [Fact]
    public void Pdf_maps_shared_glyphs_to_distinct_codes_so_text_extraction_stays_exact()
    {
        var doc = new PdfDocument();
        var font = doc.AddFont(Arabic);
        doc.AddPage(PageSize.A4).Canvas.DrawGlyphs(font, 12, 700,
            [new PositionedGlyph(14, "ب", 100, 0), new PositionedGlyph(14, "ن", 110, 0), new PositionedGlyph(316, "", 120, 0)]);
        using var ms = new MemoryStream();
        doc.Save(ms);
        var pdf = Encoding.Latin1.GetString(ms.ToArray());
        Assert.Matches(@"/CIDToGIDMap \d+ 0 R", pdf);
        var cmap = Inflated(ms.ToArray()).Single(s => s.Contains("beginbfchar"));
        Assert.Contains("<000E> <0628>", cmap); // first text keeps the glyph id as code
        Assert.Matches($"<{Arabic.NumGlyphs:X4}> <0646>", cmap); // second text gets a new code
        Assert.DoesNotContain("<013C>", cmap); // the decomposition dot has no text
    }

    private static IEnumerable<string> Inflated(byte[] pdf)
    {
        var text = Encoding.Latin1.GetString(pdf);
        for (int i = text.IndexOf("stream\n", StringComparison.Ordinal); i >= 0; i = text.IndexOf("stream\n", i + 1, StringComparison.Ordinal))
        {
            int start = i + 7, end = text.IndexOf("\nendstream", start, StringComparison.Ordinal);
            if (end < 0) break;
            string result;
            try
            {
                using var z = new ZLibStream(new MemoryStream(pdf, start, end - start), CompressionMode.Decompress);
                using var reader = new StreamReader(z, Encoding.Latin1);
                result = reader.ReadToEnd();
            }
            catch (InvalidDataException) { continue; }
            yield return result;
        }
    }
}
