using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Layout;
using DotOffice.Tests.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class NoteTests
{
    private const string Separators =
        "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
        "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";

    private static string Footnote(int id, string text) =>
        $"<w:footnote w:id=\"{id}\"><w:p><w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteRef/></w:r>{R(" " + text)}</w:p></w:footnote>";

    private static string Ref(int id, string kind = "footnote") =>
        $"<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:{kind}Reference w:id=\"{id}\"/></w:r>";

    private static string Text(Note note) =>
        string.Concat(note.Blocks.OfType<Paragraph>().SelectMany(p => p.Inlines).OfType<TextRun>().Select(t => t.Text));

    [Fact]
    public void Reads_and_numbers_footnotes_and_endnotes()
    {
        var doc = DocxReader.Read(new DocxBuilder
        {
            Settings = "<w:footnotePr><w:numFmt w:val=\"lowerLetter\"/><w:numStart w:val=\"2\"/></w:footnotePr>",
            Footnotes = Separators + Footnote(1, "first note") + Footnote(2, "second note"),
            Endnotes = "<w:endnote w:id=\"1\"><w:p><w:r><w:endnoteRef/></w:r>" + R(" the end") + "</w:p></w:endnote>",
            Body = P(R("A") + Ref(2) + R("B") + Ref(1) + R("C") + Ref(1, "endnote"))
                   + P("<w:r><w:footnoteReference w:customMarkFollows=\"1\" w:id=\"1\"/><w:t>*</w:t></w:r>"),
        }.BuildStream());

        var refs = doc.Sections[0].Blocks.OfType<Paragraph>().SelectMany(p => p.Inlines).OfType<NoteReference>().ToList();
        Assert.Equal(["b", "c", "i", ""], refs.Select(r => r.Mark));
        Assert.Equal(VerticalPosition.Superscript, refs[0].Properties.VerticalPosition);
        Assert.Equal("b second note", Text(refs[0].Note));
        Assert.Equal("c first note", Text(refs[1].Note));
        Assert.Equal(NoteKind.Endnote, refs[2].Note.Kind);
        Assert.Same(refs[2].Note, Assert.Single(doc.Endnotes));
        Assert.Equal("i the end", Text(doc.Endnotes[0]));
        Assert.Equal(" first note", Text(refs[3].Note)); // custom mark: no number in the note either
    }

    private static readonly double LineHeight =
        Math.Max(Noto.WinAscent + Noto.WinDescent, Noto.Ascender - Noto.Descender + Noto.LineGap) * 10.0 / Noto.UnitsPerEm;

    private static Section Page(int lines) => new() { PageHeight = 144 + LineHeight * lines + 0.5, Margins = new Thickness(72, 72, 72, 72) };

    private static Note Note(NoteKind kind, params string[] lines)
    {
        var note = new Note(kind);
        foreach (var line in lines) note.Blocks.Add(Para(line));
        return note;
    }

    private static Paragraph WithNote(string text, Note note)
    {
        var p = Para(text);
        p.Inlines.Add(new NoteReference(note, "1", LayoutTestHelpers.Text with { VerticalPosition = VerticalPosition.Superscript }));
        return p;
    }

    private static int PageOf(LayoutDocument layout, string text) =>
        layout.Pages.FindIndex(p => Texts(p).Any(t => t.Text == text));

    [Fact]
    public void Footnotes_go_to_the_bottom_of_the_referencing_page_below_a_separator()
    {
        var layout = Render(Doc(Page(20), Para("body"), WithNote("ref", Note(NoteKind.Footnote, "footnote"))));
        var page = Assert.Single(layout.Pages);
        var note = Texts(page).Single(t => t.Text == "footnote");
        Assert.Equal(72 + 20 * LineHeight, note.Baseline + note.Font.WinDescent * 10.0 / note.Font.UnitsPerEm, 0);
        var separator = Assert.Single(page.Items.OfType<LineItem>());
        Assert.True(separator.Y1 < note.Baseline && separator.X2 - separator.X1 < 200);
    }

    [Fact]
    public void A_reference_moves_to_the_next_page_when_its_footnote_does_not_fit()
    {
        // 9 body lines leave one line of room: the reference line fits, its footnote (plus separator) does not.
        var body = Enumerable.Range(1, 9).Select(i => (Block)Para($"line{i}")).ToList();
        body.Add(WithNote("ref", Note(NoteKind.Footnote, "footnote")));
        var layout = Render(Doc(Page(10), [.. body]));
        Assert.Equal(1, PageOf(layout, "ref"));
        Assert.Equal(1, PageOf(layout, "footnote"));
    }

    [Fact]
    public void Long_footnotes_continue_on_the_next_page()
    {
        var lines = Enumerable.Range(1, 12).Select(i => $"note{i}").ToArray();
        var layout = Render(Doc(Page(10), WithNote("ref", Note(NoteKind.Footnote, lines)), Para("after")));
        Assert.Equal(0, PageOf(layout, "ref"));
        Assert.Equal(0, PageOf(layout, "note1"));
        Assert.Equal(1, PageOf(layout, "note12"));
        // The continuation separator spans the text width.
        var separator = layout.Pages[1].Items.OfType<LineItem>().Single();
        Assert.Equal(468, separator.X2 - separator.X1, 1);
    }

    [Fact]
    public void Endnotes_follow_the_last_section()
    {
        var doc = Doc(Page(40), WithNote("ref", Note(NoteKind.Endnote, "the endnote")), Para("last paragraph"));
        doc.Endnotes.Add(((NoteReference)((Paragraph)doc.Sections[0].Blocks[0]).Inlines[1]).Note);
        var page = Assert.Single(Render(doc).Pages);
        var last = Texts(page).Single(t => t.Text == "paragraph");
        var note = Texts(page).Single(t => t.Text == "endnote");
        Assert.True(note.Baseline > last.Baseline + LineHeight);
        Assert.True(note.Baseline < 72 + 6 * LineHeight); // right after the text, not at the page bottom
    }
}
