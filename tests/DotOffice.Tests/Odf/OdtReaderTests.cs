using System.IO.Compression;
using System.Text;
using DotOffice.Core;
using DotOffice.Odf;

namespace DotOffice.Tests.Odf;

public class OdtReaderTests
{
    private const string Header = "<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
        "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" " +
        "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\">";

    private static Document Read(string autoStyles, string body, string mime = "application/vnd.oasis.opendocument.text")
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string text)
            {
                using var w = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                w.Write(text);
            }
            Add("mimetype", mime);
            Add("content.xml", Header + "<office:automatic-styles>" + autoStyles + "</office:automatic-styles><office:body><office:text>" + body +
                "</office:text></office:body></office:document-content>");
        }
        ms.Position = 0;
        return OdtReader.Read(ms);
    }

    private static string Text(Paragraph p) => string.Concat(p.Inlines.OfType<TextRun>().Select(r => r.Text));

    [Fact]
    public void CollapsesWhiteSpaceAndKeepsExplicitSpaces()
    {
        var doc = Read("", "<text:p>  Hello \n   <text:span>big</text:span><text:s text:c=\"2\"/>world</text:p>");
        var p = Assert.IsType<Paragraph>(Assert.Single(doc.Sections[0].Blocks));
        Assert.Equal("Hello big  world", Text(p));
    }

    [Fact]
    public void AppliesParagraphAndCharacterStyles()
    {
        var doc = Read(
            "<style:style style:name=\"P1\" style:family=\"paragraph\"><style:paragraph-properties fo:text-align=\"center\" fo:margin-left=\"1in\"/><style:text-properties fo:font-size=\"20pt\"/></style:style>" +
            "<style:style style:name=\"T1\" style:family=\"text\"><style:text-properties fo:font-weight=\"bold\" fo:color=\"#00ff00\"/></style:style>",
            "<text:p text:style-name=\"P1\">a<text:span text:style-name=\"T1\">b</text:span></text:p>");
        var p = (Paragraph)doc.Sections[0].Blocks[0];
        Assert.Equal(ParagraphAlignment.Center, p.Properties.Alignment);
        Assert.Equal(72, p.Properties.LeftIndent, 3);
        var runs = p.Inlines.OfType<TextRun>().ToList();
        Assert.Equal(20, runs[0].Properties.FontSize);
        Assert.False(runs[0].Properties.Bold);
        Assert.True(runs[1].Properties.Bold);
        Assert.Equal(new Color(0, 255, 0), runs[1].Properties.Color);
    }

    [Fact]
    public void NumbersListItems()
    {
        var doc = Read(
            "<text:list-style style:name=\"L1\"><text:list-level-style-number text:level=\"1\" style:num-format=\"1\" style:num-suffix=\".\">" +
            "<style:list-level-properties><style:list-level-label-alignment fo:margin-left=\"36pt\" fo:text-indent=\"-18pt\"/></style:list-level-properties></text:list-level-style-number></text:list-style>",
            "<text:list text:style-name=\"L1\"><text:list-item><text:p>one</text:p></text:list-item><text:list-item><text:p>two</text:p></text:list-item></text:list>");
        var blocks = doc.Sections[0].Blocks.Cast<Paragraph>().ToList();
        Assert.Equal("1.", blocks[0].ListLabel!.Text);
        Assert.Equal("2.", blocks[1].ListLabel!.Text);
        Assert.Equal(36, blocks[0].Properties.LeftIndent);
        Assert.Equal(-18, blocks[0].Properties.FirstLineIndent);
    }

    [Fact]
    public void ReadsTablesWithSpans()
    {
        var doc = Read(
            "<style:style style:name=\"co1\" style:family=\"table-column\"><style:table-column-properties style:column-width=\"50pt\"/></style:style>",
            "<table:table><table:table-column table:style-name=\"co1\" table:number-columns-repeated=\"2\"/>" +
            "<table:table-row><table:table-cell table:number-rows-spanned=\"2\"><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row>" +
            "<table:table-row><table:covered-table-cell/><table:table-cell><text:p>c</text:p></table:table-cell></table:table-row></table:table>");
        var t = Assert.IsType<Table>(doc.Sections[0].Blocks[0]);
        Assert.Equal([50.0, 50.0], t.Grid);
        Assert.Equal(VerticalMerge.Restart, t.Rows[0].Cells[0].VerticalMerge);
        Assert.Equal(VerticalMerge.Continue, t.Rows[1].Cells[0].VerticalMerge);
        Assert.Equal(2, t.Rows[1].Cells.Count);
    }

    [Fact]
    public void RejectsWrongTypeAndGarbage()
    {
        Assert.Throws<InvalidDataException>(() => Read("", "<text:p/>", "application/vnd.oasis.opendocument.spreadsheet"));
        Assert.Throws<InvalidDataException>(() => OdtReader.Read(new MemoryStream(Encoding.ASCII.GetBytes("nope"))));
    }
}
