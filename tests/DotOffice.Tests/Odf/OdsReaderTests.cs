using System.IO.Compression;
using System.Text;
using DotOffice.Core;
using DotOffice.Odf;

namespace DotOffice.Tests.Odf;

public class OdsReaderTests
{
    private const string Header = "<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
        "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" " +
        "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" " +
        "xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\">";

    private static MemoryStream Package(string content, string? mimetype = "application/vnd.oasis.opendocument.spreadsheet", string? manifest = null)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string name, string text)
            {
                using var w = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                w.Write(text);
            }
            if (mimetype is not null) Add("mimetype", mimetype);
            if (manifest is not null) Add("META-INF/manifest.xml", manifest);
            Add("content.xml", content);
        }
        ms.Position = 0;
        return ms;
    }

    private static Workbook Read(string autoStyles, string body) =>
        OdsReader.Read(Package(Header + "<office:automatic-styles>" + autoStyles + "</office:automatic-styles><office:body><office:spreadsheet>" + body +
            "</office:spreadsheet></office:body></office:document-content>"));

    [Fact]
    public void ReadsTextNumbersAndFormats()
    {
        var wb = Read(
            "<number:number-style style:name=\"N1\"><number:number number:decimal-places=\"2\" number:grouping=\"true\"/></number:number-style>" +
            "<style:style style:name=\"ce1\" style:family=\"table-cell\" style:data-style-name=\"N1\"><style:text-properties fo:font-weight=\"bold\" fo:color=\"#ff0000\"/></style:style>",
            "<table:table table:name=\"S\"><table:table-row>" +
            "<table:table-cell office:value-type=\"string\"><text:p>Hi <text:s text:c=\"2\"/>there</text:p></table:table-cell>" +
            "<table:table-cell table:style-name=\"ce1\" office:value-type=\"float\" office:value=\"1234.5\"/>" +
            "</table:table-row></table:table>");
        var sheet = Assert.Single(wb.Sheets);
        Assert.Equal("S", sheet.Name);
        Assert.Equal("Hi   there", sheet.Cells[(0, 0)].Text);
        var number = sheet.Cells[(0, 1)];
        Assert.Equal("1,234.50", number.Text);
        Assert.True(number.Style.Font.Bold);
        Assert.Equal(new Color(255, 0, 0), number.Style.Font.Color);
    }

    [Fact]
    public void ReadsDatesRepeatsAndMerges()
    {
        var wb = Read(
            "<number:date-style style:name=\"D1\"><number:year number:style=\"long\"/><number:text>-</number:text><number:month number:style=\"long\"/><number:text>-</number:text><number:day number:style=\"long\"/></number:date-style>" +
            "<style:style style:name=\"ce1\" style:family=\"table-cell\" style:data-style-name=\"D1\"/>",
            "<table:table table:name=\"S\"><table:table-row table:number-rows-repeated=\"2\">" +
            "<table:table-cell table:style-name=\"ce1\" office:value-type=\"date\" office:date-value=\"2020-03-07\" table:number-columns-repeated=\"2\" table:number-columns-spanned=\"2\" table:number-rows-spanned=\"2\"/>" +
            "</table:table-row></table:table>");
        var sheet = wb.Sheets[0];
        Assert.Equal("2020-03-07", sheet.Cells[(1, 1)].Text);
        Assert.Equal(4, sheet.Cells.Count);
        Assert.Equal(new CellRange(0, 0, 1, 1), sheet.Merges[0]);
    }

    [Fact]
    public void ReadsColumnWidthsAndRowHeights()
    {
        var wb = Read(
            "<style:style style:name=\"co1\" style:family=\"table-column\"><style:table-column-properties style:column-width=\"1in\"/></style:style>" +
            "<style:style style:name=\"ro1\" style:family=\"table-row\"><style:table-row-properties style:row-height=\"0.5in\"/></style:style>",
            "<table:table table:name=\"S\"><table:table-column table:style-name=\"co1\" table:number-columns-repeated=\"2\"/>" +
            "<table:table-row table:style-name=\"ro1\"><table:table-cell office:value-type=\"boolean\" office:boolean-value=\"true\"/></table:table-row></table:table>");
        var sheet = wb.Sheets[0];
        Assert.Equal(72, sheet.ColumnWidths[1], 3);
        Assert.Equal(36, sheet.RowHeights[0], 3);
        Assert.Equal("TRUE", sheet.Cells[(0, 0)].Text);
    }

    [Fact]
    public void HugeBlankRepeatsStayCheap()
    {
        var wb = Read("", "<table:table table:name=\"S\"><table:table-row table:number-rows-repeated=\"1048576\"><table:table-cell table:number-columns-repeated=\"16384\"/></table:table-row></table:table>");
        Assert.Empty(wb.Sheets[0].Cells);
    }

    [Fact]
    public void RejectsEncryptedWrongTypeAndGarbage()
    {
        const string content = Header + "<office:body><office:spreadsheet><table:table table:name=\"S\"/></office:spreadsheet></office:body></office:document-content>";
        Assert.Throws<InvalidDataException>(() => OdsReader.Read(Package(content, manifest: "<m><file-entry><encryption-data/></file-entry></m>")));
        Assert.Throws<InvalidDataException>(() => OdsReader.Read(Package(content, "application/vnd.oasis.opendocument.text")));
        Assert.Throws<InvalidDataException>(() => OdsReader.Read(new MemoryStream(Encoding.ASCII.GetBytes("not a zip"))));
        Assert.Throws<InvalidDataException>(() => OdsReader.Read(Package("<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\">")));
    }

    [Fact]
    public void ReadsPrintRangesAndHeaders()
    {
        var ms = Package(Header + "<office:automatic-styles/><office:body><office:spreadsheet><table:table table:name=\"S\" table:print-ranges=\"S.A1:S.C5 S.$E$7:.F8\">" +
            "<table:table-row><table:table-cell office:value-type=\"string\"><text:p>x</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>");
        var wb = OdsReader.Read(ms);
        Assert.Equal([new CellRange(0, 0, 4, 2), new CellRange(6, 4, 7, 5)], wb.Sheets[0].PageSetup.PrintAreas);
    }

    [Fact]
    public void ReadsFlatFileWithHeaderAndFooter()
    {
        var flat = Header.Replace("document-content", "document").Replace("<office:document ", "<office:document office:mimetype=\"application/vnd.oasis.opendocument.spreadsheet\" ") +
            "<office:automatic-styles><style:page-layout style:name=\"pm1\"><style:page-layout-properties fo:page-width=\"21cm\" fo:page-height=\"29.7cm\"/></style:page-layout></office:automatic-styles>" +
            "<office:master-styles><style:master-page style:name=\"Default\" style:page-layout-name=\"pm1\"><style:header><style:region-center><text:p>Top <text:page-number>1</text:page-number></text:p></style:region-center></style:header></style:master-page></office:master-styles>" +
            "<office:body><office:spreadsheet><table:table table:name=\"S\"><table:table-row><table:table-cell office:value-type=\"string\"><text:p>x</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document>";
        var wb = OdsReader.Read(new MemoryStream(Encoding.UTF8.GetBytes(flat)));
        Assert.Equal("Top &P", wb.Sheets[0].PageSetup.Header.Odd.Center);
        Assert.Equal("x", wb.Sheets[0].Cells[(0, 0)].Text);
    }
}
