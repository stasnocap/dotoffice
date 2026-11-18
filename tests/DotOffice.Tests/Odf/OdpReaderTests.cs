using System.IO.Compression;
using System.Text;
using DotOffice.Core;
using DotOffice.Odf;

namespace DotOffice.Tests.Odf;

public class OdpReaderTests
{
    private const string Header = "<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
        "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" " +
        "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" " +
        "xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" " +
        "xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\">";

    private static Presentation Read(string autoStyles, string pages, string masterStyles = "", string mime = "application/vnd.oasis.opendocument.presentation")
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
            Add("styles.xml", Header.Replace("document-content", "document-styles") +
                "<office:automatic-styles><style:page-layout style:name=\"PM1\"><style:page-layout-properties fo:page-width=\"28cm\" fo:page-height=\"21cm\"/></style:page-layout></office:automatic-styles>" +
                "<office:master-styles><style:master-page style:name=\"Default\" style:page-layout-name=\"PM1\"/>" + masterStyles + "</office:master-styles></office:document-styles>");
            Add("content.xml", Header + "<office:automatic-styles>" + autoStyles + "</office:automatic-styles><office:body><office:presentation>" + pages +
                "</office:presentation></office:body></office:document-content>");
        }
        ms.Position = 0;
        return OdpReader.Read(ms);
    }

    [Fact]
    public void ReadsSlideSizeAndTextFrames()
    {
        var presentation = Read("<style:style style:name=\"P1\" style:family=\"paragraph\"><style:text-properties fo:font-size=\"30pt\"/></style:style>",
            "<draw:page draw:master-page-name=\"Default\"><draw:frame svg:x=\"1cm\" svg:y=\"2cm\" svg:width=\"10cm\" svg:height=\"3cm\"><draw:text-box>" +
            "<text:p text:style-name=\"P1\">Hello</text:p></draw:text-box></draw:frame></draw:page>");
        var slide = Assert.Single(presentation.Slides);
        Assert.Equal(793.7, slide.Width, 1);
        Assert.Equal(595.3, slide.Height, 1);
        var shape = Assert.IsType<SheetShape>(Assert.Single(slide.Objects));
        Assert.Equal(28.35, shape.X, 0);
        Assert.Equal("Hello", Assert.Single(Assert.Single(shape.Paragraphs).Runs).Text);
        Assert.Equal(30, shape.Paragraphs[0].Runs[0].Font.FontSize);
    }

    [Fact]
    public void SkipsHiddenSlidesAndFillsShapes()
    {
        var presentation = Read(
            "<style:style style:name=\"dp1\" style:family=\"drawing-page\"><style:drawing-page-properties presentation:visibility=\"hidden\"/></style:style>" +
            "<style:style style:name=\"gr1\" style:family=\"graphic\"><style:graphic-properties draw:fill=\"solid\" draw:fill-color=\"#ff0000\" draw:stroke=\"none\"/></style:style>",
            "<draw:page draw:style-name=\"dp1\" draw:master-page-name=\"Default\"/>" +
            "<draw:page draw:master-page-name=\"Default\"><draw:rect draw:style-name=\"gr1\" svg:x=\"0cm\" svg:y=\"0cm\" svg:width=\"2cm\" svg:height=\"2cm\"/></draw:page>");
        var shape = Assert.IsType<SheetShape>(Assert.Single(Assert.Single(presentation.Slides).Objects));
        Assert.Equal(new Color(255, 0, 0), shape.Fill);
        Assert.Null(shape.LineColor);
    }

    [Fact]
    public void ReadsBulletsAndTableCells()
    {
        var presentation = Read(
            "<text:list-style style:name=\"L1\"><text:list-level-style-bullet text:level=\"1\" text:bullet-char=\"-\"><style:list-level-properties text:space-before=\"0.5cm\" text:min-label-width=\"0.5cm\"/></text:list-level-style-bullet></text:list-style>",
            "<draw:page draw:master-page-name=\"Default\"><draw:frame svg:x=\"0cm\" svg:y=\"0cm\" svg:width=\"10cm\" svg:height=\"3cm\"><draw:text-box>" +
            "<text:list text:style-name=\"L1\"><text:list-item><text:p>one</text:p></text:list-item></text:list></draw:text-box></draw:frame>" +
            "<draw:frame svg:x=\"0cm\" svg:y=\"5cm\" svg:width=\"6cm\" svg:height=\"2cm\"><table:table><table:table-column/><table:table-column/>" +
            "<table:table-row><table:table-cell><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row></table:table></draw:frame></draw:page>");
        var objects = presentation.Slides[0].Objects.OfType<SheetShape>().ToList();
        var bullet = objects[0].Paragraphs[0];
        Assert.StartsWith("-", bullet.Runs[0].Text);
        Assert.Equal(14.17, bullet.Indent, 0);
        Assert.Equal(28.35, bullet.Margin!.Value, 0);
        Assert.Equal(2, objects.Count(o => o.TableRows is not null && o.Paragraphs.Count > 0));
    }

    [Fact]
    public void RejectsOtherPackages() =>
        Assert.Throws<InvalidDataException>(() => Read("", "<draw:page/>", mime: "application/vnd.oasis.opendocument.text"));
}
