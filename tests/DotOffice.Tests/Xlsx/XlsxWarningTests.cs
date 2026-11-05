using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxWarningTests
{
    private const string RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    private static string[] Codes(XlsxBuilder builder)
    {
        var warnings = new ConversionWarningList();
        XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings });
        return warnings.Warnings.Select(w => w.Code).ToArray();
    }

    private static XlsxBuilder WithRelationship(string type, string target = "../x/part1.xml")
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")));
        builder.AddPart("xl/worksheets/_rels/sheet1.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}{type}\" Target=\"{target}\"/></Relationships>");
        builder.AddPart("xl/x/part1.xml", "<x/>");
        return builder;
    }

    [Fact]
    public void A_drawing_relationship_without_a_drawing_element_is_ignored_like_excel_does()
    {
        Assert.Empty(Codes(WithRelationship("drawing")));
    }

    [Fact]
    public void A_drawing_that_is_not_a_drawing_part_warns_once()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: XlsxBuilder.DrawingElement);
        builder.AddPart("xl/worksheets/_rels/sheet1.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../x/part1.xml\"/></Relationships>");
        builder.AddPart("xl/x/part1.xml", "<x/>");
        Assert.Equal(["xlsx.drawing"], Codes(builder));
    }

    [Fact]
    public void Drawing_element_without_a_relationship_part_warns()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: XlsxBuilder.DrawingElement);
        Assert.Equal(["xlsx.drawing"], Codes(builder));
    }

    [Fact]
    public void A_malformed_drawing_warns_instead_of_failing()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: XlsxBuilder.DrawingElement);
        builder.AddPart("xl/worksheets/_rels/sheet1.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/d.xml\"/></Relationships>");
        builder.AddPart("xl/drawings/d.xml", "<xdr:wsDr xmlns:xdr=\"urn:x\"><xdr:oneCellAnchor>");
        Assert.Equal(["xlsx.drawing"], Codes(builder));
    }

    [Fact]
    public void A_pivot_table_relationship_warns()
    {
        Assert.Equal(["xlsx.pivot"], Codes(WithRelationship("pivotTable")));
    }

    [Fact]
    public void Conditional_formatting_warns_once_per_sheet_and_data_validations_do_not_warn()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")),
            tail: "<conditionalFormatting sqref=\"A1\"><cfRule type=\"unknownKind\" priority=\"1\"/></conditionalFormatting>" +
                  "<conditionalFormatting sqref=\"B1\"><cfRule type=\"unknownKind\" priority=\"2\"/></conditionalFormatting>" +
                  "<dataValidations count=\"1\"><dataValidation sqref=\"A1\"/></dataValidations>");
        Assert.Equal(["xlsx.conditional-formatting"], Codes(builder));
    }

    [Fact]
    public void Elements_after_conditional_formatting_are_still_read()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")),
            tail: "<conditionalFormatting sqref=\"A1\"><cfRule type=\"unknownKind\" priority=\"1\"/></conditionalFormatting><mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells>");
        var wb = XlsxReader.Read(builder.BuildStream());
        Assert.Single(wb.Sheets[0].Merges);
    }

    [Fact]
    public void Rotated_text_no_longer_warns()
    {
        var builder = new XlsxBuilder { Styles = "<fonts><font/></fonts><fills><fill/></fills><borders><border/></borders><cellXfs><xf fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment textRotation=\"45\"/></xf></cellXfs>" }
            .AddSheet("S", Row(1, C("A1", "1", style: 0)));
        Assert.Empty(Codes(builder));
    }

    [Fact]
    public void Plain_workbooks_produce_no_warnings()
    {
        Assert.Empty(Codes(new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")))));
    }
}
