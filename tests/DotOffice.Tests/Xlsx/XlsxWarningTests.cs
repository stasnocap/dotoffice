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
    public void Drawing_relationship_warns_once()
    {
        Assert.Equal(["xlsx.drawing"], Codes(WithRelationship("drawing")));
    }

    [Fact]
    public void Drawing_element_warns_even_without_a_relationship_part()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: "<drawing r:id=\"rId1\"/>");
        Assert.Equal(["xlsx.drawing"], Codes(builder));
    }

    [Fact]
    public void Each_sheet_with_drawings_warns_once()
    {
        var builder = WithRelationship("drawing");
        builder.AddSheet("T", Row(1, C("A1", "1")), tail: "<drawing r:id=\"rId1\"/>");
        Assert.Equal(["xlsx.drawing", "xlsx.drawing"], Codes(builder)); // one per sheet
    }

    [Fact]
    public void Pivot_and_comments_relationships_warn()
    {
        Assert.Equal(["xlsx.pivot"], Codes(WithRelationship("pivotTable")));
        Assert.Equal(["xlsx.comments"], Codes(WithRelationship("comments")));
    }

    [Fact]
    public void Conditional_formatting_warns_once_per_sheet_and_data_validations_do_not_warn()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")),
            tail: "<conditionalFormatting sqref=\"A1\"><cfRule type=\"cellIs\" priority=\"1\"/></conditionalFormatting>" +
                  "<conditionalFormatting sqref=\"B1\"><cfRule type=\"cellIs\" priority=\"2\"/></conditionalFormatting>" +
                  "<dataValidations count=\"1\"><dataValidation sqref=\"A1\"/></dataValidations>");
        Assert.Equal(["xlsx.conditional-formatting"], Codes(builder));
    }

    [Fact]
    public void Elements_after_conditional_formatting_are_still_read()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")),
            tail: "<conditionalFormatting sqref=\"A1\"><cfRule type=\"cellIs\" priority=\"1\"/></conditionalFormatting><mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells>");
        var wb = XlsxReader.Read(builder.BuildStream());
        Assert.Single(wb.Sheets[0].Merges);
    }

    [Fact]
    public void Rotated_text_warns()
    {
        var builder = new XlsxBuilder { Styles = "<fonts><font/></fonts><fills><fill/></fills><borders><border/></borders><cellXfs><xf fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment textRotation=\"45\"/></xf></cellXfs>" }
            .AddSheet("S", Row(1, C("A1", "1", style: 0)));
        Assert.Equal(["xlsx.rotation"], Codes(builder));
    }

    [Fact]
    public void Plain_workbooks_produce_no_warnings()
    {
        Assert.Empty(Codes(new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")))));
    }
}
