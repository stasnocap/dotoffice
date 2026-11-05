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

    private static XlsxBuilder WithPivot(string definition, string? cache = null)
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")));
        builder.AddPart("xl/worksheets/_rels/sheet1.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}pivotTable\" Target=\"../pivotTables/pivotTable1.xml\"/></Relationships>");
        builder.AddPart("xl/pivotTables/pivotTable1.xml", definition);
        if (cache is not null)
        {
            builder.AddPart("xl/pivotTables/_rels/pivotTable1.xml.rels",
                $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition1.xml\"/></Relationships>");
            builder.AddPart("xl/pivotCache/pivotCacheDefinition1.xml", cache);
        }
        return builder;
    }

    private const string PivotNs = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

    private static string Messages(XlsxBuilder builder)
    {
        var warnings = new ConversionWarningList();
        XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings });
        return string.Join("|", warnings.Warnings.Where(w => w.Code == "xlsx.pivot").Select(w => w.Message));
    }

    [Fact]
    public void A_pivot_table_without_style_formats_or_refresh_prints_its_cells_without_a_warning()
    {
        Assert.Empty(Codes(WithPivot($"<pivotTableDefinition {PivotNs} name=\"Sales\"><pivotTableStyleInfo showRowStripes=\"1\"/></pivotTableDefinition>",
            $"<pivotCacheDefinition {PivotNs} refreshOnLoad=\"0\"/>")));
        Assert.Empty(Codes(WithPivot($"<pivotTableDefinition {PivotNs} name=\"Sales\"/>")));
    }

    [Fact]
    public void A_pivot_style_pivot_area_formats_and_refresh_on_load_are_named_in_the_warning()
    {
        var styled = Messages(WithPivot($"<pivotTableDefinition {PivotNs} name=\"Sales\"><pivotTableStyleInfo name=\"PivotStyleLight16\"/></pivotTableDefinition>"));
        Assert.Contains("'Sales'", styled);
        Assert.Contains("PivotStyleLight16", styled);
        Assert.DoesNotContain("refresh", styled);
        var all = Messages(WithPivot($"<pivotTableDefinition {PivotNs} name=\"P\"><location ref=\"A1:B3\"/><formats count=\"1\"><format dxfId=\"0\"><pivotArea field=\"0\"><references count=\"1\"><reference field=\"0\" count=\"1\"><x v=\"1\"/></reference></references></pivotArea></format></formats><pivotTableStyleInfo name=\"PivotStyleDark3\"/></pivotTableDefinition>",
            $"<pivotCacheDefinition {PivotNs} refreshOnLoad=\"1\"/>"));
        Assert.Contains("PivotStyleDark3", all);
        Assert.Contains("is not applied", all);
        Assert.Contains("pivot-area format", all);
        Assert.Contains("particular items", all);
        Assert.Contains("refresh when the file opens", all);
        Assert.Equal(["xlsx.pivot"], Codes(WithPivot($"<pivotTableDefinition {PivotNs} name=\"P\"/>", $"<pivotCacheDefinition {PivotNs} refreshOnLoad=\"true\"/>")));
    }

    [Fact]
    public void A_pivot_part_that_cannot_be_read_still_warns()
    {
        Assert.Equal(["xlsx.pivot"], Codes(WithRelationship("pivotTable")));
        Assert.Equal(["xlsx.pivot"], Codes(WithPivot("<pivotTableDefinition")));
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
