using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxPivotTests
{
    private const string PivotNs = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
    private const string RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private static readonly Color Accent1 = Color.TryParseHex("4472C4")!.Value;

    /// <summary>Row Labels / Sum header in row 3, four regions in rows 4-7 and a Grand Total in row 8 (columns A and B).</summary>
    private static (Worksheet Sheet, ConversionWarningList Warnings) Read(string definition, params string[] dxfs)
    {
        var rows = Row(3, Inline("A3", "Row Labels") + Inline("B3", "Sum of Sales"));
        string[] regions = ["North", "South", "East", "West"];
        for (int i = 0; i < regions.Length; i++) rows += Row(4 + i, Inline($"A{4 + i}", regions[i]) + C($"B{4 + i}", $"{(i + 1) * 100}"));
        rows += Row(8, Inline("A8", "Grand Total") + C("B8", "1000"));
        var builder = new XlsxBuilder { Styles = StylesWithDxfs(dxfs.Length == 0 ? [Dxf("FFC7CE")] : dxfs) }.AddSheet("S", rows);
        builder.AddPart("xl/worksheets/_rels/sheet1.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}pivotTable\" Target=\"../pivotTables/pivotTable1.xml\"/></Relationships>");
        builder.AddPart("xl/pivotTables/pivotTable1.xml", definition);
        var warnings = new ConversionWarningList();
        return (XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings }).Sheets[0], warnings);
    }

    private static string Definition(string style = "", string extra = "", string location = "<location ref=\"A3:B8\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>", string flags = "showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\"") =>
        $"<pivotTableDefinition {PivotNs} name=\"PivotTable1\">{location}{extra}" + (style.Length == 0 ? "" : $"<pivotTableStyleInfo name=\"{style}\" {flags}/>") + "</pivotTableDefinition>";

    private static Cell At(Worksheet sheet, string reference)
    {
        int column = reference[0] - 'A', row = int.Parse(reference[1..]) - 1;
        return sheet.Cells[(row, column)];
    }

    [Fact]
    public void Light16_tints_and_bolds_the_header_and_sets_off_the_grand_total()
    {
        var (sheet, _) = Read(Definition("PivotStyleLight16"));
        var header = At(sheet, "B3");
        Assert.True(header.Style.Font.Bold);
        Assert.Equal(Accent1.Blend(Color.White, 0.8), header.Style.Fill);
        Assert.True(header.Style.Borders.Bottom.IsVisible);
        Assert.False(At(sheet, "A5").Style.Font.Bold);
        Assert.Null(At(sheet, "B5").Style.Fill);
        var grand = At(sheet, "B8");
        Assert.True(grand.Style.Font.Bold);
        Assert.True(grand.Style.Borders.Top.IsVisible);
        Assert.True(At(sheet, "A8").Style.Font.Bold);
    }

    [Fact]
    public void Medium2_has_a_header_in_the_accent_color_and_bands_the_rows_when_asked()
    {
        var (plain, _) = Read(Definition("PivotStyleMedium2"));
        Assert.Equal((Accent1, Color.White), (At(plain, "A3").Style.Fill, At(plain, "A3").Style.Font.Color));
        Assert.Null(At(plain, "B4").Style.Fill); // no stripes unless showRowStripes
        var (banded, _) = Read(Definition("PivotStyleMedium2", flags: "showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"1\""));
        Assert.Equal(Accent1.Blend(Color.White, 0.8), At(banded, "B4").Style.Fill);
        Assert.Null(At(banded, "B5").Style.Fill);
        Assert.Equal(Accent1.Blend(Color.White, 0.8), At(banded, "B6").Style.Fill);
        Assert.Equal(Accent1.Blend(Color.White, 0.8), At(banded, "A4").Style.Fill); // bands cross the label column too
        Assert.Null(At(banded, "B8").Style.Fill);                                    // but not the grand total
        Assert.True(At(banded, "B8").Style.Borders.Top.IsVisible);
    }

    [Fact]
    public void The_style_number_picks_the_accent_color_and_zero_means_black()
    {
        var accent2 = Color.TryParseHex("ED7D31")!.Value;
        Assert.Equal(accent2, At(Read(Definition("PivotStyleMedium3")).Sheet, "A3").Style.Fill);
        Assert.Equal(Color.Black, At(Read(Definition("PivotStyleMedium1")).Sheet, "A3").Style.Fill);
        Assert.Equal(accent2, At(Read(Definition("PivotStyleMedium10")).Sheet, "A3").Style.Fill); // Medium 8-14 family
    }

    [Fact]
    public void Headings_are_styled_only_when_the_table_shows_them_and_unknown_styles_change_nothing()
    {
        var (noHeaders, _) = Read(Definition("PivotStyleLight16", flags: "showRowHeaders=\"1\" showColHeaders=\"0\""));
        Assert.False(At(noHeaders, "B3").Style.Font.Bold);
        var (dark, warnings) = Read(Definition("PivotStyleDark3"));
        Assert.False(At(dark, "B3").Style.Font.Bold);
        Assert.Contains("is not applied", Assert.Single(warnings.Warnings).Message);
        var (custom, _) = Read(Definition("MyOwnStyle"));
        Assert.Null(At(custom, "B3").Style.Fill);
    }

    [Fact]
    public void A_style_is_reported_as_approximated_and_a_table_without_one_stays_quiet()
    {
        var (_, warnings) = Read(Definition("PivotStyleLight16"));
        var warning = Assert.Single(warnings.Warnings);
        Assert.Equal("xlsx.pivot", warning.Code);
        Assert.Contains("drawn approximately", warning.Message);
        Assert.Empty(Read(Definition()).Warnings.Warnings);
    }

    [Fact]
    public void Subtotal_and_grand_total_rows_and_columns_come_from_the_item_lists()
    {
        var items = "<rowItems count=\"6\"><i><x/></i><i><x v=\"1\"/></i><i t=\"default\"><x/></i><i><x v=\"2\"/></i><i><x v=\"3\"/></i><i t=\"grand\"><x/></i></rowItems>" +
                    "<colItems count=\"2\"><i><x/></i><i t=\"grand\"><x/></i></colItems>";
        // rows 4-9 are the items, so the "grand" item is row 9; the table here has rows to 8, so the list is cut at the table's last row.
        var (sheet, _) = Read(Definition("PivotStyleLight16", items, "<location ref=\"A3:B8\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>"));
        Assert.True(At(sheet, "B6").Style.Font.Bold);   // the subtotal item (third row of data = row 6)
        Assert.True(At(sheet, "B6").Style.Borders.Top.IsVisible);
        Assert.False(At(sheet, "B7").Style.Font.Bold);
        Assert.False(At(sheet, "B8").Style.Font.Bold);  // the list's grand total lies outside the table
        // The column list marks column C as the grand column, which is outside A:B here; the last data column carries no special look.
        Assert.False(At(sheet, "B5").Style.Font.Bold);
    }

    [Fact]
    public void A_table_with_column_fields_has_a_taller_header_and_a_grand_total_column()
    {
        var rows = Row(3, Inline("B3", "Column Labels")) + Row(4, Inline("A4", "Row Labels") + Inline("B4", "x") + Inline("C4", "Grand Total")) +
                   Row(5, Inline("A5", "n") + C("B5", "1") + C("C5", "1")) + Row(6, Inline("A6", "Grand Total") + C("B6", "1") + C("C6", "1"));
        var builder = new XlsxBuilder { Styles = StylesWithDxfs(Dxf("FFC7CE")) }.AddSheet("S", rows);
        builder.AddPart("xl/worksheets/_rels/sheet1.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}pivotTable\" Target=\"../pivotTables/pivotTable1.xml\"/></Relationships>");
        builder.AddPart("xl/pivotTables/pivotTable1.xml", $"<pivotTableDefinition {PivotNs} name=\"P\"><location ref=\"A3:C6\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\"/>" +
            "<rowItems count=\"2\"><i><x/></i><i t=\"grand\"><x/></i></rowItems><colItems count=\"2\"><i><x/></i><i t=\"grand\"><x/></i></colItems>" +
            "<pivotTableStyleInfo name=\"PivotStyleMedium2\" showRowHeaders=\"1\" showColHeaders=\"1\"/></pivotTableDefinition>");
        var sheet = XlsxReader.Read(builder.BuildStream()).Sheets[0];
        Assert.Equal(Accent1, At(sheet, "C3").Style.Fill); // two header rows, both in the header color
        Assert.Equal(Accent1, At(sheet, "B4").Style.Fill);
        Assert.True(At(sheet, "C5").Style.Font.Bold);       // the grand total column
        Assert.Null(At(sheet, "B5").Style.Fill);
    }

    [Theory]
    [InlineData("<pivotArea type=\"all\"/>", "A3,B3,A4,B4,A5,B5,A6,B6,A7,B7,A8,B8")]
    [InlineData("<pivotArea type=\"origin\" outline=\"0\"/>", "A3")]
    [InlineData("<pivotArea type=\"topRight\"/>", "B3")]
    [InlineData("<pivotArea type=\"button\" field=\"0\" axis=\"axisRow\"/>", "A3")]
    [InlineData("<pivotArea outline=\"0\" collapsedLevelsAreSubtotals=\"1\"/>", "B4,B5,B6,B7,B8")]               // the data body
    [InlineData("<pivotArea dataOnly=\"0\" labelOnly=\"1\" outline=\"0\"/>", "A4,A5,A6,A7,A8")]                // the row labels
    [InlineData("<pivotArea grandRow=\"1\" outline=\"0\" collapsedLevelsAreSubtotals=\"1\"/>", "B8")]         // grand total row, data cells
    [InlineData("<pivotArea dataOnly=\"0\" labelOnly=\"1\" grandRow=\"1\" outline=\"0\"/>", "A8")]            // grand total row, label
    public void Pivot_area_formats_that_select_whole_parts_apply_their_dxf(string area, string expected)
    {
        var (sheet, warnings) = Read(Definition(extra: $"<formats count=\"1\"><format dxfId=\"0\">{area}</format></formats>"), Dxf("FFC7CE"));
        var filled = sheet.Cells.Where(c => c.Value.Style.Fill == Color.TryParseHex("FFC7CE")).Select(c => $"{(char)('A' + c.Key.Column)}{c.Key.Row + 1}")
            .OrderBy(x => x[0]).ThenBy(x => int.Parse(x[1..]));
        Assert.Equal(expected.Split(',').OrderBy(x => x[0]).ThenBy(x => int.Parse(x[1..])), filled);
        Assert.Empty(warnings.Warnings);
    }

    [Fact]
    public void Formats_win_over_the_style_and_formats_that_pick_items_are_counted_not_applied()
    {
        var picks = "<pivotArea field=\"0\" type=\"normal\"><references count=\"1\"><reference field=\"0\" count=\"1\"><x v=\"2\"/></reference></references></pivotArea>";
        var extra = $"<formats count=\"2\"><format dxfId=\"0\"><pivotArea type=\"origin\"/></format><format dxfId=\"0\">{picks}</format></formats>";
        var (sheet, warnings) = Read(Definition("PivotStyleLight16", extra), Dxf("FFC7CE"));
        Assert.Equal(Color.TryParseHex("FFC7CE"), At(sheet, "A3").Style.Fill); // the format's fill replaces the style's tint
        Assert.True(At(sheet, "A3").Style.Font.Bold);                          // what the format does not set stays from the style
        Assert.Null(At(sheet, "A5").Style.Fill);
        var message = Assert.Single(warnings.Warnings).Message;
        Assert.Contains("a pivot-area format that selects particular items is not applied", message);
    }

    [Fact]
    public void A_table_whose_location_cannot_be_read_applies_nothing()
    {
        var (sheet, warnings) = Read(Definition("PivotStyleLight16", location: "<location ref=\"nonsense\"/>"));
        Assert.Null(At(sheet, "B3").Style.Fill);
        Assert.Contains("is not applied", Assert.Single(warnings.Warnings).Message);
    }
}
