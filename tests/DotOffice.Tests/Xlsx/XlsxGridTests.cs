using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxGridTests
{
    private static Worksheet Read(string head = "", string rows = "", string tail = "", ConversionWarningList? warnings = null) =>
        XlsxReader.Read(new XlsxBuilder().AddSheet("S", rows, head, tail).BuildStream(), new XlsxReaderOptions { Warnings = warnings }).Sheets[0];

    [Theory]
    [InlineData(9.140625, 48)]  // 64 px
    [InlineData(10, 52.5)]      // 70 px
    [InlineData(0, 0)]
    [InlineData(20.7109375, 108.75)] // 145 px
    public void Converts_column_widths_to_points(double chars, double expected)
    {
        var sheet = Read(head: $"<cols><col min=\"2\" max=\"2\" width=\"{chars.ToString(System.Globalization.CultureInfo.InvariantCulture)}\" customWidth=\"1\"/></cols>");
        Assert.Equal(expected, sheet.ColumnWidth(1), 2);
    }

    [Fact]
    public void Column_ranges_and_hidden_columns()
    {
        var sheet = Read(head: "<cols><col min=\"1\" max=\"3\" width=\"10\"/><col min=\"5\" max=\"5\" width=\"10\" hidden=\"1\"/></cols>");
        Assert.Equal(52.5, sheet.ColumnWidth(0), 2);
        Assert.Equal(52.5, sheet.ColumnWidth(2), 2);
        Assert.Equal(48, sheet.ColumnWidth(3), 2);
        Assert.Equal(0, sheet.ColumnWidth(4));
    }

    [Fact]
    public void Default_widths_and_heights_come_from_sheetFormatPr()
    {
        Assert.Equal(48, Read().DefaultColumnWidth);
        Assert.Equal(15, Read().DefaultRowHeight);
        var sheet = Read(head: "<sheetFormatPr defaultColWidth=\"10\" defaultRowHeight=\"20.25\"/>");
        Assert.Equal(52.5, sheet.DefaultColumnWidth, 2);
        Assert.Equal(20.25, sheet.DefaultRowHeight);
        Assert.Equal(48, Read(head: "<sheetFormatPr baseColWidth=\"8\" defaultRowHeight=\"15\"/>").DefaultColumnWidth);
    }

    [Fact]
    public void Reads_row_heights_and_hidden_rows()
    {
        var sheet = Read(rows: "<row r=\"1\" ht=\"30\" customHeight=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"3\" hidden=\"1\"><c r=\"A3\"><v>2</v></c></row><row r=\"4\"/>");
        Assert.Equal(30, sheet.RowHeight(0));
        Assert.Equal(15, sheet.RowHeight(1));
        Assert.Equal(0, sheet.RowHeight(2));
        Assert.Equal(15, sheet.RowHeight(3));
    }

    [Fact]
    public void Reads_merged_cells_and_ignores_single_cell_merges()
    {
        var warnings = new ConversionWarningList();
        var sheet = Read(tail: "<mergeCells count=\"3\"><mergeCell ref=\"A1:C2\"/><mergeCell ref=\"E5\"/><mergeCell ref=\"bogus\"/></mergeCells>", warnings: warnings);
        Assert.Equal([new CellRange(0, 0, 1, 2)], sheet.Merges);
        Assert.Contains(warnings.Warnings, w => w.Code == "xlsx.bad-reference");
    }

    [Fact]
    public void Cells_are_still_read_after_the_grid_elements()
    {
        var sheet = Read(head: "<sheetFormatPr defaultRowHeight=\"15\"/><cols><col min=\"1\" max=\"1\" width=\"12\"/></cols>",
            rows: Row(1, C("A1", "7")), tail: "<mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells>");
        Assert.Equal("7", sheet.Cells[(0, 0)].Text);
        Assert.Single(sheet.Merges);
    }
}
