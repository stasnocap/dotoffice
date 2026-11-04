using DotOffice.Core;

namespace DotOffice.Tests.Core;

public class SheetModelTests
{
    [Fact]
    public void Range_dimensions_and_containment()
    {
        var r = new CellRange(1, 2, 4, 5);
        Assert.Equal(4, r.RowCount);
        Assert.Equal(4, r.ColumnCount);
        Assert.True(r.Contains(1, 2));
        Assert.True(r.Contains(4, 5));
        Assert.False(r.Contains(0, 2));
        Assert.False(r.Contains(4, 6));
    }

    [Fact]
    public void Hidden_rows_and_columns_have_zero_size()
    {
        var sheet = new Worksheet("S");
        sheet.ColumnWidths[1] = 100;
        sheet.RowHeights[2] = 30;
        Assert.Equal(100, sheet.ColumnWidth(1));
        Assert.Equal(sheet.DefaultColumnWidth, sheet.ColumnWidth(0));
        Assert.Equal(30, sheet.RowHeight(2));
        sheet.HiddenColumns.Add(1);
        sheet.HiddenRows.Add(2);
        Assert.Equal(0, sheet.ColumnWidth(1));
        Assert.Equal(0, sheet.RowHeight(2));
    }

    [Fact]
    public void Defaults_match_excel()
    {
        var setup = new SheetPageSetup();
        Assert.Equal(1.0, setup.Scale);
        Assert.Null(setup.FitToWidth);
        Assert.True(setup.Header.Odd.IsEmpty);
        Assert.Equal(HorizontalAlignment.General, CellStyle.Default.HorizontalAlignment);
        Assert.Equal("Calibri", CellStyle.Default.Font.FontFamily);
    }
}
