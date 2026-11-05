using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxPageSetupTests
{
    private static Worksheet Read(string head = "", string tail = "", string rows = "", string? definedNames = null, ConversionWarningList? warnings = null) =>
        XlsxReader.Read(new XlsxBuilder { DefinedNames = definedNames }.AddSheet("S", rows, head, tail).BuildStream(), new XlsxReaderOptions { Warnings = warnings }).Sheets[0];

    [Fact]
    public void Defaults_to_letter_portrait_with_excel_margins()
    {
        var setup = Read().PageSetup;
        Assert.Equal(Orientation.Portrait, setup.Orientation);
        Assert.Equal(612, setup.PageWidth);
        Assert.Equal(1.0, setup.Scale);
        Assert.Null(setup.FitToWidth);
    }

    [Theory]
    [InlineData(1, 612, 792)]
    [InlineData(5, 612, 1008)]
    [InlineData(8, 841.89, 1190.55)]
    [InlineData(9, 595.28, 841.89)]
    [InlineData(11, 419.53, 595.28)]
    [InlineData(3, 792, 1224)]
    [InlineData(14, 612, 936)]
    [InlineData(12, 728.5, 1031.81)]
    [InlineData(66, 1190.55, 1683.78)]
    [InlineData(67, 297.64, 419.53)]
    public void Reads_paper_sizes(int paper, double width, double height)
    {
        var setup = Read(tail: $"<pageSetup paperSize=\"{paper}\"/>").PageSetup;
        Assert.Equal(width, setup.PageWidth, 2);
        Assert.Equal(height, setup.PageHeight, 2);
    }

    [Fact]
    public void Landscape_swaps_the_page_size_and_scale_is_clamped()
    {
        var setup = Read(tail: "<pageSetup paperSize=\"9\" orientation=\"landscape\" scale=\"75\" pageOrder=\"overThenDown\"/>").PageSetup;
        Assert.Equal(Orientation.Landscape, setup.Orientation);
        Assert.Equal(841.89, setup.PageWidth, 2);
        Assert.Equal(0.75, setup.Scale);
        Assert.True(setup.PageOrderOverThenDown);
        Assert.Equal(4.0, Read(tail: "<pageSetup scale=\"900\"/>").PageSetup.Scale);
    }

    [Fact]
    public void Warns_about_unknown_paper_sizes()
    {
        var warnings = new ConversionWarningList();
        var setup = Read(tail: "<pageSetup paperSize=\"99\"/>", warnings: warnings).PageSetup;
        Assert.Equal(612, setup.PageWidth);
        Assert.Contains(warnings.Warnings, w => w.Code == "xlsx.paper-size");
    }

    [Fact]
    public void Fit_to_page_needs_the_sheetPr_flag()
    {
        var off = Read(tail: "<pageSetup fitToWidth=\"2\" fitToHeight=\"0\"/>").PageSetup;
        Assert.Null(off.FitToWidth);
        var on = Read(head: "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>", tail: "<pageSetup fitToWidth=\"2\" fitToHeight=\"0\"/>").PageSetup;
        Assert.Equal(2, on.FitToWidth);
        Assert.Equal(0, on.FitToHeight);
        var defaults = Read(head: "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>", tail: "<pageSetup/>").PageSetup;
        Assert.Equal((1, 1), (defaults.FitToWidth, defaults.FitToHeight));
    }

    [Fact]
    public void Reads_margins_in_points_and_print_options()
    {
        var setup = Read(tail: "<printOptions horizontalCentered=\"1\" gridLines=\"true\"/><pageMargins left=\"1\" right=\"0.5\" top=\"2\" bottom=\"0.25\" header=\"0.1\" footer=\"0.2\"/>").PageSetup;
        Assert.Equal(new Thickness(144, 36, 18, 72), setup.Margins);
        Assert.Equal(7.2, setup.HeaderDistance, 3);
        Assert.Equal(14.4, setup.FooterDistance, 3);
        Assert.True(setup.CenterHorizontally);
        Assert.False(setup.CenterVertically);
        Assert.True(setup.PrintGridLines);
    }

    [Fact]
    public void Reads_manual_breaks_as_zero_based_indexes_after_which_to_break()
    {
        var setup = Read(tail: "<rowBreaks count=\"2\"><brk id=\"10\" max=\"16383\" man=\"1\"/><brk id=\"20\" man=\"1\"/></rowBreaks><colBreaks count=\"1\"><brk id=\"3\" man=\"1\"/></colBreaks>").PageSetup;
        Assert.Equal([9, 19], setup.RowBreaks);
        Assert.Equal([2], setup.ColumnBreaks);
    }

    [Fact]
    public void Reads_print_area_and_titles_from_defined_names()
    {
        var setup = Read(definedNames: "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">S!$B$2:$D$20</definedName>" +
            "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">S!$A:$B,S!$1:$2</definedName>").PageSetup;
        Assert.Equal(new CellRange(1, 1, 19, 3), setup.PrintArea);
        Assert.Equal((0, 1), setup.TitleRows);
        Assert.Equal((0, 1), setup.TitleColumns);
    }

    [Fact]
    public void Splits_header_and_footer_sections()
    {
        var setup = Read(tail: "<headerFooter><oddHeader>&amp;L&amp;\"Arial,Bold\"Left&amp;CCenter &amp;A&amp;RRight</oddHeader><oddFooter>Page &amp;P of &amp;N</oddFooter></headerFooter>").PageSetup;
        Assert.Equal(new HeaderFooterParts("&\"Arial,Bold\"Left", "Center &A", "Right"), setup.Header.Odd);
        Assert.Equal(new HeaderFooterParts("", "Page &P of &N", ""), setup.Footer.Odd);
        Assert.Null(setup.Header.First);
        Assert.Null(setup.Header.Even);
    }

    [Fact]
    public void Reads_first_and_even_header_variants_only_when_enabled()
    {
        var setup = Read(tail: "<headerFooter differentOddEven=\"1\" differentFirst=\"1\"><oddHeader>&amp;Codd</oddHeader><evenHeader>&amp;Ceven</evenHeader><firstHeader>&amp;Cfirst</firstHeader><firstFooter/></headerFooter>").PageSetup;
        Assert.Equal("odd", setup.Header.Odd.Center);
        Assert.Equal("even", setup.Header.Even!.Value.Center);
        Assert.Equal("first", setup.Header.First!.Value.Center);
        Assert.True(setup.Footer.First!.Value.IsEmpty);
        var ignored = Read(tail: "<headerFooter><oddHeader>&amp;Codd</oddHeader><evenHeader>&amp;Ceven</evenHeader></headerFooter>").PageSetup;
        Assert.Null(ignored.Header.Even);
    }

    [Fact]
    public void Keeps_ampersand_escapes_and_text_before_a_section_code()
    {
        Assert.Equal(new HeaderFooterParts("", "R&&D", ""), XlsxContextProbe("R&&D"));
    }

    private static HeaderFooterParts XlsxContextProbe(string text) =>
        Read(tail: $"<headerFooter><oddHeader>{text.Replace("&", "&amp;")}</oddHeader></headerFooter>").PageSetup.Header.Odd;
}
