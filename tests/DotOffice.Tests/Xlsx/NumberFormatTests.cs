using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class NumberFormatTests
{
    private static Cell ReadCell(string value, string code, string? type = null, bool date1904 = false, int? builtIn = null, string colors = "")
    {
        var escaped = System.Security.SecurityElement.Escape(code);
        var numFmts = builtIn is null ? $"<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"{escaped}\"/></numFmts>" : "";
        var builder = new XlsxBuilder { Date1904 = date1904 };
        builder.Styles = numFmts + "<fonts><font/></fonts><fills><fill/></fills><borders><border/></borders>" +
            $"<cellXfs><xf numFmtId=\"{builtIn ?? 164}\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>" + colors;
        builder.AddSheet("S", Row(1, C("A1", value, type, 0)));
        return XlsxReader.Read(builder.BuildStream()).Sheets[0].Cells[(0, 0)];
    }

    private static string Fmt(double value, string code, bool date1904 = false) =>
        ReadCell(value.ToString("R", System.Globalization.CultureInfo.InvariantCulture), code, date1904: date1904).Text;

    [Theory]
    [InlineData(1234.5, "#,##0.00", "1,234.50")]
    [InlineData(1234.5, "#,##0", "1,235")]
    [InlineData(0, "#,##0.00", "0.00")]
    [InlineData(-5, "0;(0)", "(5)")]
    [InlineData(5, "0;(0)", "5")]
    [InlineData(-1234.567, "#,##0.0", "-1,234.6")]
    [InlineData(0.25, "0%", "25%")]
    [InlineData(0.256, "0.0%", "25.6%")]
    [InlineData(3.14159, "0.00", "3.14")]
    [InlineData(2.5, "0", "3")]
    [InlineData(0.5, "#.00", ".50")]
    [InlineData(0.5, "0.#", "0.5")]
    [InlineData(5, "0.0#", "5.0")]
    [InlineData(5, "00000", "00005")]
    [InlineData(5, "#,##0.00;[Red]-#,##0.00", "5.00")]
    [InlineData(-5, "#,##0.00;[Red]-#,##0.00", "-5.00")]
    [InlineData(1234567, "#,##0,", "1,235")]
    [InlineData(1234567, "0.0,,\"M\"", "1.2M")]
    [InlineData(12345, "0.00E+00", "1.23E+04")]
    [InlineData(0.00012, "0.00E+00", "1.20E-04")]
    [InlineData(5, "\"$\"#,##0.00", "$5.00")]
    [InlineData(5, "$#,##0", "$5")]
    [InlineData(5, "0 \"units\"", "5 units")]
    [InlineData(5, "0\\x", "5x")]
    [InlineData(5, "0_)", "5 ")]
    [InlineData(5, "0* ;", "5")]
    [InlineData(1234567.891, "#,##0.00", "1,234,567.89")]
    [InlineData(5, "0;-0;\"zero\"", "5")]
    [InlineData(0, "0;-0;\"zero\"", "zero")]
    [InlineData(-5, "0;-0;\"zero\"", "-5")]
    [InlineData(150, "[>100]\"big\";\"small\"", "big")]
    [InlineData(50, "[>100]\"big\";\"small\"", "small")]
    [InlineData(5, "[Blue]0.0", "5.0")]
    [InlineData(1e15, "General", "1E+15")]
    [InlineData(0.30000000000000004, "General", "0.3")]
    [InlineData(123456789012, "General", "1.23457E+11")]
    [InlineData(0.00001234, "General", "1.234E-05")]
    [InlineData(12345.6789, "General", "12345.6789")]
    [InlineData(1234567890, "000-0000", "123456-7890")]
    [InlineData(45000, "yyyy-mm-dd", "2023-03-15")]
    [InlineData(45000, "d mmmm yyyy", "15 March 2023")]
    [InlineData(45000, "ddd, mmm d, yy", "Wed, Mar 15, 23")]
    [InlineData(45000, "dddd", "Wednesday")]
    [InlineData(60, "m/d/yyyy", "2/29/1900")]
    [InlineData(59, "m/d/yyyy", "2/28/1900")]
    [InlineData(61, "m/d/yyyy", "3/1/1900")]
    [InlineData(1, "m/d/yyyy", "1/1/1900")]
    [InlineData(0.5, "h:mm AM/PM", "12:00 PM")]
    [InlineData(0.25, "h:mm AM/PM", "6:00 AM")]
    [InlineData(0.75, "hh:mm:ss", "18:00:00")]
    [InlineData(0.5, "h:mm A/P", "12:00 P")]
    [InlineData(1.5, "[h]:mm", "36:00")]
    [InlineData(1.5, "[mm]", "2160")]
    [InlineData(0.000694444444444, "[ss]", "60")]
    [InlineData(45000.75, "yyyy-mm-dd hh:mm", "2023-03-15 18:00")]
    [InlineData(0.0208333333333, "mm:ss", "30:00")]
    [InlineData(0.5000115740740741, "hh:mm:ss.000", "12:00:01.000")]
    [InlineData(0.99999999, "hh:mm:ss", "00:00:00")]
    [InlineData(45000, "mmm-yy", "Mar-23")]
    [InlineData(45000, "m", "3")]
    [InlineData(45000, "[$-409]d/m/yyyy", "15/3/2023")]
    [InlineData(5, "[$€-407] 0.00", "€ 5.00")]
    [InlineData(-1, "yyyy", "########")]
    public void Formats_numbers(double value, string code, string expected) => Assert.Equal(expected, Fmt(value, code));

    [Fact]
    public void Honors_the_1904_date_system()
    {
        Assert.Equal("1904-01-01", Fmt(0, "yyyy-mm-dd", date1904: true));
        Assert.Equal("2027-03-16", Fmt(45000, "yyyy-mm-dd", date1904: true));
    }

    [Theory]
    [InlineData(14, 45000, "3/15/2023")]
    [InlineData(15, 45000, "15-Mar-23")]
    [InlineData(9, 0.5, "50%")]
    [InlineData(10, 0.256, "25.60%")]
    [InlineData(2, 3, "3.00")]
    [InlineData(3, 1234, "1,234")]
    [InlineData(4, 1234, "1,234.00")]
    [InlineData(20, 0.5, "12:00")]
    [InlineData(21, 0.5, "12:00:00")]
    [InlineData(22, 45000.5, "3/15/2023 12:00")]
    [InlineData(37, -1234, "(1,234)")]
    [InlineData(46, 1.5, "36:00:00")]
    [InlineData(11, 12345, "1.23E+04")]
    [InlineData(999, 12.5, "12.5")]
    public void Formats_built_in_ids(int id, double value, string expected) =>
        Assert.Equal(expected, ReadCell(value.ToString("R", System.Globalization.CultureInfo.InvariantCulture), "", builtIn: id).Text);

    [Fact]
    public void Red_section_colors_the_font()
    {
        var cell = ReadCell("-5", "0;[Red]-0");
        Assert.Equal("-5", cell.Text);
        Assert.Equal(new Color(255, 0, 0), cell.Style.Font.Color);
        Assert.Null(ReadCell("5", "0;[Red]-0").Style.Font.Color);
    }

    [Fact]
    public void Text_section_applies_to_text_cells_only()
    {
        Assert.Equal("[abc]", ReadCell("abc", "0;0;0;\"[\"@\"]\"", "str").Text);
        Assert.Equal("abc", ReadCell("abc", "0.00", "str").Text);
        Assert.Equal("x: abc", ReadCell("abc", "\"x: \"@", "str").Text);
        Assert.Equal("TRUE", ReadCell("1", "0.00", "b").Text);
    }

    [Theory]
    [InlineData(2.5, "# ?/?", "2 1/2")]
    [InlineData(0.5, "# ?/?", " 1/2")]
    [InlineData(3, "# ?/?", "3    ")]
    [InlineData(0, "# ?/?", "0    ")]
    [InlineData(2.333333333, "# ?/?", "2 1/3")]
    [InlineData(2.75, "# ?/?", "2 3/4")]
    [InlineData(1.0 / 3, "# ?/?", " 1/3")]
    [InlineData(0.3333333, "# ?/?", " 1/3")]
    [InlineData(3.14159, "# ??/??", "3 14/99")]
    [InlineData(3.14159, "# ???/???", "3  16/113")]
    [InlineData(0.04, "# ??/??", "  1/25")]
    [InlineData(2.99999, "# ?/?", "3    ")]
    [InlineData(1.9, "# ?/?", "1 8/9")]
    [InlineData(2.5, "?/?", "5/2")]
    [InlineData(0.75, "?/?", "3/4")]
    [InlineData(10.25, "??/??", "41/4 ")]
    [InlineData(0.5, "0 ?/?", "0 1/2")]
    [InlineData(5, "0 ?/?", "5    ")]
    [InlineData(2.4375, "# ?/8", "2 4/8")]
    [InlineData(2.5, "# ?/8", "2 4/8")]
    [InlineData(3, "# ?/8", "3 0/8")]
    [InlineData(0.9999, "# ?/8", "1 0/8")]
    [InlineData(1.5, "?/8", "12/8")]
    [InlineData(-2.5, "# ?/?", "-2 1/2")]
    [InlineData(-0.25, "?/?", "-1/4")]
    [InlineData(2.5, "# ?/? \"lb\"", "2 1/2 lb")]
    [InlineData(2.5, "\"≈\" # ?/?", "≈ 2 1/2")]
    public void Formats_fractions(double value, string code, string expected) => Assert.Equal(expected, Fmt(value, code));

    [Fact]
    public void Fractions_pick_a_section_like_other_number_formats()
    {
        Assert.Equal("2 1/2", Fmt(2.5, "# ?/?;(# ?/?)"));
        Assert.Equal("(2 1/2)", Fmt(-2.5, "# ?/?;(# ?/?)"));
    }

    [Fact]
    public void Slashes_in_dates_and_quotes_are_not_fractions()
    {
        Assert.Equal("25/3/2023", Fmt(45010, "d/m/yyyy"));
        Assert.Equal("3/25", Fmt(45010, "m/d"));
        Assert.Equal("5/", Fmt(5, "0\"/\""));
    }

    [Theory]
    [InlineData("[Color3]0.0", 1.5, 255, 0, 0)]       // Color3 = palette entry 10 = red
    [InlineData("[Color4]0", 1, 0, 255, 0)]           // green
    [InlineData("[Color5]0", 1, 0, 0, 255)]           // blue
    [InlineData("0;[Color10]-0", -1, 0, 128, 0)]      // Color10 = palette entry 17 = dark green
    [InlineData("[color46]0", 1, 255, 102, 0)]        // case-insensitive; palette entry 53
    public void Indexed_colors_use_the_default_palette(string code, double value, int r, int g, int b)
    {
        var cell = ReadCell(value.ToString(System.Globalization.CultureInfo.InvariantCulture), code);
        Assert.Equal(new Color((byte)r, (byte)g, (byte)b), cell.Style.Font.Color);
    }

    [Fact]
    public void Indexed_colors_use_the_workbook_palette_when_it_has_one()
    {
        // Entries 0-9 keep their meaning; entry 10 (Color3) is replaced, entry 11 (Color4) is not listed and stays the default.
        var colors = "<colors><indexedColors>" + string.Concat(Enumerable.Repeat("<rgbColor rgb=\"FF000000\"/>", 10)) + "<rgbColor rgb=\"FF102030\"/></indexedColors></colors>";
        Assert.Equal(new Color(0x10, 0x20, 0x30), ReadCell("1.5", "[Color3]0.0", colors: colors).Style.Font.Color);
        Assert.Equal(new Color(0, 255, 0), ReadCell("1", "[Color4]0", colors: colors).Style.Font.Color);
    }

    [Theory]
    [InlineData("[Color0]0")]
    [InlineData("[Color57]0")]
    [InlineData("[Colorx]0")]
    public void Out_of_range_indexed_colors_are_ignored(string code)
    {
        var cell = ReadCell("1", code);
        Assert.Equal("1", cell.Text);
        Assert.Equal(CellStyle.Default.Font.Color, cell.Style.Font.Color);
    }
}
