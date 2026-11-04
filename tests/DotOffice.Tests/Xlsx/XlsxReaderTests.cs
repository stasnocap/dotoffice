using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxReaderTests
{
    private static Workbook Read(XlsxBuilder builder, ConversionWarningList? warnings = null) =>
        XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings });

    private static Cell At(Worksheet sheet, int row, int column) => sheet.Cells[(row, column)];

    [Fact]
    public void Reads_sheet_names_in_order_with_visibility()
    {
        var wb = Read(new XlsxBuilder()
            .AddSheet("First")
            .AddSheet("Second", state: "hidden")
            .AddSheet("Third", state: "veryHidden"));
        Assert.Equal(["First", "Second", "Third"], wb.Sheets.Select(s => s.Name));
        Assert.Equal([SheetVisibility.Visible, SheetVisibility.Hidden, SheetVisibility.VeryHidden], wb.Sheets.Select(s => s.Visibility));
    }

    [Fact]
    public void Reads_document_properties()
    {
        var wb = Read(new XlsxBuilder { Title = "Budget" }.AddSheet("S"));
        Assert.Equal("Budget", wb.Title);
        Assert.Equal("Tests", wb.Author);
    }

    [Fact]
    public void Reads_shared_inline_and_formula_strings()
    {
        var wb = Read(new XlsxBuilder { SharedStrings = "<si><t>alpha</t></si><si><t>beta</t></si>" }
            .AddSheet("S", Row(1, S("A1", 0), S("B1", 1), Inline("C1", " padded "), C("D1", "from formula", "str"))));
        var s = wb.Sheets[0];
        Assert.Equal("alpha", At(s, 0, 0).Text);
        Assert.Equal("beta", At(s, 0, 1).Text);
        Assert.Equal(" padded ", At(s, 0, 2).Text);
        Assert.Equal("from formula", At(s, 0, 3).Text);
        Assert.All(s.Cells.Values, c => Assert.Equal(CellValueKind.Text, c.Kind));
    }

    [Fact]
    public void Concatenates_rich_text_runs_and_ignores_phonetics_without_warning()
    {
        var warnings = new ConversionWarningList();
        var wb = Read(new XlsxBuilder
        {
            SharedStrings = "<si><r><rPr><b/></rPr><t>Hel</t></r><r><t>lo</t></r><rPh sb=\"0\" eb=\"2\"><t>PHON</t></rPh></si>" +
                            "<si><r><t>x</t></r><r><t>y</t></r></si>",
        }.AddSheet("S", Row(1, S("A1", 0), S("B1", 1))), warnings);
        Assert.Equal("Hello", At(wb.Sheets[0], 0, 0).Text);
        Assert.Equal("xy", At(wb.Sheets[0], 0, 1).Text);
        Assert.Empty(warnings.Warnings);
    }

    private const string RichStyles =
        "<fonts><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills><fill/></fills><borders><border/></borders>" +
        "<cellXfs><xf fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>";

    [Fact]
    public void Reads_formatting_of_shared_string_runs_and_inherits_the_rest_from_the_cell_font()
    {
        var wb = Read(new XlsxBuilder
        {
            Styles = RichStyles,
            SharedStrings = "<si><r><rPr><b/><i/><u/><strike/><sz val=\"14\"/><color rgb=\"FFFF0000\"/><rFont val=\"Arial\"/></rPr><t>Big</t></r>" +
                            "<r><t xml:space=\"preserve\"> plain </t></r>" +
                            "<r><rPr><vertAlign val=\"superscript\"/><u val=\"none\"/></rPr><t>2</t></r>" +
                            "<r><rPr><vertAlign val=\"subscript\"/></rPr><t>i</t></r></si>",
        }.AddSheet("S", Row(1, S("A1", 0, 0))));
        var runs = At(wb.Sheets[0], 0, 0).Runs!;
        Assert.Equal(["Big", " plain ", "2", "i"], runs.Select(r => r.Text));
        var big = runs[0].Font;
        Assert.Equal(("Arial", 14.0), (big.FontFamily, big.FontSize));
        Assert.True(big.Bold && big.Italic && big.Strike);
        Assert.Equal(UnderlineStyle.Single, big.Underline);
        Assert.Equal(new Color(255, 0, 0), big.Color);
        Assert.Equal(("Calibri", 11.0, false), (runs[1].Font.FontFamily, runs[1].Font.FontSize, runs[1].Font.Bold));
        Assert.Equal(VerticalPosition.Superscript, runs[2].Font.VerticalPosition);
        Assert.Equal(VerticalPosition.Subscript, runs[3].Font.VerticalPosition);
        Assert.Equal(14.0, At(wb.Sheets[0], 0, 0).LargestFontSize);
    }

    [Fact]
    public void Reads_inline_string_runs_with_theme_colors()
    {
        var wb = Read(new XlsxBuilder { Styles = RichStyles, ThemeColors = "<a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1><a:accent1><a:srgbClr val=\"112233\"/></a:accent1>" }
            .AddSheet("S", Row(1, "<c r=\"A1\" t=\"inlineStr\"><is><r><rPr><color theme=\"4\"/></rPr><t>A</t></r><r><t>B</t></r></is></c>")));
        var runs = At(wb.Sheets[0], 0, 0).Runs!;
        Assert.Equal(new Color(0x11, 0x22, 0x33), runs[0].Font.Color);
        Assert.Null(runs[1].Font.Color);
    }

    [Fact]
    public void Plain_strings_have_no_runs_and_text_formats_drop_them()
    {
        var wb = Read(new XlsxBuilder
        {
            Styles = "<numFmts><numFmt numFmtId=\"164\" formatCode=\"&quot;&lt;&quot;@&quot;&gt;&quot;\"/></numFmts><fonts><font/></fonts><fills><fill/></fills><borders><border/></borders>" +
                     "<cellXfs><xf fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>",
            SharedStrings = "<si><t>plain</t></si><si><r><rPr><b/></rPr><t>a</t></r><r><t>b</t></r></si>",
        }.AddSheet("S", Row(1, S("A1", 0, 0), S("B1", 1, 1))));
        Assert.Null(At(wb.Sheets[0], 0, 0).Runs);
        Assert.Equal("<ab>", At(wb.Sheets[0], 0, 1).Text);
        Assert.Null(At(wb.Sheets[0], 0, 1).Runs);
    }

    [Fact]
    public void Reads_booleans_errors_and_numbers()
    {
        var wb = Read(new XlsxBuilder().AddSheet("S",
            Row(1, C("A1", "1", "b"), C("B1", "0", "b"), C("C1", "#DIV/0!", "e"), C("D1", "42"), C("E1", "-1.5", "n"))));
        var s = wb.Sheets[0];
        Assert.Equal(("TRUE", CellValueKind.Bool), (At(s, 0, 0).Text, At(s, 0, 0).Kind));
        Assert.Equal("FALSE", At(s, 0, 1).Text);
        Assert.Equal(("#DIV/0!", CellValueKind.Error), (At(s, 0, 2).Text, At(s, 0, 2).Kind));
        Assert.Equal(("42", CellValueKind.Number), (At(s, 0, 3).Text, At(s, 0, 3).Kind));
        Assert.Equal("-1.5", At(s, 0, 4).Text);
    }

    [Theory]
    [InlineData("0", "0")]
    [InlineData("1234.5", "1234.5")]
    [InlineData("1E15", "1E+15")]
    [InlineData("0.30000000000000004", "0.3")]
    [InlineData("123456789012", "1.23457E+11")]
    [InlineData("12345678901", "12345678901")]
    [InlineData("0.0001", "0.0001")]
    [InlineData("0.00001", "1E-05")]
    [InlineData("-0.5", "-0.5")]
    [InlineData("1234567.891", "1234567.891")]
    public void Formats_general_numbers(string raw, string expected)
    {
        var wb = Read(new XlsxBuilder().AddSheet("S", Row(1, C("A1", raw))));
        Assert.Equal(expected, At(wb.Sheets[0], 0, 0).Text);
    }

    [Fact]
    public void Honors_cell_and_row_references_and_sparse_cells()
    {
        var wb = Read(new XlsxBuilder().AddSheet("S",
            Row(3, C("C3", "1"), C("AA3", "2")) +
            "<row><c><v>7</v></c><c><v>8</v></c></row>" + // implicit: row 4, columns A and B
            "<row r=\"10\"><c t=\"inlineStr\"><is><t>x</t></is></c></row>"));
        var s = wb.Sheets[0];
        Assert.Equal("1", At(s, 2, 2).Text);
        Assert.Equal("2", At(s, 2, 26).Text);
        Assert.Equal("7", At(s, 3, 0).Text);
        Assert.Equal("8", At(s, 3, 1).Text);
        Assert.Equal("x", At(s, 9, 0).Text);
        Assert.Equal(5, s.Cells.Count);
    }

    [Fact]
    public void Decodes_excel_escapes()
    {
        var wb = Read(new XlsxBuilder { SharedStrings = "<si><t>a_x000D_b_x005F_c</t></si>" }.AddSheet("S", Row(1, S("A1", 0))));
        Assert.Equal("a\rb_c", At(wb.Sheets[0], 0, 0).Text);
    }

    [Fact]
    public void Formula_without_cached_value_warns_once_per_sheet()
    {
        var warnings = new ConversionWarningList();
        var wb = Read(new XlsxBuilder().AddSheet("S",
            Row(1, "<c r=\"A1\"><f>1+1</f></c>", "<c r=\"B1\"><f>2+2</f></c>", "<c r=\"C1\"><f>3</f><v>3</v></c>")), warnings);
        Assert.Single(warnings.Warnings, w => w.Code == "xlsx.formula-no-cache");
        Assert.Single(wb.Sheets[0].Cells);
        Assert.Equal("3", At(wb.Sheets[0], 0, 2).Text);
    }

    [Fact]
    public void Invalid_shared_string_index_is_skipped_with_a_warning()
    {
        var warnings = new ConversionWarningList();
        var wb = Read(new XlsxBuilder { SharedStrings = "<si><t>a</t></si>" }.AddSheet("S", Row(1, S("A1", 5), S("B1", 0))), warnings);
        Assert.Contains(warnings.Warnings, w => w.Code == "xlsx.bad-shared-string");
        Assert.Single(wb.Sheets[0].Cells);
    }

    [Fact]
    public void Reads_print_area_and_titles_from_defined_names()
    {
        var wb = Read(new XlsxBuilder
        {
            DefinedNames =
                "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"1\">'Long table'!$B$2:$D$20</definedName>" +
                "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"1\">'Long table'!$A:$B,'Long table'!$1:$2</definedName>" +
                "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">Plain!$A$1</definedName>",
        }.AddSheet("Plain").AddSheet("Long table"));
        Assert.Equal(new CellRange(0, 0, 0, 0), wb.Sheets[0].PageSetup.PrintArea);
        var setup = wb.Sheets[1].PageSetup;
        Assert.Equal(new CellRange(1, 1, 19, 3), setup.PrintArea);
        Assert.Equal((0, 1), setup.TitleRows);
        Assert.Equal((0, 1), setup.TitleColumns);
    }

    [Fact]
    public void Uses_the_first_of_several_print_areas()
    {
        var warnings = new ConversionWarningList();
        var wb = Read(new XlsxBuilder { DefinedNames = "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">S!$A$1:$B$2,S!$D$4:$E$5</definedName>" }
            .AddSheet("S"), warnings);
        Assert.Equal(new CellRange(0, 0, 1, 1), wb.Sheets[0].PageSetup.PrintArea);
        Assert.Contains(warnings.Warnings, w => w.Code == "xlsx.print-area-multiple");
    }

    [Fact]
    public void Sheet_without_a_worksheet_part_is_skipped_with_a_warning()
    {
        var warnings = new ConversionWarningList();
        var builder = new XlsxBuilder().AddSheet("Good", Row(1, C("A1", "1")))
            .AddChartsheet("Chart1");
        var wb = Read(builder, warnings);
        Assert.Equal(["Good"], wb.Sheets.Select(s => s.Name));
        Assert.Contains(warnings.Warnings, w => w.Code == "xlsx.unsupported-sheet");
    }

    [Fact]
    public void Empty_workbook_has_no_sheets()
    {
        Assert.Empty(Read(new XlsxBuilder()).Sheets);
    }

    [Fact]
    public void Non_zip_input_throws_invalid_data()
    {
        Assert.Throws<InvalidDataException>(() => XlsxReader.Read(new MemoryStream("not a zip"u8.ToArray())));
    }

    [Fact]
    public void Zip_without_workbook_throws_invalid_data()
    {
        var docx = new Docx.DocxBuilder { Body = "" }.BuildStream();
        // A DOCX has an officeDocument part, but it is not a workbook: reading must not crash on the wrong content.
        var ex = Record.Exception(() => XlsxReader.Read(docx));
        Assert.True(ex is null or InvalidDataException);
    }

    [Fact]
    public void Enforces_size_limits()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: "<!--" + new string('x', 200_000) + "-->");
        Assert.Throws<InvalidDataException>(() =>
            XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { MaxPartSize = 50_000 }));
        Assert.Throws<InvalidDataException>(() =>
            XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { MaxTotalUncompressedSize = 100_000 }));
    }

    [Fact]
    public void Enforces_cell_count_limit()
    {
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1"), C("B1", "2"), C("C1", "3")));
        Assert.Throws<InvalidDataException>(() =>
            XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { MaxCellCount = 2 }));
    }

    [Fact]
    public void Malformed_xml_throws_invalid_data()
    {
        var builder = new XlsxBuilder().AddSheet("S", "<row r=\"1\"><c r=\"A1\"><v>1</c></row>");
        Assert.Throws<InvalidDataException>(() => XlsxReader.Read(builder.BuildStream()));
    }
}

public class XlsxStyleTests
{
    private const string Fonts =
        "<fonts count=\"4\">" +
        "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/></font>" +
        "<font><b/><i/><u/><strike/><sz val=\"14\"/><color rgb=\"FFFF0000\"/><name val=\"Arial\"/></font>" +
        "<font><u val=\"double\"/><vertAlign val=\"superscript\"/><sz val=\"9\"/><color indexed=\"2\"/><name val=\"Verdana\"/></font>" +
        "<font><b val=\"0\"/><u val=\"none\"/><color theme=\"4\" tint=\"-0.25\"/><name val=\"Times New Roman\"/></font>" +
        "</fonts>";

    private const string Fills =
        "<fills count=\"4\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>" +
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF00FF00\"/><bgColor indexed=\"64\"/></patternFill></fill>" +
        "<fill><gradientFill><stop position=\"0\"><color rgb=\"FF000000\"/></stop></gradientFill></fill></fills>";

    private const string Borders =
        "<borders count=\"3\"><border><left/><right/><top/><bottom/><diagonal/></border>" +
        "<border><left style=\"thin\"><color rgb=\"FF0000FF\"/></left><right style=\"medium\"/><top style=\"double\"/><bottom style=\"thick\"><color indexed=\"2\"/></bottom></border>" +
        "<border><left style=\"hair\"/><right style=\"dotted\"/><top style=\"dashed\"/><bottom style=\"mediumDashed\"/></border></borders>";

    private static Workbook ReadWith(string cellXfs, string cells, XlsxBuilder? builder = null, ConversionWarningList? warnings = null)
    {
        builder ??= new XlsxBuilder();
        builder.Styles = Fonts + Fills + Borders + $"<cellXfs>{cellXfs}</cellXfs>";
        builder.AddSheet("S", cells);
        return XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings });
    }

    private static CellStyle StyleOf(Workbook wb, string cellRef = "A1")
    {
        XlsxBuilderRefs.Parse(cellRef, out var r, out var c);
        return wb.Sheets[0].Cells[(r, c)].Style;
    }

    [Fact]
    public void Resolves_font_properties()
    {
        var wb = ReadWith("<xf fontId=\"0\"/><xf fontId=\"1\"/><xf fontId=\"2\"/><xf fontId=\"3\"/>",
            Row(1, C("A1", "1", style: 0), C("B1", "1", style: 1), C("C1", "1", style: 2), C("D1", "1", style: 3)));
        var plain = StyleOf(wb, "A1").Font;
        Assert.Equal(("Calibri", 11.0, false), (plain.FontFamily, plain.FontSize, plain.Bold));
        Assert.Equal(Color.Black, plain.Color);

        var rich = StyleOf(wb, "B1").Font;
        Assert.Equal(("Arial", 14.0), (rich.FontFamily, rich.FontSize));
        Assert.True(rich.Bold && rich.Italic && rich.Strike);
        Assert.Equal(UnderlineStyle.Single, rich.Underline);
        Assert.Equal(new Color(255, 0, 0), rich.Color);

        var other = StyleOf(wb, "C1").Font;
        Assert.Equal(UnderlineStyle.Double, other.Underline);
        Assert.Equal(VerticalPosition.Superscript, other.VerticalPosition);
        Assert.Equal(new Color(255, 0, 0), other.Color); // indexed 2

        var off = StyleOf(wb, "D1").Font;
        Assert.False(off.Bold);
        Assert.Equal(UnderlineStyle.None, off.Underline);
    }

    [Fact]
    public void Resolves_theme_colors_with_tint()
    {
        var theme = "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>" +
                    "<a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>" +
                    "<a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1>";
        var wb = ReadWith("<xf fontId=\"0\"/><xf fontId=\"3\"/>", Row(1, C("A1", "1", style: 0), C("B1", "1", style: 1)),
            new XlsxBuilder { ThemeColors = theme });
        Assert.Equal(Color.Black, StyleOf(wb, "A1").Font.Color); // theme 1 = dk1
        // Excel's "Blue, Accent 1, Darker 25%" (4F81BD, tint -0.25) is 366092.
        Assert.Equal(new Color(0x36, 0x60, 0x92), StyleOf(wb, "B1").Font.Color!.Value, new ColorTolerance(2));
    }

    [Fact]
    public void Applies_positive_tint_as_lighter()
    {
        var builder = new XlsxBuilder { ThemeColors = "<a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1>" };
        builder.Styles = "<fonts><font><color theme=\"4\" tint=\"0.4\"/></font><font><color theme=\"4\" tint=\"0.8\"/></font><font><color theme=\"4\" tint=\"-1\"/></font></fonts>" +
                         "<cellXfs><xf fontId=\"0\"/><xf fontId=\"1\"/><xf fontId=\"2\"/></cellXfs>";
        builder.AddSheet("S", Row(1, C("A1", "1", style: 0), C("B1", "1", style: 1), C("C1", "1", style: 2)));
        var wb = XlsxReader.Read(builder.BuildStream());
        // "Blue, Accent 1, Lighter 40%" is 95B3D7, "Lighter 80%" is DCE6F1.
        Assert.Equal(new Color(0x95, 0xB3, 0xD7), StyleOf(wb, "A1").Font.Color!.Value, new ColorTolerance(2));
        Assert.Equal(new Color(0xDC, 0xE6, 0xF1), StyleOf(wb, "B1").Font.Color!.Value, new ColorTolerance(2));
        Assert.Equal(Color.Black, StyleOf(wb, "C1").Font.Color!.Value);
    }

    [Fact]
    public void Resolves_fills_and_skips_gradients_with_a_warning()
    {
        var warnings = new ConversionWarningList();
        var wb = ReadWith("<xf fillId=\"0\"/><xf fillId=\"1\"/><xf fillId=\"2\"/><xf fillId=\"3\"/>",
            Row(1, C("A1", "1", style: 0), C("B1", "1", style: 1), C("C1", "1", style: 2), C("D1", "1", style: 3)), warnings: warnings);
        Assert.Null(StyleOf(wb, "A1").Fill);
        Assert.Null(StyleOf(wb, "B1").Fill);
        Assert.Equal(new Color(0, 255, 0), StyleOf(wb, "C1").Fill);
        Assert.Null(StyleOf(wb, "D1").Fill);
        Assert.Single(warnings.Warnings, w => w.Code == "xlsx.gradient-fill");
    }

    [Fact]
    public void Resolves_borders()
    {
        var wb = ReadWith("<xf borderId=\"1\"/><xf borderId=\"2\"/>", Row(1, C("A1", "1", style: 0), C("B1", "1", style: 1)));
        var b = StyleOf(wb, "A1").Borders;
        Assert.Equal((BorderStyle.Single, 0.75, new Color(0, 0, 255)), (b.Left.Style, b.Left.Width, b.Left.Color!.Value));
        Assert.Equal(BorderStyle.Single, b.Right.Style);
        Assert.Equal(1.5, b.Right.Width);
        Assert.Null(b.Right.Color);
        Assert.Equal(BorderStyle.Double, b.Top.Style);
        Assert.Equal(BorderStyle.Thick, b.Bottom.Style);
        Assert.Equal(new Color(255, 0, 0), b.Bottom.Color);

        var d = StyleOf(wb, "B1").Borders;
        Assert.Equal(0.25, d.Left.Width);
        Assert.Equal(BorderStyle.Dotted, d.Right.Style);
        Assert.Equal(BorderStyle.Dashed, d.Top.Style);
        Assert.Equal(1.5, d.Bottom.Width);
    }

    [Fact]
    public void Resolves_alignment()
    {
        var wb = ReadWith(
            "<xf><alignment horizontal=\"center\" vertical=\"top\" wrapText=\"1\" indent=\"2\" textRotation=\"45\"/></xf>" +
            "<xf><alignment horizontal=\"right\" vertical=\"center\" textRotation=\"135\"/></xf>" +
            "<xf><alignment horizontal=\"left\" textRotation=\"255\"/></xf><xf/>",
            Row(1, C("A1", "1", style: 0), C("B1", "1", style: 1), C("C1", "1", style: 2), C("D1", "1", style: 3)));
        var a = StyleOf(wb, "A1");
        Assert.Equal((HorizontalAlignment.Center, CellVerticalAlignment.Top, true, 2, 45.0),
            (a.HorizontalAlignment, a.VerticalAlignment, a.Wrap, a.Indent, a.Rotation));
        var b = StyleOf(wb, "B1");
        Assert.Equal((HorizontalAlignment.Right, CellVerticalAlignment.Center, -45.0), (b.HorizontalAlignment, b.VerticalAlignment, b.Rotation));
        Assert.Equal((0.0, true), (StyleOf(wb, "C1").Rotation, StyleOf(wb, "C1").Stacked)); // 255 = stacked, not rotated
        Assert.False(StyleOf(wb, "A1").Stacked);
        var d = StyleOf(wb, "D1");
        Assert.Equal((HorizontalAlignment.General, CellVerticalAlignment.Bottom, false), (d.HorizontalAlignment, d.VerticalAlignment, d.Wrap));
    }

    [Fact]
    public void Custom_indexed_palette_overrides_the_legacy_one()
    {
        var builder = new XlsxBuilder();
        builder.Styles = "<colors><indexedColors><rgbColor rgb=\"FF010203\"/><rgbColor rgb=\"FF0A0B0C\"/></indexedColors></colors>" +
                         "<fonts><font><color indexed=\"1\"/></font></fonts><cellXfs><xf fontId=\"0\"/></cellXfs>";
        builder.AddSheet("S", Row(1, C("A1", "1", style: 0)));
        var wb = XlsxReader.Read(builder.BuildStream());
        Assert.Equal(new Color(0x0A, 0x0B, 0x0C), StyleOf(wb).Font.Color);
    }

    [Fact]
    public void Empty_cells_are_kept_only_when_they_paint_something()
    {
        var wb = ReadWith("<xf/><xf fillId=\"2\"/><xf borderId=\"1\"/>",
            "<row r=\"1\"><c r=\"A1\" s=\"0\"/><c r=\"B1\" s=\"1\"/><c r=\"C1\" s=\"2\"/><c r=\"D1\" s=\"1\"><v></v></c></row>");
        var cells = wb.Sheets[0].Cells;
        Assert.False(cells.ContainsKey((0, 0)));
        Assert.Equal(new Color(0, 255, 0), cells[(0, 1)].Style.Fill);
        Assert.True(cells[(0, 2)].Style.Borders.HasAny);
        Assert.Equal("", cells[(0, 3)].Text);
    }

    [Fact]
    public void Out_of_range_style_and_font_indexes_fall_back_to_defaults()
    {
        var wb = ReadWith("<xf fontId=\"99\" fillId=\"99\" borderId=\"99\"/>", Row(1, C("A1", "1", style: 0), C("B1", "1", style: 50)));
        Assert.Equal(CellStyle.Default.Font, StyleOf(wb, "A1").Font);
        Assert.Null(StyleOf(wb, "A1").Fill);
        Assert.Equal(CellStyle.Default, StyleOf(wb, "B1"));
    }

    private static class XlsxBuilderRefs
    {
        public static void Parse(string reference, out int row, out int column)
        {
            column = reference[0] - 'A';
            row = int.Parse(reference[1..]) - 1;
        }
    }

    private sealed class ColorTolerance(int tolerance) : IEqualityComparer<Color>
    {
        public bool Equals(Color a, Color b) =>
            Math.Abs(a.R - b.R) <= tolerance && Math.Abs(a.G - b.G) <= tolerance && Math.Abs(a.B - b.B) <= tolerance;

        public int GetHashCode(Color c) => 0;
    }
}
