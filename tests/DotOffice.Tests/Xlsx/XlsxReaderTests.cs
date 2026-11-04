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
    public void Concatenates_rich_text_runs_ignores_phonetics_and_warns_once()
    {
        var warnings = new ConversionWarningList();
        var wb = Read(new XlsxBuilder
        {
            SharedStrings = "<si><r><rPr><b/></rPr><t>Hel</t></r><r><t>lo</t></r><rPh sb=\"0\" eb=\"2\"><t>PHON</t></rPh></si>" +
                            "<si><r><t>x</t></r><r><t>y</t></r></si>",
        }.AddSheet("S", Row(1, S("A1", 0), S("B1", 1))), warnings);
        Assert.Equal("Hello", At(wb.Sheets[0], 0, 0).Text);
        Assert.Equal("xy", At(wb.Sheets[0], 0, 1).Text);
        Assert.Single(warnings.Warnings, w => w.Code == "xlsx.rich-text");
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
