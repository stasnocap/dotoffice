using System.Text;
using DotOffice.Converter;
using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Tests.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests;

public class XlsxConverterTests
{
    private static readonly FontCollection TestFonts = new FontCollection().AddFile(TestAssets.Path("NotoSans-Regular.ttf"));

    private static async Task<byte[]> Convert(XlsxBuilder xlsx, ConversionWarningList? warnings = null)
    {
        using var output = new MemoryStream();
        await DocumentConverter.ConvertAsync(xlsx.BuildStream(), output, DocumentFormat.Xlsx, DocumentFormat.Pdf,
            new ConvertOptions { Fonts = TestFonts, UseSystemFonts = false, Warnings = warnings });
        return output.ToArray();
    }

    private static int PageCount(byte[] pdf) =>
        System.Text.RegularExpressions.Regex.Matches(Encoding.Latin1.GetString(pdf), @"/Type\s*/Page\b").Count;

    [Fact]
    public async Task Converts_a_workbook_to_a_pdf_with_one_page_per_sheet()
    {
        var builder = new XlsxBuilder { SharedStrings = "<si><t>Hello</t></si>" }
            .AddSheet("One", Row(1, S("A1", 0), C("B1", "1234.5")))
            .AddSheet("Two", Row(1, C("A1", "7")));
        var pdf = await Convert(builder);
        Assert.StartsWith("%PDF-", Encoding.Latin1.GetString(pdf, 0, 5));
        Assert.Equal(2, PageCount(pdf));
    }

    [Fact]
    public async Task Surfaces_reader_warnings()
    {
        var warnings = new ConversionWarningList();
        await Convert(new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: "<conditionalFormatting sqref=\"A1\"><cfRule type=\"timePeriod\" timePeriod=\"today\" priority=\"1\"/></conditionalFormatting>"), warnings);
        Assert.Contains(warnings.Warnings, w => w.Code == "xlsx.conditional-formatting");
    }

    [Fact]
    public async Task Formula_without_a_cached_value_warns()
    {
        var warnings = new ConversionWarningList();
        await Convert(new XlsxBuilder().AddSheet("S", "<row r=\"1\"><c r=\"A1\"><f>1+1</f></c></row>"), warnings);
        Assert.Contains(warnings.Warnings, w => w.Code == "xlsx.formula-no-cache");
    }

    [Fact]
    public async Task An_empty_workbook_still_produces_a_valid_pdf()
    {
        var pdf = await Convert(new XlsxBuilder().AddSheet("Empty"));
        Assert.Equal(1, PageCount(pdf));
        Assert.Contains("%%EOF", Encoding.Latin1.GetString(pdf));
    }

    [Fact]
    public async Task Only_hidden_sheets_still_produce_a_valid_pdf()
    {
        var pdf = await Convert(new XlsxBuilder().AddSheet("H", Row(1, C("A1", "1")), state: "hidden"));
        Assert.Equal(1, PageCount(pdf));
    }

    [Fact]
    public async Task A_non_zip_input_throws_invalid_data()
    {
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => DocumentConverter.ConvertAsync(
            new MemoryStream([1, 2, 3, 4]), output, DocumentFormat.Xlsx, DocumentFormat.Pdf,
            new ConvertOptions { Fonts = TestFonts, UseSystemFonts = false }));
    }

    [Fact]
    public async Task Converts_files_by_extension_and_leaves_no_output_on_failure()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"dotoffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        try
        {
            var input = Path.Combine(dir, "book.xlsx");
            var output = Path.Combine(dir, "book.pdf");
            await File.WriteAllBytesAsync(input, new XlsxBuilder().AddSheet("S", Row(1, C("A1", "3.5"))).Build());
            await DocumentConverter.ConvertAsync(input, output, new ConvertOptions { Fonts = TestFonts, UseSystemFonts = false });
            Assert.True(new FileInfo(output).Length > 100);

            var bad = Path.Combine(dir, "bad.xlsx");
            var badOut = Path.Combine(dir, "bad.pdf");
            await File.WriteAllBytesAsync(bad, [1, 2, 3]);
            await Assert.ThrowsAsync<InvalidDataException>(() => DocumentConverter.ConvertAsync(bad, badOut, new ConvertOptions { Fonts = TestFonts, UseSystemFonts = false }));
            Assert.False(File.Exists(badOut));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact]
    public async Task Gradient_fills_become_pdf_shadings_without_warnings()
    {
        var warnings = new ConversionWarningList();
        var builder = new XlsxBuilder
        {
            Styles = "<fonts><font/></fonts><fills><fill><patternFill patternType=\"none\"/></fill>" +
                     "<fill><gradientFill degree=\"45\"><stop position=\"0\"><color rgb=\"FFFF0000\"/></stop><stop position=\"1\"><color rgb=\"FF0000FF\"/></stop></gradientFill></fill></fills>" +
                     "<borders><border/></borders><cellXfs><xf fontId=\"0\" fillId=\"1\" borderId=\"0\"/></cellXfs>",
        }.AddSheet("S", Row(1, C("A1", "1", style: 0)));
        var text = Encoding.Latin1.GetString(await Convert(builder, warnings));
        Assert.Contains("/ShadingType 2", text);
        Assert.Contains("/Shading << /Sh1 ", text);
        Assert.DoesNotContain(warnings.Warnings, w => w.Code.StartsWith("xlsx."));
    }

    [Fact]
    public async Task Pictures_are_embedded_and_only_charts_and_shapes_warn()
    {
        var warnings = new ConversionWarningList();
        var png = PngTestEncoder.Encode(2, 2, 2, 8, [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]);
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: DrawingElement)
            .AddDrawing(1, OneCellAnchor(1, 1, 60, 40, Pic()) + OneCellAnchor(3, 3, 60, 40, Shape), png);
        var text = Encoding.Latin1.GetString(await Convert(builder, warnings));
        Assert.Contains("/Subtype /Image", text);
        var warning = Assert.Single(warnings.Warnings, w => w.Code.StartsWith("xlsx."));
        Assert.Equal("xlsx.drawing", warning.Code);
        Assert.Contains("shape", warning.Message);
    }

    [Fact]
    public async Task Conditional_formatting_icons_and_bars_reach_the_pdf_without_warnings()
    {
        var warnings = new ConversionWarningList();
        var rows = Row(1, C("A1", "1")) + Row(2, C("A2", "5")) + Row(3, C("A3", "9"));
        var tail = ConditionalFormatting("A1:A3", CfRule("iconSet", 1, "", "<iconSet iconSet=\"3Symbols\"><cfvo type=\"percent\" val=\"0\"/><cfvo type=\"percent\" val=\"33\"/><cfvo type=\"percent\" val=\"67\"/></iconSet>")) +
                   ConditionalFormatting("B1:B3", Expression("TRUE", 2, 0));
        var pdf = await Convert(new XlsxBuilder { Styles = StylesWithDxfs(Dxf("FF0000")) }.AddSheet("S", rows, tail: tail), warnings);
        Assert.DoesNotContain(warnings.Warnings, w => w.Code.StartsWith("xlsx.", StringComparison.Ordinal));
        var text = Encoding.Latin1.GetString(pdf);
        Assert.StartsWith("%PDF-", text);
        Assert.Contains("%%EOF", text);
    }
}
