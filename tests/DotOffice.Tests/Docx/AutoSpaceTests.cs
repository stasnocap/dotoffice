using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class AutoSpaceTests
{
    private const string Run = "<w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/><w:sz w:val=\"28\"/></w:rPr>";

    // The x of the last text on the line, which starts the Latin text after the East Asian text.
    private static double XAfter(string text, string pPr = "")
    {
        var body = $"<w:p><w:pPr>{pPr}</w:pPr><w:r>{Run}<w:t>{text}</w:t></w:r></w:p>" + Section();
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        var page = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
        return Texts(page).OrderBy(t => t.X).Last().X;
    }

    [Theory]
    [InlineData("日本語ABC", "<w:autoSpaceDE w:val=\"0\"/>", 1)]
    [InlineData("ABC日本語XYZ", "<w:autoSpaceDE w:val=\"0\"/>", 2)]
    [InlineData("日本語123", "<w:autoSpaceDN w:val=\"0\"/>", 1)]
    public void A_quarter_em_separates_East_Asian_text_from_Latin_text_and_digits(string text, string off, int boundaries)
    {
        Assert.Equal(boundaries * 14 / 4.0, XAfter(text) - XAfter(text, off), 3);
    }

    [Fact]
    public void Digits_are_not_spaced_by_autoSpaceDE()
    {
        Assert.Equal(XAfter("日本語123", "<w:autoSpaceDE w:val=\"0\"/>"), XAfter("日本語123") , 3);
    }

    [Fact]
    public void Latin_text_alone_gets_no_space()
    {
        Assert.Equal(XAfter("ABC DEF", "<w:autoSpaceDE w:val=\"0\"/>"), XAfter("ABC DEF"), 3);
    }
}
