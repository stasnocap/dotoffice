using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class LineNumberTests
{
    private const string Run = "<w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/></w:rPr>";

    private static string Para(string text, string pPr = "") => $"<w:p><w:pPr>{pPr}</w:pPr><w:r>{Run}<w:t>{text}</w:t></w:r></w:p>";

    private static LayoutDocument Layout(string body, string lnNumType)
    {
        var doc = DocxReader.Read(new DocxBuilder { Body = body + Section(extra: lnNumType) }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc);
    }

    private static List<string> Numbers(LayoutPage page) =>
        Texts(page).Where(t => t.Text.All(char.IsDigit) && t.X < 72).OrderBy(t => t.Baseline).Select(t => t.Text).ToList();

    [Fact]
    public void Every_line_is_numbered_left_of_the_text_with_the_distance_between()
    {
        var page = Layout(Para("one") + Para("two") + Para("three"), "<w:lnNumType w:countBy=\"1\" w:distance=\"360\"/>").Pages[0];
        Assert.Equal(["1", "2", "3"], Numbers(page));
        var number = Texts(page).First(t => t.Text == "1");
        Assert.Equal(72 - 18, Right(number), 0.01);
        Assert.Equal(Texts(page).First(t => t.Text == "one").Baseline, number.Baseline, 0.01);
    }

    [Fact]
    public void CountBy_shows_every_nth_number_and_start_sets_the_first()
    {
        var body = string.Concat(Enumerable.Range(0, 6).Select(i => Para("line" + i)));
        Assert.Equal(["5"], Numbers(Layout(body, "<w:lnNumType w:countBy=\"5\"/>").Pages[0]));
        Assert.Equal(["10", "12"], Numbers(Layout(body, "<w:lnNumType w:countBy=\"2\" w:start=\"10\"/>").Pages[0]).Where(n => n is "10" or "12").ToList());
    }

    [Fact]
    public void Wrapped_lines_each_get_a_number_and_suppressed_paragraphs_get_none()
    {
        string long_ = string.Join(' ', Enumerable.Repeat("alpha beta gamma delta", 30));
        var page = Layout(Para(long_) + Para("skipped", "<w:suppressLineNumbers/>") + Para("last"), "<w:lnNumType/>").Pages[0];
        var numbers = Numbers(page);
        Assert.True(numbers.Count >= 3);
        Assert.Equal(Enumerable.Range(1, numbers.Count).Select(i => i.ToString()), numbers);
        double skippedBaseline = Texts(page).First(t => t.Text == "skipped").Baseline;
        Assert.DoesNotContain(Texts(page), t => t.Text.All(char.IsDigit) && t.X < 72 && Math.Abs(t.Baseline - skippedBaseline) < 0.01);
        // The suppressed paragraph is not counted: "last" has the number after the last line of the long paragraph.
        double lastBaseline = Texts(page).First(t => t.Text == "last").Baseline;
        Assert.Contains(Texts(page), t => t.Text == numbers[^1] && Math.Abs(t.Baseline - lastBaseline) < 0.01);
    }

    [Fact]
    public void NewPage_restarts_each_page_and_Continuous_keeps_counting()
    {
        var body = string.Concat(Enumerable.Range(0, 120).Select(i => Para("line" + i)));
        var restart = Layout(body, "<w:lnNumType w:restart=\"newPage\"/>");
        Assert.True(restart.Pages.Count >= 2);
        Assert.Equal("1", Numbers(restart.Pages[1])[0]);
        var continuous = Layout(body, "<w:lnNumType w:restart=\"continuous\"/>");
        Assert.NotEqual("1", Numbers(continuous.Pages[1])[0]);
        Assert.Equal(int.Parse(Numbers(continuous.Pages[0])[^1]) + 1, int.Parse(Numbers(continuous.Pages[1])[0]));
    }

    [Fact]
    public void Without_lnNumType_there_are_no_numbers()
    {
        Assert.Empty(Numbers(Layout(Para("one"), "").Pages[0]));
    }
}
