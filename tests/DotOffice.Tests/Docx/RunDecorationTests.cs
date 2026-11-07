using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class RunDecorationTests
{
    private const string Font = "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/>";

    private static LayoutPage Layout(string rPr, string text = "boxed")
    {
        var body = P(R("plain ", Font) + R(text, Font + rPr) + R(" after", Font)) + Section();
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
    }

    [Fact]
    public void Run_border_draws_a_box_around_the_run_padded_by_its_space()
    {
        var page = Layout("<w:bdr w:val=\"single\" w:sz=\"8\" w:space=\"2\" w:color=\"FF0000\"/>");
        var lines = page.Items.OfType<LineItem>().Where(l => l.Color == new Color(255, 0, 0)).ToList();
        Assert.Equal(4, lines.Count);
        var text = Texts(page).First(t => t.Text == "boxed");
        Assert.Equal(text.X - 2 - 0.5, lines.Min(l => Math.Min(l.X1, l.X2)), 0.01);
        Assert.True(lines.Max(l => Math.Max(l.X1, l.X2)) > text.X + 20);
    }

    [Fact]
    public void Run_border_pushes_the_text_by_its_space_and_line_on_each_side()
    {
        var bare = Texts(Layout("")).ToList();
        var page = Layout("<w:bdr w:val=\"single\" w:sz=\"8\" w:space=\"4\" w:color=\"FF0000\"/>");
        var boxed = Texts(page).ToList();
        var bareRun = bare.First(t => t.Text == "boxed");
        var run = boxed.First(t => t.Text == "boxed");
        Assert.Equal(bareRun.X + 5, run.X, 0.01);
        var bareAfter = bare.Last();
        var after = boxed.Last();
        Assert.Equal(bareAfter.X + 10, after.X, 0.01);
    }

    [Fact]
    public void Emphasis_dot_marks_every_character_but_not_spaces()
    {
        var page = Layout("<w:em w:val=\"dot\"/>", "a b c");
        Assert.Equal(3, page.Items.OfType<EllipseItem>().Count());
    }

    [Fact]
    public void UnderDot_marks_sit_below_the_baseline_and_dot_marks_above_it()
    {
        var under = Layout("<w:em w:val=\"underDot\"/>").Items.OfType<EllipseItem>().ToList();
        var over = Layout("<w:em w:val=\"dot\"/>").Items.OfType<EllipseItem>().ToList();
        double baseline = Texts(Layout("")).First(t => t.Text == "plain").Baseline;
        Assert.Equal(5, under.Count);
        Assert.All(under, e => Assert.True(e.Y > baseline));
        Assert.All(over, e => Assert.True(e.Y < baseline - 10));
    }

    [Fact]
    public void Run_without_border_or_emphasis_draws_neither()
    {
        var page = Layout("");
        Assert.Empty(page.Items.OfType<EllipseItem>());
        Assert.Empty(page.Items.OfType<LineItem>());
    }
}
