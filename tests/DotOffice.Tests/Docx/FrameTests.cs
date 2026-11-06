using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class FrameTests
{
    private const string Font = "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/><w:sz w:val=\"24\"/>";
    private static readonly string Lorem = string.Join(' ', Enumerable.Repeat("lorem ipsum dolor sit amet", 20));

    private static LayoutPage Layout(string body)
    {
        var doc = DocxReader.Read(new DocxBuilder { Body = body + Section() }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
    }

    private static string Framed(string text, string framePr) => P(R(text, Font), $"<w:framePr {framePr}/>");

    [Fact]
    public void Frame_aligned_right_in_the_margin_sits_at_the_right_edge_and_text_wraps_to_its_left()
    {
        var page = Layout(Framed("FRAMED", "w:w=\"2400\" w:hSpace=\"180\" w:wrap=\"around\" w:hAnchor=\"margin\" w:xAlign=\"right\" w:vAnchor=\"text\" w:y=\"1\"")
            + P(R(Lorem, Font)));
        var frame = Texts(page).First(t => t.Text.StartsWith("FRAMED"));
        Assert.Equal(612 - 72 - 120, frame.X, 0.5);
        var first = Texts(page).Where(t => t.Text.Contains("lorem")).OrderBy(t => t.Baseline).First();
        Assert.Equal(frame.Baseline, first.Baseline, 0.5);
        Assert.All(Texts(page).Where(t => t.Baseline == frame.Baseline && t.X < frame.X), t => Assert.True(Right(t) <= frame.X - 9 + 0.5));
    }

    private static LayoutDocument TwoPages(string align, bool mirror)
    {
        var body = Framed("FRAMED", $"w:w=\"2400\" w:wrap=\"around\" w:hAnchor=\"margin\" w:xAlign=\"{align}\" w:vAnchor=\"text\" w:y=\"1\"")
            + P("<w:r><w:br w:type=\"page\"/></w:r>") + Framed("FRAMED", $"w:w=\"2400\" w:wrap=\"around\" w:hAnchor=\"margin\" w:xAlign=\"{align}\" w:vAnchor=\"text\" w:y=\"1\"")
            + Section();
        var doc = DocxReader.Read(new DocxBuilder { Body = body, Settings = mirror ? "<w:mirrorMargins/>" : null }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc);
    }

    [Theory]
    [InlineData("inside", true, 72, 612 - 72 - 120)]
    [InlineData("outside", true, 612 - 72 - 120, 72)]
    [InlineData("inside", false, 72, 72)]
    [InlineData("outside", false, 612 - 72 - 120, 612 - 72 - 120)]
    public void Inside_and_outside_swap_on_even_pages_with_mirrored_margins(string align, bool mirror, double odd, double even)
    {
        var doc = TwoPages(align, mirror);
        Assert.Equal(odd, Texts(doc.Pages[0]).First(t => t.Text.StartsWith("FRAMED")).X, 0.5);
        Assert.Equal(even, Texts(doc.Pages[1]).First(t => t.Text.StartsWith("FRAMED")).X, 0.5);
    }

    [Fact]
    public void Frame_without_width_is_as_wide_as_its_text_and_consecutive_equal_frames_share_one_box()
    {
        var pr = "w:wrap=\"around\" w:hAnchor=\"margin\" w:xAlign=\"left\"";
        var page = Layout(Framed("one", pr) + Framed("two", pr) + P(R(Lorem, Font)));
        var one = Texts(page).First(t => t.Text == "one");
        var two = Texts(page).First(t => t.Text == "two");
        Assert.Equal(one.X, two.X, 0.01);
        Assert.True(two.Baseline > one.Baseline);
        var body = Texts(page).Where(t => t.Text.Contains("lorem")).OrderBy(t => t.Baseline).First();
        Assert.True(body.X > Right(one));
    }

    [Fact]
    public void Drop_cap_frame_starts_the_paragraph_and_the_text_wraps_beside_it()
    {
        var page = Layout(P(R("D", "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/><w:sz w:val=\"120\"/>"),
                "<w:framePr w:dropCap=\"drop\" w:lines=\"3\" w:wrap=\"around\" w:vAnchor=\"text\" w:hAnchor=\"text\"/>")
            + P(R(Lorem, Font)));
        var cap = Texts(page).First(t => t.Text == "D");
        Assert.Equal(72, cap.X, 0.5);
        var lines = Texts(page).Where(t => t.Text.Trim().Length > 0 && t.Text != "D").GroupBy(t => Math.Round(t.Baseline, 1)).OrderBy(g => g.Key)
            .Select(g => g.Min(t => t.X)).ToList();
        Assert.True(lines[0] >= Right(cap));
        Assert.True(lines.Last() < 80);
    }

    [Fact]
    public void Drop_cap_wraps_exactly_the_number_of_lines_it_spans_even_if_the_letter_is_taller()
    {
        var page = Layout(P(R("D", "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/><w:sz w:val=\"200\"/>"),
                "<w:framePr w:dropCap=\"drop\" w:lines=\"3\" w:wrap=\"around\" w:vAnchor=\"text\" w:hAnchor=\"text\"/>")
            + P(R(Lorem + " " + Lorem, Font)));
        var cap = Texts(page).First(t => t.Text == "D");
        var lines = Texts(page).Where(t => t.Text.Trim().Length > 0 && t.Text != "D").GroupBy(t => Math.Round(t.Baseline, 1)).OrderBy(g => g.Key).ToList();
        Assert.Equal(3, lines.Count(g => g.Min(t => t.X) >= Right(cap)));
        Assert.True(lines[3].Min(t => t.X) < 80);
    }

    [Fact]
    public void Frame_with_exact_height_keeps_that_height_for_wrapping()
    {
        var page = Layout(Framed("tall", "w:w=\"2400\" w:h=\"2880\" w:hRule=\"exact\" w:wrap=\"around\" w:hAnchor=\"margin\" w:xAlign=\"right\"")
            + P(R(Lorem, Font)));
        var beside = Texts(page).Where(t => t.Text.Contains("lorem") && t.X < 300).OrderBy(t => t.Baseline).ToList();
        Assert.True(beside.Count(t => Right(t) < 612 - 72 - 120 + 1) >= 8);
    }

    private static LayoutPage WithStory(bool footer, string framePr)
    {
        var b = new DocxBuilder();
        var story = P(R("RUNNING", Font)) + Framed("FRAMED", framePr);
        var rel = footer ? b.AddFooter(story) : b.AddHeader(story);
        var tag = footer ? "footerReference" : "headerReference";
        b.Body = P(R("body", Font)) + Section(extra: $"<w:{tag} w:type=\"default\" r:id=\"{rel}\"/>");
        var doc = DocxReader.Read(b.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Frames_in_headers_and_footers_sit_at_their_position_and_stay_out_of_the_story(bool footer)
    {
        var page = WithStory(footer, "w:w=\"2400\" w:wrap=\"around\" w:hAnchor=\"margin\" w:xAlign=\"right\" w:vAnchor=\"page\" w:y=\"1440\"");
        var frame = Texts(page).Single(t => t.Text.StartsWith("FRAMED"));
        Assert.Equal(612 - 72 - 120, frame.X, 0.5);
        Assert.InRange(frame.Baseline, 72 + 1, 72 + 25);
        var running = Texts(page).Single(t => t.Text.StartsWith("RUNNING"));
        Assert.Equal(72, running.X, 0.5);
        Assert.NotEqual(frame.Baseline, running.Baseline);
    }
}
