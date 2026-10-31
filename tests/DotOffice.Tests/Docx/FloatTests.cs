using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using DotOffice.Tests.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class FloatTests
{
    private static readonly byte[] Png = PngTestEncoder.Encode(2, 2, 2, 8, new int[12]);

    private static Document Read(Func<string, string> body)
    {
        var b = new DocxBuilder();
        var rel = b.AddImage("f.png", Png);
        b.Body = body(rel) + Section();
        return DocxReader.Read(b.BuildStream());
    }

    private static LayoutPage Layout(Document doc) =>
        new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];

    private static string Body(int words) =>
        R(string.Join(' ', Enumerable.Range(1, words).Select(i => $"w{i}")), "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/>");

    [Fact]
    public void Reads_anchor_position_wrap_and_z_order()
    {
        var doc = Read(rel => P(AnchorImage(rel, 100, 50, Position("margin", "right", "paragraph", "127000"), behind: true, distance: 127000)
                                + AnchorImage(rel, 20, 10, Position("page", "254000", "page", "center"), "<wp:wrapTopAndBottom/>")));
        var floats = ((Paragraph)doc.Sections[0].Blocks[0]).Inlines.OfType<FloatingImage>().ToList();
        Assert.Equal(2, floats.Count);
        var (a, b) = (floats[0], floats[1]);
        Assert.Equal((100.0, 50.0, true, TextWrap.None), (a.Width, a.Height, a.BehindText, a.Wrap)); // behind text never wraps
        Assert.Equal(new FloatPosition(FloatReference.Margin, FloatAlignment.End, 0), a.Horizontal);
        Assert.Equal(new FloatPosition(FloatReference.Paragraph, FloatAlignment.Offset, 10), a.Vertical);
        Assert.Equal(new Thickness(10, 10, 10, 10), a.Distance);
        Assert.Equal(new FloatPosition(FloatReference.Page, FloatAlignment.Offset, 20), b.Horizontal);
        Assert.Equal(new FloatPosition(FloatReference.Page, FloatAlignment.Center, 0), b.Vertical);
        Assert.Equal(TextWrap.TopAndBottom, b.Wrap);
    }

    [Fact]
    public void Square_wrap_flows_text_beside_the_image_then_below_it()
    {
        // 150x60pt image at the right margin of the first paragraph, 9pt (114300 EMU) gap to the text.
        var page = Layout(Read(rel => P(AnchorImage(rel, 150, 60, Position("margin", "right", "paragraph", "0"), distance: 114300) + Body(200))));
        var image = Assert.Single(page.Items.OfType<ImageItem>());
        Assert.Equal((72 + 468 - 150, 72.0), (image.X, image.Y));

        var lines = Lines(page);
        var beside = lines.Where(l => l.Baseline < image.Y + image.Height).ToList();
        var below = lines.Where(l => l.Baseline > image.Y + image.Height + 15).ToList();
        Assert.NotEmpty(beside);
        Assert.NotEmpty(below);
        Assert.All(beside, l => Assert.True(Right(l.Items[^1]) <= image.X - 9 + 0.01));
        Assert.Contains(below, l => Right(l.Items[^1]) > image.X); // full width again
    }

    [Fact]
    public void Square_wrap_on_the_left_indents_lines_and_affects_following_paragraphs()
    {
        var page = Layout(Read(rel => P(AnchorImage(rel, 100, 80, Position("column", "0", "paragraph", "0")) + R("short"))
                                      + P(Body(120))));
        var image = page.Items.OfType<ImageItem>().Single();
        var beside = Lines(page).Where(l => l.Baseline < image.Y + image.Height).ToList();
        Assert.True(beside.Count >= 3);
        Assert.All(beside, l => Assert.True(l.Items[0].X >= 172 - 0.01));
    }

    [Fact]
    public void Top_and_bottom_wrap_keeps_lines_clear_of_the_image()
    {
        var page = Layout(Read(rel => P(R("before")) + P(AnchorImage(rel, 50, 100, Position("column", "center", "paragraph", "0"), "<wp:wrapTopAndBottom/>") + Body(80))));
        var image = page.Items.OfType<ImageItem>().Single();
        foreach (var t in Texts(page))
        {
            double top = t.Baseline - t.FontSize, bottom = t.Baseline + t.FontSize * 0.3;
            Assert.False(top < image.Y + image.Height && bottom > image.Y, $"'{t.Text}' overlaps the image");
        }
        Assert.Contains(Texts(page), t => t.Baseline > image.Y + image.Height);
    }

    [Fact]
    public void Behind_and_in_front_images_are_painted_in_z_order_without_moving_text()
    {
        var page = Layout(Read(rel => P(AnchorImage(rel, 50, 50, Position("page", "0", "page", "0"), "<wp:wrapNone/>", behind: true)
                                        + AnchorImage(rel, 40, 40, Position("page", "500000", "page", "500000"), "<wp:wrapNone/>") + R("text"))));
        var items = page.Items;
        int behind = items.FindIndex(i => i is ImageItem { Width: 50 });
        int text = items.FindIndex(i => i is TextItem);
        int front = items.FindIndex(i => i is ImageItem { Width: 40 });
        Assert.True(behind < text && text < front);
        Assert.Equal((0.0, 0.0), (((ImageItem)items[behind]).X, ((ImageItem)items[behind]).Y));
        Assert.Equal(500000 / 12700.0, ((ImageItem)items[front]).X, 3);
        Assert.Equal(72, Texts(page).First().X, 2);
    }
}
