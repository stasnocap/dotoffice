using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Layout;

namespace DotOffice.Tests.Layout;

public class SlideLayoutTests
{
    private static readonly RunProperties Font = RunProperties.Default with { FontFamily = "Noto Sans", FontSize = 10 };

    private static List<TextItem> Texts(double width, params ShapeParagraph[] paragraphs)
    {
        var slide = new Slide(400, 300);
        slide.Objects.Add(new SheetShape
        {
            Anchor = ImageAnchorKind.Absolute, X = 10, Y = 10, Width = width, Height = 200,
            Geometry = ShapeGeometry.Rectangle, TextInsets = new Thickness(0, 0, 0, 0), Paragraphs = [.. paragraphs],
        });
        var presentation = new Presentation();
        presentation.Slides.Add(slide);
        var engine = new SheetLayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(LayoutTestHelpers.Noto)) });
        return [.. engine.Layout(presentation).Pages[0].Items.OfType<TextItem>()];
    }

    private static ShapeParagraph Paragraph(string text, double indent = 0, double? margin = null) =>
        new(HorizontalAlignment.Left, [new CellTextRun(text, Font)], 10, indent, margin);

    [Fact]
    public void A_hanging_indent_puts_the_bullet_at_the_indent_and_every_line_of_text_at_the_margin()
    {
        var texts = Texts(120, Paragraph("•\tone two three four five six seven eight nine ten", 0, 30));
        Assert.Equal(10, texts.First(t => t.Text.Contains('•')).X, 3);
        var words = texts.Where(t => !t.Text.Contains('•') && t.Text.Trim().Length > 0).ToList();
        var lineStarts = words.GroupBy(t => t.Baseline).Select(g => g.Min(t => t.X)).ToList();
        Assert.True(lineStarts.Count > 1);
        Assert.All(lineStarts, x => Assert.Equal(40, x, 3));
    }

    [Fact]
    public void Paragraph_spacing_and_line_spacing_move_the_following_lines_down()
    {
        double Gap(ShapeParagraph second)
        {
            var baselines = Texts(300, Paragraph("first"), second).Select(t => t.Baseline).Distinct().Order().ToList();
            return baselines[1] - baselines[0];
        }
        double single = Gap(Paragraph("second"));
        Assert.Equal(2 * single + 6, Gap(Paragraph("second") with { SpaceBefore = 6, LineSpacing = 2 }), 3);
        Assert.Equal(40, Gap(Paragraph("second") with { LinePitch = 40 }), 3);
        Assert.Equal(single, Gap(Paragraph("second") with { SpaceAfter = 50 }), 3);
    }
}
