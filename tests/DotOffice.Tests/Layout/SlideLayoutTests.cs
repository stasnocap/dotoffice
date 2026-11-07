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

    private static List<LayoutItem> Lay(params SheetShape[] shapes)
    {
        var slide = new Slide(400, 300);
        slide.Objects.AddRange(shapes);
        var presentation = new Presentation();
        presentation.Slides.Add(slide);
        var engine = new SheetLayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(LayoutTestHelpers.Noto)) });
        return engine.Layout(presentation).Pages[0].Items;
    }

    /// <summary>The top and bottom of each filled rectangle or polygon, in drawing order.</summary>
    private static List<(double Top, double Bottom)> Fills(List<LayoutItem> items) =>
    [
        .. items.Select(i => i switch
        {
            RectItem r => ((double, double)?)(r.Y, r.Y + r.Height),
            PolygonItem p => (p.Points.Min(q => q.Y), p.Points.Max(q => q.Y)),
            _ => null,
        }).OfType<(double, double)>(),
    ];

    private static SheetShape Box(double y, double height, params ShapeParagraph[] paragraphs) => new()
    {
        Anchor = ImageAnchorKind.Absolute, X = 10, Y = y, Width = 100, Height = height, Fill = Color.Black,
        Geometry = ShapeGeometry.Rectangle, TextInsets = new Thickness(0, 0, 0, 0), Paragraphs = [.. paragraphs],
    };

    [Fact]
    public void A_shape_that_fits_its_text_grows_down_from_its_top_and_never_shrinks()
    {
        var lines = Enumerable.Range(0, 6).Select(i => Paragraph($"line {i}")).ToArray();
        var grown = Assert.Single(Fills(Lay(Box(20, 10, lines) with { GrowToFit = true })));
        Assert.Equal(20, grown.Top, 3);
        Assert.True(grown.Bottom - 20 > 60);
        Assert.Equal(220, Assert.Single(Fills(Lay(Box(20, 200, lines) with { GrowToFit = true }))).Bottom, 3);
        Assert.Equal(30, Assert.Single(Fills(Lay(Box(20, 10, lines)))).Bottom, 3);
    }

    [Fact]
    public void A_table_row_grows_to_fit_its_cell_text_and_moves_the_rows_below_down()
    {
        double[] heights = [10, 20];
        var tall = Enumerable.Range(0, 4).Select(i => Paragraph($"line {i}")).ToArray();
        var first = Box(0, 10, tall) with { TableRows = new TableRowSpan(0, 0, 1, heights) };
        var second = Box(10, 20, Paragraph("x")) with { TableRows = new TableRowSpan(0, 1, 1, heights), Fill = Color.White };
        var rule = new SheetShape
        {
            Anchor = ImageAnchorKind.Absolute, X = 10, Y = 10, Width = 100, Height = 0, Geometry = ShapeGeometry.Line, LineColor = Color.Black,
            TableRows = new TableRowSpan(0, 1, 0, heights),
        };
        var items = Lay(first, second, rule);
        var fills = Fills(items);
        double firstBottom = fills[0].Bottom;
        Assert.True(firstBottom > 40);
        Assert.Equal(firstBottom, fills[1].Top, 3);
        Assert.Equal(firstBottom + 20, fills[1].Bottom, 3);
        var line = items.OfType<LineItem>().Single();
        Assert.Equal(firstBottom, line.Y1, 3);
    }

    [Fact]
    public void Vertical_text_is_laid_out_across_the_height_and_turned_about_the_text_area()
    {
        var shape = Box(20, 200, Paragraph("one two three four five six seven")) with
        {
            TextRotation = -90, TextInsets = new Thickness(10, 0, 0, 0),
        };
        var items = Lay(shape);
        int push = items.FindIndex(i => i is PushItem { Rotation: -90 });
        Assert.True(push >= 0);
        var turn = (PushItem)items[push];
        Assert.Equal((60, 125), (turn.CenterX, turn.CenterY));
        // Unturned, the text runs over the area's height (190 points) from its left edge, so it fits in one line here.
        var texts = items.Skip(push).OfType<TextItem>().ToList();
        Assert.Single(texts.Select(t => t.Baseline).Distinct());
        Assert.Equal(60 - 95, texts.Min(t => t.X), 3);
        Assert.True(texts.Max(t => t.X) > 60 - 95 + 100); // wider than the frame, so not wrapped at its width
        Assert.True(texts.Max(t => t.X) < 60 + 95);
    }
}
