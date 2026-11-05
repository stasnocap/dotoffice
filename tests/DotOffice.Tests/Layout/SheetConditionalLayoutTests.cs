using DotOffice.Core;
using DotOffice.Layout;

namespace DotOffice.Tests.Layout;

public class SheetConditionalLayoutTests
{
    private static readonly Color Blue = new(0x63, 0x8E, 0xC6);

    private static LayoutPage Page(Cell cell, Action<Worksheet>? configure = null)
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, cell));
        sheet.ColumnWidths[0] = 100;
        sheet.RowHeights[0] = 20;
        configure?.Invoke(sheet);
        return SheetLayoutTests.Layout(sheet).Pages[0];
    }

    private static Cell Numbered(string text, CellDataBar? bar = null, CellIcon? icon = null, Action<SheetLayoutTests.CellStyleBuilder>? style = null) =>
        SheetLayoutTests.Number(text, SheetLayoutTests.Style(style)) with { Value = 1, DataBar = bar, Icon = icon };

    [Fact]
    public void A_solid_data_bar_is_a_rectangle_inset_by_a_point_and_drawn_below_the_text()
    {
        var page = Page(Numbered("5", new CellDataBar(0, 0.5, Blue, Gradient: false)));
        var rect = page.Items.OfType<RectItem>().Single();
        double left = 50.4;
        Assert.Equal(Blue, rect.Fill);
        Assert.Equal(left, rect.X, 6);
        Assert.Equal(50, rect.Width, 6);
        Assert.Equal(18, rect.Height, 6);
        Assert.Equal(54 + 1, rect.Y, 6);
        Assert.True(page.Items.IndexOf(rect) < page.Items.FindIndex(i => i is TextItem));
    }

    [Fact]
    public void A_gradient_data_bar_fades_from_its_base_and_a_negative_one_from_the_right()
    {
        var positive = Page(Numbered("5", new CellDataBar(0.25, 0.75, Blue, Gradient: true))).Items.OfType<GradientItem>().Single();
        Assert.Equal(0, positive.Degree);
        Assert.Equal((75.4, 50.0), (positive.X, positive.Width));
        Assert.Equal(Blue, positive.Stops[0].Color);
        Assert.True(positive.Stops[1].Color.R > Blue.R);
        var negative = Page(Numbered("-5", new CellDataBar(0, 0.25, Blue, Gradient: true, Negative: true))).Items.OfType<GradientItem>().Single();
        Assert.Equal(180, negative.Degree);
    }

    [Fact]
    public void Data_bars_scale_with_the_page_and_follow_the_cell_when_merged()
    {
        var page = Page(Numbered("5", new CellDataBar(0, 1, Blue, false)), s => { s.PageSetup.Scale = 0.5; });
        Assert.Equal(50, page.Items.OfType<RectItem>().Single().Width, 6);
        var merged = Page(Numbered("5", new CellDataBar(0, 1, Blue, false)), s => { s.ColumnWidths[1] = 40; s.Merges.Add(new CellRange(0, 0, 0, 1)); });
        Assert.Equal(140, merged.Items.OfType<RectItem>().Single().Width, 6);
    }

    [Theory]
    [InlineData(IconShape.ArrowUp, 7)]
    [InlineData(IconShape.ArrowDown, 7)]
    [InlineData(IconShape.ArrowSideways, 7)]
    public void Arrow_icons_are_polygons_in_the_icon_color(IconShape shape, int points)
    {
        var color = new Color(0, 0xB0, 0x50);
        var polygon = Page(Numbered("5", icon: new CellIcon(shape, color))).Items.OfType<PolygonItem>().Single();
        Assert.Equal(color, polygon.Fill);
        Assert.Equal(points, polygon.Points.Count);
        Assert.All(polygon.Points, p => Assert.InRange(p.X, 50.4 + 2 - 0.01, 50.4 + 2 + 10 + 0.01));
    }

    [Fact]
    public void Symbols_are_a_disc_with_a_white_glyph_and_lights_are_plain_discs()
    {
        var color = new Color(0xE0, 0x30, 0x30);
        var check = Page(Numbered("5", icon: new CellIcon(IconShape.Check, color))).Items;
        Assert.Equal(color, check.OfType<EllipseItem>().Single().Fill);
        Assert.Equal(Color.White, check.OfType<PolygonItem>().Single().Fill);
        var exclamation = Page(Numbered("5", icon: new CellIcon(IconShape.Exclamation, color))).Items;
        Assert.Equal(2, exclamation.OfType<EllipseItem>().Count());
        var light = Page(Numbered("5", icon: new CellIcon(IconShape.Circle, color))).Items;
        Assert.Single(light.OfType<EllipseItem>());
        Assert.Empty(light.OfType<PolygonItem>());
    }

    [Fact]
    public void The_icon_sits_in_the_cell_before_the_text_and_pushes_a_left_aligned_text_right()
    {
        var icon = new CellIcon(IconShape.Circle, Blue);
        double without = Page(SheetLayoutTests.Text("hi")).Items.OfType<TextItem>().Single().X;
        var page = Page(SheetLayoutTests.Text("hi") with { Icon = icon });
        var disc = page.Items.OfType<EllipseItem>().Single();
        var text = page.Items.OfType<TextItem>().Single();
        Assert.True(disc.X < text.X);
        Assert.True(disc.X + disc.Width <= text.X + 0.01);
        Assert.Equal(10, disc.Width, 6); // the size of the text
        Assert.Equal(without + 10 + 2, text.X, 6);
        Assert.InRange(disc.Y, 54, 54 + 20 - 10);
    }

    [Fact]
    public void A_right_aligned_text_stays_at_the_right_edge_despite_the_icon()
    {
        var icon = new CellIcon(IconShape.Circle, Blue);
        double without = Page(Numbered("5")).Items.OfType<TextItem>().Single().X;
        double with = Page(Numbered("5", icon: icon)).Items.OfType<TextItem>().Single().X;
        Assert.Equal(without, with, 6);
    }

    [Fact]
    public void An_icon_without_text_is_still_drawn()
    {
        var page = Page(Numbered("", icon: new CellIcon(IconShape.ArrowUp, Blue)));
        Assert.Single(page.Items.OfType<PolygonItem>());
        Assert.Empty(page.Items.OfType<TextItem>());
    }

    [Fact]
    public void Icons_follow_the_page_scale()
    {
        var page = Page(Numbered("5", icon: new CellIcon(IconShape.Circle, Blue)), s => s.PageSetup.Scale = 0.5);
        Assert.Equal(5, page.Items.OfType<EllipseItem>().Single().Width, 6);
    }

    [Fact]
    public void An_icon_makes_a_wrapping_cell_wrap_earlier()
    {
        var style = SheetLayoutTests.Style(s => s.Wrap = true);
        int Lines(Cell cell) => Page(cell, s => { s.ColumnWidths[0] = 60; s.RowHeights[0] = 100; }).Items.OfType<TextItem>().Select(t => t.Baseline).Distinct().Count();
        var text = SheetLayoutTests.Text("aaa bbb ccc", style);
        Assert.True(Lines(text with { Icon = new CellIcon(IconShape.Circle, Blue) }) > Lines(text));
    }
}
