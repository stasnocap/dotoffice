using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class TextDirectionTests
{
    private const string Font = "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/><w:sz w:val=\"24\"/>";
    private static readonly string Lorem = string.Join(' ', Enumerable.Repeat("lorem ipsum dolor", 8));

    private static LayoutPage Layout(string direction, string rowPr = "<w:trPr><w:trHeight w:val=\"3000\" w:hRule=\"exact\"/></w:trPr>", string cellPr = "", string? text = null)
    {
        var cell = $"<w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/>{(direction.Length > 0 ? $"<w:textDirection w:val=\"{direction}\"/>" : "")}{cellPr}</w:tcPr>{P(R(text ?? Lorem, Font))}</w:tc>";
        var body = $"<w:tbl><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr>{rowPr}{cell}</w:tr></w:tbl>{P("")}{Section()}";
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        return new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
    }

    [Theory]
    [InlineData("btLr", 90)]
    [InlineData("tbRl", -90)]
    public void Cell_text_is_turned_a_quarter_about_the_center_of_the_cell(string direction, double rotation)
    {
        var page = Layout(direction);
        var turn = page.Items.OfType<PushItem>().Single(p => p.Rotation != 0);
        Assert.Equal(rotation, turn.Rotation);
        Assert.Equal(72 + 100, turn.CenterX, 0.5);
        Assert.Equal(72 + 75, turn.CenterY, 0.5);
    }

    [Fact]
    public void Lines_are_as_long_as_the_row_is_high_and_stacked_across_the_cell()
    {
        var page = Layout("btLr");
        var texts = Texts(page).ToList();
        double centerX = page.Items.OfType<PushItem>().Single(p => p.Rotation != 0).CenterX;
        // Before the turn the lines run horizontally through the center, 150 pt long.
        Assert.All(texts, t => Assert.InRange(t.X, centerX - 75 - 0.5, centerX + 75));
        Assert.All(texts, t => Assert.True(t.X + 1 <= centerX + 75));
        Assert.True(texts.Select(t => Math.Round(t.Baseline, 1)).Distinct().Count() > 1);
    }

    [Fact]
    public void Row_without_a_fixed_height_is_as_high_as_the_unwrapped_text()
    {
        var page = Layout("btLr", rowPr: "", text: "lorem ipsum dolor lorem ipsum dolor");
        var turn = page.Items.OfType<PushItem>().Single(p => p.Rotation != 0);
        Assert.True(turn.Clip!.Value.Height > 150);
        Assert.Single(Texts(page).Select(t => Math.Round(t.Baseline, 1)).Distinct());
    }

    [Fact]
    public void Text_longer_than_the_page_wraps_at_the_body_height_and_the_row_grows_across()
    {
        var long_ = string.Join(' ', Enumerable.Repeat("lorem ipsum dolor", 400));
        var cell = $"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:textDirection w:val=\"btLr\"/></w:tcPr>{P(R(long_, Font))}</w:tc>";
        var body = $"<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr>{cell}</w:tr></w:tbl>{P("")}{Section()}";
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        var page = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)) }).Layout(doc).Pages[0];
        var turn = page.Items.OfType<PushItem>().Single(p => p.Rotation != 0);
        Assert.True(turn.Clip!.Value.Height <= page.Height - 144 + 0.5);
        Assert.True(Texts(page).Select(t => Math.Round(t.Baseline, 1)).Distinct().Count() > 1);
    }

    [Fact]
    public void Horizontal_cells_are_not_turned()
    {
        Assert.DoesNotContain(Layout("").Items, i => i is PushItem { Rotation: not 0 });
    }
}
