using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

public class FloatingTableTests
{
    private const string Run = "<w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/></w:rPr>";
    private static readonly string Words = string.Join(' ', Enumerable.Repeat("alpha beta gamma delta", 40));

    private static string Cell(string text) => $"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r>{Run}<w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static string Table(string tblpPr, int rows = 1) =>
        $"<w:tbl><w:tblPr><w:tblpPr {tblpPr}/><w:tblW w:w=\"2000\" w:type=\"dxa\"/><w:tblLayout w:type=\"fixed\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
        string.Concat(Enumerable.Range(0, rows).Select(i => "<w:tr>" + Cell(i == 0 ? "boxed" : "row" + i) + "</w:tr>")) + "</w:tbl>";

    private sealed class ListSink(List<ConversionWarning> list) : IConversionWarningSink
    {
        public void Warn(ConversionWarning warning) => list.Add(warning);
    }

    private static (Document Doc, LayoutDocument Layout, List<ConversionWarning> Warnings) Run1(string tblpPr, int rows = 1)
    {
        var warnings = new List<ConversionWarning>();
        var sink = new ListSink(warnings);
        string body = Table(tblpPr, rows) + $"<w:p><w:r>{Run}<w:t>{Words}</w:t></w:r></w:p>" + Section();
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream(), new DocxReaderOptions { Warnings = sink });
        var layout = new LayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(Noto)), Warnings = sink }).Layout(doc);
        return (doc, layout, warnings);
    }

    private static TextItem Boxed(LayoutPage page) => Texts(page).First(t => t.Text == "boxed");

    private static IEnumerable<TextItem> Body(LayoutPage page) => Texts(page).Where(t => t.Text != "boxed" && !t.Text.StartsWith("row"));

    [Fact]
    public void Reads_tblpPr_into_a_position_with_distances()
    {
        var (doc, _, _) = Run1("w:leftFromText=\"180\" w:rightFromText=\"360\" w:topFromText=\"40\" w:bottomFromText=\"80\" w:horzAnchor=\"page\" w:vertAnchor=\"margin\" w:tblpX=\"1000\" w:tblpY=\"2000\"");
        var position = ((Table)doc.Sections[0].Blocks[0]).Properties.Position!;
        Assert.Equal(new FloatPosition(FloatReference.Page, FloatAlignment.Offset, 50), position.Horizontal);
        Assert.Equal(new FloatPosition(FloatReference.Margin, FloatAlignment.Offset, 100), position.Vertical);
        Assert.Equal(new Thickness(2, 18, 4, 9), position.Distance);
    }

    [Fact]
    public void Specs_become_alignments_and_inline_is_not_floating()
    {
        var (doc, _, _) = Run1("w:horzAnchor=\"margin\" w:tblpXSpec=\"right\" w:vertAnchor=\"margin\" w:tblpYSpec=\"center\"");
        var position = ((Table)doc.Sections[0].Blocks[0]).Properties.Position!;
        Assert.Equal(FloatAlignment.End, position.Horizontal.Alignment);
        Assert.Equal(FloatAlignment.Center, position.Vertical.Alignment);
        Assert.Null(((Table)Run1("w:tblpYSpec=\"inline\"").Doc.Sections[0].Blocks[0]).Properties.Position);
    }

    [Fact]
    public void A_table_at_page_offsets_sits_there_and_ignores_the_text_flow_position()
    {
        var page = Run1("w:horzAnchor=\"page\" w:vertAnchor=\"page\" w:tblpX=\"4000\" w:tblpY=\"3000\"").Layout.Pages[0];
        var box = Boxed(page);
        Assert.Equal(200 + 5.4, box.X, 0.5);
        Assert.True(box.Baseline > 150 && box.Baseline < 150 + 30);
    }

    [Fact]
    public void Text_wraps_beside_a_floating_table_and_continues_full_width_below_it()
    {
        // Table flush right in the margin area, 100 pt wide, three rows tall.
        var page = Run1("w:horzAnchor=\"margin\" w:tblpXSpec=\"right\" w:vertAnchor=\"text\" w:tblpY=\"0\"", rows: 3).Layout.Pages[0];
        var box = Boxed(page);
        double tableLeft = 72 + 468 - 100;
        Assert.Equal(tableLeft + 5.4, box.X, 0.5);
        double tableBottom = Texts(page).Where(t => t.Text.StartsWith("row")).Max(t => t.Baseline) + 20;
        var lines = Body(page).GroupBy(t => Math.Round(t.Baseline, 1)).OrderBy(g => g.Key).ToList();
        var beside = lines.Where(g => g.Key < tableBottom - 25).ToList();
        var below = lines.Where(g => g.Key > tableBottom + 25).ToList();
        Assert.NotEmpty(beside);
        Assert.NotEmpty(below);
        // Beside the table a line holds fewer characters than the full-width lines below it.
        Assert.True(beside.Max(g => g.Sum(t => t.Text.Length)) < below.Max(g => g.Sum(t => t.Text.Length)) - 8);
        Assert.Equal(72, beside[0].Min(t => t.X), 0.5);
    }

    [Fact]
    public void Distance_from_text_keeps_text_away_from_the_table()
    {
        int FirstLineLength(string extra)
        {
            var page = Run1("w:horzAnchor=\"margin\" w:tblpXSpec=\"right\" w:vertAnchor=\"text\" " + extra, rows: 3).Layout.Pages[0];
            return Body(page).GroupBy(t => Math.Round(t.Baseline, 1)).OrderBy(g => g.Key).First().Sum(t => t.Text.Length);
        }
        Assert.True(FirstLineLength("w:leftFromText=\"1440\"") < FirstLineLength("w:leftFromText=\"0\"") - 5);
    }

    [Fact]
    public void A_table_taller_than_the_page_falls_back_to_the_flow_with_a_warning()
    {
        var (_, layout, warnings) = Run1("w:horzAnchor=\"margin\" w:tblpXSpec=\"right\" w:vertAnchor=\"text\"", rows: 120);
        Assert.Contains(warnings, w => w.Code == "docx.floating-table");
        Assert.True(layout.Pages.Count > 1);
        Assert.Equal(72 + 5.4, Boxed(layout.Pages[0]).X, 0.5);
    }
}
