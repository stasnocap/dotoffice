using DotOffice.Core;
using DotOffice.Layout;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Layout;

public class TableLayoutTests
{
    private static readonly Border Line = new(BorderStyle.Single, 0.5, Color.Black);
    private static readonly BorderSet Box = new() { Top = Line, Bottom = Line, Left = Line, Right = Line };

    private static TableCell Cell(string text, int span = 1, VerticalMerge merge = VerticalMerge.None, TableWidth? width = null,
        CellVerticalAlignment valign = CellVerticalAlignment.Top)
    {
        var cell = new TableCell { GridSpan = span, VerticalMerge = merge, Borders = Box, Width = width ?? TableWidth.Auto, VerticalAlignment = valign };
        cell.Blocks.Add(Para(text));
        return cell;
    }

    private static Table Table(IEnumerable<double> grid, TableProperties? props = null, params TableRow[] rows)
    {
        var t = new Table { Properties = props ?? new TableProperties() };
        t.Grid.AddRange(grid);
        t.Rows.AddRange(rows);
        return t;
    }

    private static TableRow Row(params TableCell[] cells) => Row(new TableRow(), cells);

    private static TableRow Row(TableRow row, params TableCell[] cells)
    {
        row.Cells.AddRange(cells);
        return row;
    }

    private static TextItem Find(LayoutPage page, string text) => Texts(page).Single(t => t.Text == text);

    [Fact]
    public void Uses_grid_columns_and_cell_margins()
    {
        var page = Render(Doc(Table([100, 200], null, Row(Cell("a"), Cell("b"))))).Pages.Single();
        Assert.Equal(72 + 5.4, Find(page, "a").X, 2);
        Assert.Equal(72 + 100 + 5.4, Find(page, "b").X, 2);
        var verticals = page.Items.OfType<LineItem>().Where(l => l.X1 == l.X2).Select(l => l.X1).Distinct().Order().ToArray();
        Assert.Equal([72.0, 172.0, 372.0], verticals);
    }

    [Fact]
    public void Percentage_width_scales_the_grid()
    {
        var page = Render(Doc(Table([100, 100], new TableProperties { Width = new TableWidth(100, TableWidthUnit.Percent) },
            Row(Cell("a"), Cell("b"))))).Pages.Single();
        Assert.Equal(72 + 234 + 5.4, Find(page, "b").X, 2); // 468pt split evenly
    }

    [Fact]
    public void Auto_layout_distributes_width_by_content()
    {
        var page = Render(Doc(Table([], null, Row(Cell("x"), Cell(Words(60)))))).Pages.Single();
        var x = Find(page, "x");
        var first = Find(page, "word1");
        double firstColumn = first.X - 5.4 - 72;
        Assert.True(firstColumn < 40, $"narrow content should get a narrow column, got {firstColumn}");
        Assert.True(Right(Texts(page).Where(t => t.Text.StartsWith("word")).MaxBy(Right)!) <= 72 + 468 + 0.01);
        Assert.Equal(72 + 5.4, x.X, 2);
    }

    [Fact]
    public void Auto_layout_respects_preferred_cell_widths()
    {
        var page = Render(Doc(Table([], null, Row(Cell("a", width: new TableWidth(150, TableWidthUnit.Points)), Cell("b"))))).Pages.Single();
        Assert.Equal(72 + 150 + 5.4, Find(page, "b").X, 2);
    }

    [Fact]
    public void Column_spans_and_alignment()
    {
        var page = Render(Doc(Table([100, 100, 100], new TableProperties { Alignment = TableAlignment.Center },
            Row(Cell("wide", span: 2), Cell("c")), Row(Cell("a"), Cell("b"), Cell("c2"))))).Pages.Single();
        double left = 72 + (468 - 300) / 2.0;
        Assert.Equal(left + 5.4, Find(page, "wide").X, 2);
        Assert.Equal(left + 200 + 5.4, Find(page, "c").X, 2);
    }

    [Fact]
    public void Row_height_rules_and_vertical_alignment()
    {
        var exact = new TableRow { Height = 50, HeightRule = RowHeightRule.Exact };
        var page = Render(Doc(Table([100, 100], null,
            Row(exact, Cell("top"), Cell("mid", valign: CellVerticalAlignment.Center)), Row(Cell("next"), Cell("n2"))))).Pages.Single();
        var top = Find(page, "top");
        var mid = Find(page, "mid");
        var next = Find(page, "next");
        Assert.True(mid.Baseline > top.Baseline + 10);
        Assert.True(next.Baseline - top.Baseline >= 50 - 0.01);
    }

    [Fact]
    public void Vertically_merged_cells_span_rows_and_grow_the_last_row()
    {
        var page = Render(Doc(Table([100, 100], null,
            Row(Cell(Words(30), merge: VerticalMerge.Restart), Cell("r1")),
            Row(Cell("", merge: VerticalMerge.Continue), Cell("r2")),
            Row(Cell("after"), Cell("r3"))))).Pages.Single();
        var merged = Texts(page).Where(t => t.Text.StartsWith("word")).ToList();
        var after = Find(page, "after");
        Assert.True(merged.Max(t => t.Baseline) < after.Baseline, "rows below the merge start after its content");
        Assert.True(Find(page, "r2").Baseline > Find(page, "r1").Baseline);
        // No horizontal border inside the merged cell between rows 1 and 2.
        double r2Top = Find(page, "r2").Baseline - 12;
        Assert.DoesNotContain(page.Items.OfType<LineItem>(), l => l.Y1 == l.Y2 && l.X1 == 72 && l.Y1 > Find(page, "r1").Baseline && l.Y1 < r2Top + 12);
    }

    [Fact]
    public void Long_tables_split_between_rows_and_repeat_header_rows()
    {
        var rows = new List<TableRow> { Row(new TableRow { IsHeader = true }, Cell("Header"), Cell("H2")) };
        for (int i = 1; i <= 80; i++) rows.Add(Row(Cell($"row{i}"), Cell("x")));
        var layout = Render(Doc(Table([100, 100], null, [.. rows])));
        Assert.True(layout.Pages.Count >= 2);
        foreach (var page in layout.Pages)
        {
            var lines = LayoutTestHelpers.Lines(page);
            Assert.Equal("Header", lines[0].Items[0].Text);
            Assert.True(lines.Max(l => l.Baseline) <= 792 - 72);
        }
        var all = layout.Pages.SelectMany(p => Texts(p)).Where(t => t.Text.StartsWith("row")).Select(t => t.Text).ToList();
        Assert.Equal(Enumerable.Range(1, 80).Select(i => $"row{i}"), all); // every row exactly once, in order
    }

    [Fact]
    public void Tall_rows_split_across_pages_unless_cant_split()
    {
        var tall = Row(Cell(string.Join(' ', Enumerable.Repeat("filler", 700))), Cell("side"));
        var layout = Render(Doc(Para("intro"), Table([234, 234], null, tall)));
        Assert.True(layout.Pages.Count >= 2);
        Assert.Contains(Texts(layout.Pages[0]), t => t.Text == "filler"); // row starts on the first page
        Assert.Contains(Texts(layout.Pages[1]), t => t.Text == "filler");

        var kept = Row(new TableRow { CantSplit = true }, Cell(string.Join(' ', Enumerable.Repeat("filler", 300))));
        var keptLayout = Render(Doc(Para(Words(400)), Table([468], null, kept)));
        var firstFillerPage = keptLayout.Pages.FindIndex(p => Texts(p).Any(t => t.Text == "filler"));
        // The whole row moved to a page of its own instead of splitting.
        Assert.Equal(firstFillerPage, keptLayout.Pages.FindLastIndex(p => Texts(p).Any(t => t.Text == "filler")));
        Assert.Equal(300, Texts(keptLayout.Pages[firstFillerPage]).Count(t => t.Text == "filler"));
    }

    [Fact]
    public void Nested_tables_are_laid_out_inside_cells()
    {
        var inner = Table([60], null, Row(Cell("inner")));
        var cell = new TableCell { Borders = Box };
        cell.Blocks.Add(inner);
        cell.Blocks.Add(Para(""));
        var page = Render(Doc(Table([200], null, Row(cell)))).Pages.Single();
        Assert.Equal(72 + 5.4 + 5.4, Find(page, "inner").X, 2);
    }
}