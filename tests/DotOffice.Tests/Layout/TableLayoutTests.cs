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
        // The 0.5 pt left border has its outer edge at the margin, so every column edge is 0.25 pt further right.
        Assert.Equal(72.25 + 5.4, Find(page, "a").X, 2);
        Assert.Equal(72.25 + 100 + 5.4, Find(page, "b").X, 2);
        var verticals = page.Items.OfType<LineItem>().Where(l => l.X1 == l.X2).Select(l => l.X1).Distinct().Order().ToArray();
        Assert.Equal([72.25, 172.25, 372.25], verticals);
    }

    [Fact]
    public void Percentage_width_scales_the_grid()
    {
        var page = Render(Doc(Table([100, 100], new TableProperties { Width = new TableWidth(100, TableWidthUnit.Percent) },
            Row(Cell("a"), Cell("b"))))).Pages.Single();
        Assert.Equal(72.25 + 234 + 5.4, Find(page, "b").X, 2); // 468pt split evenly
    }

    [Fact]
    public void Auto_layout_distributes_width_by_content()
    {
        var page = Render(Doc(Table([], null, Row(Cell("x"), Cell(Words(60)))))).Pages.Single();
        var x = Find(page, "x");
        var first = Find(page, "word1");
        double firstColumn = first.X - 5.4 - 72.25;
        Assert.True(firstColumn < 40, $"narrow content should get a narrow column, got {firstColumn}");
        Assert.True(Right(Texts(page).Where(t => t.Text.StartsWith("word")).MaxBy(Right)!) <= 72 + 468 + 0.01);
        Assert.Equal(72.25 + 5.4, x.X, 2);
    }

    [Fact]
    public void Auto_layout_respects_preferred_cell_widths()
    {
        var page = Render(Doc(Table([], null, Row(Cell("a", width: new TableWidth(150, TableWidthUnit.Points)), Cell("b"))))).Pages.Single();
        Assert.Equal(72.25 + 150 + 5.4, Find(page, "b").X, 2);
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
        Assert.DoesNotContain(page.Items.OfType<LineItem>(), l => l.Y1 == l.Y2 && l.X1 == 72.25 && l.Y1 > Find(page, "r1").Baseline && l.Y1 < r2Top + 12);
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
        // Both tables are a quarter point right of their indent, for their 0.5 pt left borders.
        Assert.Equal(72.25 + 5.4 + 0.25 + 5.4, Find(page, "inner").X, 2);
    }

    private static TableCell Cell(string text, BorderSet borders)
    {
        var cell = new TableCell { Borders = borders };
        cell.Blocks.Add(Para(text));
        return cell;
    }

    /// <summary>before, a one-column table of r1..rN whose cells have the given borders, after.</summary>
    private static LayoutPage Bordered(int rows, Func<int, BorderSet> borders) =>
        Render(Doc(Para("before"),
            Table([200], null, [.. Enumerable.Range(1, rows).Select(i => Row(Cell($"r{i}", borders(i))))]),
            Para("after"))).Pages.Single();

    private static BorderSet Edges(double top, double bottom)
    {
        Border Of(double w) => w > 0 ? new Border(BorderStyle.Single, w, Color.Black) : Border.None;
        return new BorderSet { Top = Of(top), Bottom = Of(bottom), Left = Of(Math.Max(top, bottom)), Right = Of(Math.Max(top, bottom)) };
    }

    // Word 2016: with Arial 12 (13.8 pt lines) the row pitch is 14.05, 14.3, 14.8, 15.3, 16.8 and 19.8 pt for borders of
    // 0.25, 0.5, 1, 1.5, 3 and 6 pt, and the top and bottom border add their width before the first and after the last row.
    [Theory]
    [InlineData(0.25)]
    [InlineData(0.5)]
    [InlineData(1)]
    [InlineData(1.5)]
    [InlineData(3)]
    [InlineData(6)]
    public void Horizontal_borders_add_their_width_to_the_height(double width)
    {
        var page = Bordered(3, _ => Edges(width, width));
        double line = NotoLineHeight();
        Assert.Equal(line + width, Find(page, "r1").Baseline - Find(page, "before").Baseline, 6);
        Assert.Equal(line + width, Find(page, "r2").Baseline - Find(page, "r1").Baseline, 6);
        Assert.Equal(line + width, Find(page, "r3").Baseline - Find(page, "r2").Baseline, 6);
        Assert.Equal(line + width, Find(page, "after").Baseline - Find(page, "r3").Baseline, 6);
    }

    [Fact]
    public void Tables_without_horizontal_borders_keep_the_line_pitch()
    {
        var page = Bordered(3, _ => BorderSet.Empty);
        double line = NotoLineHeight();
        Assert.Equal(line, Find(page, "r1").Baseline - Find(page, "before").Baseline, 6);
        Assert.Equal(line, Find(page, "r2").Baseline - Find(page, "r1").Baseline, 6);
        Assert.Equal(line, Find(page, "after").Baseline - Find(page, "r3").Baseline, 6);
    }

    // Word 2016, 3 pt outer borders and no inside ones: 16.8 pt to the first row, 13.8 pt between rows, 16.8 pt to the next paragraph.
    [Fact]
    public void Outer_borders_only_add_height_above_and_below_the_table()
    {
        var page = Bordered(3, i => Edges(i == 1 ? 3 : 0, i == 3 ? 3 : 0));
        double line = NotoLineHeight();
        Assert.Equal(line + 3, Find(page, "r1").Baseline - Find(page, "before").Baseline, 6);
        Assert.Equal(line, Find(page, "r2").Baseline - Find(page, "r1").Baseline, 6);
        Assert.Equal(line, Find(page, "r3").Baseline - Find(page, "r2").Baseline, 6);
        Assert.Equal(line + 3, Find(page, "after").Baseline - Find(page, "r3").Baseline, 6);
    }

    // Word 2016, 0.5 pt borders and 3 pt top and bottom cell margins: the pitch is 13.8 + 6 + 0.5.
    [Fact]
    public void Border_height_adds_to_cell_margins()
    {
        TableCell Padded(string text)
        {
            var cell = new TableCell { Borders = Box, Margins = new Thickness(3, 5, 3, 5) };
            cell.Blocks.Add(Para(text));
            return cell;
        }
        var page = Render(Doc(Table([200], null, Row(Padded("r1")), Row(Padded("r2"))))).Pages.Single();
        Assert.Equal(NotoLineHeight() + 6 + 0.5, Find(page, "r2").Baseline - Find(page, "r1").Baseline, 6);
    }

    [Fact]
    public void Horizontal_border_lines_are_centered_in_the_space_they_take()
    {
        var page = Bordered(3, _ => Edges(4, 4));
        double line = NotoLineHeight();
        var verticals = page.Items.OfType<LineItem>().Where(l => l.X1 == l.X2).ToList();
        double top = verticals.Min(l => l.Y1), bottom = verticals.Max(l => l.Y2);
        Assert.Equal(4 + 3 * (line + 4), bottom - top, 6);
        var horizontals = page.Items.OfType<LineItem>().Where(l => l.Y1 == l.Y2).Select(l => Math.Round(l.Y1 - top, 6)).Distinct().Order().ToArray();
        // Top edge, two inside edges (the bottom border of one row and the top border of the next on the same line), bottom edge.
        Assert.Equal(4, horizontals.Length);
        for (int i = 0; i < 4; i++) Assert.Equal(2 + i * (line + 4), horizontals[i], 5);
        // The vertical lines of the rows join: each row's ends where the next one's starts.
        var left = verticals.Where(l => l.X1 == 74).OrderBy(l => l.Y1).ToList();
        Assert.Equal(3, left.Count);
        Assert.Equal(left[0].Y2, left[1].Y1, 6);
        Assert.Equal(left[1].Y2, left[2].Y1, 6);
        // The text sits below the band of the top border.
        Assert.Equal(top + 4 + line - NotoDescent(), Find(page, "r1").Baseline, 6);
    }

    // Not measured in Word: a row edge whose cells have borders of different widths is as high as the widest of them.
    [Fact]
    public void A_row_edge_is_as_high_as_its_widest_border()
    {
        var thin = Edges(1, 1);
        var page = Render(Doc(Table([100, 100], null,
            Row(Cell("a1", thin), Cell("b1", thin with { Bottom = thin.Bottom with { Width = 5 } })),
            Row(Cell("a2", thin with { Top = thin.Top with { Width = 2 } }), Cell("b2", thin))))).Pages.Single();
        Assert.Equal(Find(page, "a1").Baseline, Find(page, "b1").Baseline, 6);
        Assert.Equal(Find(page, "a2").Baseline, Find(page, "b2").Baseline, 6);
        Assert.Equal(NotoLineHeight() + 5, Find(page, "a2").Baseline - Find(page, "a1").Baseline, 6);
        // All four border lines of that edge lie on its middle.
        double first = Find(page, "a1").Baseline + NotoDescent();
        var inside = page.Items.OfType<LineItem>().Where(l => l.Y1 == l.Y2 && l.Y1 > first - 0.01 && l.Y1 < first + 5.01).ToList();
        Assert.Equal(4, inside.Count);
        Assert.All(inside, l => Assert.Equal(first + 2.5, l.Y1, 6));
    }

    // Not measured in Word: the edges between the rows of a vertically merged cell count as part of that cell's height.
    [Fact]
    public void Borders_between_the_rows_of_a_merged_cell_do_not_grow_the_table()
    {
        var box = Edges(2, 2);
        TableCell Merged(string text, VerticalMerge merge)
        {
            var cell = new TableCell { Borders = box, VerticalMerge = merge };
            cell.Blocks.Add(Para(text));
            return cell;
        }
        var page = Render(Doc(Table([100, 100], null,
            Row(Merged("m", VerticalMerge.Restart), Cell("r1", box)),
            Row(Merged("", VerticalMerge.Continue), Cell("r2", box))), Para("after"))).Pages.Single();
        double line = NotoLineHeight();
        Assert.Equal(Find(page, "r1").Baseline, Find(page, "m").Baseline, 6);
        Assert.Equal(line + 2, Find(page, "r2").Baseline - Find(page, "r1").Baseline, 6);
        Assert.Equal(line + 2, Find(page, "after").Baseline - Find(page, "r2").Baseline, 6);
        // The edge between the rows is drawn beside the merged cell only.
        double edge = Find(page, "r1").Baseline + NotoDescent() + 1;
        var lines = page.Items.OfType<LineItem>().Where(l => l.Y1 == l.Y2 && Math.Abs(l.Y1 - edge) < 0.001).ToList();
        Assert.NotEmpty(lines);
        Assert.All(lines, l => Assert.Equal(173, Math.Min(l.X1, l.X2), 6));
    }

    // Word 2016: a table that starts a page has its first row lower than a first line of text by the top border's width.
    // Not measured: a row that follows a page break inside the table keeps the space of the edge above it.
    [Fact]
    public void Rows_at_the_top_of_a_page_stand_below_their_top_border()
    {
        var box = Edges(6, 6);
        var rows = Enumerable.Range(1, 50).Select(i => Row(Cell($"row{i}", box))).ToArray();
        var layout = Render(Doc(Table([200], null, rows), Para("after")));
        Assert.True(layout.Pages.Count >= 2);
        double first = 72 + 6 + NotoLineHeight() - NotoDescent();
        foreach (var page in layout.Pages.Where(p => Texts(p).Any(t => t.Text.StartsWith("row"))))
        {
            Assert.Equal(first, Texts(page).Where(t => t.Text.StartsWith("row")).Min(t => t.Baseline), 6);
            // Every row fits above the bottom margin with the border under it.
            Assert.True(page.Items.OfType<LineItem>().Where(l => l.X1 == l.X2).Max(l => l.Y2) <= 792 - 72 + 0.01);
        }
        var last = layout.Pages[^1];
        Assert.Equal(NotoLineHeight() + 6, Find(last, "after").Baseline - Find(last, "row50").Baseline, 6);
    }

    // Not measured in Word: the second part of a row split across pages has no top border and keeps the table's bottom border.
    [Fact]
    public void Split_rows_keep_the_bottom_border_space()
    {
        var box = Edges(6, 6);
        var layout = Render(Doc(Table([468], null, Row(Cell(string.Join(' ', Enumerable.Repeat("filler", 900)), box))), Para("after")));
        Assert.Equal(2, layout.Pages.Count);
        Assert.Equal(72 + 6 + NotoLineHeight() - NotoDescent(), Texts(layout.Pages[0]).Min(t => t.Baseline), 6);
        var rest = layout.Pages[1];
        Assert.Equal(72 + NotoLineHeight() - NotoDescent(), Texts(rest).Min(t => t.Baseline), 6);
        double lastFiller = Texts(rest).Where(t => t.Text == "filler").Max(t => t.Baseline);
        Assert.Equal(NotoLineHeight() + 6, Find(rest, "after").Baseline - lastFiller, 6);
        var bottom = rest.Items.OfType<LineItem>().Single(l => l.Y1 == l.Y2);
        Assert.Equal(lastFiller + NotoDescent() + 3, bottom.Y1, 6);
    }

    private static Border Wide(double w) => w > 0 ? new Border(BorderStyle.Single, w, Color.Black) : Border.None;

    private static BorderSet Sides(double left, double right) => new() { Left = Wide(left), Right = Wide(right) };

    private static TableCell Cell(string text, BorderSet borders, Thickness margins)
    {
        var cell = new TableCell { Borders = borders, Margins = margins };
        cell.Blocks.Add(Para(text));
        return cell;
    }

    private static Document WithSettings(Document doc, DocumentSettings settings)
    {
        var copy = new Document { Settings = settings };
        copy.Sections.AddRange(doc.Sections);
        return copy;
    }

    private static double[] Verticals(LayoutPage page) =>
        [.. page.Items.OfType<LineItem>().Where(l => l.X1 == l.X2).Select(l => Math.Round(l.X1, 6)).Distinct().Order()];

    // Word 2016, compatibility mode 15, left margin 72 pt, two 150 pt columns, all borders of one width: the left border's outer
    // edge is at 72.02 for 0.25, 0.5, 1, 1.5, 3 and 6 pt, the inside border starts at 222.05 and the right one at 372.07, and the
    // text of the first cell (left margin 5 pt) is at 77.30, 78.50 and 80.06 for 0.5, 3 and 6 pt (77.06 without borders).
    [Theory]
    [InlineData(0)]
    [InlineData(0.25)]
    [InlineData(0.5)]
    [InlineData(1)]
    [InlineData(1.5)]
    [InlineData(3)]
    [InlineData(6)]
    public void Left_border_has_its_outer_edge_at_the_indent(double width)
    {
        var box = Sides(width, width);
        var margins = new Thickness(0, 5, 0, 5);
        var page = Render(Doc(Table([150, 150], null, Row(Cell("a", box, margins), Cell("b", box, margins))))).Pages.Single();
        Assert.Equal(72 + width / 2 + 5, Find(page, "a").X, 6);
        Assert.Equal(72 + width / 2 + 150 + 5, Find(page, "b").X, 6);
        if (width > 0) Assert.Equal([72 + width / 2, 222 + width / 2, 372 + width / 2], Verticals(page));
    }

    // Word 2016, mode 15: with tblInd 10 pt and 0.5 pt borders the left border starts at 81.98 and the text at 87.26.
    [Fact]
    public void Left_border_edge_follows_the_table_indent()
    {
        var box = Sides(0.5, 0.5);
        var page = Render(Doc(Table([150, 150], new TableProperties { Indent = 10 },
            Row(Cell("a", box, new Thickness(0, 5, 0, 5)), Cell("b", box, new Thickness(0, 5, 0, 5)))))).Pages.Single();
        Assert.Equal(82.25, Verticals(page)[0], 6);
        Assert.Equal(87.25, Find(page, "a").X, 6);
    }

    // Word 2016, mode 15: a centered table with 0.5 pt borders has its text where a borderless one has it (161.06).
    // Right-aligned tables were not measured and are left alone as well.
    [Theory]
    [InlineData(TableAlignment.Center, 72 + 84.0)]
    [InlineData(TableAlignment.Right, 72 + 168.0)]
    public void Centered_and_right_aligned_tables_are_not_shifted_by_their_border(TableAlignment alignment, double left)
    {
        var box = Sides(3, 3);
        var page = Render(Doc(Table([150, 150], new TableProperties { Alignment = alignment },
            Row(Cell("a", box, new Thickness(0, 5, 0, 5)), Cell("b", box, new Thickness(0, 5, 0, 5)))))).Pages.Single();
        Assert.Equal([left, left + 150, left + 300], Verticals(page));
        Assert.Equal(left + 5, Find(page, "a").X, 6);
    }

    // Word 2016, mode 15, 0.5 pt borders: cell margins of 5, 10 and 20 pt put the text at 77.30, 82.22 and 92.30 (border center
    // plus margin), and no margin puts it at 72.50, clear of the border (one sample).
    [Theory]
    [InlineData(0, 0.5, 72.5)]
    [InlineData(5, 0.5, 77.25)]
    [InlineData(10, 0.5, 82.25)]
    [InlineData(20, 0.5, 92.25)]
    [InlineData(1, 6, 78)] // not measured: half of a wide border is more than a small margin
    public void Cell_content_is_inset_by_the_margin_or_half_the_left_border(double margin, double border, double x)
    {
        var page = Render(Doc(Table([150], null, Row(Cell("a", Sides(border, border), new Thickness(0, margin, 0, margin)))))).Pages.Single();
        Assert.Equal(x, Find(page, "a").X, 6);
    }

    // Not measured in Word: the right side and inner cells are taken to behave like the left side of the first cell.
    [Fact]
    public void Inner_cells_and_right_sides_keep_clear_of_their_borders_too()
    {
        var none = new Thickness(0, 0, 0, 0);
        var right = Para("b", new ParagraphProperties { Alignment = ParagraphAlignment.Right });
        var second = new TableCell { Borders = Sides(4, 6), Margins = none };
        second.Blocks.Add(right);
        var page = Render(Doc(Table([150, 150], null, Row(Cell("a", Sides(0, 4), none), second)))).Pages.Single();
        Assert.Equal(72, Find(page, "a").X, 6); // no left border: no shift and no inset
        Assert.Equal(72 + 300 - 3, Right(Find(page, "b")), 6);
        var left = new TableCell { Borders = Sides(4, 6), Margins = none };
        left.Blocks.Add(Para("c"));
        page = Render(Doc(Table([150, 150], null, Row(Cell("a", Sides(0, 4), none), left)))).Pages.Single();
        Assert.Equal(72 + 150 + 2, Find(page, "c").X, 6);
    }

    // Not measured in Word: the shift is half the left border of the first row's first cell, whatever the other rows have, and
    // a table without a left border is not shifted by its other borders.
    [Fact]
    public void The_first_cell_of_the_first_row_decides_the_shift()
    {
        var margins = new Thickness(0, 5, 0, 5);
        var page = Render(Doc(Table([150], null, Row(Cell("r1", Sides(2, 2), margins)), Row(Cell("r2", Sides(6, 6), margins))))).Pages.Single();
        Assert.Equal(78, Find(page, "r1").X, 6);
        Assert.Equal(78, Find(page, "r2").X, 6);
        Assert.Equal([73.0, 223.0], Verticals(page));

        page = Render(Doc(Table([150], null, Row(Cell("r1", Sides(0, 6) with { Top = Wide(6), Bottom = Wide(6) }, margins))))).Pages.Single();
        Assert.Equal(77, Find(page, "r1").X, 6);
        Assert.Equal([222.0], Verticals(page));
    }

    // Word 2016, modes 14 and lower: the table is outdented by the first cell's left margin and its border width changes nothing.
    [Theory]
    [InlineData(0)]
    [InlineData(0.5)]
    [InlineData(6)]
    public void Legacy_tables_ignore_the_border_width(double width)
    {
        var box = Sides(width, width);
        var doc = WithSettings(Doc(
            Table([150, 150], null, Row(Cell("a", box, new Thickness(0, 5, 0, 5)), Cell("b", box, new Thickness(0, 5, 0, 5)))),
            Table([150], null, Row(Cell("flush", box, new Thickness(0, 0, 0, 0))))), new DocumentSettings { TableEdgeAtIndent = false });
        var page = Render(doc).Pages.Single();
        Assert.Equal(72, Find(page, "a").X, 6);
        Assert.Equal(72 + 150, Find(page, "b").X, 6);
        Assert.Equal(72, Find(page, "flush").X, 6);
        if (width > 0) Assert.Equal([67.0, 72.0, 217.0, 222.0, 367.0], Verticals(page));
    }
}
