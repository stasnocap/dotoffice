using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Layout;

namespace DotOffice.Tests.Layout;

public class SheetLayoutTests
{
    private static readonly RunProperties Font = new() { FontFamily = "Noto Sans", FontSize = 10 };

    internal static CellStyle Style(Action<CellStyleBuilder>? configure = null)
    {
        var b = new CellStyleBuilder();
        configure?.Invoke(b);
        return b.Build();
    }

    internal sealed class CellStyleBuilder
    {
        public Color? Fill;
        public BorderSet Borders = BorderSet.Empty;
        public HorizontalAlignment Horizontal;
        public CellVerticalAlignment Vertical = CellVerticalAlignment.Bottom;
        public bool Wrap;
        public int Indent;
        public RunProperties Font = SheetLayoutTests.Font;

        public CellStyle Build() => new()
        {
            Font = Font, Fill = Fill, Borders = Borders, HorizontalAlignment = Horizontal, VerticalAlignment = Vertical, Wrap = Wrap, Indent = Indent,
        };
    }

    internal static Cell Text(string text, CellStyle? style = null) => new() { Text = text, Kind = CellValueKind.Text, Style = style ?? Style() };

    internal static Cell Number(string text, CellStyle? style = null) => new() { Text = text, Kind = CellValueKind.Number, Style = style ?? Style() };

    internal static Worksheet Sheet(params (int Row, int Column, Cell Cell)[] cells)
    {
        var sheet = new Worksheet("S");
        foreach (var (r, c, cell) in cells) sheet.Cells[(r, c)] = cell;
        return sheet;
    }

    internal static LayoutDocument Layout(Worksheet sheet, ConversionWarningList? warnings = null)
    {
        var workbook = new Workbook();
        workbook.Sheets.Add(sheet);
        return Layout(workbook, warnings);
    }

    internal static LayoutDocument Layout(Workbook workbook, ConversionWarningList? warnings = null) =>
        new SheetLayoutEngine(new LayoutOptions { Fonts = new FontResolver(new FontCollection().Add(LayoutTestHelpers.Noto)), Warnings = warnings }).Layout(workbook);

    internal static string PageText(LayoutPage page) => string.Concat(page.Items.OfType<TextItem>().Select(t => t.Text));

    [Fact]
    public void Empty_workbook_yields_one_blank_page()
    {
        var doc = Layout(new Workbook());
        Assert.Single(doc.Pages);
        Assert.Empty(doc.Pages[0].Items);
    }

    [Fact]
    public void Empty_sheet_and_hidden_sheets_print_nothing()
    {
        var workbook = new Workbook();
        workbook.Sheets.Add(new Worksheet("Empty"));
        var hidden = Sheet((0, 0, Text("secret")));
        hidden.Visibility = SheetVisibility.Hidden;
        workbook.Sheets.Add(hidden);
        var doc = Layout(workbook);
        Assert.Single(doc.Pages);
        Assert.DoesNotContain("secret", PageText(doc.Pages[0]));
    }

    [Fact]
    public void Cell_text_is_placed_in_its_column_and_row()
    {
        var sheet = Sheet((0, 0, Text("A")), (0, 1, Text("B")), (1, 0, Text("C")));
        sheet.ColumnWidths[0] = 100;
        var page = Layout(sheet).Pages[0];
        var texts = page.Items.OfType<TextItem>().ToDictionary(t => t.Text);
        double left = sheet.PageSetup.Margins.Left, top = sheet.PageSetup.Margins.Top;
        Assert.InRange(texts["A"].X, left, left + 5);
        Assert.InRange(texts["B"].X, left + 100, left + 105);
        Assert.True(texts["C"].Baseline > texts["A"].Baseline + 14);
        Assert.InRange(texts["A"].Baseline, top, top + sheet.DefaultRowHeight);
    }

    [Fact]
    public void Numbers_align_right_text_left_and_booleans_center()
    {
        var sheet = Sheet((0, 0, Number("42")), (1, 0, Text("ab")), (2, 0, new Cell { Text = "TRUE", Kind = CellValueKind.Bool, Style = Style() }));
        sheet.ColumnWidths[0] = 100;
        var items = Layout(sheet).Pages[0].Items.OfType<TextItem>().ToDictionary(t => t.Text);
        double left = sheet.PageSetup.Margins.Left;
        double numberWidth = LayoutTestHelpers.Noto.MeasureText("42", 10);
        Assert.Equal(left + 100 - 2 - numberWidth, items["42"].X, 1);
        Assert.Equal(left + 2, items["ab"].X, 1);
        double boolWidth = LayoutTestHelpers.Noto.MeasureText("TRUE", 10);
        Assert.Equal(left + 50 - boolWidth / 2, items["TRUE"].X, 1);
    }

    [Fact]
    public void Explicit_alignment_and_indent_override_general()
    {
        var sheet = Sheet(
            (0, 0, Number("42", Style(s => s.Horizontal = HorizontalAlignment.Left))),
            (1, 0, Text("ab", Style(s => s.Horizontal = HorizontalAlignment.Right))),
            (2, 0, Text("cd", Style(s => s.Indent = 2))));
        sheet.ColumnWidths[0] = 100;
        var items = Layout(sheet).Pages[0].Items.OfType<TextItem>().ToDictionary(t => t.Text);
        double left = sheet.PageSetup.Margins.Left;
        Assert.Equal(left + 2, items["42"].X, 1);
        Assert.Equal(left + 100 - 2 - LayoutTestHelpers.Noto.MeasureText("ab", 10), items["ab"].X, 1);
        Assert.Equal(left + 2 + 13.5, items["cd"].X, 1);
    }

    [Fact]
    public void Fills_and_borders_are_drawn_around_the_cell()
    {
        var border = new Border(BorderStyle.Single, 1, new Color(255, 0, 0));
        var style = Style(s =>
        {
            s.Fill = new Color(0, 255, 0);
            s.Borders = new BorderSet { Top = border, Bottom = border, Left = border, Right = border };
        });
        var sheet = Sheet((0, 0, Text("x", style)), (1, 1, new Cell { Style = Style(s => s.Fill = new Color(1, 2, 3)) }));
        var page = Layout(sheet).Pages[0];
        var fill = page.Items.OfType<RectItem>().First(r => r.Fill == new Color(0, 255, 0));
        Assert.Equal(sheet.PageSetup.Margins.Left, fill.X, 3);
        Assert.Equal(sheet.ColumnWidth(0), fill.Width, 3);
        Assert.Equal(sheet.RowHeight(0), fill.Height, 3);
        Assert.Equal(4, page.Items.OfType<LineItem>().Count(l => l.Color == new Color(255, 0, 0)));
        Assert.Contains(page.Items.OfType<RectItem>(), r => r.Fill == new Color(1, 2, 3)); // styled empty cell
    }

    [Fact]
    public void Text_is_clipped_at_the_cell_edge()
    {
        var sheet = Sheet((0, 0, Text("a very long text indeed")), (0, 1, Text("next")));
        sheet.ColumnWidths[0] = 40;
        var page = Layout(sheet).Pages[0];
        double right = sheet.PageSetup.Margins.Left + 40;
        foreach (var t in page.Items.OfType<TextItem>().Where(t => t.Text != "next"))
            Assert.True(t.X + LayoutTestHelpers.Noto.MeasureText(t.Text, 10) <= right + 0.1, $"{t.Text} exceeds the cell");
        Assert.Contains("a ver", PageText(page).Replace(" ", " "));
    }

    [Fact]
    public void Numbers_that_do_not_fit_show_hashes()
    {
        var sheet = Sheet((0, 0, Number("123456789012")));
        sheet.ColumnWidths[0] = 30;
        var text = PageText(Layout(sheet).Pages[0]);
        Assert.NotEmpty(text);
        Assert.All(text, ch => Assert.Equal('#', ch));
        Assert.True(LayoutTestHelpers.Noto.MeasureText(text, 10) <= 30);
    }

    [Fact]
    public void Wrapped_text_grows_the_row_and_breaks_into_lines()
    {
        var sheet = Sheet((0, 0, Text("alpha beta gamma delta epsilon", Style(s => s.Wrap = true))), (1, 0, Text("below")));
        sheet.ColumnWidths[0] = 50;
        var items = Layout(sheet).Pages[0].Items.OfType<TextItem>().ToList();
        var baselines = items.Where(t => t.Text != "below").Select(t => Math.Round(t.Baseline, 1)).Distinct().ToList();
        Assert.True(baselines.Count >= 3);
        var below = items.First(t => t.Text == "below");
        Assert.True(below.Baseline > baselines.Max());
    }

    [Fact]
    public void Newlines_break_wrapped_text_and_vanish_in_unwrapped_text()
    {
        var wrapped = Sheet((0, 0, Text("one\ntwo", Style(s => s.Wrap = true))));
        wrapped.ColumnWidths[0] = 200;
        var lines = Layout(wrapped).Pages[0].Items.OfType<TextItem>().Select(t => Math.Round(t.Baseline, 1)).Distinct();
        Assert.Equal(2, lines.Count());
        var plain = Sheet((0, 0, Text("one\ntwo")));
        plain.ColumnWidths[0] = 200;
        Assert.Single(Layout(plain).Pages[0].Items.OfType<TextItem>().Select(t => Math.Round(t.Baseline, 1)).Distinct());
    }

    [Fact]
    public void Vertical_alignment_moves_the_text_within_the_row()
    {
        var sheet = Sheet((0, 0, Text("T", Style(s => s.Vertical = CellVerticalAlignment.Top))),
            (0, 1, Text("C", Style(s => s.Vertical = CellVerticalAlignment.Center))),
            (0, 2, Text("B")));
        sheet.RowHeights[0] = 60;
        var items = Layout(sheet).Pages[0].Items.OfType<TextItem>().ToDictionary(t => t.Text);
        Assert.True(items["T"].Baseline < items["C"].Baseline);
        Assert.True(items["C"].Baseline < items["B"].Baseline);
    }

    [Fact]
    public void Merged_cells_draw_across_the_range_and_hide_covered_cells()
    {
        var sheet = Sheet((0, 0, Text("Title", Style(s => { s.Fill = new Color(9, 9, 9); s.Horizontal = HorizontalAlignment.Center; }))),
            (0, 1, Text("hidden")));
        sheet.ColumnWidths[0] = 60;
        sheet.ColumnWidths[1] = 60;
        sheet.Merges.Add(new CellRange(0, 0, 0, 1));
        var page = Layout(sheet).Pages[0];
        Assert.DoesNotContain("hidden", PageText(page));
        Assert.Equal(120, page.Items.OfType<RectItem>().Single().Width, 3);
        var title = page.Items.OfType<TextItem>().Single();
        double width = LayoutTestHelpers.Noto.MeasureText("Title", 10);
        Assert.Equal(sheet.PageSetup.Margins.Left + 60 - width / 2, title.X, 1);
    }

    [Fact]
    public void Hidden_rows_and_columns_are_not_printed()
    {
        var sheet = Sheet((0, 0, Text("keep")), (0, 1, Text("hiddencol")), (1, 0, Text("hiddenrow")), (2, 0, Text("end")));
        sheet.HiddenColumns.Add(1);
        sheet.HiddenRows.Add(1);
        var page = Layout(sheet).Pages[0];
        var text = PageText(page);
        Assert.Contains("keep", text);
        Assert.Contains("end", text);
        Assert.DoesNotContain("hidden", text);
        var texts = page.Items.OfType<TextItem>().ToDictionary(t => t.Text);
        Assert.True(texts["end"].Baseline - texts["keep"].Baseline < 20); // adjacent rows
    }

    [Fact]
    public void Wide_sheets_split_into_column_bands_that_go_down_then_over()
    {
        var sheet = Sheet();
        for (int r = 0; r < 100; r++)
            for (int c = 0; c < 12; c++)
                sheet.Cells[(r, c)] = Text($"r{r}c{c}");
        for (int c = 0; c < 12; c++) sheet.ColumnWidths[c] = 100;
        var doc = Layout(sheet);
        // 504 pt wide text area → 5 columns per band; 684 pt tall → 45 rows per band.
        Assert.Equal(3 * 3, doc.Pages.Count);
        Assert.Contains("r0c0", PageText(doc.Pages[0]));
        Assert.Contains("r45c0", PageText(doc.Pages[1]));
        Assert.Contains("r0c5", PageText(doc.Pages[3]));

        sheet.PageSetup.PageOrderOverThenDown = true;
        var over = Layout(sheet);
        Assert.Contains("r0c5", PageText(over.Pages[1]));
        Assert.Contains("r45c0", PageText(over.Pages[3]));
    }

    [Fact]
    public void Print_area_limits_the_printed_cells()
    {
        var sheet = Sheet((0, 0, Text("in")), (5, 5, Text("out")));
        sheet.PageSetup.PrintArea = new CellRange(0, 0, 2, 2);
        var text = PageText(Layout(sheet).Pages[0]);
        Assert.Contains("in", text);
        Assert.DoesNotContain("out", text);
    }

    [Fact]
    public void Gridlines_are_drawn_only_when_requested()
    {
        var sheet = Sheet((0, 0, Text("a")), (1, 1, Text("b")));
        Assert.Empty(Layout(sheet).Pages[0].Items.OfType<LineItem>());
        sheet.PageSetup.PrintGridLines = true;
        Assert.NotEmpty(Layout(sheet).Pages[0].Items.OfType<LineItem>());
    }

    [Fact]
    public void Scale_and_landscape_apply_to_the_page()
    {
        var sheet = Sheet((0, 0, Text("x")));
        sheet.PageSetup.Scale = 0.5;
        sheet.PageSetup.PageWidth = 792;
        sheet.PageSetup.PageHeight = 612;
        var doc = Layout(sheet);
        Assert.Equal(792, doc.Pages[0].Width);
        Assert.Equal(5, doc.Pages[0].Items.OfType<TextItem>().Single().FontSize);
    }

    [Fact]
    public void Centering_puts_the_grid_in_the_middle_of_the_page()
    {
        var sheet = Sheet((0, 0, new Cell { Style = Style(s => s.Fill = new Color(1, 1, 1)) }));
        sheet.PageSetup.CenterHorizontally = true;
        sheet.PageSetup.CenterVertically = true;
        var fill = Layout(sheet).Pages[0].Items.OfType<RectItem>().Single();
        Assert.Equal((612 - fill.Width) / 2, fill.X, 1);
        Assert.Equal((792 - fill.Height) / 2, fill.Y, 1);
    }

    [Fact]
    public void Carries_the_workbook_title_and_author()
    {
        var workbook = new Workbook { Title = "T", Author = "A" };
        var doc = Layout(workbook);
        Assert.Equal("T", doc.Title);
        Assert.Equal("A", doc.Author);
    }
}
