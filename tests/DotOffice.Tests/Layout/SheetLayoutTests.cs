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
        public double Rotation;
        public bool Stacked;
        public GradientFill? Gradient;
        public RunProperties Font = SheetLayoutTests.Font;

        public CellStyle Build() => new()
        {
            Font = Font, Fill = Fill, Borders = Borders, HorizontalAlignment = Horizontal, VerticalAlignment = Vertical, Wrap = Wrap, Indent = Indent,
            Rotation = Rotation, Stacked = Stacked, Gradient = Gradient,
        };
    }

    internal static Cell Rich(CellStyle style, params (string Text, RunProperties Font)[] runs) => new()
    {
        Text = string.Concat(runs.Select(r => r.Text)), Kind = CellValueKind.Text, Style = style,
        Runs = runs.Select(r => new CellTextRun(r.Text, r.Font)).ToList(),
    };

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

public class SheetOverflowTests
{
    [Fact]
    public void Text_wider_than_a_wide_column_still_extends_the_printed_range_past_the_last_used_column()
    {
        const string text = "thirty characters of overflow.";
        var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text(text)));
        sheet.ColumnWidths[0] = 100; // 30 characters at 10 pt need about 135 pt: a length-based lower bound would call this a fit
        Assert.Contains(text, SheetLayoutTests.PageText(SheetLayoutTests.Layout(sheet).Pages[0]));
    }

    private static Cell Text(string text, HorizontalAlignment align = HorizontalAlignment.General) =>
        SheetLayoutTests.Text(text, SheetLayoutTests.Style(s => s.Horizontal = align));

    private static double Right(TextItem t) => t.X + LayoutTestHelpers.Noto.MeasureText(t.Text, t.FontSize);

    private const string Long = "overflowing text spills";

    [Fact]
    public void Left_aligned_text_spills_into_empty_neighbors()
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, Text(Long)));
        sheet.ColumnWidths[0] = 40;
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().ToList();
        Assert.Equal(Long, string.Concat(items.Select(t => t.Text)));
    }

    [Fact]
    public void Text_stops_at_the_first_non_empty_neighbor()
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, Text(Long)), (0, 2, Text("block")));
        sheet.ColumnWidths[0] = 40;
        sheet.ColumnWidths[1] = 40;
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        double limit = sheet.PageSetup.Margins.Left + 80;
        foreach (var t in page.Items.OfType<TextItem>().Where(t => t.Text != "block")) Assert.True(Right(t) <= limit + 0.1);
        Assert.True(page.Items.OfType<TextItem>().Where(t => t.Text != "block").Max(Right) > sheet.PageSetup.Margins.Left + 45);
    }

    [Fact]
    public void A_neighbor_with_only_a_fill_does_not_block_the_text()
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, Text(Long)), (0, 1, new Cell { Style = SheetLayoutTests.Style(s => s.Fill = new Color(5, 5, 5)) }));
        sheet.ColumnWidths[0] = 40;
        var text = string.Concat(SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().Select(t => t.Text));
        Assert.Equal(Long, text);
    }

    [Fact]
    public void Right_aligned_text_spills_to_the_left()
    {
        var sheet = SheetLayoutTests.Sheet((0, 1, Text(Long, HorizontalAlignment.Right)));
        sheet.ColumnWidths[0] = 200;
        sheet.ColumnWidths[1] = 40;
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().ToList();
        Assert.Equal(Long, string.Concat(items.Select(t => t.Text)));
        Assert.True(items.Min(t => t.X) < sheet.PageSetup.Margins.Left + 200);
        Assert.True(items.Max(Right) <= sheet.PageSetup.Margins.Left + 240 + 0.1);
    }

    [Fact]
    public void Centered_text_spills_both_ways_but_is_limited_by_the_closer_blocker()
    {
        var sheet = SheetLayoutTests.Sheet((0, 1, Text(Long, HorizontalAlignment.Center)), (0, 2, Text("x")));
        sheet.ColumnWidths[0] = 100;
        sheet.ColumnWidths[1] = 40;
        sheet.ColumnWidths[2] = 40;
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().Where(t => t.Text != "x").ToList();
        double left = sheet.PageSetup.Margins.Left + 100;
        Assert.True(items.Max(Right) <= left + 40 + 0.1); // the right neighbor is occupied, so no room on that side
        Assert.True(items.Min(t => t.X) >= left - 0.1);
    }

    [Fact]
    public void Wrapped_text_and_numbers_never_spill()
    {
        var wrapped = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text(Long, SheetLayoutTests.Style(s => s.Wrap = true))));
        wrapped.ColumnWidths[0] = 40;
        wrapped.ColumnWidths[1] = 200;
        foreach (var t in SheetLayoutTests.Layout(wrapped).Pages[0].Items.OfType<TextItem>())
            Assert.True(Right(t) <= wrapped.PageSetup.Margins.Left + 40 + 0.1);

        var number = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Number("123456789")));
        number.ColumnWidths[0] = 30;
        number.ColumnWidths[1] = 200;
        Assert.All(SheetLayoutTests.PageText(SheetLayoutTests.Layout(number).Pages[0]), ch => Assert.Equal('#', ch));
    }

    [Fact]
    public void Spill_stops_at_the_page_edge()
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, Text(new string('w', 200))));
        sheet.ColumnWidths[0] = 40;
        sheet.ColumnWidths[1] = 40;
        var doc = SheetLayoutTests.Layout(sheet);
        Assert.Single(doc.Pages);
        var items = doc.Pages[0].Items.OfType<TextItem>().ToList();
        Assert.True(items.Max(Right) <= 612 - sheet.PageSetup.Margins.Right + 0.1);
        Assert.True(items.Max(Right) > sheet.PageSetup.Margins.Left + 80); // spilled past the original two columns
    }
}

public class SheetPaginationTests
{
    private static Worksheet Grid(int rows, int columns, double columnWidth = 100)
    {
        var sheet = SheetLayoutTests.Sheet();
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < columns; c++)
                sheet.Cells[(r, c)] = SheetLayoutTests.Text($"r{r}c{c}");
        for (int c = 0; c < columns; c++) sheet.ColumnWidths[c] = columnWidth;
        return sheet;
    }

    private static string[] Pages(Worksheet sheet) =>
        SheetLayoutTests.Layout(sheet).Pages.Select(p => string.Join(" ", p.Items.OfType<TextItem>().Select(t => t.Text))).ToArray();

    [Fact]
    public void Manual_row_breaks_start_a_new_page()
    {
        var sheet = Grid(10, 1);
        sheet.PageSetup.RowBreaks.Add(3); // break after the fourth row
        var pages = Pages(sheet);
        Assert.Equal(2, pages.Length);
        Assert.Contains("r3c0", pages[0]);
        Assert.DoesNotContain("r4c0", pages[0]);
        Assert.StartsWith("r4c0", pages[1]);
    }

    [Fact]
    public void Manual_column_breaks_start_a_new_page()
    {
        var sheet = Grid(2, 4);
        sheet.PageSetup.ColumnBreaks.Add(1);
        var pages = Pages(sheet);
        Assert.Equal(2, pages.Length);
        Assert.Contains("r1c1", pages[0]);
        Assert.DoesNotContain("r0c2", pages[0]);
        Assert.Contains("r0c2", pages[1]);
    }

    [Fact]
    public void A_break_before_a_hidden_row_applies_to_the_next_visible_row()
    {
        var sheet = Grid(6, 1);
        sheet.HiddenRows.Add(3);
        sheet.PageSetup.RowBreaks.Add(2);
        var pages = Pages(sheet);
        Assert.Equal(2, pages.Length);
        Assert.StartsWith("r4c0", pages[1]);
    }

    [Fact]
    public void Title_rows_repeat_on_every_following_page()
    {
        var sheet = Grid(100, 1);
        sheet.PageSetup.TitleRows = (0, 1);
        var pages = Pages(sheet);
        Assert.True(pages.Length >= 2);
        foreach (var page in pages)
        {
            Assert.StartsWith("r0c0 r1c0", page);
        }
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(pages[0], @"\br0c0\b"));
        Assert.Contains("r99c0", pages[^1]);
        // every data row appears exactly once
        var all = string.Join(" ", pages);
        for (int r = 2; r < 100; r++) Assert.Single(System.Text.RegularExpressions.Regex.Matches(all, $@"\br{r}c0\b"));
    }

    [Fact]
    public void Title_columns_repeat_on_every_column_band()
    {
        var sheet = Grid(2, 12);
        sheet.PageSetup.TitleColumns = (0, 0);
        var pages = Pages(sheet);
        Assert.True(pages.Length >= 2);
        foreach (var page in pages) Assert.Contains("r0c0", page);
        Assert.Contains("r0c11", pages[^1]);
    }

    [Fact]
    public void Fit_to_one_page_wide_shrinks_the_sheet()
    {
        var sheet = Grid(3, 12); // 1200 pt wide on a 504 pt text area
        sheet.PageSetup.FitToWidth = 1;
        sheet.PageSetup.FitToHeight = 0;
        var doc = SheetLayoutTests.Layout(sheet);
        Assert.Single(doc.Pages);
        Assert.Equal(10 * 504.0 / 1200, doc.Pages[0].Items.OfType<TextItem>().First().FontSize, 0.5);
        var text = SheetLayoutTests.PageText(doc.Pages[0]);
        Assert.Contains("r0c0", text);
        Assert.Contains("r2c11", text);
        foreach (var t in doc.Pages[0].Items.OfType<TextItem>())
            Assert.True(t.X + 1 <= 612 - sheet.PageSetup.Margins.Right + 0.5);
    }

    [Fact]
    public void Fit_to_page_never_enlarges_and_ignores_manual_breaks()
    {
        var sheet = Grid(3, 2);
        sheet.PageSetup.FitToWidth = 1;
        sheet.PageSetup.FitToHeight = 1;
        sheet.PageSetup.RowBreaks.Add(0);
        var doc = SheetLayoutTests.Layout(sheet);
        Assert.Single(doc.Pages);
        Assert.Equal(10, doc.Pages[0].Items.OfType<TextItem>().First().FontSize);
    }

    [Fact]
    public void Fit_to_two_pages_tall_produces_at_most_two_pages()
    {
        var sheet = Grid(120, 1);
        sheet.PageSetup.FitToWidth = 0;
        sheet.PageSetup.FitToHeight = 2;
        var doc = SheetLayoutTests.Layout(sheet);
        Assert.True(doc.Pages.Count <= 2);
        Assert.True(doc.Pages[0].Items.OfType<TextItem>().First().FontSize < 10);
    }
}

public class SheetHeaderFooterTests
{
    private static Worksheet WithText(string name, int rows = 1)
    {
        var sheet = new Worksheet(name);
        for (int r = 0; r < rows; r++) sheet.Cells[(r, 0)] = SheetLayoutTests.Text($"{name}{r}");
        return sheet;
    }

    private static string Joined(LayoutPage page, double? aboveY = null, double? belowY = null) =>
        string.Concat(page.Items.OfType<TextItem>()
            .Where(t => (aboveY is null || t.Baseline < aboveY) && (belowY is null || t.Baseline > belowY)).Select(t => t.Text));

    [Fact]
    public void Fields_are_substituted_and_page_numbers_continue_across_sheets()
    {
        var workbook = new Workbook { Title = "Book" };
        var first = WithText("One", 80); // two pages
        first.PageSetup.Header = new HeaderFooterText { Odd = new HeaderFooterParts("", "&A", "") };
        first.PageSetup.Footer = new HeaderFooterText { Odd = new HeaderFooterParts("", "Page &P of &N", "") };
        var second = WithText("Two");
        second.PageSetup.Header = first.PageSetup.Header;
        second.PageSetup.Footer = first.PageSetup.Footer;
        workbook.Sheets.Add(first);
        workbook.Sheets.Add(second);
        var doc = SheetLayoutTests.Layout(workbook);
        Assert.Equal(3, doc.Pages.Count);
        for (int i = 0; i < 3; i++)
        {
            var footer = Joined(doc.Pages[i], belowY: 740);
            Assert.Equal($"Page {i + 1} of 3", footer);
        }
        Assert.Contains("One", Joined(doc.Pages[1], aboveY: 40));
        Assert.Contains("Two", Joined(doc.Pages[2], aboveY: 40));
    }

    [Fact]
    public void Left_center_and_right_parts_are_anchored_to_the_margins()
    {
        var sheet = WithText("S");
        sheet.PageSetup.Footer = new HeaderFooterText { Odd = new HeaderFooterParts("L", "C", "R") };
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().Where(t => t.Baseline > 740).ToDictionary(t => t.Text);
        var m = sheet.PageSetup.Margins;
        Assert.Equal(m.Left, items["L"].X, 1);
        Assert.True(Math.Abs(items["C"].X + items["C"].FontSize * 0.3 - 306) < 6);
        Assert.True(items["R"].X + 10 > 612 - m.Right - 10 && items["R"].X < 612 - m.Right);
    }

    [Fact]
    public void Footers_sit_above_the_bottom_edge_and_headers_below_the_top_edge()
    {
        var sheet = WithText("S");
        sheet.PageSetup.Header = new HeaderFooterText { Odd = new HeaderFooterParts("", "head", "") };
        sheet.PageSetup.Footer = new HeaderFooterText { Odd = new HeaderFooterParts("", "foot", "") };
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().ToList();
        var head = items.First(t => t.Text == "head");
        var foot = items.First(t => t.Text == "foot");
        Assert.InRange(head.Baseline, sheet.PageSetup.HeaderDistance, sheet.PageSetup.HeaderDistance + 15);
        Assert.InRange(foot.Baseline, 792 - sheet.PageSetup.FooterDistance - 15, 792 - sheet.PageSetup.FooterDistance);
    }

    [Fact]
    public void Font_codes_change_size_and_weight_and_double_ampersand_is_literal()
    {
        var sheet = WithText("S");
        sheet.PageSetup.Header = new HeaderFooterText { Odd = new HeaderFooterParts("", "&\"Arial,Bold\"&20Big&B plain &&", "") };
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().Where(t => t.Baseline < 52).ToList();
        var big = items.First(t => t.Text == "Big");
        Assert.Equal(20, big.FontSize);
        Assert.Contains("&", string.Concat(items.Select(t => t.Text)));
        Assert.Contains("plain", string.Concat(items.Select(t => t.Text)));
    }

    [Fact]
    public void First_page_and_even_page_variants_are_used()
    {
        var sheet = WithText("S", 80);
        sheet.PageSetup.Header = new HeaderFooterText
        {
            Odd = new HeaderFooterParts("", "odd", ""),
            Even = new HeaderFooterParts("", "even", ""),
            First = new HeaderFooterParts("", "first", ""),
        };
        var doc = SheetLayoutTests.Layout(sheet);
        Assert.Equal(2, doc.Pages.Count);
        Assert.Equal("first", Joined(doc.Pages[0], aboveY: 40));
        Assert.Equal("even", Joined(doc.Pages[1], aboveY: 40));
    }

    [Fact]
    public void Multiple_lines_stack_downwards_in_a_header()
    {
        var sheet = WithText("S");
        sheet.PageSetup.Header = new HeaderFooterText { Odd = new HeaderFooterParts("", "one\ntwo", "") };
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().ToDictionary(t => t.Text);
        Assert.True(items["two"].Baseline > items["one"].Baseline + 10);
    }

    [Fact]
    public void Header_and_footer_are_not_scaled_with_the_sheet()
    {
        var sheet = WithText("S");
        sheet.PageSetup.Scale = 0.5;
        sheet.PageSetup.Header = new HeaderFooterText { Odd = new HeaderFooterParts("", "head", "") };
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().ToDictionary(t => t.Text);
        Assert.Equal(11, items["head"].FontSize);
        Assert.Equal(5, items["S0"].FontSize);
    }
}

public class SheetRichTextTests
{
    private static readonly RunProperties Base = new() { FontFamily = "Noto Sans", FontSize = 10 };

    [Fact]
    public void Each_run_is_drawn_with_its_own_size_weight_and_color()
    {
        var red = Base with { FontSize = 16, Color = new Color(255, 0, 0) };
        var cell = SheetLayoutTests.Rich(SheetLayoutTests.Style(), ("Big", red), (" small", Base with { Underline = UnderlineStyle.Single }));
        var sheet = SheetLayoutTests.Sheet((0, 0, cell));
        sheet.ColumnWidths[0] = 200;
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        var big = page.Items.OfType<TextItem>().Single(t => t.Text == "Big");
        var small = page.Items.OfType<TextItem>().Single(t => t.Text.Contains("small"));
        Assert.Equal((16.0, new Color(255, 0, 0)), (big.FontSize, big.Color));
        Assert.Equal(10.0, small.FontSize);
        Assert.True(small.X >= big.X + LayoutTestHelpers.Noto.MeasureText("Big", 16) - 0.5);
        Assert.Equal(big.Baseline, small.Baseline, 3); // same baseline
        Assert.Contains(page.Items.OfType<LineItem>(), l => l.Y1 > small.Baseline); // underline under the small run only
    }

    [Fact]
    public void Superscript_runs_are_smaller_and_raised()
    {
        var cell = SheetLayoutTests.Rich(SheetLayoutTests.Style(), ("m", Base), ("2", Base with { VerticalPosition = VerticalPosition.Superscript }));
        var items = SheetLayoutTests.Layout(SheetLayoutTests.Sheet((0, 0, cell))).Pages[0].Items.OfType<TextItem>().ToList();
        var m = items.Single(t => t.Text == "m");
        var two = items.Single(t => t.Text == "2");
        Assert.True(two.FontSize < m.FontSize);
        Assert.True(two.Baseline < m.Baseline);
    }

    [Fact]
    public void The_largest_run_sets_the_autofit_row_height()
    {
        var cell = SheetLayoutTests.Rich(SheetLayoutTests.Style(), ("a", Base), ("b", Base with { FontSize = 30 }));
        var sheet = SheetLayoutTests.Sheet((0, 0, cell), (1, 0, SheetLayoutTests.Text("below")));
        var items = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<TextItem>().ToList();
        double rowBottom = items.Single(t => t.Text == "below").Baseline - sheet.PageSetup.Margins.Top;
        Assert.True(rowBottom > 30 * 15.0 / 11, $"row 0 should be about 41 pt tall, 'below' sits {rowBottom} pt down");
    }

    [Fact]
    public void Rich_text_wraps_across_runs_and_still_overflows_when_not_wrapping()
    {
        var bold = Base with { Bold = true };
        var wrapStyle = SheetLayoutTests.Style(s => s.Wrap = true);
        var wrapped = SheetLayoutTests.Rich(wrapStyle, ("alpha beta ", Base), ("gamma delta epsilon", bold));
        var sheet = SheetLayoutTests.Sheet((0, 0, wrapped));
        sheet.ColumnWidths[0] = 60;
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        var baselines = page.Items.OfType<TextItem>().Select(t => Math.Round(t.Baseline, 2)).Distinct().ToList();
        Assert.True(baselines.Count >= 3, "text should wrap onto several lines");

        const string tail = " keeps going until a value stops it";
        var overflow = SheetLayoutTests.Rich(SheetLayoutTests.Style(), ("Start", bold), (tail, Base));
        var sheet2 = SheetLayoutTests.Sheet((0, 0, overflow), (0, 3, SheetLayoutTests.Text("STOP")));
        var text = SheetLayoutTests.PageText(SheetLayoutTests.Layout(sheet2).Pages[0]);
        Assert.Contains("Start", text);
        Assert.Contains("STOP", text);
        Assert.DoesNotContain("stops it", text); // cut at the occupied neighbour
        Assert.Contains("keeps", text); // but spilled into the empty cells first
    }

    [Fact]
    public void Newlines_inside_runs_start_new_lines_in_wrapping_cells()
    {
        var cell = SheetLayoutTests.Rich(SheetLayoutTests.Style(s => s.Wrap = true), ("one\n", Base), ("two", Base with { Italic = true }));
        var items = SheetLayoutTests.Layout(SheetLayoutTests.Sheet((0, 0, cell))).Pages[0].Items.OfType<TextItem>().ToList();
        Assert.True(items.Single(t => t.Text == "two").Baseline > items.Single(t => t.Text == "one").Baseline + 5);
    }
}

public class SheetRotationTests
{
    private static (LayoutPage Page, Worksheet Sheet) Run(Cell cell, Action<Worksheet>? configure = null)
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, cell));
        sheet.ColumnWidths[0] = 100;
        configure?.Invoke(sheet);
        return (SheetLayoutTests.Layout(sheet).Pages[0], sheet);
    }

    private static Cell Rotated(string text, double degrees, Action<SheetLayoutTests.CellStyleBuilder>? configure = null) =>
        SheetLayoutTests.Text(text, SheetLayoutTests.Style(s => { s.Rotation = degrees; configure?.Invoke(s); }));

    [Fact]
    public void Rotated_text_is_drawn_inside_a_group_turned_about_its_center_and_clipped_to_the_cell()
    {
        var (page, sheet) = Run(Rotated("Hello", 45), s => s.RowHeights[0] = 80);
        var push = page.Items.OfType<PushItem>().Single();
        Assert.Equal(45, push.Rotation);
        double left = sheet.PageSetup.Margins.Left, top = sheet.PageSetup.Margins.Top;
        Assert.Equal((left, top, 100.0, 80.0), push.Clip!.Value);
        Assert.InRange(push.CenterX, left, left + 100);
        Assert.InRange(push.CenterY, top, top + 80);
        var items = page.Items.ToList();
        int pushAt = items.IndexOf(push), popAt = items.FindIndex(i => i is PopItem);
        Assert.True(pushAt < items.FindIndex(i => i is TextItem) && items.FindIndex(i => i is TextItem) < popAt);
    }

    [Fact]
    public void Unrotated_text_creates_no_groups()
    {
        var (page, _) = Run(SheetLayoutTests.Text("Hello"));
        Assert.Empty(page.Items.OfType<PushItem>());
    }

    [Fact]
    public void Rotating_to_90_degrees_grows_an_autofit_row_to_the_text_length()
    {
        const string text = "A fairly long heading";
        var (page, sheet) = Run(Rotated(text, 90));
        double length = LayoutTestHelpers.Noto.MeasureText(text, 10);
        var push = page.Items.OfType<PushItem>().Single();
        Assert.Equal(length, push.Clip!.Value.Height, 0); // the row is as tall as the text is long (the cell is clipped to it)
        Assert.True(push.Clip!.Value.Height > sheet.DefaultRowHeight * 3);
    }

    [Fact]
    public void An_explicit_row_height_is_kept_and_text_is_clipped_not_enlarged()
    {
        var (page, _) = Run(Rotated("A fairly long heading", 90), s => s.RowHeights[0] = 20);
        Assert.Equal(20.0, page.Items.OfType<PushItem>().Single().Clip!.Value.Height, 3);
    }

    [Fact]
    public void Rotation_of_45_needs_width_times_sine_plus_height_times_cosine()
    {
        const string text = "Sample text";
        var (page, _) = Run(Rotated(text, 45));
        double w = LayoutTestHelpers.Noto.MeasureText(text, 10), h = 10 * 15.0 / 11;
        double expected = (w + h) * Math.Sin(Math.PI / 4);
        Assert.Equal(expected, page.Items.OfType<PushItem>().Single().Clip!.Value.Height, 1);
    }

    [Fact]
    public void Clockwise_rotation_uses_a_negative_angle()
    {
        var (page, _) = Run(Rotated("Down", -90));
        Assert.Equal(-90, page.Items.OfType<PushItem>().Single().Rotation);
    }

    [Fact]
    public void Rotated_text_is_placed_by_its_bounding_box_per_alignment()
    {
        Cell At(HorizontalAlignment h, CellVerticalAlignment v) => Rotated("Hi", 90, s => { s.Horizontal = h; s.Vertical = v; });
        var (leftTop, sheet) = Run(At(HorizontalAlignment.Left, CellVerticalAlignment.Top), s => s.RowHeights[0] = 60);
        var (rightBottom, _) = Run(At(HorizontalAlignment.Right, CellVerticalAlignment.Bottom), s => s.RowHeights[0] = 60);
        var a = leftTop.Items.OfType<PushItem>().Single();
        var b = rightBottom.Items.OfType<PushItem>().Single();
        double left = sheet.PageSetup.Margins.Left, top = sheet.PageSetup.Margins.Top;
        Assert.True(a.CenterX < left + 20 && b.CenterX > left + 80);
        Assert.True(a.CenterY < top + 20 && b.CenterY > top + 40);
    }

    [Fact]
    public void Stacked_text_puts_each_character_on_its_own_upright_line()
    {
        var cell = SheetLayoutTests.Text("ABC", SheetLayoutTests.Style(s => s.Stacked = true));
        var (page, sheet) = Run(cell);
        var texts = page.Items.OfType<TextItem>().ToList();
        Assert.Equal(["A", "B", "C"], texts.Select(t => t.Text));
        Assert.True(texts[1].Baseline > texts[0].Baseline + 10 && texts[2].Baseline > texts[1].Baseline + 10);
        Assert.Equal(0, page.Items.OfType<PushItem>().Single().Rotation);
        Assert.Equal(3 * 10 * 15.0 / 11, page.Items.OfType<PushItem>().Single().Clip!.Value.Height, 1); // the row grew to fit three lines
        Assert.Equal(sheet.PageSetup.Margins.Top, page.Items.OfType<PushItem>().Single().Clip!.Value.Y, 3);
    }

    [Fact]
    public void Wrapped_rotated_text_breaks_at_the_cell_width_over_the_cosine()
    {
        var (page, _) = Run(Rotated("alpha beta gamma delta epsilon zeta eta theta", 30, s => s.Wrap = true), s => s.ColumnWidths[0] = 60);
        var lines = page.Items.OfType<TextItem>().Select(t => Math.Round(t.Baseline, 2)).Distinct().Count();
        Assert.True(lines >= 2, "the text should wrap onto several lines");
    }

    [Fact]
    public void Rotated_cells_do_not_spill_into_neighbors()
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, Rotated("a very long rotated heading that cannot spill", 15)), (0, 3, SheetLayoutTests.Text("x")));
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        Assert.Single(page.Items.OfType<PushItem>());
        Assert.Equal(sheet.DefaultColumnWidth, page.Items.OfType<PushItem>().Single().Clip!.Value.Width, 3); // clipped to its own column
    }
}

public class SheetGradientTests
{
    private static GradientFill Gradient(GradientKind kind, double degree = 0, double focusX = 0, double focusY = 0) => new()
    {
        Kind = kind, Degree = degree, Left = focusX, Right = focusX, Top = focusY, Bottom = focusY,
        Stops = [new GradientStop(0, new Color(255, 0, 0)), new GradientStop(1, new Color(0, 0, 255))],
    };

    private static LayoutPage Page(GradientFill gradient, Action<Worksheet>? configure = null)
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text("x", SheetLayoutTests.Style(s => s.Gradient = gradient))));
        sheet.ColumnWidths[0] = 100;
        sheet.RowHeights[0] = 40;
        configure?.Invoke(sheet);
        return SheetLayoutTests.Layout(sheet).Pages[0];
    }

    [Fact]
    public void A_linear_gradient_is_one_gradient_item_over_the_cell_below_its_text()
    {
        var page = Page(Gradient(GradientKind.Linear, 90));
        var item = page.Items.OfType<GradientItem>().Single();
        Assert.Equal((100.0, 40.0, 90.0), (item.Width, item.Height, item.Degree));
        Assert.Equal(2, item.Stops.Count);
        Assert.True(page.Items.IndexOf(item) < page.Items.FindIndex(i => i is TextItem));
        Assert.Empty(page.Items.OfType<RectItem>());
    }

    [Fact]
    public void A_gradient_replaces_the_solid_fill()
    {
        var style = SheetLayoutTests.Style(s => { s.Fill = Color.White; s.Gradient = Gradient(GradientKind.Linear); });
        var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text("x", style)));
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        Assert.Single(page.Items.OfType<GradientItem>());
        Assert.Empty(page.Items.OfType<RectItem>());
    }

    [Fact]
    public void A_path_gradient_is_drawn_as_nested_rectangles_growing_from_the_focus_inside_a_clip()
    {
        var page = Page(Gradient(GradientKind.Path, focusX: 0.5, focusY: 0.5));
        var items = page.Items;
        var rects = items.OfType<RectItem>().ToList();
        Assert.True(rects.Count >= 16);
        Assert.Contains(items, i => i is PushItem { Clip: not null });
        Assert.Contains(items, i => i is PopItem);
        // Painted outermost first, so each following rectangle is smaller and nearer the focus color (red).
        Assert.True(rects[0].Width > rects[^1].Width);
        Assert.True(rects[^1].Fill.R > rects[0].Fill.R && rects[^1].Fill.B < rects[0].Fill.B);
        Assert.Equal(rects[^1].X + rects[^1].Width / 2, rects[0].X + rects[0].Width / 2, 3); // concentric around the cell center
    }

    [Fact]
    public void A_path_gradient_from_the_corner_starts_at_that_corner()
    {
        var page = Page(Gradient(GradientKind.Path, 0, 0));
        var innermost = page.Items.OfType<RectItem>().Last();
        double left = new SheetPageSetup().Margins.Left, top = new SheetPageSetup().Margins.Top;
        Assert.Equal((left, top), (innermost.X + innermost.Width / 2, innermost.Y + innermost.Height / 2));
    }

    [Fact]
    public void Gradients_follow_the_page_scale()
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text("x", SheetLayoutTests.Style(s => s.Gradient = Gradient(GradientKind.Linear)))));
        sheet.ColumnWidths[0] = 100;
        sheet.PageSetup.Scale = 0.5;
        Assert.Equal(50.0, SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<GradientItem>().Single().Width, 3);
    }
}

public class SheetImageTests
{
    private static readonly ImageData Png = new(PngTestEncoder.Encode(2, 2, 2, 8, [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]), "t.png");

    private static SheetImage TwoCell(int fromColumn, int fromRow, int toColumn, int toRow, double fromColumnOffset = 0, double fromRowOffset = 0) => new()
    {
        Image = Png, Anchor = ImageAnchorKind.TwoCell,
        From = new AnchorPoint(fromColumn, fromRow, fromColumnOffset, fromRowOffset), To = new AnchorPoint(toColumn, toRow, 0, 0),
    };

    private static Worksheet SheetWith(params SheetImage[] images)
    {
        var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text("a")), (9, 9, SheetLayoutTests.Text("z")));
        sheet.DefaultColumnWidth = 50;
        sheet.DefaultRowHeight = 20;
        sheet.Images.AddRange(images);
        return sheet;
    }

    private static (double Left, double Top) Origin(Worksheet sheet) => (sheet.PageSetup.Margins.Left, sheet.PageSetup.Margins.Top);

    [Fact]
    public void A_two_cell_picture_covers_its_cells_and_is_clipped_to_them()
    {
        var sheet = SheetWith(TwoCell(1, 2, 3, 4));
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        var (left, top) = Origin(sheet);
        var image = page.Items.OfType<ImageItem>().Single();
        Assert.Equal((left + 50, top + 40, 100.0, 40.0), (image.X, image.Y, image.Width, image.Height));
        var push = Assert.IsType<PushItem>(page.Items[page.Items.IndexOf(image) - 1]);
        Assert.Equal((left + 50, top + 40, 100.0, 40.0), push.Clip!.Value);
        Assert.IsType<PopItem>(page.Items[page.Items.IndexOf(image) + 1]);
    }

    [Fact]
    public void Pictures_are_painted_above_cell_text()
    {
        var page = SheetLayoutTests.Layout(SheetWith(TwoCell(0, 0, 2, 2))).Pages[0];
        Assert.True(page.Items.FindIndex(i => i is ImageItem) > page.Items.FindIndex(i => i is TextItem));
    }

    [Fact]
    public void One_cell_and_absolute_anchors_use_their_own_size_and_offsets()
    {
        var oneCell = new SheetImage { Image = Png, Anchor = ImageAnchorKind.OneCell, From = new AnchorPoint(2, 1, 5, 6), Width = 33, Height = 11 };
        var absolute = new SheetImage { Image = Png, Anchor = ImageAnchorKind.Absolute, X = 120, Y = 70, Width = 22, Height = 12 };
        var sheet = SheetWith(oneCell, absolute);
        var images = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<ImageItem>().ToList();
        var (left, top) = Origin(sheet);
        Assert.Equal((left + 105, top + 26, 33.0, 11.0), (images[0].X, images[0].Y, images[0].Width, images[0].Height));
        Assert.Equal((left + 120, top + 70, 22.0, 12.0), (images[1].X, images[1].Y, images[1].Width, images[1].Height));
    }

    [Fact]
    public void A_picture_that_sticks_out_of_the_print_area_is_clipped_at_it()
    {
        var sheet = SheetWith(TwoCell(0, 0, 6, 6));
        sheet.PageSetup.PrintArea = new CellRange(0, 0, 1, 1);
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        var (left, top) = Origin(sheet);
        var clip = page.Items.OfType<PushItem>().First().Clip!.Value;
        Assert.Equal((left, top, 100.0, 40.0), clip);
        var image = page.Items.OfType<ImageItem>().Single();
        Assert.Equal((300.0, 120.0), (image.Width, image.Height)); // still drawn at full size, only clipped
    }

    [Fact]
    public void A_picture_outside_the_print_area_is_not_drawn()
    {
        var sheet = SheetWith(TwoCell(5, 5, 7, 7));
        sheet.PageSetup.PrintArea = new CellRange(0, 0, 1, 1);
        Assert.Empty(SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<ImageItem>());
    }

    [Fact]
    public void Pictures_extend_the_used_range_and_can_be_the_only_content()
    {
        var onlyImage = new Worksheet("S");
        onlyImage.Images.Add(TwoCell(0, 0, 2, 2));
        var page = SheetLayoutTests.Layout(onlyImage).Pages[0];
        Assert.Single(page.Items.OfType<ImageItem>());

        var sheet = SheetLayoutTests.Sheet((0, 0, SheetLayoutTests.Text("a")));
        sheet.Images.Add(TwoCell(3, 0, 20, 2)); // right of the only cell, wider than a page
        sheet.DefaultColumnWidth = 50;
        var doc = SheetLayoutTests.Layout(sheet);
        Assert.True(doc.Pages.Count >= 2, "the picture's columns are printed too, so the sheet needs more than one page");
    }

    [Fact]
    public void A_picture_spanning_a_page_break_continues_on_the_next_page_and_pieces_add_up()
    {
        var sheet = SheetWith(TwoCell(0, 3, 2, 8));
        sheet.PageSetup.RowBreaks.Add(5); // break after row index 5
        var doc = SheetLayoutTests.Layout(sheet);
        Assert.Equal(2, doc.Pages.Count);
        var first = doc.Pages[0].Items.OfType<PushItem>().First().Clip!.Value;
        var second = doc.Pages[1].Items.OfType<PushItem>().First().Clip!.Value;
        Assert.Equal(3 * 20.0, first.Height, 3);  // rows 3-5 on the first page
        Assert.Equal(2 * 20.0, second.Height, 3); // rows 6-7 on the second
        var secondImage = doc.Pages[1].Items.OfType<ImageItem>().Single();
        var firstImage = doc.Pages[0].Items.OfType<ImageItem>().Single();
        Assert.Equal(100.0, firstImage.Width);
        Assert.Equal(firstImage.Height, secondImage.Height);
        double top = sheet.PageSetup.Margins.Top;
        Assert.Equal(top + 3 * 20.0, firstImage.Y, 3);  // rows 0-5 are on the first page, the picture starts at row 3
        Assert.Equal(top - 3 * 20.0, secondImage.Y, 3); // rows 6-9 on the second: the picture began three rows above its top
        Assert.Equal(top, second.Y, 3);
    }

    [Fact]
    public void A_picture_in_the_print_title_rows_repeats_on_every_page()
    {
        var sheet = SheetWith(TwoCell(0, 0, 1, 1));
        sheet.PageSetup.TitleRows = (0, 0);
        sheet.PageSetup.RowBreaks.Add(4);
        var doc = SheetLayoutTests.Layout(sheet);
        Assert.True(doc.Pages.Count >= 2);
        Assert.All(doc.Pages, p => Assert.Single(p.Items.OfType<ImageItem>()));
    }

    [Fact]
    public void Rows_that_grow_to_fit_wrapped_text_push_anchored_pictures_down()
    {
        var wrapped = SheetLayoutTests.Text("one two three four five six seven eight nine ten", SheetLayoutTests.Style(s => s.Wrap = true));
        var sheet = SheetLayoutTests.Sheet((0, 0, wrapped));
        sheet.ColumnWidths[0] = 40;
        sheet.Images.Add(TwoCell(1, 1, 2, 2));
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        var image = page.Items.OfType<ImageItem>().Single();
        var text = page.Items.OfType<TextItem>().Max(t => t.Baseline);
        Assert.True(image.Y > text, "the picture sits below the last line of the grown row");
        Assert.True(image.Y - sheet.PageSetup.Margins.Top > sheet.DefaultRowHeight * 2, "row 0 is taller than the default");
    }

    [Fact]
    public void Hidden_columns_take_no_width_in_a_two_cell_picture()
    {
        var sheet = SheetWith(TwoCell(0, 0, 3, 1));
        sheet.HiddenColumns.Add(1);
        Assert.Equal(100.0, SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<ImageItem>().Single().Width, 3);
    }

    [Fact]
    public void A_cropped_picture_is_drawn_enlarged_and_clipped_to_its_frame()
    {
        var image = TwoCell(1, 1, 3, 3) with { CropLeft = 0.25, CropRight = 0.25, CropTop = 0.1 };
        var sheet = SheetWith(image);
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        var (left, top) = Origin(sheet);
        var drawn = page.Items.OfType<ImageItem>().Single();
        Assert.Equal(200.0, drawn.Width, 3); // the kept half of the width fills 100 pt
        Assert.Equal(40.0 / 0.9, drawn.Height, 3);
        Assert.Equal(left + 50 - 0.25 * 200, drawn.X, 3);
        Assert.Equal(top + 20 - 0.1 * (40 / 0.9), drawn.Y, 3);
        Assert.Contains(page.Items.OfType<PushItem>(), p => p.Clip == (left + 50, top + 20, 100.0, 40.0));
    }

    [Fact]
    public void A_rotated_picture_turns_about_the_center_of_its_frame()
    {
        var sheet = SheetWith(TwoCell(1, 1, 3, 3) with { Rotation = 30 });
        var page = SheetLayoutTests.Layout(sheet).Pages[0];
        var (left, top) = Origin(sheet);
        var turn = page.Items.OfType<PushItem>().Single(p => p.Rotation != 0);
        Assert.Equal((30.0, left + 100, top + 40), (turn.Rotation, turn.CenterX, turn.CenterY));
        // The corners of the turned picture reach outside its frame, so it is clipped at the printed band (here the whole used range), not at the frame.
        var clip = page.Items.OfType<PushItem>().First().Clip!.Value;
        Assert.True(clip.Width > 100 && clip.Height > 40);
    }

    [Fact]
    public void Pictures_follow_the_page_scale()
    {
        var sheet = SheetWith(TwoCell(0, 0, 2, 2));
        sheet.PageSetup.Scale = 0.5;
        var image = SheetLayoutTests.Layout(sheet).Pages[0].Items.OfType<ImageItem>().Single();
        Assert.Equal((50.0, 20.0), (image.Width, image.Height));
    }

    [Fact]
    public void Zero_sized_pictures_are_skipped()
    {
        Assert.Empty(SheetLayoutTests.Layout(SheetWith(TwoCell(2, 2, 2, 2))).Pages[0].Items.OfType<ImageItem>());
    }
}
