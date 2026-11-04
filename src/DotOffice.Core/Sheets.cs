namespace DotOffice.Core;

/// <summary>
/// A spreadsheet workbook. Unlike <see cref="Document"/>, a worksheet is a grid rather than a flow of blocks, so it has its
/// own model. All cell text is already formatted (number and date formats are applied by the reader). Rows and columns are
/// zero-based throughout.
/// </summary>
public sealed class Workbook
{
    public List<Worksheet> Sheets { get; } = [];

    public string? Title { get; set; }

    public string? Author { get; set; }
}

public enum SheetVisibility { Visible, Hidden, VeryHidden }

/// <summary>A rectangular cell range, inclusive on both ends, zero-based.</summary>
public readonly record struct CellRange(int FirstRow, int FirstColumn, int LastRow, int LastColumn)
{
    public int RowCount => LastRow - FirstRow + 1;

    public int ColumnCount => LastColumn - FirstColumn + 1;

    public bool Contains(int row, int column) =>
        row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
}

public sealed class Worksheet(string name)
{
    public string Name { get; } = name;

    public SheetVisibility Visibility { get; set; }

    /// <summary>Cells keyed by (row, column). Empty cells are absent.</summary>
    public Dictionary<(int Row, int Column), Cell> Cells { get; } = [];

    /// <summary>Explicit column widths in points, keyed by column index. Other columns use <see cref="DefaultColumnWidth"/>.</summary>
    public Dictionary<int, double> ColumnWidths { get; } = [];

    public HashSet<int> HiddenColumns { get; } = [];

    /// <summary>Explicit row heights in points, keyed by row index. Other rows use <see cref="DefaultRowHeight"/>.</summary>
    public Dictionary<int, double> RowHeights { get; } = [];

    public HashSet<int> HiddenRows { get; } = [];

    public double DefaultColumnWidth { get; set; } = 48.75;

    public double DefaultRowHeight { get; set; } = 15;

    public List<CellRange> Merges { get; } = [];

    public SheetPageSetup PageSetup { get; set; } = new();

    /// <summary>Whether the sheet is displayed with gridlines on screen (printing is controlled by <see cref="SheetPageSetup.PrintGridLines"/>).</summary>
    public bool ShowGridLines { get; set; } = true;

    public double ColumnWidth(int column) =>
        HiddenColumns.Contains(column) ? 0 : ColumnWidths.GetValueOrDefault(column, DefaultColumnWidth);

    public double RowHeight(int row) =>
        HiddenRows.Contains(row) ? 0 : RowHeights.GetValueOrDefault(row, DefaultRowHeight);
}

public enum CellValueKind { Text, Number, Bool, Error }

/// <summary>One formatted run of a rich-text cell; its <see cref="Font"/> is fully resolved (cell font plus the run's overrides).</summary>
public sealed record CellTextRun(string Text, RunProperties Font);

public sealed class Cell
{
    /// <summary>The display text, with number/date formatting already applied.</summary>
    public string Text { get; init; } = "";

    /// <summary>
    /// Rich-text runs whose texts concatenate to <see cref="Text"/>; null for cells that use the cell style's font throughout.
    /// </summary>
    public IReadOnlyList<CellTextRun>? Runs { get; init; }

    /// <summary>The largest font size used by the cell's text (it decides the line height).</summary>
    public double LargestFontSize => Runs is { Count: > 0 } runs ? runs.Max(r => r.Font.FontSize) : Style.Font.FontSize;

    public CellStyle Style { get; init; } = CellStyle.Default;

    public CellValueKind Kind { get; init; }
}

public enum HorizontalAlignment { General, Left, Center, Right, Fill, Justify }

/// <summary>Fully resolved cell formatting.</summary>
public sealed record CellStyle
{
    public static CellStyle Default { get; } = new();

    public RunProperties Font { get; init; } = new() { FontFamily = "Calibri", FontSize = 11 };

    /// <summary>Solid background fill; null means none.</summary>
    public Color? Fill { get; init; }

    /// <summary>Only <see cref="BorderSet.Top"/>, <see cref="BorderSet.Bottom"/>, <see cref="BorderSet.Left"/> and <see cref="BorderSet.Right"/> are used.</summary>
    public BorderSet Borders { get; init; } = BorderSet.Empty;

    public HorizontalAlignment HorizontalAlignment { get; init; }

    public CellVerticalAlignment VerticalAlignment { get; init; } = CellVerticalAlignment.Bottom;

    public bool Wrap { get; init; }

    /// <summary>Indent in levels (each is about three spaces).</summary>
    public int Indent { get; init; }

    /// <summary>Text rotation in degrees. Read but not rendered yet.</summary>
    public double Rotation { get; init; }
}

/// <summary>Left, center and right strings of one header or footer, still containing Excel's <c>&amp;</c> codes.</summary>
public readonly record struct HeaderFooterParts(string Left, string Center, string Right)
{
    public bool IsEmpty => string.IsNullOrEmpty(Left) && string.IsNullOrEmpty(Center) && string.IsNullOrEmpty(Right);
}

/// <summary>A header or footer with its optional first-page and even-page variants.</summary>
public sealed record HeaderFooterText
{
    public static HeaderFooterText Empty { get; } = new();

    public HeaderFooterParts Odd { get; init; }

    /// <summary>Used on the first page when set (differentFirst).</summary>
    public HeaderFooterParts? First { get; init; }

    /// <summary>Used on even pages when set (differentOddEven).</summary>
    public HeaderFooterParts? Even { get; init; }
}

/// <summary>Print settings of a worksheet. Lengths in points; rows and columns zero-based.</summary>
public sealed class SheetPageSetup
{
    public Orientation Orientation { get; set; }

    public double PageWidth { get; set; } = 612;

    public double PageHeight { get; set; } = 792;

    public Thickness Margins { get; set; } = new(54, 50.4, 54, 50.4);

    public double HeaderDistance { get; set; } = 21.6;

    public double FooterDistance { get; set; } = 21.6;

    /// <summary>The range to print; null means the used range.</summary>
    public CellRange? PrintArea { get; set; }

    /// <summary>Scale factor, 1.0 = 100%.</summary>
    public double Scale { get; set; } = 1.0;

    /// <summary>Fit-to-page width in pages (0 = unconstrained); null when fit-to-page is off.</summary>
    public int? FitToWidth { get; set; }

    /// <summary>Fit-to-page height in pages (0 = unconstrained); null when fit-to-page is off.</summary>
    public int? FitToHeight { get; set; }

    /// <summary>Manual page breaks: a break is placed after each listed row index.</summary>
    public List<int> RowBreaks { get; } = [];

    /// <summary>Manual page breaks: a break is placed after each listed column index.</summary>
    public List<int> ColumnBreaks { get; } = [];

    /// <summary>Rows repeated at the top of every page (inclusive range).</summary>
    public (int From, int To)? TitleRows { get; set; }

    /// <summary>Columns repeated at the left of every page (inclusive range).</summary>
    public (int From, int To)? TitleColumns { get; set; }

    public HeaderFooterText Header { get; set; } = HeaderFooterText.Empty;

    public HeaderFooterText Footer { get; set; } = HeaderFooterText.Empty;

    public bool CenterHorizontally { get; set; }

    public bool CenterVertically { get; set; }

    public bool PrintGridLines { get; set; }

    /// <summary>When true pages run left to right before moving down.</summary>
    public bool PageOrderOverThenDown { get; set; }
}
