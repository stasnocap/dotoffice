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

/// <summary>A position inside a cell: the cell (zero-based) and the offset from its top-left corner in points.</summary>
public readonly record struct AnchorPoint(int Column, int Row, double ColumnOffset, double RowOffset);

public enum ImageAnchorKind
{
    /// <summary>Corners at <see cref="SheetObject.From"/> and <see cref="SheetObject.To"/>; the size follows the cells.</summary>
    TwoCell,
    /// <summary>Top-left corner at <see cref="SheetObject.From"/>, size <see cref="SheetObject.Width"/> × <see cref="SheetObject.Height"/>.</summary>
    OneCell,
    /// <summary>Top-left corner at (<see cref="SheetObject.X"/>, <see cref="SheetObject.Y"/>) measured from the corner of A1, with the given size.</summary>
    Absolute,
}

/// <summary>A comment attached to a cell: <paramref name="Reference"/> is its A1 address, <paramref name="Text"/> the plain text.</summary>
public sealed record SheetComment(int Row, int Column, string Reference, string? Author, string Text);

/// <summary>Where cell comments are printed (the "Comments" option of the page setup).</summary>
public enum CommentPrinting { None, AtEnd, AsDisplayed }

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

    /// <summary>Pictures in drawing order (later ones are on top).</summary>
    public List<SheetImage> Images { get; } = [];

    /// <summary>Shapes, text boxes and lines of the drawing.</summary>
    public List<SheetShape> Shapes { get; } = [];

    /// <summary>Charts of the drawing.</summary>
    public List<SheetChart> Charts { get; } = [];

    /// <summary>Cell comments (notes); printed after the sheet when <see cref="SheetPageSetup.Comments"/> asks for it.</summary>
    public List<SheetComment> Comments { get; } = [];

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

/// <summary>
/// A conditional-format data bar: the bar spans <see cref="Start"/> to <see cref="End"/> (fractions of the cell width, left to right).
/// Gradient bars fade from the full color at their base (the end next to the axis) to a light tint. A <see cref="BorderColor"/> outlines the bar.
/// <see cref="Axis"/> is where the axis lies (a fraction of the cell width) when the bar grows from one, and is drawn in <see cref="AxisColor"/>;
/// <see cref="Mirrored"/> bars grow from right to left (their <see cref="Start"/> and <see cref="End"/> are already mirrored).
/// </summary>
public sealed record CellDataBar(double Start, double End, Color Color, bool Gradient, bool Negative = false, Color? BorderColor = null,
    double? Axis = null, Color? AxisColor = null, bool Mirrored = false);

public enum IconShape
{
    ArrowUp, ArrowSideways, ArrowDown, ArrowUpRight, ArrowDownRight,
    Circle, RimmedCircle, Check, Exclamation, Cross,
    BareCheck, BareExclamation, BareCross,
    Flag, Diamond, TriangleUp, TriangleDown, Dash,
    /// <summary>A five-pointed star: filled for <c>Level</c> 2, half for 1, empty for 0.</summary>
    Star,
    /// <summary>A disc with <c>Level</c> (0–4) of its quarters filled, clockwise from the top.</summary>
    Quarters,
    /// <summary>Four ascending bars of which the first <c>Level</c> (0–4) are filled.</summary>
    Bars,
    /// <summary>A square with <c>Level</c> (0–4) of its quarters filled.</summary>
    Boxes,
}

/// <summary>A conditional-format icon drawn at the left of the cell.</summary>
public sealed record CellIcon(IconShape Shape, Color Color, int Level = 0);

public sealed record Cell
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

    /// <summary>Whether the cell holds an empty string (a formula returning <c>""</c>, an empty shared string) rather than nothing: formulas see it as text, not as a blank.</summary>
    public bool IsEmptyString { get; init; }

    /// <summary>The raw numeric value (before number formatting) of number cells; null for everything else.</summary>
    public double? Value { get; init; }

    /// <summary>A conditional-format data bar drawn behind the text, if any.</summary>
    public CellDataBar? DataBar { get; init; }

    /// <summary>A conditional-format icon; the text is shifted right by the icon's width.</summary>
    public CellIcon? Icon { get; init; }
}

/// <summary>A color stop of a gradient: <see cref="Position"/> runs from 0 to 1.</summary>
public readonly record struct GradientStop(double Position, Color Color);

public enum GradientKind { Linear, Path }

/// <summary>A cell background that blends between colors.</summary>
public sealed record GradientFill
{
    public GradientKind Kind { get; init; }

    /// <summary>Linear gradients: the direction of the blend in degrees, clockwise from left-to-right (90 runs top to bottom).</summary>
    public double Degree { get; init; }

    /// <summary>Path gradients: the focus, as fractions of the cell from its left and top edge; stop 0 is at the focus, stop 1 at the far edge.</summary>
    public double Left { get; init; }
    public double Right { get; init; }
    public double Top { get; init; }
    public double Bottom { get; init; }

    /// <summary>Path gradients: whether the colors run in concentric rectangles (<c>path="rect"</c>) instead of ellipses.</summary>
    public bool Rectangular { get; init; }

    /// <summary>At least two stops, sorted by position.</summary>
    public IReadOnlyList<GradientStop> Stops { get; init; } = [];

    /// <summary>The color at <paramref name="t"/> (0–1) along the gradient.</summary>
    public Color ColorAt(double t)
    {
        t = Math.Clamp(t, 0, 1);
        if (Stops.Count == 0) return Color.White;
        if (t <= Stops[0].Position) return Stops[0].Color;
        for (int i = 1; i < Stops.Count; i++)
        {
            if (t > Stops[i].Position) continue;
            double span = Stops[i].Position - Stops[i - 1].Position;
            return span <= 1e-9 ? Stops[i].Color : Stops[i - 1].Color.Blend(Stops[i].Color, (t - Stops[i - 1].Position) / span);
        }
        return Stops[^1].Color;
    }
}

public enum HorizontalAlignment { General, Left, Center, Right, Fill, Justify }

/// <summary>Fully resolved cell formatting.</summary>
public sealed record CellStyle
{
    public static CellStyle Default { get; } = new();

    public RunProperties Font { get; init; } = new() { FontFamily = "Calibri", FontSize = 11 };

    /// <summary>Solid background fill; null means none.</summary>
    public Color? Fill { get; init; }

    /// <summary>Gradient background; takes the place of <see cref="Fill"/> when set.</summary>
    public GradientFill? Gradient { get; init; }

    /// <summary>Only <see cref="BorderSet.Top"/>, <see cref="BorderSet.Bottom"/>, <see cref="BorderSet.Left"/> and <see cref="BorderSet.Right"/> are used.</summary>
    public BorderSet Borders { get; init; } = BorderSet.Empty;

    public HorizontalAlignment HorizontalAlignment { get; init; }

    public CellVerticalAlignment VerticalAlignment { get; init; } = CellVerticalAlignment.Bottom;

    public bool Wrap { get; init; }

    /// <summary>Indent in levels (each is about three spaces).</summary>
    public int Indent { get; init; }

    /// <summary>Text rotation in degrees counter-clockwise (−90 reads downwards, 90 upwards).</summary>
    public double Rotation { get; init; }

    /// <summary>Vertical stacked text (textRotation 255): upright letters one below the other.</summary>
    public bool Stacked { get; init; }
}

/// <summary>Left, center and right strings of one header or footer, still containing Excel's <c>&amp;</c> codes.</summary>
public readonly record struct HeaderFooterParts(string Left, string Center, string Right)
{
    public bool IsEmpty => string.IsNullOrEmpty(Left) && string.IsNullOrEmpty(Center) && string.IsNullOrEmpty(Right);
}

/// <summary>A picture placed in a header or footer section with <c>&amp;G</c>; the size is in points.</summary>
public sealed record HeaderFooterPicture(ImageData Image, double Width, double Height);

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

    /// <summary>Headers and footers are scaled with the page scale (<c>scaleWithDoc</c>, on by default).</summary>
    public bool HeaderFooterScaleWithDocument { get; set; } = true;

    /// <summary>Pictures for the <c>&amp;G</c> codes, keyed by section: <c>LH</c>, <c>CH</c>, <c>RH</c> (header left/center/right) and <c>LF</c>, <c>CF</c>, <c>RF</c> (footer).</summary>
    public Dictionary<string, HeaderFooterPicture> HeaderFooterPictures { get; } = new(StringComparer.Ordinal);

    public double HeaderDistance { get; set; } = 21.6;

    public double FooterDistance { get; set; } = 21.6;

    /// <summary>The number the sheet's first page gets (<c>firstPageNumber</c>) when <see cref="UseFirstPageNumber"/> is set; otherwise numbering continues.</summary>
    public int FirstPageNumber { get; set; } = 1;

    public bool UseFirstPageNumber { get; set; }

    /// <summary>The ranges to print, each as its own page set, in order; empty means the used range.</summary>
    public List<CellRange> PrintAreas { get; } = [];

    /// <summary>The first range to print; null means the used range. Setting it replaces all print areas.</summary>
    public CellRange? PrintArea
    {
        get => PrintAreas.Count > 0 ? PrintAreas[0] : null;
        set
        {
            PrintAreas.Clear();
            if (value is { } area) PrintAreas.Add(area);
        }
    }

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

    /// <summary>
    /// Whether and where comments are printed. The layout lists them on pages after the sheet for <see cref="CommentPrinting.AtEnd"/> and
    /// <see cref="CommentPrinting.AsDisplayed"/> alike (floating notes are not drawn).
    /// </summary>
    public CommentPrinting Comments { get; set; }

    /// <summary>When true pages run left to right before moving down.</summary>
    public bool PageOrderOverThenDown { get; set; }
}
