namespace DotOffice.Core;

public enum ChartKind
{
    /// <summary>Vertical bars.</summary>
    Column,
    /// <summary>Horizontal bars.</summary>
    Bar,
    Line,
    Pie,
    /// <summary>Pie slices in rings, one ring per series.</summary>
    Doughnut,
    Area,
    /// <summary>Points at (x, y) values; lines are optional.</summary>
    Scatter,
    /// <summary>A scatter chart whose points are circles sized by a third value.</summary>
    Bubble,
    Radar,
}

public enum ChartGrouping { Clustered, Stacked, PercentStacked }

public enum LegendPosition { None, Right, Left, Top, Bottom }

/// <summary>
/// One data series with everything needed to draw it. <paramref name="Labels"/> are the data-label texts per point (null when the
/// series shows none); <paramref name="PointColors"/> are the per-point colors of pie slices.
/// </summary>
/// <remarks>
/// <paramref name="XValues"/> are the x coordinates of scatter and bubble points (null: 1, 2, 3…), <paramref name="Sizes"/> the bubble
/// sizes, and <paramref name="Line"/> whether the points of a line, scatter or radar series are joined.
/// </remarks>
public sealed record ChartSeries(string Name, IReadOnlyList<double?> Values, Color Color, bool Marker = false,
    IReadOnlyList<Color>? PointColors = null, IReadOnlyList<string>? Labels = null,
    IReadOnlyList<double?>? XValues = null, IReadOnlyList<double?>? Sizes = null, bool Line = true);

/// <summary>A tick of the value axis with its formatted label.</summary>
public readonly record struct ChartTick(double Value, string Label);

/// <summary>
/// A chart with its data already resolved: categories, series, and the value-axis scale with formatted tick labels. The layout only
/// has to place and draw it. Lengths in points.
/// </summary>
public sealed record SheetChart : SheetObject
{
    public ChartKind Kind { get; init; }

    public ChartGrouping Grouping { get; init; }

    public string? Title { get; init; }

    public double TitleSize { get; init; } = 14;

    public IReadOnlyList<string> Categories { get; init; } = [];

    public IReadOnlyList<ChartSeries> Series { get; init; } = [];

    public LegendPosition Legend { get; init; } = LegendPosition.Right;

    /// <summary>The ticks of the value axis, lowest first; <see cref="ValueMin"/> and <see cref="ValueMax"/> are its ends.</summary>
    public IReadOnlyList<ChartTick> Ticks { get; init; } = [];

    public double ValueMin { get; init; }

    public double ValueMax { get; init; } = 1;

    public bool ShowValueAxis { get; init; } = true;

    public bool ShowCategoryAxis { get; init; } = true;

    public bool MajorGridlines { get; init; } = true;

    public string? ValueAxisTitle { get; init; }

    public string? CategoryAxisTitle { get; init; }

    /// <summary>The ticks of the horizontal value axis of scatter and bubble charts, with its ends.</summary>
    public IReadOnlyList<ChartTick> XTicks { get; init; } = [];

    public double XMin { get; init; }

    public double XMax { get; init; } = 1;

    public bool XGridlines { get; init; }

    /// <summary>
    /// Whether the points of line and area charts sit in the middle of their category band (<c>between</c>) rather than on its edges
    /// (<c>midCat</c>, the default of area charts).
    /// </summary>
    public bool CategoriesBetween { get; init; } = true;

    /// <summary>The hole of a doughnut as a fraction of its radius.</summary>
    public double HoleSize { get; init; } = 0.5;

    /// <summary>Where the first slice of a pie or doughnut starts, in degrees clockwise from 12 o'clock.</summary>
    public double FirstSliceAngle { get; init; }

    /// <summary>The size of the largest bubble in percent of the default.</summary>
    public double BubbleScale { get; init; } = 100;

    /// <summary>Whether radar series are filled polygons.</summary>
    public bool Filled { get; init; }

    /// <summary>Gap between the groups of bars as a percentage of a bar's width.</summary>
    public double GapWidth { get; init; } = 150;

    /// <summary>How far the bars of one group overlap, in percent of a bar's width (negative values leave space).</summary>
    public double Overlap { get; init; }

    /// <summary>Whether pie slices are labeled with their share (<see cref="ChartSeries.Labels"/> carry the text).</summary>
    public bool ShowPercent { get; init; }

    /// <summary>The chart area's fill; null means transparent.</summary>
    public Color? Background { get; init; } = Color.White;

    /// <summary>The chart area's outline; null means none.</summary>
    public Color? Border { get; init; } = new Color(0xD9, 0xD9, 0xD9);

    /// <summary>The size of the text of axes, legend and labels.</summary>
    public double FontSize { get; init; } = 9;
}
