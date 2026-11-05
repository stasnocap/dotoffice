namespace DotOffice.Core;

public enum ChartKind
{
    /// <summary>Vertical bars.</summary>
    Column,
    /// <summary>Horizontal bars.</summary>
    Bar,
    Line,
    Pie,
}

public enum ChartGrouping { Clustered, Stacked, PercentStacked }

public enum LegendPosition { None, Right, Left, Top, Bottom }

/// <summary>
/// One data series with everything needed to draw it. <paramref name="Labels"/> are the data-label texts per point (null when the
/// series shows none); <paramref name="PointColors"/> are the per-point colors of pie slices.
/// </summary>
public sealed record ChartSeries(string Name, IReadOnlyList<double?> Values, Color Color, bool Marker = false,
    IReadOnlyList<Color>? PointColors = null, IReadOnlyList<string>? Labels = null);

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
