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
    IReadOnlyList<double?>? XValues = null, IReadOnlyList<double?>? Sizes = null, bool Line = true, bool Smooth = false,
    double Alpha = 1, ChartTrendline? Trendline = null, IReadOnlyList<ChartErrorBars>? ErrorBars = null);

public enum TrendlineKind { Linear, Exponential, Logarithmic, Power, Polynomial, MovingAverage }

/// <summary>
/// A trendline fitted to a series (over its x values, or 1, 2, 3… for category charts) and drawn <paramref name="Forward"/> units beyond
/// its last x and <paramref name="Backward"/> before its first. <paramref name="Order"/> is the degree of a polynomial and
/// <paramref name="Period"/> the window of a moving average.
/// </summary>
public sealed record ChartTrendline(TrendlineKind Kind, Color Color, double Width = 2.25, bool Dashed = true, int Order = 2, int Period = 2,
    double Forward = 0, double Backward = 0, bool ShowEquation = false, bool ShowRSquared = false, ChartTextStyle? LabelStyle = null);

/// <summary>Text formatting that differs from the chart's base: any part that is null keeps the default.</summary>
public sealed record ChartTextStyle(double? Size = null, Color? Color = null, bool? Bold = null, bool? Italic = null, string? Family = null);

/// <summary>The text styles of the parts of a chart (title, axis labels, axis titles, legend, data labels); null parts use the defaults.</summary>
public sealed record ChartTextStyles
{
    public ChartTextStyle? Title { get; init; }
    public ChartTextStyle? ValueAxis { get; init; }
    public ChartTextStyle? CategoryAxis { get; init; }
    public ChartTextStyle? SecondaryAxis { get; init; }
    public ChartTextStyle? ValueAxisTitle { get; init; }
    public ChartTextStyle? CategoryAxisTitle { get; init; }
    public ChartTextStyle? SecondaryAxisTitle { get; init; }
    public ChartTextStyle? Legend { get; init; }
    public ChartTextStyle? DataLabels { get; init; }
}

/// <summary>
/// Error bars of a series: per point the distance drawn above (<paramref name="Plus"/>) and below (<paramref name="Minus"/>) the value,
/// or to the right and left of an x value when <paramref name="Horizontal"/>.
/// </summary>
public sealed record ChartErrorBars(bool Horizontal, IReadOnlyList<double> Plus, IReadOnlyList<double> Minus, bool EndCap = true, Color? Color = null, double Width = 1);

/// <summary>A manually placed box in fractions of the chart's width and height; <paramref name="Inner"/> excludes the axis labels.</summary>
public readonly record struct ChartBox(double X, double Y, double Width, double Height, bool Inner);

/// <summary>
/// A further chart type drawn on the same category axis as the first (a combination chart), on the primary or the secondary value axis.
/// </summary>
public sealed record ChartGroup(ChartKind Kind, ChartGrouping Grouping, IReadOnlyList<ChartSeries> Series, bool Secondary, double GapWidth = 150, double Overlap = 0);

/// <summary>Where the values of an axis fall: linear or logarithmic, and from the start or (reversed) from the end.</summary>
public readonly record struct AxisScale(double Min, double Max, double LogBase = 0, bool Reversed = false)
{
    /// <summary>How far along the axis (0 at its start, 1 at its end) a value lies; values outside the scale are clamped.</summary>
    public double Fraction(double value)
    {
        double f;
        if (LogBase > 1 && Min > 0 && Max > Min)
            f = Math.Log(Math.Max(value, Min) / Min) / Math.Log(Max / Min);
        else
            f = Max > Min ? (value - Min) / (Max - Min) : 0;
        f = Math.Clamp(f, 0, 1);
        return Reversed ? 1 - f : f;
    }
}

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

    /// <summary>The base of a logarithmic value axis (0: linear); the same for the horizontal axis of scatter charts and the secondary axis.</summary>
    public double LogBase { get; init; }

    public double XLogBase { get; init; }

    public double SecondaryLogBase { get; init; }

    /// <summary>Whether the axes run backwards: values from the top (or right), categories from the right (or top), x from the right.</summary>
    public bool ReverseValues { get; init; }

    public bool ReverseCategories { get; init; }

    public bool ReverseX { get; init; }

    public bool ReverseSecondary { get; init; }

    /// <summary>Further chart types of a combination chart, drawn over the first.</summary>
    public IReadOnlyList<ChartGroup> ExtraGroups { get; init; } = [];

    /// <summary>The secondary value axis (on the right), used by the groups marked <see cref="ChartGroup.Secondary"/>.</summary>
    public IReadOnlyList<ChartTick> SecondaryTicks { get; init; } = [];

    public double SecondaryMin { get; init; }

    public double SecondaryMax { get; init; } = 1;

    public string? SecondaryAxisTitle { get; init; }

    public bool ShowSecondaryAxis { get; init; } = false;

    public AxisScale ValueScale => new(ValueMin, ValueMax, LogBase, ReverseValues);

    public AxisScale XScale => new(XMin, XMax, XLogBase, ReverseX);

    public AxisScale SecondaryScale => new(SecondaryMin, SecondaryMax, SecondaryLogBase, ReverseSecondary);

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

    /// <summary>A manual layout (<c>c:layout</c>) of the plot area; null means automatic.</summary>
    public ChartBox? PlotArea { get; init; }

    /// <summary>A manual position (top-left, in fractions of the chart) of the title; null means centered on top.</summary>
    public (double X, double Y)? TitlePosition { get; init; }

    /// <summary>A manual position of the legend (top-left in fractions of the chart; the size, when given, is not used); null means automatic.</summary>
    public ChartBox? LegendBox { get; init; }

    /// <summary>Formatting of the text of the title, axes, legend and labels beyond the base size.</summary>
    public ChartTextStyles TextStyles { get; init; } = new();

    /// <summary>The size of the text of axes, legend and labels.</summary>
    public double FontSize { get; init; } = 9;
}
