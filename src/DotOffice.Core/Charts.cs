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

/// <summary>The shape of a series' markers (<c>c:marker/c:symbol</c>, ISO 29500 ST_MarkerStyle). <c>none</c> is <see cref="ChartSeries.Marker"/> = false; <c>picture</c> and unknown values read as <see cref="Auto"/>.</summary>
public enum ChartMarkerSymbol { Auto, Circle, Dash, Diamond, Dot, Plus, Square, Star, Triangle, X }

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
    double Alpha = 1, ChartTrendline? Trendline = null, IReadOnlyList<ChartErrorBars>? ErrorBars = null,
    IReadOnlyList<ChartTextStyle?>? LabelStyles = null, IReadOnlyList<ChartLabelPlacement?>? LabelPlacements = null,
    double? MarkerSize = null, ChartMarkerSymbol MarkerSymbol = ChartMarkerSymbol.Auto, ChartMarkerStyle? MarkerStyle = null,
    IReadOnlyDictionary<int, ChartPointMarker>? PointMarkers = null, ChartAutoMarker? Auto = null)
{
    /// <summary>The ISO 29500 default of <c>c:marker/c:size</c>, in points.</summary>
    public const double DefaultMarkerSize = 5;

    /// <summary>What an automatic marker of this series looks like; a series built without it (not read from a file) is automatic in its own color.</summary>
    public ChartAutoMarker AutomaticMarker => Auto ?? new ChartAutoMarker(0, Color, Color, 0.75, DefaultMarkerSize);

    /// <summary>
    /// The marker's diameter in points: the series' explicit <c>c:size</c>, else the automatic size, which follows the line's width
    /// (see <see cref="ChartAutoMarker.Size"/>).
    /// </summary>
    public double MarkerDiameter => MarkerSize ?? AutomaticMarker.Size;
}

/// <summary>
/// What Office draws for the parts of a line, scatter or radar marker that the file leaves automatic. <paramref name="Index"/> is the
/// series' <c>c:idx</c> (not its position), which picks the automatic symbol (modulo 9) and the theme color; <paramref name="Color"/> is
/// the automatic fill; <paramref name="Outline"/> and <paramref name="OutlineWidth"/> the automatic outline (and the strokes of plus, x and
/// star); <paramref name="Size"/> the automatic size in points, inferred from the series line's width.
/// </summary>
public sealed record ChartAutoMarker(int Index, Color Color, Color Outline, double OutlineWidth, double Size);

/// <summary>
/// What <c>c:marker/c:spPr</c> says about a marker. <paramref name="Fill"/> null with <paramref name="NoFill"/> false means the fill is not
/// stated (the series color); likewise for the outline, whose absence draws none. <paramref name="FillAlpha"/> and <paramref name="LineAlpha"/>
/// are the opacities (0-1) of the colors' <c>alpha</c> modifiers.
/// </summary>
public sealed record ChartMarkerStyle(Color? Fill = null, bool NoFill = false, double FillAlpha = 1,
    Color? Line = null, bool NoLine = false, double? LineWidth = null, double LineAlpha = 1);

/// <summary>
/// What a data point's own <c>c:dPt/c:marker</c> states, keyed by point index in <see cref="ChartSeries.PointMarkers"/>. Every aspect is
/// null when the point does not state it and then comes from the series: <paramref name="Visible"/> is false for <c>symbol="none"</c>
/// and true for any other symbol; <paramref name="Style"/> replaces the series' as a whole (a fill-only point has an automatic outline).
/// </summary>
public sealed record ChartPointMarker(bool? Visible = null, ChartMarkerSymbol? Symbol = null, double? Size = null, ChartMarkerStyle? Style = null);

/// <summary>
/// Where a data label sits relative to its point or bar (<c>dLblPos</c>). <see cref="InsideEnd"/>, <see cref="InsideBase"/> and
/// <see cref="OutsideEnd"/> belong to bars and columns; <see cref="Above"/>, <see cref="Below"/>, <see cref="Left"/> and
/// <see cref="Right"/> to points of lines and scatter charts; <see cref="Center"/> to both. Pie slices take <see cref="Center"/>, <see cref="InsideEnd"/>,
/// <see cref="OutsideEnd"/> and <see cref="BestFit"/> (inside when the text fits the slice, else outside).
/// </summary>
public enum ChartLabelPlacement { Center, InsideEnd, InsideBase, OutsideEnd, Above, Below, Left, Right, BestFit }

public enum TrendlineKind { Linear, Exponential, Logarithmic, Power, Polynomial, MovingAverage }

/// <summary>
/// A trendline fitted to a series (over its x values, or 1, 2, 3… for category charts) and drawn <paramref name="Forward"/> units beyond
/// its last x and <paramref name="Backward"/> before its first. <paramref name="Order"/> is the degree of a polynomial and
/// <paramref name="Period"/> the window of a moving average.
/// </summary>
public sealed record ChartTrendline(TrendlineKind Kind, Color Color, double Width = 2.25, bool Dashed = true, int Order = 2, int Period = 2,
    double Forward = 0, double Backward = 0, bool ShowEquation = false, bool ShowRSquared = false, ChartTextStyle? LabelStyle = null,
    Func<double, string>? NumberFormatter = null, ChartLabelOffset? LabelPosition = null);

/// <summary>
/// Where a trendline's label sits when the chart places it by hand: either at a position measured from the chart's top-left corner
/// (<paramref name="FromEdge"/>) or moved by that fraction of the chart's size from where it would be.
/// </summary>
public readonly record struct ChartLabelOffset(double X, double Y, bool FromEdge);

/// <summary>Text formatting that differs from the chart's base: any part that is null keeps the default.</summary>
public sealed record ChartTextStyle(double? Size = null, Color? Color = null, bool? Bold = null, bool? Italic = null, string? Family = null,
    double? Rotation = null);

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
    public ChartTextStyle? SecondaryXAxisTitle { get; init; }
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

    /// <summary>Whether a vertical line joins the highest and lowest value of every category (stock charts, <c>hiLowLines</c>).</summary>
    public bool HiLowLines { get; init; }

    /// <summary>Whether a body joins the first and last series of every category (candlesticks, <c>upDownBars</c>): white when the last is above the first, black otherwise.</summary>
    public bool UpDownBars { get; init; }

    /// <summary>The gap between up/down bars as a percentage of their width.</summary>
    public double UpDownGap { get; init; } = 150;

    public string? Title { get; init; }

    public double TitleSize { get; init; } = 14;

    /// <summary>
    /// The color of text that states none; null draws Excel's dark gray. When set, data labels inside bars and slices take it too, as
    /// PowerPoint and LibreOffice draw presentation charts, instead of a color picked for contrast with the fill.
    /// </summary>
    public Color? TextColor { get; init; }

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
    /// <summary>The title of the secondary horizontal axis (along the top) of a scatter or bubble chart.</summary>
    public string? SecondaryXAxisTitle { get; init; }

    public bool ShowSecondaryAxis { get; init; } = false;

    /// <summary>The secondary horizontal axis of a scatter or bubble chart (on top), used by the XY groups marked <see cref="ChartGroup.Secondary"/>.</summary>
    public IReadOnlyList<ChartTick> SecondaryXTicks { get; init; } = [];

    public double SecondaryXMin { get; init; }

    public double SecondaryXMax { get; init; } = 1;

    public double SecondaryXLogBase { get; init; }

    public bool ReverseSecondaryX { get; init; }

    public bool ShowSecondaryXAxis { get; init; } = false;

    public AxisScale SecondaryXScale => new(SecondaryXMin, SecondaryXMax, SecondaryXLogBase, ReverseSecondaryX);

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
    /// <summary>Whether the legend is drawn over the plot area (<c>overlay</c>) instead of making the plot area give way to it.</summary>
    public bool LegendOverlay { get; init; }

    /// <summary>Formatting of the text of the title, axes, legend and labels beyond the base size.</summary>
    public ChartTextStyles TextStyles { get; init; } = new();

    /// <summary>The size of the text of axes, legend and labels.</summary>
    public double FontSize { get; init; } = 9;
}
