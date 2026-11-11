using DotOffice.Core;

namespace DotOffice.Layout;

// Word 2016's chart layout (SheetChart.WordLayout). The numbers were measured on charts with Arial 10 pt text in Word 2016's PDF output; where a
// rule is stretched to other text sizes or to a case that was not measured, the comment says INFERRED. Everything else about a chart (title
// position, series, axis lines) is drawn by the shared code.
public sealed partial class SheetLayoutEngine
{
    /// <summary>The plot rectangle of a Word chart being drawn (absolute), which the axis and scatter layouts use instead of their own estimate; null for every other chart.</summary>
    private (double X, double Y, double Width, double Height)? _wordPlot;

    /// <summary>The gap between the plot's edge and the right edge of the value-axis labels: 5 pt, 8.63 in Word (measured).</summary>
    private double AxisLabelGap => _wordPlot is null ? 5 : 8.63;

    /// <summary>
    /// A label that is centered on a gridline or a slot sits 0.22 pt higher in Word than the shared layout puts it (bounding box bottom 5.2 pt below
    /// the center at 10 pt, against 5.42); the scaling with the text size is INFERRED.
    /// </summary>
    private double CenterShift(RunProperties font) => _wordPlot is null ? 0 : -0.22 * font.FontSize / 10;

    /// <summary>Where a label under the horizontal axis starts below the axis: its line top at 4 pt in the shared layout; in Word its text top (ascent) is 6.0 below the axis.</summary>
    private double LabelDrop(RunProperties font)
    {
        if (_wordPlot is null) return 4;
        var (pitch, ascent, descent) = LineMetrics(font);
        return WordLabelGap - (pitch - (ascent + descent)) / 2;
    }

    private const double WordLabelGap = 5.96;

    /// <summary>The height of a line of text as Word takes it (ascent plus descent with the external leading), not the spreadsheet engine's taller row pitch.</summary>
    private double WordLineHeight(RunProperties font) { var (_, ascent, descent) = LineMetrics(font); return ascent + descent; }

    private (double Pitch, double Ascent, double Descent) LineMetrics(RunProperties font) => BuildBlock(ChartCell("Hg", font), "Hg", 1e7).Metrics[0];

    private enum WordKey { Square, Line, Marker }

    /// <summary>
    /// Whether a chart is laid out by <see cref="EmitWordChart"/>: a Word chart of a measured type with an automatic layout, no axis titles,
    /// no turned or hidden axes and no secondary axes. Other charts keep the shared layout.
    /// </summary>
    private static bool UseWordLayout(SheetChart chart) =>
        chart.WordLayout && chart.PlotArea is null && chart.LegendBox is null && chart.TitlePosition is null && !chart.LegendOverlay
        && chart.Kind is ChartKind.Column or ChartKind.Bar or ChartKind.Line or ChartKind.Scatter or ChartKind.Pie
        && chart.ExtraGroups.Count == 0 && !chart.ShowSecondaryAxis
        && string.IsNullOrEmpty(chart.ValueAxisTitle) && string.IsNullOrEmpty(chart.CategoryAxisTitle)
        && (chart.Kind == ChartKind.Pie || chart is { ShowValueAxis: true, ShowCategoryAxis: true, ReverseValues: false, ReverseCategories: false })
        && chart.TextStyles.ValueAxis?.Rotation is null or 0 && chart.TextStyles.CategoryAxis?.Rotation is null or 0
        && chart.Series.Count > 0;

    private static List<(string Name, Color Color, ChartSeries? Series, ChartKind Kind)> WordLegendEntries(SheetChart chart)
    {
        var entries = chart.Kind == ChartKind.Pie
            ? chart.Categories.Select((c, i) => (Name: c, Color: chart.Series[0].PointColors is { } pc && i < pc.Count ? pc[i] : ChartColor(i), Series: (ChartSeries?)null, Kind: chart.Kind)).ToList()
            : chart.Series.Select(s => (Name: s.Name, Color: s.Color, Series: (ChartSeries?)s, Kind: chart.Kind)).ToList();
        // A clustered bar chart and a stacked column chart list their series backwards in a vertical legend (the bar chart's first series is the lowest one,
        // the stack's first series the lowest segment). Measured on the right; the left legend and percent-stacked columns are INFERRED.
        bool vertical = chart.Legend is LegendPosition.Right or LegendPosition.Left;
        if (vertical && (chart.Kind == ChartKind.Bar && chart.Grouping == ChartGrouping.Clustered || chart.Kind == ChartKind.Column && chart.Grouping != ChartGrouping.Clustered)) entries.Reverse();
        return entries;
    }

    /// <summary>
    /// Lays out and draws a Word chart below its title (already drawn; <paramref name="band"/> is the title's height plus 9 pt, 0 without a title):
    /// the legend, then the plot area, from these measured rules (W and H are the chart's size). The plot's right edge is W - 11 and its top
    /// 11 + band; its bottom is 6.5 + the axis labels' line height + 5.71 above the bottom edge; its left edge is 6.5 + the widest value label + 8.63.
    /// A right legend is 22.04 + its widest text wide, takes that much off the right and has the text ending 10.05 from the edge; entries are 17.07
    /// apart at 10 pt and centered on (H + band) / 2. A line's key (19.2 long, marker in the middle) also moves the plot 14 pt further.
    /// Top and bottom legends are one row, 23.07 tall. Returns false, drawing nothing, when the legend does not fit that way.
    /// </summary>
    private bool EmitWordChart(SheetChart chart, double x, double y, double w, double h, double band, RunProperties font, List<LayoutItem> items)
    {
        var styles = chart.TextStyles;
        var valueFont = Styled(font, styles.ValueAxis);
        var categoryFont = Styled(font, styles.CategoryAxis);
        bool bar = chart.Kind == ChartKind.Bar, pie = chart.Kind == ChartKind.Pie, scatter = chart.Kind == ChartKind.Scatter;
        var entries = WordLegendEntries(chart);
        bool hasLegend = chart.Legend != LegendPosition.None && entries.Count > 0;
        var legendFont = Styled(font, styles.Legend);
        var widths = entries.Select(e => MeasureChartText(e.Name, legendFont)).ToList();
        double widest = widths.DefaultIfEmpty(0).Max();
        WordKey KeyOf(int i) => entries[i].Kind == ChartKind.Line || entries[i].Kind == ChartKind.Scatter && (entries[i].Series?.Line ?? true) ? WordKey.Line
            : entries[i].Kind == ChartKind.Scatter ? WordKey.Marker : WordKey.Square;
        // From the key's left edge to the text: the key and its gap (the square 5.03 + 2.2, the line 19.2 + 2.09, the marker alone 5.04 + 2.3).
        double Advance(int i) => KeyOf(i) switch { WordKey.Line => 21.29, WordKey.Marker => 7.34, _ => 7.23 };
        bool anyLine = Enumerable.Range(0, entries.Count).Any(i => KeyOf(i) == WordKey.Line);
        double entryPitch = WordLineHeight(legendFont) + 5.57; // 17.07 at 10 pt; the 5.57 is kept at other sizes (INFERRED)
        double rowHeight = entryPitch + 6.0;

        // ---- the plot rectangle (relative to the chart's corner)
        double labelWidth = pie ? 0 : bar ? chart.Categories.Select(c => MeasureChartText(c, categoryFont)).DefaultIfEmpty(0).Max()
            : chart.Ticks.Select(t => MeasureChartText(t.Label, valueFont)).DefaultIfEmpty(0).Max();
        var bottomFont = bar ? valueFont : categoryFont;
        
        double left = pie ? 11 : 6.5 + labelWidth + 8.63, right = w - 11, top = 11 + band, bottom = pie ? h - 11 : h - (6.5 + WordLineHeight(bottomFont) + 5.71 * bottomFont.FontSize / 10);
        double legendShift = hasLegend ? widest + 22.04 + (anyLine ? 14 : 0) : 0;
        if (hasLegend)
            switch (chart.Legend)
            {
                case LegendPosition.Right:
                    right -= legendShift;
                    // The last label of the horizontal value axis is centered on the plot's edge and the plot gives way by half of it (bars and scatter).
                    if (bar && chart.Ticks.Count > 0) right -= MeasureChartText(chart.Ticks[^1].Label, valueFont) / 2;
                    else if (scatter && chart.XTicks.Count > 0) right -= MeasureChartText(chart.XTicks[^1].Label, categoryFont) / 2;
                    break;
                case LegendPosition.Left: left += legendShift; break;
                case LegendPosition.Top: top += rowHeight; break;
                default: bottom -= rowHeight; break;
            }
        if (right - left < 20 || bottom - top < 20) return false;

        // ---- the legend's entries: (key left, centre y, text left)
        var placed = new List<(int Index, double KeyLeft, double Cy, double TextLeft)>();
        if (hasLegend && chart.Legend is LegendPosition.Right or LegendPosition.Left)
        {
            int shown = Math.Clamp((int)Math.Floor((h - band - 12 + 0.5) / entryPitch), 1, entries.Count);
            double textLeft = chart.Legend == LegendPosition.Right ? w - 10.05 - widest : 10.76 + Enumerable.Range(0, entries.Count).Max(Advance);
            double centre = shown < entries.Count ? band + 6 + shown * entryPitch / 2 : (h + band) / 2;
            for (int i = 0; i < shown; i++)
                placed.Add((i, textLeft - Advance(i), centre + (i - (shown - 1) / 2.0) * entryPitch, textLeft));
        }
        else if (hasLegend)
        {
            double total = 0;
            for (int i = 0; i < entries.Count; i++) total += Advance(i) + widths[i] + (i > 0 ? KeyOf(i) == WordKey.Line ? 10.42 : 8.45 : 0);
            if (total > w - 22) return false; // entries that would wrap are not measured
            double cx = w / 2 + 1.84 - total / 2, cy = chart.Legend == LegendPosition.Top ? band + 6.0 + entryPitch / 2 : h - 6.0 - entryPitch / 2;
            for (int i = 0; i < entries.Count; i++)
            {
                if (i > 0) cx += KeyOf(i) == WordKey.Line ? 10.42 : 8.45;
                placed.Add((i, cx, cy, cx + Advance(i)));
                cx += Advance(i) + widths[i];
            }
        }

        // ---- draw: the legend, then the plot
        foreach (var (i, keyLeft, cy, textLeft) in placed)
        {
            var (name, color, series, _) = entries[i];
            double ky = y + cy, kx = x + keyLeft;
            switch (KeyOf(i))
            {
                case WordKey.Line:
                    items.Add(new LineItem(kx, ky, kx + 19.2, ky, series?.StrokeWidth ?? 1.5, series?.StrokeColor ?? color));
                    if (series is null || series.Marker)
                        DrawMarker(items, kx + 9.6, ky, series?.MarkerSymbol ?? ChartMarkerSymbol.Auto, series?.MarkerDiameter ?? 5, series?.AutomaticMarker ?? new ChartAutoMarker(0, color, color, 0.75, 5), series?.MarkerStyle);
                    break;
                case WordKey.Marker:
                    if (series is null || series.Marker)
                        DrawMarker(items, kx + 2.52, ky, series?.MarkerSymbol ?? ChartMarkerSymbol.Auto, series?.MarkerDiameter ?? 5, series?.AutomaticMarker ?? new ChartAutoMarker(0, color, color, 0.75, 5), series?.MarkerStyle);
                    break;
                default:
                    items.Add(new RectItem(kx, ky - 2.515, 5.03, 5.03, color));
                    break;
            }
            PlaceText(name, legendFont, x + textLeft, ky + CenterShift(legendFont) - 0.21 * legendFont.FontSize / 10, Anchor.Start, true, 0, items); // 0.21 pt above the other labels: Word's legend text baseline is 3.15 below the key's center
        }
        double pl = x + left, pt = y + top, pr = x + right, pb = y + bottom;
        _manualPlot = null;
        _wordPlot = (pl, pt, pr - pl, pb - pt);
        try
        {
            if (pie) EmitPie(chart, pl, pt, pr, pb, font, items);
            else if (scatter) EmitScatter(chart, pl, pt, pr, pb, font, items);
            else EmitAxisChart(chart, pl, pt, pr, pb, font, items);
        }
        finally { _wordPlot = null; }
        return true;
    }
}
