using DotOffice.Core;

namespace DotOffice.Layout;

// Word 2016's chart layout (SheetChart.WordLayout). The numbers were measured on charts with Arial 10 pt text in Word 2016's PDF output; where a
// rule is stretched to other text sizes or to a case that was not measured, the comment says INFERRED. Everything else about a chart (title
// position, series, axis lines) is drawn by the shared code.
public sealed partial class SheetLayoutEngine
{
    /// <summary>The plot rectangle of a Word chart being drawn (absolute), which the axis and scatter layouts use instead of their own estimate; null for every other chart.</summary>
    private (double X, double Y, double Width, double Height)? _wordPlot;

    /// <summary>The gap between the plot's edge and the right edge of the axis labels: 5 pt, in Word 8.63 at 10 pt (measured) and 6.90 / 12.05 at 8 / 14 pt, so 0.863 em.</summary>
    private double AxisLabelGap(RunProperties font) => _wordPlot is null ? 5 : WordAxisGap(font);

    private static double WordAxisGap(RunProperties font) => 0.863 * font.FontSize;

    /// <summary>
    /// A label that is centered on a gridline or a slot sits 0.22 pt higher in Word than the shared layout puts it (bounding box bottom 5.2 pt below
    /// the center at 10 pt, against 5.42); the scaling with the text size is INFERRED.
    /// </summary>
    private double CenterShift(RunProperties font) => _wordPlot is null ? 0 : -0.22 * font.FontSize / 10 + WordTextDrift(font);

    /// <summary>Word's labels sit 0.2 pt lower than the 10 pt scaling gives at 14 pt and as predicted at 8 and 10 pt (fitted to those three sizes; INFERRED beyond them).</summary>
    private static double WordTextDrift(RunProperties font) => Math.Max(0, 0.05 * (font.FontSize - 10));

    /// <summary>Where a label under the horizontal axis starts below the axis: its line top at 4 pt in the shared layout; in Word its text top (ascent) is 6.0 below the axis.</summary>
    private double LabelDrop(RunProperties font)
    {
        if (_wordPlot is null) return 4;
        var (pitch, ascent, descent) = LineMetrics(font);
        return WordLabelGap * font.FontSize / 10 + WordTextDrift(font) - (pitch - (ascent + descent)) / 2;
    }

    private const double WordLabelGap = 5.96;

    /// <summary>The height of a line of text as Word takes it (ascent plus descent with the external leading), not the spreadsheet engine's taller row pitch.</summary>
    private double WordLineHeight(RunProperties font) { var (_, ascent, descent) = LineMetrics(font); return ascent + descent; }

    private (double Pitch, double Ascent, double Descent) LineMetrics(RunProperties font) => BuildBlock(ChartCell("Hg", font), "Hg", 1e7).Metrics[0];

    /// <summary>How far a tick mark of the given kind reaches into the plot and out of it from the axis line.</summary>
    private static (double Inward, double Outward) TickSpan(ChartTickMark mark, double length) => mark switch
    {
        ChartTickMark.Out => (0, length), ChartTickMark.In => (length, 0), ChartTickMark.Cross => (length, length), _ => (0, 0),
    };

    private enum WordKey { Square, Line, Marker }

    /// <summary>
    /// Whether a chart is laid out by <see cref="EmitWordChart"/>: a Word chart of a measured type with an automatic layout, no axis titles,
    /// no turned or hidden axes and no secondary axes. Other charts keep the shared layout.
    /// </summary>
    private static bool UseWordLayout(SheetChart chart) =>
        chart.WordLayout && chart.PlotArea is null && chart.LegendBox is null && chart.TitlePosition is null && (!chart.LegendOverlay || chart.Legend == LegendPosition.Right && !chart.WordLegendTopRight)
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
        // the stack's first series the lowest segment). Measured on the right; a bottom legend of clustered bars lists them backwards too (measured at 8 and 14 pt);
        // the top and left legends and percent-stacked columns are INFERRED.
        bool vertical = chart.Legend is LegendPosition.Right or LegendPosition.Left;
        if (chart.Kind == ChartKind.Bar && chart.Grouping == ChartGrouping.Clustered || vertical && chart.Kind == ChartKind.Column && chart.Grouping != ChartGrouping.Clustered) entries.Reverse();
        return entries;
    }

    /// <summary>Greedy word wrap of a legend name to <paramref name="limit"/>; a word wider than the limit breaks by character.</summary>
    private List<string> WrapLegendName(string name, RunProperties font, double limit)
    {
        var result = new List<string>();
        string current = "";
        foreach (var word in name.Split(' '))
        {
            string tryLine = current.Length == 0 ? word : current + " " + word;
            if (MeasureChartText(tryLine, font) <= limit) { current = tryLine; continue; }
            if (current.Length > 0) { result.Add(current); current = ""; }
            var rest = word;
            while (rest.Length > 0 && MeasureChartText(rest, font) > limit)
            {
                int k = 1;
                while (k < rest.Length && MeasureChartText(rest[..(k + 1)], font) <= limit) k++;
                result.Add(rest[..k]);
                rest = rest[k..];
            }
            current = rest;
        }
        result.Add(current);
        return result;
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
        double square = 0.503 * legendFont.FontSize; // the key is 0.503 em square (4.02 / 5.03 / 7.04 pt measured at 8 / 10 / 14 pt)
        double Advance(int i) => KeyOf(i) switch { WordKey.Line => 21.29, WordKey.Marker => 7.34, _ => square + 0.25 * legendFont.FontSize - 0.3 }; // the gap to the text: 1.72 / 2.2 / 3.22 pt measured; the line and marker keys are only measured at 10 pt
        bool anyLine = Enumerable.Range(0, entries.Count).Any(i => KeyOf(i) == WordKey.Line);
        double entryPitch = WordLineHeight(legendFont) + 6.0 - 0.043 * legendFont.FontSize; // 14.85 / 17.07 / 21.49 pt measured at 8 / 10 / 14 pt
        double rowHeight = entryPitch + 6.0;

        // Right and left legends wrap a name wider than W / 3 - 11.5 (measured, 40+ samples at W 236..480 and 8 / 10 pt: greedy at spaces, a single word that is too
        // wide breaks by character; the 11.5 not scaling with the font size is UNMEASURED). The legend's text width is then the widest line, and the legend box is at
        // most W / 3 + 9.0 wide (so a line up to 1.5 pt over the limit is not wrapped, it just moves plot and legend less). Top and bottom legends never wrap (measured).
        bool verticalLegend = hasLegend && chart.Legend is LegendPosition.Right or LegendPosition.Left;
        var lines = entries.Select((e, i) => verticalLegend ? WrapLegendName(e.Name, legendFont, w / 3 - 11.5) : [e.Name]).ToList();
        int maxLines = lines.Select(l => l.Count).DefaultIfEmpty(1).Max();
        double lineStep = 1.15 * legendFont.FontSize; // lines inside an entry: 11.50 / 9.20 pt measured at 10 / 8 pt
        if (verticalLegend)
        {
            widths = lines.Select(l => l.Max(t => MeasureChartText(t, legendFont))).ToList();
            widest = widths.DefaultIfEmpty(0).Max();
        }
        // Every entry takes the height of the tallest: 17.07 / 28.12 / 39.19 pt for 1 / 2 / 3 lines at 10 pt (14.85 / 23.69 at 8 pt), i.e. 1.106 em per extra line.
        double stackPitch = entryPitch + (maxLines - 1) * 1.106 * legendFont.FontSize;

        // Top and bottom legends (measured, 80+ samples at 10 pt, W 240..480): Lmax = 326.85 at W 360 (9 capped rows); its dependence on W (0.9 W + 2.85) is INFERRED.
        // One row when T + 0.283 em x (n + 1) <= Lmax; otherwise a grid. The row's gaps are all equal (see below); the 8 and 14 pt constants come from one sample each.
        int gridRows = 1, gridCols = entries.Count;
        double gridPitch = 0, rowGap = 0, rowTotal = 0;
        if (hasLegend && !verticalLegend)
        {
            int n = entries.Count;
            double fs = legendFont.FontSize, lmax = 0.9 * w + 2.85;
            var e = Enumerable.Range(0, n).Select(i => Advance(i) + widths[i]).ToList();
            rowTotal = e.Sum();
            if (rowTotal + 0.283 * fs * (n + 1) <= lmax)
                rowGap = Math.Min(0.2 * widths.Sum() + 0.482 * fs * n + (1.28 + 0.16 * fs), lmax - rowTotal) / (n + 1);
            else
            {
                gridPitch = e.Max() + 0.283 * fs;
                int maxCols = Math.Max(1, (int)Math.Floor((lmax - 2.83) / gridPitch));
                gridRows = (n + maxCols - 1) / maxCols;
                gridCols = (n + gridRows - 1) / gridRows;
            }
        }
        // The plot gives way by one row plus 17.06 per further row, less 1.50 once there are several rows (10 pt; whether the 1.50 scales with the font is UNMEASURED,
        // the 8 pt sample c-b-s8-n12 is checked against it).
        double extraRows = gridRows > 1 ? (gridRows - 1) * entryPitch - 1.5 : 0;

        // ---- the plot rectangle (relative to the chart's corner)
        double labelWidth = pie ? 0 : bar ? chart.Categories.Select(c => MeasureChartText(c, categoryFont)).DefaultIfEmpty(0).Max()
            : chart.Ticks.Select(t => MeasureChartText(t.Label, valueFont)).DefaultIfEmpty(0).Max();
        var bottomFont = bar ? valueFont : categoryFont;
        
        // The first value label is centered on the plot's top, which gives way to it: 12.82 at 14 pt, 11 at 8 and 10 pt (one sample at 14 pt; 0.9157 em is INFERRED).
        double topMargin = bar || pie ? 11 : Math.Max(11, 0.9157 * valueFont.FontSize);
        double left = pie ? 11 : 6.5 + labelWidth + WordAxisGap(bar ? categoryFont : valueFont), right = w - 11, top = topMargin + band, bottom = pie ? h - 11 : h - (6.5 + WordLineHeight(bottomFont) + 5.71 * bottomFont.FontSize / 10);
        double legendShift = hasLegend ? widest + 12.04 + legendFont.FontSize + (anyLine ? 14 : 0) : 0; // 22.04 at 10 pt; 20.04 and 26.04 measured at 8 and 14 pt
        if (verticalLegend) legendShift = Math.Min(legendShift, w / 3 + 9.0 + (anyLine ? 14 : 0)); // the box is at most W / 3 + 9.0 (measured at 10 pt)
        double textWidth = legendShift - 12.04 - legendFont.FontSize - (anyLine ? 14 : 0); // the text width the box leaves (the widest line, at most the cap)
        // Overlay (measured, 1 sample: right legend, column chart, with title): the legend stands where a normal right legend does but the plot keeps its full
        // width and the legend is drawn after it. Overlay with any other legend position is not measured and keeps the shared layout (UseWordLayout).
        bool overlay = hasLegend && chart.LegendOverlay;
        // legendPos "tr" (measured, 2 samples: 3-entry column charts at 10 pt, with and without a title): the entries stack at the top right at the right legend's x,
        // the plot gives way on the right as for a right legend, on top by one entry pitch per entry and on the bottom by 6.0 pt. That the top offset is n x pitch,
        // other text sizes and other chart kinds are INFERRED; so is a tr legend with more entries than fit (the plot under 20 pt falls back to the shared layout).
        bool topRight = hasLegend && chart.WordLegendTopRight;
        if (hasLegend && !overlay)
            switch (chart.Legend)
            {
                case LegendPosition.Right:
                    right -= legendShift;
                    // The last label of the horizontal value axis is centered on the plot's edge and the plot gives way by half of it (bars and scatter).
                    if (scatter && chart.XTicks.Count > 0) right -= MeasureChartText(chart.XTicks[^1].Label, categoryFont) / 2;
                    if (topRight) { top += entries.Count * entryPitch; bottom -= 6.0; }
                    break;
                case LegendPosition.Left: left += legendShift; break;
                case LegendPosition.Top: top += rowHeight + extraRows; break;
                default: bottom -= rowHeight + extraRows; break;
            }
        // Bars: the plot gives way by half of the last value label (measured with a bottom legend at 8 and 14 pt; with a top or left legend or none it is INFERRED).
        if (bar && chart.Ticks.Count > 0) right -= MeasureChartText(chart.Ticks[^1].Label, valueFont) / 2;
        if (right - left < 20 || bottom - top < 20) return false;

        // ---- the legend's entries: (key left, centre y, text left)
        var placed = new List<(int Index, double KeyLeft, double Cy, double TextLeft)>();
        if (verticalLegend)
        {
            // The count that fits (measured, 26 samples): floor((0.9 (H - band) + 2.76) / pitch), the constant known to lie in 2.40..3.12. With wrapped names the multi-line
            // pitch is used (INFERRED: how wrapping interacts with the cut is UNMEASURED). Right-aligned tr entries are never cut.
            int shown = topRight ? entries.Count : Math.Clamp((int)Math.Floor((0.9 * (h - band) + 2.76) / stackPitch), 1, entries.Count);
            double textLeft = chart.Legend == LegendPosition.Right ? w - 10.05 - textWidth : 10.76 + 0.25 * (legendFont.FontSize - 10) + Enumerable.Range(0, entries.Count).Max(Advance);
            // Centered on (H + band) / 2 whether or not entries are cut (measured: 10 of 12 shown; the left legend is INFERRED the same). A tr legend starts at the top.
            // The first key sits beside the first line: its centre is (H + band) / 2 - shown x pitch / 2 + single pitch / 2 (measured with 1, 2 and 3 lines).
            double first = (h + band) / 2 - shown * stackPitch / 2 + entryPitch / 2, trCentre = band + 6 + shown * entryPitch / 2;
            for (int i = 0; i < shown; i++)
                placed.Add((i, textLeft - Advance(i), topRight ? trCentre + (i - (shown - 1) / 2.0) * entryPitch : first + i * stackPitch, textLeft));
        }
        else if (hasLegend)
        {
            double fs = legendFont.FontSize, start = w / 2 + 1.84 + 0.12 * (fs - 10);
            double bottomShift = gridRows > 1 && chart.Legend == LegendPosition.Bottom ? 1.5 : 0; // the last row's key top is 1.50 lower than a single row's
            double firstCy = chart.Legend == LegendPosition.Top ? band + 6.0 + entryPitch / 2
                : h - 6.0 - entryPitch / 2 + bottomShift - (gridRows - 1) * entryPitch;
            if (gridRows == 1)
            {
                double cx = start - (rowTotal + (entries.Count - 1) * rowGap) / 2;
                for (int i = 0; i < entries.Count; i++)
                {
                    placed.Add((i, cx, firstCy, cx + Advance(i)));
                    cx += Advance(i) + widths[i] + rowGap;
                }
            }
            else
            {
                // Grid (measured): column pitch = widest entry + 0.283 em, filled row by row; the order inside a grid of a reversed bar legend is UNMEASURED.
                double maxE = Enumerable.Range(0, entries.Count).Max(i => Advance(i) + widths[i]);
                double left0 = start - ((gridCols - 1) * gridPitch + maxE) / 2;
                for (int i = 0; i < entries.Count; i++)
                {
                    double kx = left0 + i % gridCols * gridPitch;
                    placed.Add((i, kx, firstCy + i / gridCols * entryPitch, kx + Advance(i)));
                }
            }
        }

        // ---- draw: the legend, then the plot (an overlaid legend after it)
        void DrawLegend()
        {
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
                        items.Add(new RectItem(kx, ky - square / 2, square, square, color));
                        break;
                }
                for (int li = 0; li < lines[i].Count; li++)
                    PlaceText(lines[i][li], legendFont, x + textLeft, ky + li * lineStep + CenterShift(legendFont) + WordTextDrift(legendFont) - 0.21 * legendFont.FontSize / 10, Anchor.Start, true, 0, items); // 0.21 pt above the other labels: Word's legend text baseline is 3.15 below the key's center
            }
        }
        if (!overlay) DrawLegend();
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
        if (overlay) DrawLegend();
        return true;
    }
}
