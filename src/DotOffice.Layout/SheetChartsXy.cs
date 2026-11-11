using DotOffice.Core;

namespace DotOffice.Layout;

public sealed partial class SheetLayoutEngine
{
    // ------------------------------------------------------------------ scatter and bubble

    /// <summary>Draws a scatter or bubble chart: two value axes, then the joined points or the bubbles of each series.</summary>
    private void EmitScatter(SheetChart chart, double left, double top, double right, double bottom, RunProperties font, List<LayoutItem> items)
    {
        double fontSize = chart.FontSize;
        var styles = chart.TextStyles;
        var yFont = Styled(font, styles.ValueAxis);
        var xFont = Styled(font, styles.CategoryAxis);
        if (chart.CategoryAxisTitle is { Length: > 0 } xTitle)
        {
            var titleFont = Styled(font with { FontSize = fontSize + 1 }, styles.CategoryAxisTitle);
            PlaceText(xTitle, titleFont, (left + right) / 2, bottom - titleFont.FontSize - 2, Anchor.Middle, false, 0, items);
            bottom -= titleFont.FontSize + 7;
        }
        if (chart.ValueAxisTitle is { Length: > 0 } yTitle)
        {
            var titleFont = Styled(font with { FontSize = fontSize + 1 }, styles.ValueAxisTitle);
            PlaceText(yTitle, titleFont, left + titleFont.FontSize / 2 + 1, (top + bottom) / 2, Anchor.Middle, true, 90, items);
            left += titleFont.FontSize + 7;
        }
        // The secondary vertical axis (right) and horizontal axis (top) of groups on their own pair of value axes.
        bool secondaryY = chart.ShowSecondaryAxis && chart.ExtraGroups.Any(g => g.Secondary), secondaryX = chart.ShowSecondaryXAxis && chart.ExtraGroups.Any(g => g.Secondary);
        var secondaryYFont = Styled(font, styles.SecondaryAxis);
        if (secondaryY && chart.SecondaryAxisTitle is { Length: > 0 } secondaryTitle)
        {
            var titleFont = Styled(font with { FontSize = fontSize + 1 }, styles.SecondaryAxisTitle);
            PlaceText(secondaryTitle, titleFont, right - titleFont.FontSize / 2 - 1, (top + bottom) / 2, Anchor.Middle, true, -90, items);
            right -= titleFont.FontSize + 7;
        }
        double secondaryWidth = secondaryY ? chart.SecondaryTicks.Select(t => MeasureChartText(t.Label, secondaryYFont)).DefaultIfEmpty(0).Max() + 6 : 0;
        if (secondaryX && chart.SecondaryXAxisTitle is { Length: > 0 } secondaryXTitle)
        {
            var titleFont = Styled(font with { FontSize = fontSize + 1 }, styles.SecondaryXAxisTitle);
            PlaceText(secondaryXTitle, titleFont, (left + right) / 2, top, Anchor.Middle, false, 0, items);
            top += titleFont.FontSize + 7;
        }
        if (secondaryX) top += xFont.FontSize * 1.25 + 4;
        right -= secondaryWidth;
        double tickWidth = chart.ShowValueAxis ? chart.Ticks.Select(t => MeasureChartText(t.Label, yFont)).DefaultIfEmpty(0).Max() : 0;
        double lastWidth = chart.ShowCategoryAxis ? chart.XTicks.Select(t => MeasureChartText(t.Label, xFont)).DefaultIfEmpty(0).Max() : 0;
        double labelHeight = xFont.FontSize * 1.25;
        double xAngle = Math.Clamp(styles.CategoryAxis?.Rotation ?? 0, -90, 90) * Math.PI / 180;
        double xLabelSpace = chart.ShowCategoryAxis ? lastWidth * Math.Abs(Math.Sin(xAngle)) + labelHeight * Math.Abs(Math.Cos(xAngle)) + 5 : 0;
        double plotLeft = left + (chart.ShowValueAxis ? tickWidth + 6 : 0) + 2, plotRight = right - Math.Max(4, lastWidth / 2);
        double plotTop = top + labelHeight / 2, plotBottom = bottom - xLabelSpace;
        if (_wordPlot is { } wp) (plotLeft, plotTop, plotRight, plotBottom) = (wp.X, wp.Y, wp.X + wp.Width, wp.Y + wp.Height);
        if (_manualPlot is { } manual) (plotLeft, plotTop, plotRight, plotBottom) = (manual.X, manual.Y, manual.X + manual.Width, manual.Y + manual.Height);
        double plotWidth = plotRight - plotLeft, plotHeight = plotBottom - plotTop;
        if (chart.PlotBackground is { } plotFill && plotWidth > 0 && plotHeight > 0) items.Add(new RectItem(plotLeft, plotTop, plotWidth, plotHeight, plotFill));
        var xScale = chart.XScale;
        var yScale = chart.ValueScale;
        double X(double v) => plotLeft + plotWidth * xScale.Fraction(v);
        double Y(double v) => plotBottom - plotHeight * yScale.Fraction(v);

        foreach (var tick in chart.Ticks)
        {
            double ty = Y(tick.Value);
            if (chart.MajorGridlines) items.Add(new LineItem(plotLeft, ty, plotRight, ty, chart.GridlineWidth, chart.GridlineColor));
            if (chart.ShowValueAxis) PlaceText(tick.Label, yFont, plotLeft - AxisLabelGap, ty + CenterShift(yFont), Anchor.End, true, 0, items);
        }
        foreach (var tick in chart.XTicks)
        {
            double tx = X(tick.Value);
            if (chart.XGridlines) items.Add(new LineItem(tx, plotTop, tx, plotBottom, chart.GridlineWidth, chart.GridlineColor));
            if (chart.ShowCategoryAxis) PlaceXLabel(tick.Label, xFont, tx, plotBottom + LabelDrop(xFont), styles.CategoryAxis?.Rotation ?? 0, items);
        }
        // Word's outside major tick marks at every tick value (3.02 pt; the axis line's width and color are INFERRED).
        if (chart.ValueAxisTickOut && chart.ShowValueAxis)
            foreach (var tick in chart.Ticks) items.Add(new LineItem(X(0) - 3.02, Y(tick.Value), X(0), Y(tick.Value), chart.AxisLineWidth, Color.Black));
        if (chart.CategoryAxisTickOut && chart.ShowCategoryAxis)
            foreach (var tick in chart.XTicks) items.Add(new LineItem(X(tick.Value), Y(0), X(tick.Value), Y(0) + 3.02, chart.AxisLineWidth, Color.Black));
        if (chart.ShowCategoryAxis) items.Add(new LineItem(plotLeft, Y(0), plotRight, Y(0), chart.AxisLineWidth, Color.Black));
        if (chart.ShowValueAxis) items.Add(new LineItem(X(0), plotTop, X(0), plotBottom, chart.AxisLineWidth, Color.Black));

        var labels = new List<(string Text, double X, double Y, Color? Fill, ChartTextStyle? Style, LabelSide Side)>();
        var trendLabels = new List<LayoutItem>();
        // The groups: the chart's own series on the primary axes, then the further groups on the primary or the secondary pair.
        var secondaryXScale = chart.SecondaryXScale;
        var secondaryYScale = chart.SecondaryScale;
        double SecondaryX(double v) => plotLeft + plotWidth * secondaryXScale.Fraction(v);
        double SecondaryY(double v) => plotBottom - plotHeight * secondaryYScale.Fraction(v);
        var drawGroups = new List<(IReadOnlyList<ChartSeries> Series, Func<double, double> X, Func<double, double> Y)> { (chart.Series, X, Y) };
        foreach (var extra in chart.ExtraGroups) drawGroups.Add((extra.Series, extra.Secondary ? SecondaryX : X, extra.Secondary ? SecondaryY : Y));
        if (secondaryY)
            foreach (var tick in chart.SecondaryTicks) PlaceText(tick.Label, secondaryYFont, plotRight + 5, SecondaryY(tick.Value), Anchor.Start, true, 0, items);
        if (secondaryX)
            foreach (var tick in chart.SecondaryXTicks) PlaceText(tick.Label, xFont, SecondaryX(tick.Value), plotTop - 4 - xFont.FontSize * 1.25, Anchor.Middle, false, 0, items);
        double maxSize = chart.Kind == ChartKind.Bubble ? drawGroups.SelectMany(g => g.Series).SelectMany(s => s.Sizes ?? []).Where(v => v is > 0).Select(v => v!.Value).DefaultIfEmpty(0).Max() : 0;
        var bubbles = new List<(double Size, double X, double Y, Color Color, string? Label, double Alpha, ChartTextStyle? Style)>();
        foreach (var (groupSeries, GX, GY) in drawGroups)
        {
            if (chart.Kind == ChartKind.Bubble)
            {
                foreach (var s in groupSeries)
                    for (int i = 0; i < s.Values.Count; i++)
                        if (s.Values[i] is { } y && s.XValues is { } xs && i < xs.Count && xs[i] is { } x && s.Sizes is { } sizes && i < sizes.Count && sizes[i] is > 0 and var size)
                            bubbles.Add((size, GX(x), GY(y), s.Color, s.Labels is { } l && i < l.Count ? l[i] : null, s.Alpha, LabelStyleAt(s, i)));
                continue;
            }
            foreach (var s in groupSeries)
            {
                var points = new (double X, double Y)?[s.Values.Count];
                for (int i = 0; i < s.Values.Count; i++)
                    if (s.Values[i] is { } y && s.XValues is { } xs && i < xs.Count && xs[i] is { } x)
                    {
                        points[i] = (GX(x), GY(y));
                        if (s.Labels is { } l && i < l.Count && l[i].Length > 0)
                        {
                            double gx = GX(x), gy = GY(y);
                            labels.Add(LabelPlacementAt(s, i) switch
                            {
                                ChartLabelPlacement.Center => (l[i], gx, gy, null, LabelStyleAt(s, i), LabelSide.Center),
                                ChartLabelPlacement.Below => (l[i], gx, gy + 8, null, LabelStyleAt(s, i), LabelSide.Center),
                                ChartLabelPlacement.Left => (l[i], gx - 6, gy, null, LabelStyleAt(s, i), LabelSide.Left),
                                ChartLabelPlacement.Right => (l[i], gx + 6, gy, null, LabelStyleAt(s, i), LabelSide.Right),
                                _ => (l[i], gx, gy - 8, null, LabelStyleAt(s, i), LabelSide.Center),
                            });
                        }
                    }
                var overlays = new List<LayoutItem>();
                foreach (var bars in s.ErrorBars ?? [])
                    for (int i = 0; i < s.Values.Count && i < bars.Plus.Count; i++)
                    {
                        if (s.Values[i] is not { } y || s.XValues is not { } xs || i >= xs.Count || xs[i] is not { } x) continue;
                        if (bars.Horizontal) DrawErrorBar(overlays, true, GY(y), GX(x - bars.Minus[i]), GX(x + bars.Plus[i]), bars);
                        else DrawErrorBar(overlays, false, GX(x), GY(y - bars.Minus[i]), GY(y + bars.Plus[i]), bars);
                    }
                if (s.Trendline is { } trend && s.XValues is { } trendX)
                {
                    var known = new List<(double X, double Y)>();
                    for (int i = 0; i < s.Values.Count && i < trendX.Count; i++) if (s.Values[i] is { } y && trendX[i] is { } x) known.Add((x, y));
                    if (known.Count > 0)
                    {
                        var drawn = ChartTrendlines.Fit(trend, known, known.Min(k => k.X), known.Max(k => k.X)).Select(q => (GX(q.X), GY(q.Y))).ToList();
                        DrawTrendline(overlays, drawn, trend);
                        if (drawn.Count > 0) AddTrendlineLabel(trendLabels, trend, known, drawn[^1], font, plotLeft, plotTop, plotRight, plotBottom);
                    }
                }
                if (overlays.Count > 0) { items.Add(new PushItem((plotLeft, plotTop, plotWidth, plotHeight))); items.AddRange(overlays); items.Add(new PopItem()); }
                if (s.Line) DrawPolyline(points, s.Smooth, s.StrokeWidth, s.StrokeColor, items);
                if (s.Marker || (!s.Line && !s.MarkerNone) || s.PointMarkers is not null)
                    for (int i = 0; i < points.Length; i++)
                        if (points[i] is { } q)
                            DrawMarkerAt(items, q.X, q.Y, s, i, s.Marker || (!s.Line && !s.MarkerNone));
            }
        }
        if (chart.Kind == ChartKind.Bubble)
        {
            // Bubble area follows the size; the largest bubble is a quarter of the plot's shorter side. Larger bubbles go underneath.
            double maxDiameter = Math.Min(plotWidth, plotHeight) * 0.25 * chart.BubbleScale / 100;
            foreach (var b in bubbles.OrderByDescending(b => b.Size))
            {
                double d = maxSize > 0 ? maxDiameter * Math.Sqrt(b.Size / maxSize) : 0;
                if (d < 1) continue;
                if (b.Alpha >= 1) items.Add(new EllipseItem(b.X - d / 2 - 0.75, b.Y - d / 2 - 0.75, d + 1.5, d + 1.5, Color.White));
                AddFaded(items, b.Alpha, new EllipseItem(b.X - d / 2, b.Y - d / 2, d, d, b.Color));
                if (!string.IsNullOrEmpty(b.Label)) labels.Add((b.Label, b.X, b.Y, b.Color, b.Style, LabelSide.Center));
            }
        }
        items.AddRange(trendLabels);
        foreach (var (text, lx, ly, fill, pointStyle, side) in labels)
            PlaceText(text, Styled(Styled(font with { Color = fill is { } f ? LabelColor(chart, f) : chart.TextColor ?? ChartText }, styles.DataLabels), pointStyle), lx, ly,
                side switch { LabelSide.Left => Anchor.End, LabelSide.Right => Anchor.Start, _ => Anchor.Middle }, true, 0, items);
    }

    /// <summary>A label under a horizontal axis, turned by <paramref name="angle"/> degrees counter-clockwise (negative: clockwise) from its tick.</summary>
    private void PlaceXLabel(string text, RunProperties font, double x, double y, double angle, List<LayoutItem> items)
    {
        angle = Math.Clamp(angle, -90, 90);
        if (angle == 0) { PlaceText(text, font, x, y, Anchor.Middle, false, 0, items); return; }
        double turn = Math.Pow(Math.Sin(angle * Math.PI / 180), 2) * (3 + font.FontSize * 0.625);
        if (angle > 0) PlaceText(text, font, x + 3 - turn, y + 1, Anchor.End, false, angle, items);
        else PlaceText(text, font, x - 3 + turn, y + 1, Anchor.Start, false, angle, items);
    }

    // ------------------------------------------------------------------ radar

    /// <summary>Draws a radar chart: a web of value rings, one spoke per category, and a closed polygon per series.</summary>
    private void EmitRadar(SheetChart chart, double left, double top, double right, double bottom, RunProperties font, List<LayoutItem> items)
    {
        int n = Math.Max(chart.Categories.Count, 3);
        double labelHeight = chart.FontSize * 1.25;
        var categoryFont = Styled(font, chart.TextStyles.CategoryAxis);
        var valueFont = Styled(font, chart.TextStyles.ValueAxis);
        double labelWidth = chart.Categories.Select(c => MeasureChartText(c, categoryFont)).DefaultIfEmpty(0).Max();
        double radius = Math.Min((right - left) / 2 - labelWidth - 6, (bottom - top) / 2 - labelHeight - 4);
        if (radius < 15) return;
        double cx = (left + right) / 2, cy = (top + bottom) / 2;
        var valueScale = chart.ValueScale;
        (double X, double Y) Point(int k, double r)
        {
            double a = (-90 + 360.0 * k / n) * Math.PI / 180;
            return (cx + r * Math.Cos(a), cy + r * Math.Sin(a));
        }
        double Radius(double v) => radius * valueScale.Fraction(v);

        double lastLabel = -1e9;
        foreach (var tick in chart.Ticks)
        {
            double r = Radius(tick.Value);
            if (r <= 0) continue;
            if (chart.MajorGridlines)
                for (int k = 0; k < n; k++)
                {
                    var a = Point(k, r);
                    var b = Point((k + 1) % n, r);
                    items.Add(new LineItem(a.X, a.Y, b.X, b.Y, chart.GridlineWidth, chart.GridlineColor));
                }
            if (chart.ShowValueAxis && r - lastLabel >= labelHeight)
            {
                PlaceText(tick.Label, valueFont, cx - 4, cy - r, Anchor.End, true, 0, items);
                lastLabel = r;
            }
        }
        for (int k = 0; k < n; k++)
        {
            var end = Point(k, radius);
            items.Add(new LineItem(cx, cy, end.X, end.Y, 0.75, ChartRadarSpoke));
            if (!chart.ShowCategoryAxis || k >= chart.Categories.Count) continue;
            var at = Point(k, radius + 6);
            double cos = (at.X - cx) / (radius + 6), sin = (at.Y - cy) / (radius + 6);
            var anchor = cos > 0.3 ? Anchor.Start : cos < -0.3 ? Anchor.End : Anchor.Middle;
            PlaceText(chart.Categories[k], categoryFont, at.X, at.Y + sin * labelHeight / 2, anchor, true, 0, items);
        }

        foreach (var s in chart.Series)
        {
            var polygon = new List<(double X, double Y)>();
            for (int k = 0; k < n; k++)
                polygon.Add(Point(k, k < s.Values.Count && s.Values[k] is { } v ? Radius(v) : 0));
            if (chart.Filled) items.Add(new PolygonItem(polygon, s.Color.Blend(Color.White, 0.35)));
            if (s.Line || chart.Filled)
                for (int k = 0; k < n; k++)
                    items.Add(new LineItem(polygon[k].X, polygon[k].Y, polygon[(k + 1) % n].X, polygon[(k + 1) % n].Y, s.StrokeWidth, s.StrokeColor));
            if (s.Marker || s.PointMarkers is not null)
                for (int k = 0; k < polygon.Count; k++)
                    DrawMarkerAt(items, polygon[k].X, polygon[k].Y, s, k, s.Marker);
        }
    }
}
