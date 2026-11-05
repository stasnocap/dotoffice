using DotOffice.Core;

namespace DotOffice.Layout;

public sealed partial class SheetLayoutEngine
{
    // ------------------------------------------------------------------ scatter and bubble

    /// <summary>Draws a scatter or bubble chart: two value axes, then the joined points or the bubbles of each series.</summary>
    private void EmitScatter(SheetChart chart, double left, double top, double right, double bottom, RunProperties font, List<LayoutItem> items)
    {
        double fontSize = chart.FontSize;
        if (chart.CategoryAxisTitle is { Length: > 0 } xTitle)
        {
            PlaceText(xTitle, font with { FontSize = fontSize + 1 }, (left + right) / 2, bottom - fontSize - 2, Anchor.Middle, false, 0, items);
            bottom -= fontSize + 8;
        }
        if (chart.ValueAxisTitle is { Length: > 0 } yTitle)
        {
            PlaceText(yTitle, font with { FontSize = fontSize + 1 }, left + fontSize / 2 + 1, (top + bottom) / 2, Anchor.Middle, true, 90, items);
            left += fontSize + 8;
        }
        double tickWidth = chart.ShowValueAxis ? chart.Ticks.Select(t => MeasureChartText(t.Label, font)).DefaultIfEmpty(0).Max() : 0;
        double lastWidth = chart.ShowCategoryAxis ? chart.XTicks.Select(t => MeasureChartText(t.Label, font)).DefaultIfEmpty(0).Max() : 0;
        double labelHeight = fontSize * 1.25;
        double plotLeft = left + (chart.ShowValueAxis ? tickWidth + 6 : 0) + 2, plotRight = right - Math.Max(4, lastWidth / 2);
        double plotTop = top + labelHeight / 2, plotBottom = bottom - (chart.ShowCategoryAxis ? labelHeight + 5 : 0);
        if (_manualPlot is { } manual) (plotLeft, plotTop, plotRight, plotBottom) = (manual.X, manual.Y, manual.X + manual.Width, manual.Y + manual.Height);
        double plotWidth = plotRight - plotLeft, plotHeight = plotBottom - plotTop;
        var xScale = chart.XScale;
        var yScale = chart.ValueScale;
        double X(double v) => plotLeft + plotWidth * xScale.Fraction(v);
        double Y(double v) => plotBottom - plotHeight * yScale.Fraction(v);

        foreach (var tick in chart.Ticks)
        {
            double ty = Y(tick.Value);
            if (chart.MajorGridlines) items.Add(new LineItem(plotLeft, ty, plotRight, ty, 0.75, ChartGrid));
            if (chart.ShowValueAxis) PlaceText(tick.Label, font, plotLeft - 5, ty, Anchor.End, true, 0, items);
        }
        foreach (var tick in chart.XTicks)
        {
            double tx = X(tick.Value);
            if (chart.XGridlines) items.Add(new LineItem(tx, plotTop, tx, plotBottom, 0.75, ChartGrid));
            if (chart.ShowCategoryAxis) PlaceText(tick.Label, font, tx, plotBottom + 4, Anchor.Middle, false, 0, items);
        }
        if (chart.ShowCategoryAxis) items.Add(new LineItem(plotLeft, Y(0), plotRight, Y(0), 0.75, ChartGrid));
        if (chart.ShowValueAxis) items.Add(new LineItem(X(0), plotTop, X(0), plotBottom, 0.75, ChartGrid));

        var labels = new List<(string Text, double X, double Y, Color? Fill)>();
        if (chart.Kind == ChartKind.Bubble)
        {
            // Bubble area follows the size; the largest bubble is a quarter of the plot's shorter side. Larger bubbles go underneath.
            double maxSize = chart.Series.SelectMany(s => s.Sizes ?? []).Where(v => v is > 0).Select(v => v!.Value).DefaultIfEmpty(0).Max();
            double maxDiameter = Math.Min(plotWidth, plotHeight) * 0.25 * chart.BubbleScale / 100;
            var bubbles = new List<(double Size, double X, double Y, Color Color, string? Label, double Alpha)>();
            foreach (var s in chart.Series)
                for (int i = 0; i < s.Values.Count; i++)
                    if (s.Values[i] is { } y && s.XValues is { } xs && i < xs.Count && xs[i] is { } x && s.Sizes is { } sizes && i < sizes.Count && sizes[i] is > 0 and var size)
                        bubbles.Add((size, X(x), Y(y), s.Color, s.Labels is { } l && i < l.Count ? l[i] : null, s.Alpha));
            foreach (var b in bubbles.OrderByDescending(b => b.Size))
            {
                double d = maxSize > 0 ? maxDiameter * Math.Sqrt(b.Size / maxSize) : 0;
                if (d < 1) continue;
                if (b.Alpha >= 1) items.Add(new EllipseItem(b.X - d / 2 - 0.75, b.Y - d / 2 - 0.75, d + 1.5, d + 1.5, Color.White));
                AddFaded(items, b.Alpha, new EllipseItem(b.X - d / 2, b.Y - d / 2, d, d, b.Color));
                if (!string.IsNullOrEmpty(b.Label)) labels.Add((b.Label, b.X, b.Y, b.Color));
            }
        }
        else
        {
            foreach (var s in chart.Series)
            {
                var points = new (double X, double Y)?[s.Values.Count];
                for (int i = 0; i < s.Values.Count; i++)
                    if (s.Values[i] is { } y && s.XValues is { } xs && i < xs.Count && xs[i] is { } x)
                    {
                        points[i] = (X(x), Y(y));
                        if (s.Labels is { } l && i < l.Count && l[i].Length > 0) labels.Add((l[i], X(x), Y(y) - 8, null));
                    }
                var overlays = new List<LayoutItem>();
                foreach (var bars in s.ErrorBars ?? [])
                    for (int i = 0; i < s.Values.Count && i < bars.Plus.Count; i++)
                    {
                        if (s.Values[i] is not { } y || s.XValues is not { } xs || i >= xs.Count || xs[i] is not { } x) continue;
                        if (bars.Horizontal) DrawErrorBar(overlays, true, Y(y), X(x - bars.Minus[i]), X(x + bars.Plus[i]), bars);
                        else DrawErrorBar(overlays, false, X(x), Y(y - bars.Minus[i]), Y(y + bars.Plus[i]), bars);
                    }
                if (s.Trendline is { } trend && s.XValues is { } trendX)
                {
                    var known = new List<(double X, double Y)>();
                    for (int i = 0; i < s.Values.Count && i < trendX.Count; i++) if (s.Values[i] is { } y && trendX[i] is { } x) known.Add((x, y));
                    if (known.Count > 0)
                        DrawTrendline(overlays, ChartTrendlines.Fit(trend, known, known.Min(k => k.X), known.Max(k => k.X)).Select(q => (X(q.X), Y(q.Y))).ToList(), trend);
                }
                if (overlays.Count > 0) { items.Add(new PushItem((plotLeft, plotTop, plotWidth, plotHeight))); items.AddRange(overlays); items.Add(new PopItem()); }
                if (s.Line) DrawPolyline(points, s.Smooth, 2.25, s.Color, items);
                if (s.Marker || !s.Line)
                    foreach (var p in points)
                        if (p is { } q)
                        {
                            items.Add(new EllipseItem(q.X - 3.2, q.Y - 3.2, 6.4, 6.4, Color.White));
                            items.Add(new EllipseItem(q.X - 2.5, q.Y - 2.5, 5, 5, s.Color));
                        }
            }
        }
        foreach (var (text, lx, ly, fill) in labels)
            PlaceText(text, font with { Color = fill is { } f ? ContrastColor(f) : ChartText }, lx, ly, Anchor.Middle, true, 0, items);
    }

    // ------------------------------------------------------------------ radar

    /// <summary>Draws a radar chart: a web of value rings, one spoke per category, and a closed polygon per series.</summary>
    private void EmitRadar(SheetChart chart, double left, double top, double right, double bottom, RunProperties font, List<LayoutItem> items)
    {
        int n = Math.Max(chart.Categories.Count, 3);
        double labelHeight = chart.FontSize * 1.25;
        double labelWidth = chart.Categories.Select(c => MeasureChartText(c, font)).DefaultIfEmpty(0).Max();
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
                    items.Add(new LineItem(a.X, a.Y, b.X, b.Y, 0.75, ChartGrid));
                }
            if (chart.ShowValueAxis && r - lastLabel >= labelHeight)
            {
                PlaceText(tick.Label, font, cx - 4, cy - r, Anchor.End, true, 0, items);
                lastLabel = r;
            }
        }
        for (int k = 0; k < n; k++)
        {
            var end = Point(k, radius);
            items.Add(new LineItem(cx, cy, end.X, end.Y, 0.75, ChartGrid));
            if (!chart.ShowCategoryAxis || k >= chart.Categories.Count) continue;
            var at = Point(k, radius + 6);
            double cos = (at.X - cx) / (radius + 6), sin = (at.Y - cy) / (radius + 6);
            var anchor = cos > 0.3 ? Anchor.Start : cos < -0.3 ? Anchor.End : Anchor.Middle;
            PlaceText(chart.Categories[k], font, at.X, at.Y + sin * labelHeight / 2, anchor, true, 0, items);
        }

        foreach (var s in chart.Series)
        {
            var polygon = new List<(double X, double Y)>();
            for (int k = 0; k < n; k++)
                polygon.Add(Point(k, k < s.Values.Count && s.Values[k] is { } v ? Radius(v) : 0));
            if (chart.Filled) items.Add(new PolygonItem(polygon, s.Color.Blend(Color.White, 0.35)));
            if (s.Line || chart.Filled)
                for (int k = 0; k < n; k++)
                    items.Add(new LineItem(polygon[k].X, polygon[k].Y, polygon[(k + 1) % n].X, polygon[(k + 1) % n].Y, 2.25, s.Color));
            if (s.Marker)
                foreach (var p in polygon)
                {
                    items.Add(new EllipseItem(p.X - 3.2, p.Y - 3.2, 6.4, 6.4, Color.White));
                    items.Add(new EllipseItem(p.X - 2.5, p.Y - 2.5, 5, 5, s.Color));
                }
        }
    }
}
