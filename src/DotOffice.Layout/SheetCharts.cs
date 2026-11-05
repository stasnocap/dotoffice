using DotOffice.Core;

namespace DotOffice.Layout;

public sealed partial class SheetLayoutEngine
{
    private static readonly Color ChartText = new(0x59, 0x59, 0x59);
    private static readonly Color ChartGrid = new(0xD9, 0xD9, 0xD9);
    private const double ChartPad = 7;

    private enum Anchor { Start, Middle, End }

    /// <summary>
    /// Draws a chart into its frame: background, title, legend, axes with gridlines and the series (columns, bars, lines or pie slices).
    /// Everything is clipped to the frame.
    /// </summary>
    private void EmitChart(SheetChart chart, double x, double y, double w, double h,
        (double X, double Y, double Width, double Height) visible, List<LayoutItem> items)
    {
        items.Add(new PushItem(visible));
        items.Add(new PushItem((x, y, w, h)));
        if (chart.Background is { } background) items.Add(new RectItem(x, y, w, h, background));
        if (chart.Border is { } border)
        {
            items.Add(new RectItem(x, y, w, 0.75, border));
            items.Add(new RectItem(x, y + h - 0.75, w, 0.75, border));
            items.Add(new RectItem(x, y, 0.75, h, border));
            items.Add(new RectItem(x + w - 0.75, y, 0.75, h, border));
        }

        var font = RunProperties.Default with { FontFamily = "Calibri", FontSize = chart.FontSize, Color = ChartText };
        double left = x + ChartPad, top = y + ChartPad, right = x + w - ChartPad, bottom = y + h - ChartPad;

        // ---- title
        if (!string.IsNullOrEmpty(chart.Title))
        {
            var titleFont = font with { FontSize = chart.TitleSize };
            double used = PlaceText(chart.Title, titleFont, (left + right) / 2, top, Anchor.Middle, false, 0, items);
            top += used + 4;
        }

        // ---- legend
        var legendEntries = chart.Kind is ChartKind.Pie or ChartKind.Doughnut
            ? chart.Categories.Select((c, i) => (Name: c, Color: chart.Series.FirstOrDefault()?.PointColors is { } pc && i < pc.Count ? pc[i] : ChartColor(i), Series: (ChartSeries?)null)).ToList()
            : chart.Series.Select(s => (Name: s.Name, Color: s.Color, Series: (ChartSeries?)s)).ToList();
        if (chart.Legend != LegendPosition.None && legendEntries.Count > 0 && right > left && bottom > top)
        {
            double entryHeight = chart.FontSize * 1.25 + 3, swatch = chart.FontSize * 0.8;
            var widths = legendEntries.Select(e => MeasureChartText(e.Name, font)).ToList();
            if (chart.Legend is LegendPosition.Right or LegendPosition.Left)
            {
                double boxWidth = widths.Max() + swatch + 12;
                double lx = chart.Legend == LegendPosition.Right ? right - boxWidth : left;
                double ly = (top + bottom) / 2 - legendEntries.Count * entryHeight / 2;
                for (int i = 0; i < legendEntries.Count; i++)
                    DrawLegendEntry(chart, legendEntries[i].Name, legendEntries[i].Color, legendEntries[i].Series, lx, ly + i * entryHeight, swatch, entryHeight, font, items);
                if (chart.Legend == LegendPosition.Right) right -= boxWidth + 6; else left += boxWidth + 6;
            }
            else
            {
                double total = widths.Sum() + legendEntries.Count * (swatch + 14);
                double lx = (left + right) / 2 - total / 2, ly = chart.Legend == LegendPosition.Top ? top : bottom - entryHeight;
                for (int i = 0; i < legendEntries.Count; i++)
                {
                    DrawLegendEntry(chart, legendEntries[i].Name, legendEntries[i].Color, legendEntries[i].Series, lx, ly, swatch, entryHeight, font, items);
                    lx += widths[i] + swatch + 14;
                }
                if (chart.Legend == LegendPosition.Top) top += entryHeight + 4; else bottom -= entryHeight + 4;
            }
        }
        if (right - left < 20 || bottom - top < 20)
        {
            items.Add(new PopItem());
            items.Add(new PopItem());
            return;
        }

        if (chart.Kind is ChartKind.Pie or ChartKind.Doughnut) EmitPie(chart, left, top, right, bottom, font, items);
        else if (chart.Kind == ChartKind.Radar) EmitRadar(chart, left, top, right, bottom, font, items);
        else if (chart.Kind is ChartKind.Scatter or ChartKind.Bubble) EmitScatter(chart, left, top, right, bottom, font, items);
        else EmitAxisChart(chart, left, top, right, bottom, font, items);
        items.Add(new PopItem());
        items.Add(new PopItem());
    }

    private static Color ChartColor(int index) => new[] { new Color(0x44, 0x72, 0xC4), new Color(0xED, 0x7D, 0x31), new Color(0xA5, 0xA5, 0xA5), new Color(0xFF, 0xC0, 0x00), new Color(0x5B, 0x9B, 0xD5), new Color(0x70, 0xAD, 0x47) }[index % 6];

    private void DrawLegendEntry(SheetChart chart, string name, Color color, ChartSeries? series, double x, double y, double swatch, double height, RunProperties font, List<LayoutItem> items)
    {
        double cy = y + height / 2;
        bool lineStyle = chart.Kind is ChartKind.Line or ChartKind.Scatter || (chart.Kind == ChartKind.Radar && !chart.Filled);
        if (lineStyle)
        {
            bool joined = series?.Line ?? true;
            if (joined) items.Add(new LineItem(x, cy, x + swatch + 4, cy, 2.25, color));
            if (series is null || series.Marker || !joined)
                items.Add(new EllipseItem(x + (swatch + 4) / 2 - 2.5, cy - 2.5, 5, 5, color));
            PlaceText(name, font, x + swatch + 8, cy, Anchor.Start, true, 0, items);
        }
        else if (chart.Kind == ChartKind.Bubble)
        {
            items.Add(new EllipseItem(x, cy - swatch / 2, swatch, swatch, color));
            PlaceText(name, font, x + swatch + 4, cy, Anchor.Start, true, 0, items);
        }
        else
        {
            items.Add(new RectItem(x, cy - swatch / 2, swatch, swatch, color));
            PlaceText(name, font, x + swatch + 4, cy, Anchor.Start, true, 0, items);
        }
    }

    private double MeasureChartText(string text, RunProperties font) => BuildBlock(ChartCell(text, font), text, 1e7).Width;

    private static Cell ChartCell(string text, RunProperties font) =>
        new() { Text = text, Kind = CellValueKind.Text, Style = CellStyle.Default with { Font = font } };

    /// <summary>
    /// Draws one line of text anchored at (<paramref name="ax"/>, <paramref name="ay"/>): horizontally by <paramref name="anchor"/>, and
    /// either with its top at <paramref name="ay"/> or, when <paramref name="centerVertically"/>, centered on it. The text may be turned
    /// counter-clockwise about the anchor. Returns the height of the text.
    /// </summary>
    private double PlaceText(string text, RunProperties font, double ax, double ay, Anchor anchor, bool centerVertically, double rotation, List<LayoutItem> items)
    {
        if (text.Length == 0) return 0;
        var block = BuildBlock(ChartCell(text, font), text, 1e7);
        if (block.Lines.Count == 0) return 0;
        var (pitch, ascent, descent) = block.Metrics[0];
        double width = block.Width;
        double x0 = anchor switch { Anchor.Middle => ax - width / 2, Anchor.End => ax - width, _ => ax };
        double top = centerVertically ? ay - pitch / 2 : ay;
        if (rotation != 0) items.Add(new PushItem(null, rotation, ax, ay));
        if (block.Lines[0].Atoms.Count > 0)
            EmitLine(block.Lines[0], x0, top + (pitch - (ascent + descent)) / 2 + ascent, ax - 1e5, ax + 1e5, items);
        if (rotation != 0) items.Add(new PopItem());
        return pitch;
    }

    // ------------------------------------------------------------------ axes, columns, bars and lines

    private void EmitAxisChart(SheetChart chart, double left, double top, double right, double bottom, RunProperties font, List<LayoutItem> items)
    {
        bool horizontalBars = chart.Kind == ChartKind.Bar;
        int categories = Math.Max(chart.Categories.Count, 1);
        double fontSize = chart.FontSize;

        // ---- room for axis titles and labels
        if (chart.CategoryAxisTitle is { Length: > 0 } categoryTitle && !horizontalBars)
        {
            PlaceText(categoryTitle, font with { FontSize = fontSize + 1 }, (left + right) / 2, bottom - fontSize - 2, Anchor.Middle, false, 0, items);
            bottom -= fontSize + 8;
        }
        if (chart.ValueAxisTitle is { Length: > 0 } valueTitle && !horizontalBars)
        {
            PlaceText(valueTitle, font with { FontSize = fontSize + 1 }, left + fontSize / 2 + 1, (top + bottom) / 2, Anchor.Middle, true, 90, items);
            left += fontSize + 8;
        }
        if (horizontalBars)
        {
            if (chart.ValueAxisTitle is { Length: > 0 } t1)
            {
                PlaceText(t1, font with { FontSize = fontSize + 1 }, (left + right) / 2, bottom - fontSize - 2, Anchor.Middle, false, 0, items);
                bottom -= fontSize + 8;
            }
            if (chart.CategoryAxisTitle is { Length: > 0 } t2)
            {
                PlaceText(t2, font with { FontSize = fontSize + 1 }, left + fontSize / 2 + 1, (top + bottom) / 2, Anchor.Middle, true, 90, items);
                left += fontSize + 8;
            }
        }

        double tickWidth = chart.ShowValueAxis ? chart.Ticks.Select(t => MeasureChartText(t.Label, font)).DefaultIfEmpty(0).Max() : 0;
        double categoryWidth = chart.ShowCategoryAxis ? chart.Categories.Select(c => MeasureChartText(c, font)).DefaultIfEmpty(0).Max() : 0;
        double labelHeight = fontSize * 1.25;

        double plotLeft, plotRight, plotTop, plotBottom;
        bool rotateLabels = false;
        double stride = 1;
        if (!horizontalBars)
        {
            plotLeft = left + (chart.ShowValueAxis ? tickWidth + 6 : 0);
            plotRight = right - 4;
            plotTop = top + labelHeight / 2;
            double bandWidth = (plotRight - plotLeft) / (chart.CategoriesBetween ? categories : Math.Max(categories - 1, 1));
            rotateLabels = chart.ShowCategoryAxis && categoryWidth > bandWidth * 0.95 - 2;
            double below = !chart.ShowCategoryAxis ? 0 : rotateLabels ? categoryWidth * 0.7071 + labelHeight * 0.7071 + 6 : labelHeight + 5;
            plotBottom = bottom - below;
            stride = !chart.ShowCategoryAxis ? 1 : Math.Max(1, Math.Ceiling((rotateLabels ? labelHeight * 1.1 : categoryWidth + 6) / Math.Max(bandWidth, 1)));
        }
        else
        {
            plotLeft = left + (chart.ShowCategoryAxis ? categoryWidth + 6 : 0);
            plotRight = right - 8;
            plotTop = top + 4;
            plotBottom = bottom - (chart.ShowValueAxis ? labelHeight + 5 : 0);
        }
        if (plotRight - plotLeft < 10 || plotBottom - plotTop < 10) return;
        double plotWidth = plotRight - plotLeft, plotHeight = plotBottom - plotTop;
        double range = chart.ValueMax - chart.ValueMin;
        if (range <= 0) range = 1;

        // Positions along the value axis and the category axis.
        double ValueX(double v) => plotLeft + plotWidth * (v - chart.ValueMin) / range;
        double ValueY(double v) => plotBottom - plotHeight * (v - chart.ValueMin) / range;
        double zero = Math.Clamp(0, chart.ValueMin, chart.ValueMax);
        bool between = chart.CategoriesBetween || chart.Kind is ChartKind.Column or ChartKind.Bar;
        double CategoryX(int i) => between ? plotLeft + (i + 0.5) * plotWidth / categories
            : categories > 1 ? plotLeft + i * plotWidth / (categories - 1) : plotLeft + plotWidth / 2;

        // ---- gridlines and value labels
        foreach (var tick in chart.Ticks)
        {
            if (horizontalBars)
            {
                double tx = ValueX(tick.Value);
                if (chart.MajorGridlines) items.Add(new LineItem(tx, plotTop, tx, plotBottom, 0.75, ChartGrid));
                if (chart.ShowValueAxis) PlaceText(tick.Label, font, tx, plotBottom + 4, Anchor.Middle, false, 0, items);
            }
            else
            {
                double ty = ValueY(tick.Value);
                if (chart.MajorGridlines) items.Add(new LineItem(plotLeft, ty, plotRight, ty, 0.75, ChartGrid));
                if (chart.ShowValueAxis) PlaceText(tick.Label, font, plotLeft - 5, ty, Anchor.End, true, 0, items);
            }
        }

        // ---- series
        int seriesCount = chart.Series.Count;
        var totals = new double[categories];
        if (chart.Grouping == ChartGrouping.PercentStacked)
            foreach (var s in chart.Series)
                for (int i = 0; i < categories; i++) totals[i] += i < s.Values.Count && s.Values[i] is { } v ? Math.Abs(v) : 0;
        double Scaled(ChartSeries s, int i) => i < s.Values.Count && s.Values[i] is { } v ? chart.Grouping == ChartGrouping.PercentStacked ? (totals[i] > 0 ? v / totals[i] : 0) : v : 0;

        var labels = new List<(string Text, double X, double Y, bool Inside, Color Fill)>();
        if (chart.Kind is ChartKind.Column or ChartKind.Bar)
        {
            double band = (horizontalBars ? plotHeight : plotWidth) / categories;
            bool stacked = chart.Grouping != ChartGrouping.Clustered;
            double overlap = stacked ? 1 : chart.Overlap / 100, gap = chart.GapWidth / 100;
            int perGroup = stacked ? 1 : Math.Max(seriesCount, 1);
            double barWidth = band / (perGroup - (perGroup - 1) * overlap + gap);
            double groupWidth = barWidth * (perGroup - (perGroup - 1) * overlap);
            var positive = new double[categories];
            var negative = new double[categories];
            for (int j = 0; j < seriesCount; j++)
            {
                var s = chart.Series[j];
                for (int i = 0; i < categories; i++)
                {
                    if (i >= s.Values.Count || s.Values[i] is null) continue;
                    double value = Scaled(s, i), from, to;
                    if (stacked)
                    {
                        if (value >= 0) { from = positive[i]; to = from + value; positive[i] = to; }
                        else { from = negative[i]; to = from + value; negative[i] = to; }
                    }
                    else (from, to) = (0, value);
                    double c0 = Math.Clamp(from, chart.ValueMin, chart.ValueMax), c1 = Math.Clamp(to, chart.ValueMin, chart.ValueMax);
                    double center = (horizontalBars ? plotBottom - (i + 0.5) * band : plotLeft + (i + 0.5) * band);
                    double offset = stacked ? -barWidth / 2 : -groupWidth / 2 + j * barWidth * (1 - overlap);
                    var color = s.Color;
                    if (horizontalBars)
                    {
                        double yTop = center + offset, ya = Math.Min(ValueX(c0), ValueX(c1)), yb = Math.Max(ValueX(c0), ValueX(c1));
                        // In horizontal bar charts the first series sits lowest in each group.
                        double slot = stacked ? yTop : center - offset - barWidth;
                        items.Add(new RectItem(ya, slot, Math.Max(yb - ya, 0), barWidth, color));
                        if (s.Labels is { } l && i < l.Count && l[i].Length > 0) labels.Add((l[i], stacked ? (ya + yb) / 2 : yb + 3, slot + barWidth / 2, stacked, color));
                    }
                    else
                    {
                        double xLeft = center + offset, ya = Math.Min(ValueY(c0), ValueY(c1)), yb = Math.Max(ValueY(c0), ValueY(c1));
                        items.Add(new RectItem(xLeft, ya, barWidth, Math.Max(yb - ya, 0), color));
                        if (s.Labels is { } l && i < l.Count && l[i].Length > 0) labels.Add((l[i], xLeft + barWidth / 2, stacked ? (ya + yb) / 2 : (value >= 0 ? ya - 2 : yb + 2), stacked, color));
                    }
                }
            }
        }
        else if (chart.Kind == ChartKind.Area)
        {
            var cumulative = new double[categories];
            for (int j = 0; j < seriesCount; j++)
            {
                var s = chart.Series[j];
                var baseline = (double[])cumulative.Clone();
                var tops = new double[categories];
                for (int i = 0; i < categories; i++)
                {
                    double value = i < s.Values.Count && s.Values[i] is not null ? Scaled(s, i) : 0;
                    tops[i] = chart.Grouping == ChartGrouping.Clustered ? value : baseline[i] + value;
                    if (chart.Grouping != ChartGrouping.Clustered) cumulative[i] = tops[i];
                    if (s.Labels is { } l && i < l.Count && l[i].Length > 0) labels.Add((l[i], CategoryX(i), ValueY(Math.Clamp(tops[i], chart.ValueMin, chart.ValueMax)) + font.FontSize, false, s.Color));
                }
                var outline = new List<(double X, double Y)>();
                for (int i = 0; i < categories; i++) outline.Add((CategoryX(i), ValueY(Math.Clamp(tops[i], chart.ValueMin, chart.ValueMax))));
                var polygon = new List<(double X, double Y)>(outline);
                for (int i = categories - 1; i >= 0; i--)
                    polygon.Add((CategoryX(i), ValueY(Math.Clamp(chart.Grouping == ChartGrouping.Clustered ? zero : baseline[i], chart.ValueMin, chart.ValueMax))));
                if (polygon.Count >= 3) items.Add(new PolygonItem(polygon, s.Color));
            }
        }
        else
        {
            double band = plotWidth / categories;
            var cumulative = new double[categories];
            for (int j = 0; j < seriesCount; j++)
            {
                var s = chart.Series[j];
                var points = new (double X, double Y)?[categories];
                for (int i = 0; i < categories; i++)
                {
                    if (i >= s.Values.Count || s.Values[i] is null) continue;
                    double value = Scaled(s, i);
                    if (chart.Grouping != ChartGrouping.Clustered) { cumulative[i] += value; value = cumulative[i]; }
                    points[i] = (CategoryX(i), ValueY(Math.Clamp(value, chart.ValueMin, chart.ValueMax)));
                    if (s.Labels is { } l && i < l.Count && l[i].Length > 0) labels.Add((l[i], points[i]!.Value.X, points[i]!.Value.Y - 6, false, s.Color));
                }
                for (int i = 0; i + 1 < categories; i++)
                    if (s.Line && points[i] is { } a && points[i + 1] is { } b) items.Add(new LineItem(a.X, a.Y, b.X, b.Y, 2.25, s.Color));
                if (s.Marker)
                    foreach (var p in points)
                        if (p is { } q)
                        {
                            items.Add(new EllipseItem(q.X - 3.2, q.Y - 3.2, 6.4, 6.4, Color.White));
                            items.Add(new EllipseItem(q.X - 2.5, q.Y - 2.5, 5, 5, s.Color));
                        }
            }
        }

        // ---- category axis line and labels
        if (chart.ShowCategoryAxis)
        {
            if (horizontalBars)
            {
                double ax = ValueX(zero);
                items.Add(new LineItem(ax, plotTop, ax, plotBottom, 0.75, ChartGrid));
                double band = plotHeight / categories;
                for (int i = 0; i < categories && i < chart.Categories.Count; i++)
                    PlaceText(chart.Categories[i], font, plotLeft - 5, plotBottom - (i + 0.5) * band, Anchor.End, true, 0, items);
            }
            else
            {
                double ay = ValueY(zero);
                items.Add(new LineItem(plotLeft, ay, plotRight, ay, 0.75, ChartGrid));
                double band = plotWidth / categories;
                for (int i = 0; i < categories && i < chart.Categories.Count; i += (int)stride)
                {
                    double cx = CategoryX(i);
                    if (rotateLabels) PlaceText(chart.Categories[i], font, cx + 3, plotBottom + 5, Anchor.End, false, 45, items);
                    else PlaceText(chart.Categories[i], font, cx, plotBottom + 4, Anchor.Middle, false, 0, items);
                }
            }
        }

        // ---- data labels last, on top of the series
        foreach (var (text, lx, ly, inside, fill) in labels)
        {
            var color = inside ? ContrastColor(fill) : ChartText;
            if (horizontalBars && !inside) PlaceText(text, font with { Color = color }, lx, ly, Anchor.Start, true, 0, items);
            else PlaceText(text, font with { Color = color }, lx, inside ? ly : ly - font.FontSize * 1.25, Anchor.Middle, inside, 0, items);
        }
    }

    private static Color ContrastColor(Color fill) =>
        0.299 * fill.R + 0.587 * fill.G + 0.114 * fill.B > 160 ? new Color(0x40, 0x40, 0x40) : Color.White;

    // ------------------------------------------------------------------ pie

    private void EmitPie(SheetChart chart, double left, double top, double right, double bottom, RunProperties font, List<LayoutItem> items)
    {
        if (chart.Series.Count == 0) return;
        double radius = Math.Min(right - left, bottom - top) / 2 - 4;
        if (radius < 5) return;
        double cx = (left + right) / 2, cy = (top + bottom) / 2;
        var labels = new List<(string Text, double X, double Y, Color Fill)>();
        if (chart.Kind == ChartKind.Pie)
            DrawSlices(chart, chart.Series[0], cx, cy, 0, radius, labels, items);
        else
        {
            // Doughnut: the first series is the innermost ring.
            double hole = radius * chart.HoleSize, ring = (radius - hole) / chart.Series.Count;
            for (int k = 0; k < chart.Series.Count; k++)
                DrawSlices(chart, chart.Series[k], cx, cy, hole + k * ring, hole + (k + 1) * ring, labels, items);
        }
        foreach (var (text, lx, ly, fill) in labels)
            PlaceText(text, font with { Color = ContrastColor(fill) }, lx, ly, Anchor.Middle, true, 0, items);
    }

    /// <summary>Draws the slices of one series between two radii (the inner one is 0 for a pie).</summary>
    private static void DrawSlices(SheetChart chart, ChartSeries series, double cx, double cy, double inner, double outer,
        List<(string Text, double X, double Y, Color Fill)> labels, List<LayoutItem> items)
    {
        var values = series.Values.Select(v => v is { } x && x > 0 ? x : 0).ToList();
        double total = values.Sum();
        if (total <= 0) return;
        double start = -90 + chart.FirstSliceAngle; // the first slice begins at 12 o'clock and runs clockwise
        for (int i = 0; i < values.Count; i++)
        {
            if (values[i] <= 0) continue;
            double sweep = values[i] / total * 360;
            var color = series.PointColors is { } pc && i < pc.Count ? pc[i] : ChartColor(i);
            int steps = Math.Max(2, (int)Math.Ceiling(sweep / 3));
            var points = new List<(double X, double Y)>();
            if (inner <= 0) points.Add((cx, cy));
            for (int k = 0; k <= steps; k++)
            {
                double angle = (start + sweep * k / steps) * Math.PI / 180;
                points.Add((cx + outer * Math.Cos(angle), cy + outer * Math.Sin(angle)));
            }
            if (inner > 0)
                for (int k = steps; k >= 0; k--)
                {
                    double angle = (start + sweep * k / steps) * Math.PI / 180;
                    points.Add((cx + inner * Math.Cos(angle), cy + inner * Math.Sin(angle)));
                }
            items.Add(new PolygonItem(points, color));
            if (values.Count > 1)
            {
                double a0 = start * Math.PI / 180, a1 = (start + sweep) * Math.PI / 180;
                items.Add(new LineItem(cx + inner * Math.Cos(a0), cy + inner * Math.Sin(a0), cx + outer * Math.Cos(a0), cy + outer * Math.Sin(a0), 0.75, Color.White));
                items.Add(new LineItem(cx + inner * Math.Cos(a1), cy + inner * Math.Sin(a1), cx + outer * Math.Cos(a1), cy + outer * Math.Sin(a1), 0.75, Color.White));
            }
            if (series.Labels is { } l && i < l.Count && l[i].Length > 0)
            {
                double mid = (start + sweep / 2) * Math.PI / 180, r = inner > 0 ? (inner + outer) / 2 : outer * 0.65;
                labels.Add((l[i], cx + r * Math.Cos(mid), cy + r * Math.Sin(mid), color));
            }
            start += sweep;
        }
    }
}
