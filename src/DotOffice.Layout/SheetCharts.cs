using DotOffice.Core;

namespace DotOffice.Layout;

public sealed partial class SheetLayoutEngine
{
    private static readonly Color ChartText = new(0x59, 0x59, 0x59);
    private static readonly Color ChartGrid = new(0xD9, 0xD9, 0xD9);
    private const double ChartPad = 7;

    private enum Anchor { Start, Middle, End }

    /// <summary>The font with the parts that a chart text style sets replaced.</summary>
    private static RunProperties Styled(RunProperties font, ChartTextStyle? style) => style is null ? font : font with
    {
        FontSize = style.Size ?? font.FontSize, Color = style.Color ?? font.Color, Bold = style.Bold ?? font.Bold,
        Italic = style.Italic ?? font.Italic, FontFamily = style.Family ?? font.FontFamily,
    };

    /// <summary>The text style of one data label that differs from the series' labels (null when it does not).</summary>
    private static ChartTextStyle? LabelStyleAt(ChartSeries series, int index) => series.LabelStyles is { } styles && index < styles.Count ? styles[index] : null;

    /// <summary>The plot rectangle of the chart being drawn when its layout is manual and inner (set by <see cref="EmitChart"/>).</summary>
    private (double X, double Y, double Width, double Height)? _manualPlot;

    /// <summary>The frame of the chart being drawn (set by <see cref="EmitChart"/>), for labels that are placed relative to the chart.</summary>
    private (double X, double Y, double Width, double Height) _chartRect;

    /// <summary>
    /// Draws a chart into its frame: background, title, legend, axes with gridlines and the series (columns, bars, lines or pie slices).
    /// Everything is clipped to the frame.
    /// </summary>
    private void EmitChart(SheetChart chart, double x, double y, double w, double h,
        (double X, double Y, double Width, double Height) visible, List<LayoutItem> items)
    {
        items.Add(new PushItem(visible));
        items.Add(new PushItem((x, y, w, h)));
        _chartRect = (x, y, w, h);
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
            var titleFont = Styled(font with { FontSize = chart.TitleSize }, chart.TextStyles.Title);
            if (chart.TitlePosition is { } at) PlaceText(chart.Title, titleFont, x + at.X * w, y + at.Y * h, Anchor.Start, false, 0, items);
            else top += PlaceText(chart.Title, titleFont, (left + right) / 2, top, Anchor.Middle, false, 0, items) + 4;
        }

        // ---- legend
        var legendEntries = chart.Kind is ChartKind.Pie or ChartKind.Doughnut
            ? chart.Categories.Select((c, i) => (Name: c, Color: chart.Series.FirstOrDefault()?.PointColors is { } pc && i < pc.Count ? pc[i] : ChartColor(i), Series: (ChartSeries?)null, Kind: chart.Kind)).ToList()
            : chart.Series.Select(s => (Name: s.Name, Color: s.Color, Series: (ChartSeries?)s, Kind: chart.Kind))
                .Concat(chart.ExtraGroups.SelectMany(g => g.Series.Select(s => (Name: s.Name, Color: s.Color, Series: (ChartSeries?)s, Kind: g.Kind)))).ToList();
        if (chart.Legend != LegendPosition.None && legendEntries.Count > 0 && right > left && bottom > top)
        {
            var legendFont = Styled(font, chart.TextStyles.Legend);
            double entryHeight = legendFont.FontSize * 1.25 + 3, swatch = legendFont.FontSize * 0.8;
            var widths = legendEntries.Select(e => MeasureChartText(e.Name, legendFont)).ToList();
            bool vertical = chart.Legend is LegendPosition.Right or LegendPosition.Left;
            if (chart.LegendBox is { } legendBox)
            {
                // A manual legend sits where the layout puts it, keeps the plot's automatic room (as if it were not there) and runs along its position's direction.
                double lx = x + legendBox.X * w, ly = y + legendBox.Y * h;
                double boxWidth = legendBox.Width * w, boxHeight = legendBox.Height * h;
                if (boxWidth > swatch + 8 && boxHeight >= entryHeight)
                {
                    // A manual size makes the entries flow inside the box: down the columns of a vertical legend, along the rows of a horizontal one.
                    double columnWidth = widths.Max() + swatch + 14, cursorX = 0, cursorY = 0;
                    for (int i = 0; i < legendEntries.Count; i++)
                    {
                        double entryWidth = widths[i] + swatch + 14;
                        if (vertical && cursorY + entryHeight > boxHeight + 0.5 && cursorY > 0) { cursorX += columnWidth; cursorY = 0; }
                        else if (!vertical && cursorX + entryWidth > boxWidth + 0.5 && cursorX > 0) { cursorY += entryHeight; cursorX = 0; }
                        DrawLegendEntry(chart, legendEntries[i].Name, legendEntries[i].Color, legendEntries[i].Series, legendEntries[i].Kind, lx + cursorX, ly + cursorY, swatch, entryHeight, legendFont, items);
                        if (vertical) cursorY += entryHeight; else cursorX += entryWidth;
                    }
                }
                else if (vertical)
                {
                    for (int i = 0; i < legendEntries.Count; i++)
                        DrawLegendEntry(chart, legendEntries[i].Name, legendEntries[i].Color, legendEntries[i].Series, legendEntries[i].Kind, lx, ly + i * entryHeight, swatch, entryHeight, legendFont, items);
                }
                else
                {
                    for (int i = 0; i < legendEntries.Count; i++)
                    {
                        DrawLegendEntry(chart, legendEntries[i].Name, legendEntries[i].Color, legendEntries[i].Series, legendEntries[i].Kind, lx, ly, swatch, entryHeight, legendFont, items);
                        lx += widths[i] + swatch + 14;
                    }
                }
            }
            else if (vertical)
            {
                double boxWidth = widths.Max() + swatch + 12;
                double lx = chart.Legend == LegendPosition.Right ? right - boxWidth : left;
                double ly = (top + bottom) / 2 - legendEntries.Count * entryHeight / 2;
                for (int i = 0; i < legendEntries.Count; i++)
                    DrawLegendEntry(chart, legendEntries[i].Name, legendEntries[i].Color, legendEntries[i].Series, legendEntries[i].Kind, lx, ly + i * entryHeight, swatch, entryHeight, legendFont, items);
                if (chart.Legend == LegendPosition.Right) right -= boxWidth + 6; else left += boxWidth + 6;
            }
            else
            {
                double total = widths.Sum() + legendEntries.Count * (swatch + 14);
                double lx = (left + right) / 2 - total / 2, ly = chart.Legend == LegendPosition.Top ? top : bottom - entryHeight;
                for (int i = 0; i < legendEntries.Count; i++)
                {
                    DrawLegendEntry(chart, legendEntries[i].Name, legendEntries[i].Color, legendEntries[i].Series, legendEntries[i].Kind, lx, ly, swatch, entryHeight, legendFont, items);
                    lx += widths[i] + swatch + 14;
                }
                if (chart.Legend == LegendPosition.Top) top += entryHeight + 4; else bottom -= entryHeight + 4;
            }
        }
        // A manual plot-area layout replaces the automatic one; an inner layout fixes the plot itself (below the axis labels).
        _manualPlot = null;
        if (chart.PlotArea is { } box)
        {
            var rect = (X: x + box.X * w, Y: y + box.Y * h, Width: box.Width * w, Height: box.Height * h);
            if (box.Inner && chart.Kind is not (ChartKind.Pie or ChartKind.Doughnut or ChartKind.Radar)) _manualPlot = rect;
            else (left, top, right, bottom) = (rect.X, rect.Y, rect.X + rect.Width, rect.Y + rect.Height);
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

    private void DrawLegendEntry(SheetChart chart, string name, Color color, ChartSeries? series, ChartKind kind, double x, double y, double swatch, double height, RunProperties font, List<LayoutItem> items)
    {
        double cy = y + height / 2;
        bool lineStyle = kind is ChartKind.Line or ChartKind.Scatter || (kind == ChartKind.Radar && !chart.Filled);
        if (lineStyle)
        {
            bool joined = series?.Line ?? true;
            if (joined) items.Add(new LineItem(x, cy, x + swatch + 4, cy, 2.25, color));
            if (series is null || series.Marker || !joined)
                items.Add(new EllipseItem(x + (swatch + 4) / 2 - 2.5, cy - 2.5, 5, 5, color));
            PlaceText(name, font, x + swatch + 8, cy, Anchor.Start, true, 0, items);
        }
        else if (kind == ChartKind.Bubble)
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

    // ------------------------------------------------------------------ axes, columns, bars, lines and areas

    private void EmitAxisChart(SheetChart chart, double left, double top, double right, double bottom, RunProperties font, List<LayoutItem> items)
    {
        bool horizontalBars = chart.Kind == ChartKind.Bar;
        int categories = Math.Max(chart.Categories.Count, 1);
        double fontSize = chart.FontSize;
        bool secondary = chart.ShowSecondaryAxis && chart.ExtraGroups.Any(g => g.Secondary);
        var styles = chart.TextStyles;
        var valueFont = Styled(font, styles.ValueAxis);
        var categoryFont = Styled(font, styles.CategoryAxis);
        var secondaryFont = Styled(font, styles.SecondaryAxis);

        // ---- room for axis titles and labels
        var horizontalTitleFont = Styled(font with { FontSize = fontSize + 1 }, horizontalBars ? styles.ValueAxisTitle : styles.CategoryAxisTitle);
        var verticalTitleFont = Styled(font with { FontSize = fontSize + 1 }, horizontalBars ? styles.CategoryAxisTitle : styles.ValueAxisTitle);
        var secondaryTitleFont = Styled(font with { FontSize = fontSize + 1 }, styles.SecondaryAxisTitle);
        string? horizontalTitle = horizontalBars ? chart.ValueAxisTitle : chart.CategoryAxisTitle, verticalTitle = horizontalBars ? chart.CategoryAxisTitle : chart.ValueAxisTitle;
        if (horizontalTitle is { Length: > 0 })
        {
            PlaceText(horizontalTitle, horizontalTitleFont, (left + right) / 2, bottom - horizontalTitleFont.FontSize - 2, Anchor.Middle, false, 0, items);
            bottom -= horizontalTitleFont.FontSize + 7;
        }
        if (verticalTitle is { Length: > 0 })
        {
            PlaceText(verticalTitle, verticalTitleFont, left + verticalTitleFont.FontSize / 2 + 1, (top + bottom) / 2, Anchor.Middle, true, 90, items);
            left += verticalTitleFont.FontSize + 7;
        }
        if (secondary && chart.SecondaryAxisTitle is { Length: > 0 } secondaryTitle)
        {
            PlaceText(secondaryTitle, secondaryTitleFont, right - secondaryTitleFont.FontSize / 2 - 1, (top + bottom) / 2, Anchor.Middle, true, -90, items);
            right -= secondaryTitleFont.FontSize + 7;
        }

        double tickWidth = chart.ShowValueAxis ? chart.Ticks.Select(t => MeasureChartText(t.Label, valueFont)).DefaultIfEmpty(0).Max() : 0;
        double secondaryWidth = secondary ? chart.SecondaryTicks.Select(t => MeasureChartText(t.Label, secondaryFont)).DefaultIfEmpty(0).Max() + 6 : 0;
        double categoryWidth = chart.ShowCategoryAxis ? chart.Categories.Select(c => MeasureChartText(c, categoryFont)).DefaultIfEmpty(0).Max() : 0;
        // Room for one line of the axis labels along the bottom (categories, or the values of horizontal bars) and of those along the side.
        double labelHeight = (horizontalBars ? valueFont : categoryFont).FontSize * 1.25;

        // With a reversed value axis the category axis crosses at the value axis' minimum, which now lies at the top (or, for bars, right).
        bool labelsAtFar = chart.ReverseValues && chart.ShowCategoryAxis;
        double plotLeft, plotRight, plotTop, plotBottom;
        bool rotateLabels = false;
        double stride = 1, labelAngle = 0;
        if (!horizontalBars)
        {
            plotLeft = left + (chart.ShowValueAxis ? tickWidth + 6 : 0);
            plotRight = right - 4 - secondaryWidth;
            plotTop = top + labelHeight / 2;
            double bandWidth = (plotRight - plotLeft) / (chart.CategoriesBetween ? categories : Math.Max(categories - 1, 1));
            // The labels turn by the angle the axis asks for, or by 45° when they would overlap.
            if (styles.CategoryAxis?.Rotation is { } asked && !labelsAtFar) labelAngle = Math.Clamp(asked, -90, 90);
            rotateLabels = chart.ShowCategoryAxis && (labelAngle != 0 || styles.CategoryAxis?.Rotation is null && categoryWidth > bandWidth * 0.95 - 2);
            if (rotateLabels && labelAngle == 0) labelAngle = 45;
            double sine = Math.Abs(Math.Sin(labelAngle * Math.PI / 180)), cosine = Math.Abs(Math.Cos(labelAngle * Math.PI / 180));
            double below = !chart.ShowCategoryAxis ? 0 : rotateLabels ? categoryWidth * sine + labelHeight * cosine + 6 : labelHeight + 5;
            if (labelsAtFar) { plotTop = top + below; plotBottom = bottom - labelHeight / 2; }
            else plotBottom = bottom - below;
            stride = !chart.ShowCategoryAxis ? 1 : Math.Max(1, Math.Ceiling((rotateLabels ? labelHeight * 1.1 / Math.Max(sine, 0.25) : categoryWidth + 6) / Math.Max(bandWidth, 1)));
        }
        else
        {
            plotLeft = left + (chart.ShowCategoryAxis && !labelsAtFar ? categoryWidth + 6 : 0);
            plotRight = right - (labelsAtFar ? categoryWidth + 12 : 8);
            plotTop = top + 4;
            plotBottom = bottom - (chart.ShowValueAxis ? labelHeight + 5 : 0);
        }
        if (_manualPlot is { } manual) (plotLeft, plotTop, plotRight, plotBottom) = (manual.X, manual.Y, manual.X + manual.Width, manual.Y + manual.Height);
        if (plotRight - plotLeft < 10 || plotBottom - plotTop < 10) return;
        double plotWidth = plotRight - plotLeft, plotHeight = plotBottom - plotTop;
        var primary = chart.ValueScale;
        var secondaryScale = chart.SecondaryScale;

        // Positions along the value axis (of the primary or the secondary scale) and the category axis.
        double ValueX(AxisScale scale, double v) => plotLeft + plotWidth * scale.Fraction(v);
        double ValueY(AxisScale scale, double v) => plotBottom - plotHeight * scale.Fraction(v);
        double ValuePosition(AxisScale scale, double v) => horizontalBars ? ValueX(scale, v) : ValueY(scale, v);
        double Baseline(AxisScale scale) => scale.LogBase > 1 ? scale.Min : Math.Clamp(0, scale.Min, scale.Max);
        int Slot(int i) => chart.ReverseCategories ? categories - 1 - i : i;
        bool between = chart.CategoriesBetween || chart.Kind is ChartKind.Column or ChartKind.Bar;
        // t is the zero-based category position, fractional for points between categories (trendlines).
        double SlotAt(double t) => chart.ReverseCategories ? categories - 1 - t : t;
        double CategoryAt(double t) => between ? plotLeft + (SlotAt(t) + 0.5) * plotWidth / categories
            : categories > 1 ? plotLeft + SlotAt(t) * plotWidth / (categories - 1) : plotLeft + plotWidth / 2;
        double CategoryX(int i) => CategoryAt(i);
        var errorBars = new List<(ChartSeries Series, int Index, double Across, double Value, AxisScale Scale)>();

        // ---- gridlines and value labels
        foreach (var tick in chart.Ticks)
        {
            if (horizontalBars)
            {
                double tx = ValueX(primary, tick.Value);
                if (chart.MajorGridlines) items.Add(new LineItem(tx, plotTop, tx, plotBottom, 0.75, ChartGrid));
                if (chart.ShowValueAxis) PlaceText(tick.Label, valueFont, tx, plotBottom + 4, Anchor.Middle, false, 0, items);
            }
            else
            {
                double ty = ValueY(primary, tick.Value);
                if (chart.MajorGridlines) items.Add(new LineItem(plotLeft, ty, plotRight, ty, 0.75, ChartGrid));
                if (chart.ShowValueAxis) PlaceText(tick.Label, valueFont, plotLeft - 5, ty, Anchor.End, true, 0, items);
            }
        }
        if (secondary && !horizontalBars)
            foreach (var tick in chart.SecondaryTicks)
                PlaceText(tick.Label, secondaryFont, plotRight + 5, ValueY(secondaryScale, tick.Value), Anchor.Start, true, 0, items);

        // ---- series: the first chart type, then the further types of a combination chart
        var labels = new List<(string Text, double X, double Y, bool Inside, Color Fill, ChartTextStyle? Style)>();
        var groups = new List<ChartGroup> { new(chart.Kind, chart.Grouping, chart.Series, false, chart.GapWidth, chart.Overlap) };
        groups.AddRange(chart.ExtraGroups);
        foreach (var group in groups)
        {
            var scale = group.Secondary ? secondaryScale : primary;
            var list = group.Series;
            int seriesCount = list.Count;
            var totals = new double[categories];
            if (group.Grouping == ChartGrouping.PercentStacked)
                foreach (var s in list)
                    for (int i = 0; i < categories; i++) totals[i] += i < s.Values.Count && s.Values[i] is { } v ? Math.Abs(v) : 0;
            double Scaled(ChartSeries s, int i) => i < s.Values.Count && s.Values[i] is { } v ? group.Grouping == ChartGrouping.PercentStacked ? (totals[i] > 0 ? v / totals[i] : 0) : v : 0;

            if (group.Kind is ChartKind.Column or ChartKind.Bar)
            {
                double band = (horizontalBars ? plotHeight : plotWidth) / categories;
                bool stacked = group.Grouping != ChartGrouping.Clustered;
                double overlap = stacked ? 1 : group.Overlap / 100, gap = group.GapWidth / 100;
                int perGroup = stacked ? 1 : Math.Max(seriesCount, 1);
                double barWidth = band / (perGroup - (perGroup - 1) * overlap + gap);
                double groupWidth = barWidth * (perGroup - (perGroup - 1) * overlap);
                double baseline = Baseline(scale);
                var positive = new double[categories];
                var negative = new double[categories];
                for (int j = 0; j < seriesCount; j++)
                {
                    var s = list[j];
                    for (int i = 0; i < categories; i++)
                    {
                        if (i >= s.Values.Count || s.Values[i] is null) continue;
                        double value = Scaled(s, i), from, to;
                        if (stacked)
                        {
                            if (value >= 0) { from = positive[i]; to = from + value; positive[i] = to; }
                            else { from = negative[i]; to = from + value; negative[i] = to; }
                        }
                        else (from, to) = (baseline, value);
                        double offset = stacked ? -barWidth / 2 : -groupWidth / 2 + j * barWidth * (1 - overlap);
                        var color = s.Color;
                        double p0 = ValuePosition(scale, from), p1 = ValuePosition(scale, to);
                        if (horizontalBars)
                        {
                            double center = plotBottom - (Slot(i) + 0.5) * band;
                            double ya = Math.Min(p0, p1), yb = Math.Max(p0, p1);
                            // In horizontal bar charts the first series sits lowest in each group.
                            double slot = stacked ? center + offset : center - offset - barWidth;
                            AddFaded(items, s.Alpha, new RectItem(ya, slot, Math.Max(yb - ya, 0), barWidth, color));
                            errorBars.Add((s, i, slot + barWidth / 2, stacked ? to : value, scale));
                            if (s.Labels is { } l && i < l.Count && l[i].Length > 0) labels.Add((l[i], stacked ? (ya + yb) / 2 : yb + 3, slot + barWidth / 2, stacked, color, LabelStyleAt(s, i)));
                        }
                        else
                        {
                            double center = plotLeft + (Slot(i) + 0.5) * band;
                            double xLeft = center + offset, ya = Math.Min(p0, p1), yb = Math.Max(p0, p1);
                            AddFaded(items, s.Alpha, new RectItem(xLeft, ya, barWidth, Math.Max(yb - ya, 0), color));
                            errorBars.Add((s, i, xLeft + barWidth / 2, stacked ? to : value, scale));
                            if (s.Labels is { } l && i < l.Count && l[i].Length > 0) labels.Add((l[i], xLeft + barWidth / 2, stacked ? (ya + yb) / 2 : (value >= 0 ? ya - 2 : yb + 2), stacked, color, LabelStyleAt(s, i)));
                        }
                    }
                }
            }
            else if (group.Kind == ChartKind.Area)
            {
                var cumulative = new double[categories];
                double baseline = Baseline(scale);
                for (int j = 0; j < seriesCount; j++)
                {
                    var s = list[j];
                    var below = (double[])cumulative.Clone();
                    var tops = new double[categories];
                    for (int i = 0; i < categories; i++)
                    {
                        double value = i < s.Values.Count && s.Values[i] is not null ? Scaled(s, i) : 0;
                        tops[i] = group.Grouping == ChartGrouping.Clustered ? value : below[i] + value;
                        if (group.Grouping != ChartGrouping.Clustered) cumulative[i] = tops[i];
                        if (s.Labels is { } l && i < l.Count && l[i].Length > 0) labels.Add((l[i], CategoryX(i), ValueY(scale, tops[i]) + font.FontSize, false, s.Color, LabelStyleAt(s, i)));
                    }
                    var order = Enumerable.Range(0, categories).OrderBy(CategoryX).ToList();
                    var polygon = order.Select(i => (CategoryX(i), ValueY(scale, tops[i]))).ToList();
                    foreach (int i in Enumerable.Reverse(order))
                        polygon.Add((CategoryX(i), ValueY(scale, group.Grouping == ChartGrouping.Clustered ? baseline : below[i])));
                    if (polygon.Count >= 3) AddFaded(items, s.Alpha, new PolygonItem(polygon, s.Color));
                    for (int i = 0; i < categories; i++) if (i < s.Values.Count && s.Values[i] is not null) errorBars.Add((s, i, CategoryX(i), tops[i], scale));
                }
            }
            else if (group.Kind == ChartKind.Line)
            {
                var cumulative = new double[categories];
                for (int j = 0; j < seriesCount; j++)
                {
                    var s = list[j];
                    var points = new (double X, double Y)?[categories];
                    for (int i = 0; i < categories; i++)
                    {
                        if (i >= s.Values.Count || s.Values[i] is null) continue;
                        double value = Scaled(s, i);
                        if (group.Grouping != ChartGrouping.Clustered) { cumulative[i] += value; value = cumulative[i]; }
                        points[i] = (CategoryX(i), ValueY(scale, value));
                        errorBars.Add((s, i, CategoryX(i), value, scale));
                        if (s.Labels is { } l && i < l.Count && l[i].Length > 0) labels.Add((l[i], points[i]!.Value.X, points[i]!.Value.Y - 6, false, s.Color, LabelStyleAt(s, i)));
                    }
                    if (s.Line) DrawPolyline(points, s.Smooth, 2.25, s.Color, items);
                    if (s.Marker)
                        foreach (var p in points)
                            if (p is { } q)
                            {
                                items.Add(new EllipseItem(q.X - 3.2, q.Y - 3.2, 6.4, 6.4, Color.White));
                                items.Add(new EllipseItem(q.X - 2.5, q.Y - 2.5, 5, 5, s.Color));
                            }
                }
            }
        }

        // ---- error bars and trendlines, clipped to the plot
        var overlays = new List<LayoutItem>();
        foreach (var (s, i, across, value, scale) in errorBars)
            foreach (var bars in s.ErrorBars ?? [])
            {
                if (bars.Horizontal != horizontalBars || i >= bars.Plus.Count) continue;
                double a = ValuePosition(scale, value - bars.Minus[i]), b = ValuePosition(scale, value + bars.Plus[i]);
                DrawErrorBar(overlays, horizontalBars, across, a, b, bars);
            }
        // A trendline follows what is plotted: the stacked totals of stacked groups, the values themselves otherwise. Horizontal bars
        // run it along the value axis with the categories up the side.
        double BarCategoryAt(double t) => plotBottom - (SlotAt(t) + 0.5) * plotHeight / categories;
        var trendLabels = new List<LayoutItem>();
        foreach (var group in groups)
            foreach (var s in group.Series)
                if (s.Trendline is { } trend)
                {
                    var scale = group.Secondary ? secondaryScale : primary;
                    var known = errorBars.Where(e => ReferenceEquals(e.Series, s)).OrderBy(e => e.Index).Select(e => ((double)e.Index + 1, e.Value)).ToList();
                    if (known.Count == 0) continue;
                    var curve = ChartTrendlines.Fit(trend, known, known.Min(k => k.Item1), known.Max(k => k.Item1));
                    var drawn = curve.Select(q => horizontalBars ? (ValueX(scale, q.Y), BarCategoryAt(q.X - 1)) : (CategoryAt(q.X - 1), ValueY(scale, q.Y))).ToList();
                    DrawTrendline(overlays, drawn, trend);
                    if (drawn.Count > 0) AddTrendlineLabel(trendLabels, trend, known, drawn[^1], font, plotLeft, plotTop, plotRight, plotBottom);
                }
        if (overlays.Count > 0)
        {
            items.Add(new PushItem((plotLeft, plotTop, plotWidth, plotHeight)));
            items.AddRange(overlays);
            items.Add(new PopItem());
        }
        items.AddRange(trendLabels);

        // ---- category axis line and labels
        if (chart.ShowCategoryAxis)
        {
            double crossing = Baseline(primary);
            if (horizontalBars)
            {
                double ax = ValueX(primary, crossing);
                items.Add(new LineItem(ax, plotTop, ax, plotBottom, 0.75, ChartGrid));
                double band = plotHeight / categories;
                for (int i = 0; i < categories && i < chart.Categories.Count; i++)
                    if (labelsAtFar) PlaceText(chart.Categories[i], categoryFont, plotRight + 5, plotBottom - (Slot(i) + 0.5) * band, Anchor.Start, true, 0, items);
                    else PlaceText(chart.Categories[i], categoryFont, plotLeft - 5, plotBottom - (Slot(i) + 0.5) * band, Anchor.End, true, 0, items);
            }
            else
            {
                double ay = ValueY(primary, crossing);
                items.Add(new LineItem(plotLeft, ay, plotRight, ay, 0.75, ChartGrid));
                for (int i = 0; i < categories && i < chart.Categories.Count; i += (int)stride)
                {
                    double cx = CategoryX(i);
                    if (labelsAtFar && rotateLabels) PlaceText(chart.Categories[i], categoryFont, cx - 3, plotTop - 6, Anchor.Start, true, 45, items);
                    else if (labelsAtFar) PlaceText(chart.Categories[i], categoryFont, cx, plotTop - 4 - labelHeight, Anchor.Middle, false, 0, items);
                    else if (rotateLabels)
                    {
                        // Turned counter-clockwise (positive) the text ends at its tick; clockwise it starts there. A upright column of text is centered on the tick.
                        double turn = Math.Pow(Math.Sin(labelAngle * Math.PI / 180), 2) * (3 + labelHeight / 2);
                        if (labelAngle > 0) PlaceText(chart.Categories[i], categoryFont, cx + 3 - turn, plotBottom + 5, Anchor.End, false, labelAngle, items);
                        else PlaceText(chart.Categories[i], categoryFont, cx - 3 + turn, plotBottom + 5, Anchor.Start, false, labelAngle, items);
                    }
                    else PlaceText(chart.Categories[i], categoryFont, cx, plotBottom + 4, Anchor.Middle, false, 0, items);
                }
            }
        }

        // ---- data labels last, on top of the series
        foreach (var (text, lx, ly, inside, fill, pointStyle) in labels)
        {
            var labelFont = Styled(Styled(font with { Color = inside ? ContrastColor(fill) : ChartText }, styles.DataLabels), pointStyle);
            if (horizontalBars && !inside) PlaceText(text, labelFont, lx, ly, Anchor.Start, true, 0, items);
            else PlaceText(text, labelFont, lx, inside ? ly : ly - labelFont.FontSize * 1.25, Anchor.Middle, inside, 0, items);
        }
    }

    /// <summary>Adds a filled item, painted with the series' opacity when it is translucent.</summary>
    private static void AddFaded(List<LayoutItem> items, double alpha, LayoutItem item)
    {
        if (alpha >= 1) { items.Add(item); return; }
        items.Add(new PushItem(null, Opacity: alpha));
        items.Add(item);
        items.Add(new PopItem());
    }

    /// <summary>
    /// An error bar with end caps: along the value axis from <paramref name="from"/> to <paramref name="to"/> (positions on that axis) at
    /// <paramref name="across"/> on the other one; <paramref name="alongX"/> when the bar runs horizontally.
    /// </summary>
    private static void DrawErrorBar(List<LayoutItem> items, bool alongX, double across, double from, double to, ChartErrorBars bars)
    {
        var color = bars.Color ?? ChartText;
        const double cap = 3;
        if (alongX)
        {
            items.Add(new LineItem(from, across, to, across, bars.Width, color));
            if (bars.EndCap) { items.Add(new LineItem(from, across - cap, from, across + cap, bars.Width, color)); items.Add(new LineItem(to, across - cap, to, across + cap, bars.Width, color)); }
        }
        else
        {
            items.Add(new LineItem(across, from, across, to, bars.Width, color));
            if (bars.EndCap) { items.Add(new LineItem(across - cap, from, across + cap, from, bars.Width, color)); items.Add(new LineItem(across - cap, to, across + cap, to, bars.Width, color)); }
        }
    }

    /// <summary>
    /// Prints the equation and R² of a trendline (when it asks for them) just above its far end, kept inside the plot, right-aligned to the end.
    /// </summary>
    private void AddTrendlineLabel(List<LayoutItem> items, ChartTrendline trend, IReadOnlyList<(double X, double Y)> data, (double X, double Y) end,
        RunProperties font, double plotLeft, double plotTop, double plotRight, double plotBottom)
    {
        if (!trend.ShowEquation && !trend.ShowRSquared) return;
        if (ChartTrendlines.Describe(trend, data) is not var (equation, r2)) return;
        var lines = new List<string>();
        if (trend.ShowEquation) lines.Add(equation);
        if (trend.ShowRSquared) lines.Add("R² = " + (trend.NumberFormatter is { } format ? format(r2) : ChartTrendlines.Number(Math.Round(r2, 4))));
        var labelFont = Styled(font, trend.LabelStyle);
        double pitch = labelFont.FontSize * 1.25;
        double width = lines.Max(l => MeasureChartText(l, labelFont));
        double x1 = Math.Clamp(end.X, plotLeft + width, Math.Max(plotRight, plotLeft + width));
        double y0 = Math.Clamp(end.Y - 4 - lines.Count * pitch, plotTop, Math.Max(plotBottom - lines.Count * pitch, plotTop));
        if (trend.LabelPosition is { } by)
        {
            // Placed by hand: from the chart's corner, or moved from the automatic spot by a fraction of the chart's size.
            var (cx, cy, cw, ch) = _chartRect;
            double left = by.FromEdge ? cx + by.X * cw : x1 - width + by.X * cw, top = by.FromEdge ? cy + by.Y * ch : y0 + by.Y * ch;
            for (int i = 0; i < lines.Count; i++) PlaceText(lines[i], labelFont, left, top + i * pitch, Anchor.Start, false, 0, items);
            return;
        }
        for (int i = 0; i < lines.Count; i++) PlaceText(lines[i], labelFont, x1, y0 + i * pitch, Anchor.End, false, 0, items);
    }

    private static void DrawTrendline(List<LayoutItem> items, List<(double X, double Y)> points, ChartTrendline trend)
    {
        for (int k = 0; k + 1 < points.Count; k++)
            items.Add(new LineItem(points[k].X, points[k].Y, points[k + 1].X, points[k + 1].Y, trend.Width, trend.Color, trend.Dashed ? LineStyle.Dashed : LineStyle.Solid));
    }

    /// <summary>
    /// Joins points with straight segments, or with a smooth curve through them (a Catmull-Rom spline, which Excel's smoothing
    /// resembles). Missing points break the line.
    /// </summary>
    private static void DrawPolyline((double X, double Y)?[] points, bool smooth, double width, Color color, List<LayoutItem> items)
    {
        int i = 0;
        while (i < points.Length)
        {
            if (points[i] is null) { i++; continue; }
            int j = i;
            while (j + 1 < points.Length && points[j + 1] is not null) j++;
            var run = points[i..(j + 1)].Select(p => p!.Value).ToList();
            for (int k = 0; k + 1 < run.Count; k++)
            {
                if (!smooth)
                {
                    items.Add(new LineItem(run[k].X, run[k].Y, run[k + 1].X, run[k + 1].Y, width, color));
                    continue;
                }
                var p0 = run[Math.Max(k - 1, 0)]; var p1 = run[k]; var p2 = run[k + 1]; var p3 = run[Math.Min(k + 2, run.Count - 1)];
                var c1 = (X: p1.X + (p2.X - p0.X) / 6, Y: p1.Y + (p2.Y - p0.Y) / 6);
                var c2 = (X: p2.X - (p3.X - p1.X) / 6, Y: p2.Y - (p3.Y - p1.Y) / 6);
                var last = p1;
                for (int step = 1; step <= 12; step++)
                {
                    double t = step / 12.0, u = 1 - t;
                    var q = (X: u * u * u * p1.X + 3 * u * u * t * c1.X + 3 * u * t * t * c2.X + t * t * t * p2.X,
                             Y: u * u * u * p1.Y + 3 * u * u * t * c1.Y + 3 * u * t * t * c2.Y + t * t * t * p2.Y);
                    items.Add(new LineItem(last.X, last.Y, q.X, q.Y, width, color));
                    last = q;
                }
            }
            i = j + 1;
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
        var labels = new List<(string Text, double X, double Y, Color Fill, ChartTextStyle? Style)>();
        if (chart.Kind == ChartKind.Pie)
            DrawSlices(chart, chart.Series[0], cx, cy, 0, radius, labels, items);
        else
        {
            // Doughnut: the first series is the innermost ring.
            double hole = radius * chart.HoleSize, ring = (radius - hole) / chart.Series.Count;
            for (int k = 0; k < chart.Series.Count; k++)
                DrawSlices(chart, chart.Series[k], cx, cy, hole + k * ring, hole + (k + 1) * ring, labels, items);
        }
        foreach (var (text, lx, ly, fill, pointStyle) in labels)
            PlaceText(text, Styled(Styled(font with { Color = ContrastColor(fill) }, chart.TextStyles.DataLabels), pointStyle), lx, ly, Anchor.Middle, true, 0, items);
    }

    /// <summary>Draws the slices of one series between two radii (the inner one is 0 for a pie).</summary>
    private static void DrawSlices(SheetChart chart, ChartSeries series, double cx, double cy, double inner, double outer,
        List<(string Text, double X, double Y, Color Fill, ChartTextStyle? Style)> labels, List<LayoutItem> items)
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
                labels.Add((l[i], cx + r * Math.Cos(mid), cy + r * Math.Sin(mid), color, LabelStyleAt(series, i)));
            }
            start += sweep;
        }
    }
}
