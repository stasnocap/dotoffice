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
        chart.WordLayout && chart.PlotArea is null && chart.LegendBox is null && chart.TitlePosition is null
        && chart.Kind is ChartKind.Column or ChartKind.Bar or ChartKind.Line or ChartKind.Scatter or ChartKind.Pie
        && chart.ExtraGroups.Count == 0 && !chart.ShowSecondaryAxis
        && string.IsNullOrEmpty(chart.ValueAxisTitle) && string.IsNullOrEmpty(chart.CategoryAxisTitle)
        && (chart.Kind == ChartKind.Pie || chart is { ShowValueAxis: true, ShowCategoryAxis: true, ReverseValues: false, ReverseCategories: false })
        && chart.TextStyles.ValueAxis?.Rotation is null or 0 && (chart.TextStyles.CategoryAxis?.Rotation is null or 0 || chart.Kind is ChartKind.Column or ChartKind.Line && chart.TextStyles.CategoryAxis?.Rotation is 90 or -90 or 45)
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

    /// <summary>How Word 2021 sets the category labels under a column or line chart's horizontal axis (<see cref="PlanWordCategoryLabels"/>).</summary>
    private sealed record WordCategoryPlan(double Angle, int Stride, List<List<string>> Lines, double Width, double[]? Overhang = null, double[]? Widths = null, bool Asked = false)
    {
        public int MaxLines => Lines.Select(l => l.Count).DefaultIfEmpty(1).Max();
    }

    /// <summary>Set by <see cref="EmitWordChart"/> for the axis layout that draws the labels; null when the labels keep the shared layout.</summary>
    private WordCategoryPlan? _wordCategories;

    /// <summary>Set by <see cref="EmitWordChart"/> for a bar chart: each category label's lines, wrapped as Word wraps them; null when the labels keep the shared layout.</summary>
    private (List<List<string>> Lines, double Cap)? _wordBarLabels;

    // Word 2021's category labels (measured at Arial 10 pt, bar and line charts of 3 to 80 categories, 200 to 430 pt wide; every size below is therefore stated per
    // pt of font size and INFERRED for other sizes). The slot is the plot's width over the category count, taken before the labels narrow the plot.
    private const double WordTurnLine = 1.113;         // em: a turned label needs this much slot for its line; 11.1 at 10 pt; 1.113 fits the thresholds at 8, 10 and 14 pt (every case within 0.1 pt of the limit)
    private const double WordTurnBottom = 1.40;        // em: the plot's bottom is 6.5 + this + 0.7071 W above the chart's bottom when turned by 45 degrees (20.5 at 10 pt; 17.7 / 26.1 + ... measured at 8 / 14 pt)
    private const double WordUprightBottom = 0.61;     // em: ... 6.5 + this + W when turned by 90 degrees (12.6 at 10 pt; 11.4 / 15.0 + ... measured at 8 / 14 pt)
    private const double WordTurnLeft = 10.9;          // pt, not scaling with the font (measured at 8, 10 and 14 pt): a 45 degree label's start keeps the plot's left edge at 0.7071 W + this - R / 2n

    /// <summary>
    /// Chooses how the category labels are drawn, from these rules (measured in Word 2021 at 10 pt): the labels stay upright when the widest fits its slot; else
    /// a label of several words wraps at its spaces when its longest word fits the slot (lines 1.15 em apart); else the labels turn by 45 degrees when the slot is
    /// at least 1.11 em / sin 45 (reading upward to the right), by 90 degrees when it is at least 1.11 em, and by 45 degrees showing every Nth label below that, with
    /// N = ceil(1.11 em / slot). Showing every Nth label upright (N = ceil(W / slot)) wins when it shows at least as many labels as turning would (ties: upright, 45, 90).
    /// </summary>
    private WordCategoryPlan PlanWordCategoryLabels(IReadOnlyList<string> categories, RunProperties font, double slot, double chartHeight = double.PositiveInfinity, double band = 0)
    {
        int n = Math.Max(categories.Count, 1);
        var plain = categories.Select(c => new List<string> { c.Replace('\n', ' ') }).ToList();
        double widest = categories.Select(c => MeasureChartText(c.Replace('\n', ' '), font)).DefaultIfEmpty(0).Max();
        if (widest <= slot) return new WordCategoryPlan(0, 1, plain, widest);
        var words = categories.SelectMany(c => c.Replace('\n', ' ').Split(' ', StringSplitOptions.RemoveEmptyEntries)).ToList();
        if (words.Count > categories.Count && words.All(w => MeasureChartText(w, font) <= slot))
            return new WordCategoryPlan(0, 1, categories.Select(c => WrapLegendName(c.Replace('\n', ' '), font, slot)).ToList(), widest);
        double turnLine = WordTurnLine * font.FontSize;
        int uprightStride = (int)Math.Ceiling(widest / slot), uprightShown = (n + uprightStride - 1) / uprightStride;
        (double Angle, int Stride) turn = slot >= turnLine / Math.Sin(Math.PI / 4) ? (45, 1) : slot >= turnLine ? (90, 1) : (45, (int)Math.Ceiling(turnLine / slot));
        int turnShown = (n + turn.Stride - 1) / turn.Stride;
        if (uprightShown >= turnShown) return new WordCategoryPlan(0, uprightStride, plain, widest);
        return TruncateTurnedLabels(categories, font, turn.Angle, turn.Stride, chartHeight, band);
    }

    /// <summary>
    /// Turned labels (measured in Word 2021 at 8, 10 and 14 pt, 45 and 90 degrees, H 100 to 400, with and without a title; width and legend do not matter): the label plus its "…" may take
    /// the plot's bottom gap (6.5 + the turn's constant + the label's extent) up to (H - title band - top or bottom legend rows) / 2 + 0.5 / 1.7 / 2.4 pt at 8 / 10 / 14 pt; a longer label keeps the most whole characters that fit with the "…".
    /// The plot gives way for the kept characters only (the "…" is not counted), so the "…" reaches past the label's end into the plot.
    /// </summary>
    private WordCategoryPlan TruncateTurnedLabels(IReadOnlyList<string> categories, RunProperties font, double angle, int stride, double chartHeight, double band)
    {
        double fs = font.FontSize;
        // The constant fits every sample at 8, 10 and 14 pt (0.38..0.58, 1.60..1.74, 2.24..2.61 pt); between and beyond them it is interpolated and INFERRED.
        double extra = fs <= 10 ? 0.48 + 0.595 * (fs - 8) : 1.67 + 0.19 * (fs - 10);
        double total = 0.5 * (chartHeight - band) + extra;
        double cap = angle == 45 ? (total - 6.5 - WordTurnBottom * fs) / 0.7071 : total - 6.5 - WordUprightBottom * fs;
        double ellipsis = MeasureChartText("…", font);
        var lines = new List<List<string>>();
        var overhang = new double[categories.Count];
        double width = 0;
        var widths = new double[categories.Count];
        for (int i = 0; i < categories.Count; i++)
        {
            string text = categories[i].Replace('\n', ' ');
            double full = MeasureChartText(text, font);
            if (full <= cap) { lines.Add([text]); widths[i] = full; width = Math.Max(width, full); continue; }
            // A label of several words is cut after the last whole word that fits (measured: "Quarterly revenue north" -> "Quarterly…"); a first word that does not fit alone is cut by character.
            int keep = text.Length;
            while (keep > 0 && MeasureChartText(text[..keep], font) + ellipsis > cap) keep--;
            int wordEnd = keep;
            while (wordEnd > 0 && !(wordEnd < text.Length && text[wordEnd] == ' ' && text[wordEnd - 1] != ' ')) wordEnd--;
            if (wordEnd > 0) keep = wordEnd;
            lines.Add([text[..keep] + "…"]);
            overhang[i] = ellipsis;
            widths[i] = MeasureChartText(text[..keep], font);
            width = Math.Max(width, widths[i]);
        }
        return new WordCategoryPlan(angle, stride, lines, width, overhang, widths);
    }

    /// <summary>Draws category label <paramref name="i"/> under the axis at <paramref name="cx"/> as <paramref name="plan"/> says; <paramref name="plotBottom"/> is the axis.</summary>
    private void PlaceWordCategory(WordCategoryPlan plan, int i, RunProperties font, double cx, double plotBottom, List<LayoutItem> items)
    {
        double fs = font.FontSize;
        if (plan.Angle == 0)
        {
            // Every line is centered and every label starts on the first row (top aligned: a label of fewer lines than the tallest leaves its last rows empty; measured at 8, 10 and 14 pt); with the plot giving way by a line per extra row, the tallest label's last baseline stays where an upright single line has it.
            var lines = plan.Lines[i];
            double step = 1.15 * fs, top = plotBottom + LabelDrop(font);
            for (int k = 0; k < lines.Count; k++) PlaceText(lines[k], font, cx, top + k * step, Anchor.Middle, false, 0, items);
        }
        else if (plan.Angle == 45)
        {
            // The text's end (its ink, measured) stands 2.4 right of the slot (1.9 to the baseline's corner)'s centre; the anchor is the line box's top, so the baseline's end lies (to the baseline) x sin 45 further right.
            var (pitch, ascent, descent) = LineMetrics(font);
            double toBaseline = (pitch - (ascent + descent)) / 2 + ascent;
            double over = plan.Overhang?[i] ?? 0; // a cut label's "…" runs on past the end its kept text has
            PlaceText(plan.Lines[i][0], font, cx + 2.4 * fs / 10 - toBaseline * 0.7071 + over * 0.7071, plotBottom + 5.3 * fs / 10 - over * 0.7071, Anchor.End, false, 45, items);
        }
        else if (plan.Asked)
        {
            // An asked turn by 90 degrees either way (measured in Word 2021 at 10 pt, 1 sample each): the lines stand 1.15 em apart, centred on the slot (3.45 right of
            // its centre for one line reading bottom-up, 3.39 left reading top-down, the first line the leftmost resp. rightmost), each centred along a block as long as
            // the label's widest line whose near end is 6.05 (bottom-up) or 6.09 (top-down) below the axis.
            var (pitch, ascent, descent) = LineMetrics(font);
            double toBaseline = (pitch - (ascent + descent)) / 2 + ascent, step = 1.15 * fs, scale = fs / 10;
            var lines = plan.Lines[i];
            double own = lines.Max(l => MeasureChartText(l, font));
            for (int k = 0; k < lines.Count; k++)
            {
                double inset = (own - MeasureChartText(lines[k], font)) / 2, shift = (k - (lines.Count - 1) / 2.0) * step;
                if (plan.Angle == 90) PlaceText(lines[k], font, cx + 3.45 * scale + shift - toBaseline, plotBottom + 6.05 * scale + inset, Anchor.End, false, 90, items);
                else PlaceText(lines[k], font, cx - 3.39 * scale - shift + toBaseline, plotBottom + 6.09 * scale + inset, Anchor.Start, false, -90, items);
            }
        }
        else
        {
            // Reading bottom-up: the baseline stands 3.45 right of the slot's centre and the text ends 6.0 below the axis (measured).
            var (pitch, ascent, descent) = LineMetrics(font);
            PlaceText(plan.Lines[i][0], font, cx + 3.45 * fs / 10 - ((pitch - (ascent + descent)) / 2 + ascent), plotBottom + 6.0 * fs / 10 - (plan.Overhang?[i] ?? 0), Anchor.End, false, 90, items);
        }
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

        // Right and left legends wrap a name wider than W / 3 - 1.15 em (measured, 40+ samples at W 236..480 and 10 pt, then at 8, 9, 12 and 14 pt: greedy at spaces, a single word that is too
        // wide breaks by character; 11.5 pt at 10 pt, but 8 and 14 pt show it scales with the font, as one line height). The legend's text width is then the widest line, and the legend box is at
        // most W / 3 + 9.0 wide (so a line up to 1.5 pt over the limit is not wrapped, it just moves plot and legend less). A top or bottom entry wraps when its key, text and two 0.283 em gaps exceed Lmax (below),
        // at that width less the key (measured in Word 2021 at 10 pt, W 200: "… number 1 here" 175.4 wide wraps, one line from W 220; the cap lies in 152.8..175.4).
        bool verticalLegend = hasLegend && chart.Legend is LegendPosition.Right or LegendPosition.Left;
        double rowCap = 0.9 * w + 2.85 - 0.566 * legendFont.FontSize;
        var lines = entries.Select((e, i) => verticalLegend ? WrapLegendName(e.Name, legendFont, w / 3 - 1.15 * legendFont.FontSize)
            : hasLegend && Advance(i) + widths[i] > rowCap ? WrapLegendName(e.Name, legendFont, rowCap - Advance(i)) : [e.Name]).ToList();
        int maxLines = lines.Select(l => l.Count).DefaultIfEmpty(1).Max();
        double lineStep = 1.15 * legendFont.FontSize; // lines inside an entry: 11.50 / 9.20 pt measured at 10 / 8 pt
        if (hasLegend)
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
        // Wrapped entries (measured at W 200, 1 and 3 entries of 2 lines): every row takes the tallest entry's pitch; a single row grows by 1.106 em per extra line.
        double extraLines = (maxLines - 1) * 1.106 * legendFont.FontSize;
        double extraRows = gridRows > 1 ? (gridRows - 1) * stackPitch - 1.5 : extraLines;

        // ---- the plot rectangle (relative to the chart's corner)
        double labelWidth = pie ? 0 : bar ? chart.Categories.Select(c => MeasureChartText(c, categoryFont)).DefaultIfEmpty(0).Max()
            : chart.Ticks.Select(t => MeasureChartText(t.Label, valueFont)).DefaultIfEmpty(0).Max();
        var bottomFont = bar ? valueFont : categoryFont;
        
        // The first value label is centered on the plot's top, which gives way to it: 12.82 at 14 pt, 11.71 at 12 pt, 11 at 8 and 10 pt (a line through the 12 and 14 pt samples; INFERRED beyond them).
        double topMargin = bar || pie ? 11 : Math.Max(11, 0.555 * valueFont.FontSize + 5.05);
        double left = pie ? 11 : 6.5 + labelWidth + WordAxisGap(bar ? categoryFont : valueFont), right = w - 11, top = topMargin + band, bottom = pie ? h - 11 : h - (6.5 + WordLineHeight(bottomFont) + 5.71 * bottomFont.FontSize / 10);
        double legendShift = hasLegend ? widest + 12.04 + legendFont.FontSize + (anyLine ? 14 : 0) : 0; // 22.04 at 10 pt; 20.04 and 26.04 measured at 8 and 14 pt
        if (verticalLegend) legendShift = Math.Min(legendShift, w / 3 + 9.0 + (anyLine ? 14 : 0)); // the box is at most W / 3 + 9.0 (measured at 10 pt)
        double textWidth = legendShift - 12.04 - legendFont.FontSize - (anyLine ? 14 : 0); // the text width the box leaves (the widest line, at most the cap)
        // A bar chart's category labels wrap at their spaces to lines of at most 0.484 (W - the side legend's box) - 10.9 (fitted to the brackets of 7 samples in Word 2021
        // at 10 pt, W 200 to 500, with and without a right legend; INFERRED at other sizes); the plot's left edge then follows the widest line.
        _wordBarLabels = null;
        if (bar && chart.Categories.Count > 0)
        {
            double cap = 0.484 * (w - (verticalLegend && !chart.LegendOverlay ? legendShift : 0)) - 10.9 * categoryFont.FontSize / 10;
            var barLines = chart.Categories.Select(c => WrapLegendName(c.Replace('\n', ' '), categoryFont, cap)).ToList();
            _wordBarLabels = (barLines, cap);
            double wrapped = barLines.SelectMany(l => l).Select(l => MeasureChartText(l, categoryFont)).DefaultIfEmpty(0).Max();
            left += wrapped - labelWidth;
        }
        // Overlay (measured in Word 2021, column charts at r, l, t, b and tr): the legend stands where the normal one does but the plot keeps the legend-less
        // layout and the legend is drawn after it.
        bool overlay = hasLegend && chart.LegendOverlay;
        // legendPos "tr" (measured, 2 samples: 3-entry column charts at 10 pt, with and without a title): the entries stack at the top right at the right legend's x,
        // the plot gives way on the right as for a right legend, on top by one entry pitch per entry and on the bottom by 6.0 pt. That the top offset is n x pitch,
        // other text sizes and other chart kinds are INFERRED; so is a tr legend with more entries than fit (the plot under 20 pt falls back to the shared layout).
        // The entries stack above the plot (measured in Word 2021, 18 samples at 10 pt, W 360, 3 to 5 entries, H 144 / 216 / 288, with and without a title): the first baseline is
        // 6.7 below the normal plot top, the plot top 10.4 below the last baseline (so (n - 1) x pitch + 11.1 below the normal top), the bottom 6.0 up. Word does that only while
        // the plot top stays within 0.43 H (stacked up to 0.425, beside from 0.432); otherwise the entries sit beside the plot: the plot keeps its normal top and
        // bottom and the first baseline is 6.0 lower (12.7 below the normal top). Normal tops here are measured 6.0 below this code's top (11 + band).
        // The 6.0 pt is the legend's own padding above its first entry, the same pad a top legend has above its row and a bottom legend below it (rowHeight = pitch + 6.0):
        // a tr legend's band starts 6.0 below the normal plot top whether it stands beside the plot or above it, and the stacked plot also gives way by 6.0 at the bottom.
        const double legendPad = 6.0;
        bool topRight = hasLegend && chart.WordLegendTopRight;
        // Stacked, the plot top is n x entry pitch below the normal top (55.54 / 77.27 measured at 8 / 14 pt, 3 entries); the first key's centre is the pad plus half a pitch below the band (6.0 beside the plot more).
        bool trStacked = topRight && top + entries.Count * entryPitch <= 0.43 * h;
        double trFirstCentre = band + legendPad + entryPitch / 2 + (trStacked ? 0 : 6.0);
        if (hasLegend && !overlay)
            switch (chart.Legend)
            {
                case LegendPosition.Right:
                    right -= legendShift;
                    // The last label of the horizontal value axis is centered on the plot's edge and the plot gives way by half of it (bars and scatter).
                    if (scatter && chart.XTicks.Count > 0) right -= MeasureChartText(chart.XTicks[^1].Label, categoryFont) / 2;
                    if (topRight) { if (trStacked) { top += entries.Count * entryPitch; bottom -= legendPad; } else top += legendPad; }
                    break;
                case LegendPosition.Left: left += legendShift; break;
                case LegendPosition.Top: top += rowHeight + extraRows; break;
                default: bottom -= rowHeight + extraRows; break;
            }
        // Bars: the plot gives way by half of the last value label (measured with a bottom legend at 8 and 14 pt; with a top or left legend or none it is INFERRED).
        if (bar && chart.Ticks.Count > 0) right -= MeasureChartText(chart.Ticks[^1].Label, valueFont) / 2;
        if (right - left < 20 || bottom - top < 20) return false;

        // ---- the category labels (columns and lines): they may wrap, turn or be skipped, which moves the plot's bottom, left and, when skipping, right edge
        _wordCategories = null;
        if (!bar && !pie && !scatter && chart.ShowCategoryAxis && chart.Categories.Count > 0 && styles.CategoryAxis?.Rotation is 90 or -90)
        {
            // An asked turn by 90 degrees (measured in Word 2021 at 10 pt, H 144 / 216 / 288): the labels wrap at their spaces to lines of at most 0.5 H - 12 (fitted, INFERRED)
            // and the plot's bottom stands 12.51 + the widest line (as this code measures it) above the chart's bottom.
            double fs = categoryFont.FontSize, cap = 0.5 * h - 12;
            var asked = chart.Categories.Select(c => WrapLegendName(c.Replace('\n', ' '), categoryFont, cap)).ToList();
            double askedWidest = asked.SelectMany(l => l).Select(l => MeasureChartText(l, categoryFont)).DefaultIfEmpty(0).Max();
            _wordCategories = new WordCategoryPlan(styles.CategoryAxis!.Rotation!.Value, 1, asked, askedWidest, Asked: true);
            bottom = h - (12.51 * fs / 10 + askedWidest);
            if (right - left < 20 || bottom - top < 20) { _wordCategories = null; return false; }
        }
        else if (!bar && !pie && !scatter && chart.ShowCategoryAxis && chart.Categories.Count > 0 && styles.CategoryAxis?.Rotation is null or 0 or 45)
        {
            double fs = categoryFont.FontSize, scale = fs / 10;
            int slots = chart.CategoriesBetween ? chart.Categories.Count : Math.Max(chart.Categories.Count - 1, 1);
            double slot = (right - left) / slots, restBottom = 6.5 + WordLineHeight(categoryFont) + 5.71 * scale;
            double labelBand = band + (hasLegend && !overlay && chart.Legend is LegendPosition.Top or LegendPosition.Bottom ? rowHeight + extraRows : 0);
            // An asked turn by 45 degrees (rot -2700000; measured in Word 2021 at 10 pt, H 216 / 288): never wraps or skips, cut as an automatic turn is.
            var plan = styles.CategoryAxis?.Rotation == 45 ? TruncateTurnedLabels(chart.Categories, categoryFont, 45, 1, h, labelBand)
                : PlanWordCategoryLabels(chart.Categories, categoryFont, slot, h, labelBand);
            _wordCategories = plan;
            if (plan.MaxLines > 1) bottom -= (plan.MaxLines - 1) * 1.15 * fs;
            else if (plan.Angle == 45) bottom += restBottom - (6.5 + WordTurnBottom * fs + 0.7071 * plan.Width);
            else if (plan.Angle == 90) bottom += restBottom - (6.5 + WordUprightBottom * fs + plan.Width);
            // Upright labels shown every Nth also take a little off the right: 0.41 and 1.14 pt at slots of 5.4 and 4.0 for 6.7 wide labels, none at 6.5 (fitted as half
            // of what the label's width exceeds its slot less 0.55 pt, the dependence on the width is INFERRED).
            if (plan.Angle == 0 && plan.Stride > 1) right -= Math.Max(0, 0.5 * (plan.Width - 0.55 * scale - slot));
            if (plan.Angle == 45)
            {
                // Skipped 45 degree labels also take a little off the right (2.0 to 3.5 pt measured at 8 to 16 skipped labels; fitted, INFERRED elsewhere).
                if (plan.Stride > 1) right -= Math.Max(0, 7.23 * scale - 0.5 * slot + 0.02 * (chart.Categories.Count - 20));
                // Every shown label's start must stay clear of the left edge, so a short first label lets the plot start further left (measured: "Quarterly…" then "Annual total west").
                for (int i = 0; i < chart.Categories.Count; i += plan.Stride)
                    left = Math.Max(left, 0.7071 * (plan.Widths?[i] ?? plan.Width) + WordTurnLeft - (2 * i + 1) * right / (2 * slots));
            }
            if (right - left < 20 || bottom - top < 20) { _wordCategories = null; return false; }
        }

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
            double first = (h + band) / 2 - shown * stackPitch / 2 + entryPitch / 2, trFirst = trFirstCentre;
            for (int i = 0; i < shown; i++)
                placed.Add((i, textLeft - Advance(i), topRight ? trFirst + i * entryPitch : first + i * stackPitch, textLeft));
        }
        else if (hasLegend)
        {
            double fs = legendFont.FontSize, start = w / 2 + 1.84 + 0.12 * (fs - 10);
            double bottomShift = gridRows > 1 && chart.Legend == LegendPosition.Bottom ? 1.5 : 0; // the last row's key top is 1.50 lower than a single row's
            // A bottom grid of wrapped entries keeps its last row's first line where a single line stands, the rest running past the chart's edge (measured).
            double firstCy = chart.Legend == LegendPosition.Top ? band + 6.0 + entryPitch / 2
                : h - 6.0 - entryPitch / 2 + bottomShift - (gridRows > 1 ? (gridRows - 1) * stackPitch : extraLines);
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
                    placed.Add((i, kx, firstCy + i / gridCols * stackPitch, kx + Advance(i)));
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
                            // The key's marker is capped at 0.6 em (measured in Word 2021: 7, 9 and 12 drawn 6.0 at 10 pt, 7 kept at 14 pt).
                            DrawMarker(items, kx + 9.6, ky, series?.MarkerSymbol ?? ChartMarkerSymbol.Auto, Math.Min(series?.MarkerDiameter ?? 5, 0.6 * legendFont.FontSize), series?.AutomaticMarker ?? new ChartAutoMarker(0, color, color, 0.75, 5), series?.MarkerStyle);
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
        finally { _wordPlot = null; _wordCategories = null; _wordBarLabels = null; }
        if (overlay) DrawLegend();
        return true;
    }
}
