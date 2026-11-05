using System.Globalization;
using System.Xml.Linq;
using DotOffice.Core;
using static DotOffice.Xlsx.DrawingMl;

namespace DotOffice.Xlsx;

/// <summary>
/// Reads a chart part (<c>c:chartSpace</c>) into a <see cref="SheetChart"/>: column, bar, line, area, pie, doughnut, scatter, bubble and radar charts. The data comes from the
/// chart's cached values when it has them and from the cells it refers to otherwise; the axis scale and the labels are computed here, so
/// that the layout only has to draw.
/// </summary>
internal static class XlsxCharts
{
    private static readonly HashSet<string> Supported = ["barChart", "lineChart", "pieChart", "doughnutChart", "areaChart", "scatterChart", "bubbleChart", "radarChart"];

    /// <summary>What a chart type is called in warnings.</summary>
    private static string Describe(string element) => element switch
    {
        "area3DChart" => "3-D area", "areaChart" => "area", "bar3DChart" => "3-D bar", "line3DChart" => "3-D line", "pie3DChart" => "3-D pie",
        "doughnutChart" => "doughnut", "scatterChart" => "scatter", "bubbleChart" => "bubble", "radarChart" => "radar", "stockChart" => "stock",
        "surfaceChart" or "surface3DChart" => "surface", "ofPieChart" => "pie-of-pie", _ => element.EndsWith("Chart", StringComparison.Ordinal) ? element[..^5] : element,
    };

    /// <summary>
    /// Prepares a chart: the unsupported chart type's name when it cannot be drawn, otherwise a function that builds the chart once all
    /// sheets are known (the chart may refer to cells of other sheets). <paramref name="notes"/> receives approximations.
    /// </summary>
    public static (string? Unsupported, Func<IReadOnlyList<Worksheet>, SheetChart>? Build) Prepare(XlsxContext context, Worksheet owner, XDocument doc, SheetChart template, List<string> notes)
    {
        var space = doc.Root;
        var plot = Path(space, "chart", "plotArea");
        if (space is null || plot is null) return ("empty", null);
        var groups = plot.Elements().Where(e => e.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToList();
        if (groups.Count == 0) return ("empty", null);
        var first = groups[0];
        if (!Supported.Contains(first.Name.LocalName)) return (Describe(first.Name.LocalName), null);
        // A combination of bar, line and area groups on one category axis is drawn whole (on the primary and the secondary value axis);
        // anything else after the first group is named and left out.
        var categoryTypes = new[] { "barChart", "lineChart", "areaChart" };
        var extras = categoryTypes.Contains(first.Name.LocalName) ? groups.Skip(1).Where(g => categoryTypes.Contains(g.Name.LocalName)).ToList() : [];
        var dropped = groups.Skip(1).Where(g => !extras.Contains(g)).ToList();
        if (dropped.Count > 0)
            notes.Add($"a combination chart draws only its {Describe(first.Name.LocalName)} part and not the {string.Join(" and ", dropped.Select(g => Describe(g.Name.LocalName)))} part");
        if (first.Name.LocalName is not ("scatterChart" or "bubbleChart") && extras.Count == 0 && plot.Elements().Count(e => e.Name.LocalName == "valAx") > 1) notes.Add("secondary axes are ignored");
        return (null, sheets => Build(context, owner, sheets, space, plot, first, extras, template));
    }

    private static SheetChart Build(XlsxContext context, Worksheet owner, IReadOnlyList<Worksheet> sheets, XElement space, XElement plot, XElement group, IReadOnlyList<XElement> extras, SheetChart template)
    {
        var theme = context.Theme;
        var chart = Child(space, "chart");
        // ---- data
        IEnumerable<Cell?> CellsOf(string formula)
        {
            foreach (var area in CellReference.ParseAreas(formula))
            {
                if (area.Range is not { } range) continue;
                var sheet = area.Sheet is null ? owner : sheets.FirstOrDefault(s => string.Equals(s.Name, area.Sheet, StringComparison.OrdinalIgnoreCase));
                for (int r = range.FirstRow; r <= range.LastRow; r++)
                    for (int c = range.FirstColumn; c <= range.LastColumn; c++)
                        yield return sheet is not null && sheet.Cells.TryGetValue((r, c), out var cell) ? cell : null;
            }
        }

        string Format(double value, string? code) =>
            string.IsNullOrEmpty(code) || code == "General" ? NumberFormat.Parse("General").FormatNumber(value, context.Date1904).Text : NumberFormat.Parse(code, null).FormatNumber(value, context.Date1904).Text;

        // The cached points of a reference or literal: index → (text, number), plus the format code.
        (List<(string? Text, double? Number)> Points, string? Format)? Cached(XElement? data)
        {
            var holder = data?.Elements().FirstOrDefault(e => e.Name.LocalName is "strCache" or "numCache" or "multiLvlStrCache")
                         ?? (data?.Name.LocalName is "strLit" or "numLit" ? data : null);
            if (holder is null) return null;
            int count = (int)Number(Child(holder, "ptCount"), "val");
            var points = new List<(string?, double?)>();
            foreach (var pt in holder.Elements().Where(e => e.Name.LocalName == "pt"))
            {
                int index = (int)Number(pt, "idx", -1);
                if (index < 0 || index > 100_000) continue;
                while (points.Count <= index) points.Add((null, null));
                var text = Child(pt, "v")?.Value;
                points[index] = (text, double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) ? v : null);
            }
            while (points.Count < count) points.Add((null, null));
            return points.Count == 0 && data!.Name.LocalName is not ("strLit" or "numLit") ? null : (points, Child(holder, "formatCode")?.Value);
        }

        List<string> Strings(XElement? data)
        {
            if (data is null) return [];
            var reference = Child(data, "strRef") ?? Child(data, "numRef") ?? Child(data, "multiLvlStrRef") ?? Child(data, "strLit") ?? Child(data, "numLit") ?? data;
            if (Cached(reference) is { } cache)
                return cache.Points.Select(p => p.Number is { } n && reference.Name.LocalName == "numRef" ? Format(n, cache.Format) : p.Text ?? "").ToList();
            return Child(reference, "f")?.Value is { } formula ? CellsOf(formula).Select(c => c?.Text ?? "").ToList() : [];
        }

        (List<double?> Values, string? Format) Numbers(XElement? data)
        {
            var reference = Child(data, "numRef") ?? Child(data, "numLit") ?? data;
            if (reference is not null && Cached(reference) is { } cache) return (cache.Points.Select(p => p.Number).ToList(), cache.Format);
            if (Child(reference, "f")?.Value is { } formula)
                return (CellsOf(formula).Select(c => c is { Kind: CellValueKind.Number } ? c.Value ?? (double.TryParse(c.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : null) : (double?)null).ToList(), null);
            return ([], null);
        }

        var categories = new List<string>();
        foreach (var ser in new[] { group }.Concat(extras).SelectMany(g => g.Elements().Where(e => e.Name.LocalName == "ser")))
        {
            categories = Strings(Child(ser, "cat"));
            if (categories.Count > 0) break;
        }

        string? sourceFormat = null;
        ChartGrouping GroupingOf(XElement g, ChartKind k) =>
            k is ChartKind.Pie or ChartKind.Doughnut or ChartKind.Radar or ChartKind.Scatter or ChartKind.Bubble ? ChartGrouping.Clustered
            : (string?)Child(g, "grouping")?.Attribute("val") switch { "stacked" => ChartGrouping.Stacked, "percentStacked" => ChartGrouping.PercentStacked, _ => ChartGrouping.Clustered };
        ChartKind KindOf(XElement g) => g.Name.LocalName switch
        {
            "pieChart" => ChartKind.Pie, "doughnutChart" => ChartKind.Doughnut, "lineChart" => ChartKind.Line, "areaChart" => ChartKind.Area,
            "scatterChart" => ChartKind.Scatter, "bubbleChart" => ChartKind.Bubble, "radarChart" => ChartKind.Radar,
            _ => ((string?)Child(g, "barDir")?.Attribute("val") ?? "col") == "bar" ? ChartKind.Bar : ChartKind.Column,
        };

        List<ChartSeries> ReadSeries(XElement group, ChartKind kind, int firstIndex)
        {
        bool slices = kind is ChartKind.Pie or ChartKind.Doughnut, xy = kind is ChartKind.Scatter or ChartKind.Bubble;
        var sers = group.Elements().Where(e => e.Name.LocalName == "ser").ToList();
        string? radarStyle = (string?)Child(group, "radarStyle")?.Attribute("val");
        bool filled = kind == ChartKind.Radar && radarStyle == "filled";
        var groupLabels = Child(group, "dLbls");
        var series = new List<ChartSeries>();
        for (int i = 0; i < sers.Count; i++)
        {
            var ser = sers[i];
            var (values, format) = Numbers(Child(ser, "val") ?? Child(ser, "yVal"));
            List<double?>? xValues = null, sizes = null;
            if (xy)
            {
                var xv = Numbers(Child(ser, "xVal")).Values;
                xValues = xv.Count > 0 ? xv : Enumerable.Range(1, values.Count).Select(k => (double?)k).ToList();
                if (kind == ChartKind.Bubble) sizes = Numbers(Child(ser, "bubbleSize")).Values;
            }
            sourceFormat ??= format;
            bool smooth = kind is ChartKind.Line or ChartKind.Scatter && Child(ser, "smooth") is { } sm && ((string?)sm.Attribute("val") is null || Flag(sm, "val"));
            string name = Child(Child(ser, "tx"), "v")?.Value is { } literal ? literal : Strings(Child(ser, "tx")).FirstOrDefault() is { Length: > 0 } cached ? cached : $"Series{firstIndex + i + 1}";
            var spPr = Child(ser, "spPr");
            bool lineLike = kind is ChartKind.Line or ChartKind.Scatter || (kind == ChartKind.Radar && !filled);
            var lineProperties = Child(spPr, "ln");
            Color color = (lineLike ? SolidFill(lineProperties, theme) ?? SolidFill(Child(Child(ser, "marker"), "spPr"), theme) : null) ?? SolidFill(spPr, theme) ?? PaletteColor(theme, firstIndex + i);
            bool line = !(lineLike && lineProperties is not null && Child(lineProperties, "noFill") is not null);

            bool marker = true;
            if (kind is ChartKind.Line or ChartKind.Scatter or ChartKind.Radar)
            {
                var symbol = (string?)Path(ser, "marker", "symbol")?.Attribute("val");
                marker = symbol is not null ? symbol != "none"
                    : kind == ChartKind.Radar ? radarStyle == "marker"
                    : (string?)Child(group, "marker")?.Attribute("val") is not ("0" or "false");
            }

            List<Color>? pointColors = null;
            if (slices)
            {
                int count = Math.Max(values.Count, categories.Count);
                pointColors = Enumerable.Range(0, count).Select(k => PaletteColor(theme, k)).ToList();
                foreach (var point in ser.Elements().Where(e => e.Name.LocalName == "dPt"))
                    if ((int)Number(Child(point, "idx"), "val", -1) is var k && k >= 0 && k < count && SolidFill(Child(point, "spPr"), theme) is { } c) pointColors[k] = c;
            }

            // Data labels: values and/or percentages, formatted like the label's number format or the source cells.
            var dLbls = Child(ser, "dLbls") ?? groupLabels;
            IReadOnlyList<string>? labels = null;
            if (dLbls is not null && (Flag(Child(dLbls, "showVal"), "val") || Flag(Child(dLbls, "showPercent"), "val")))
            {
                bool showValue = Flag(Child(dLbls, "showVal"), "val"), showPercent = Flag(Child(dLbls, "showPercent"), "val") && slices;
                var numFmt = Child(dLbls, "numFmt");
                string? code = numFmt is not null && !Flag(numFmt, "sourceLinked") ? (string?)numFmt.Attribute("formatCode") : format;
                double total = values.Sum(v => v is { } x && x > 0 ? x : 0);
                labels = values.Select(v =>
                {
                    if (v is not { } x) return "";
                    var parts = new List<string>();
                    if (showValue) parts.Add(Format(x, code));
                    if (showPercent) parts.Add(total > 0 ? (x / total).ToString("0%", CultureInfo.InvariantCulture) : "0%");
                    return string.Join("; ", parts);
                }).ToList();
            }
            double alpha = Opacity(spPr);
            if (alpha >= 1 && lineLike) alpha = 1;
            series.Add(new ChartSeries(name, values, color, marker, pointColors, labels, xValues, sizes, line, smooth, alpha,
                ReadTrendline(ser, color), ReadErrorBars(ser, values, xValues, kind)));
        }
        return series;
        }

        ChartTrendline? ReadTrendline(XElement ser, Color seriesColor)
        {
            if (Child(ser, "trendline") is not { } t) return null;
            var type = (string?)Child(t, "trendlineType")?.Attribute("val") switch
            {
                "exp" => TrendlineKind.Exponential, "log" => TrendlineKind.Logarithmic, "power" => TrendlineKind.Power, "poly" => TrendlineKind.Polynomial,
                "movingAvg" => TrendlineKind.MovingAverage, _ => TrendlineKind.Linear,
            };
            var ln = Path(t, "spPr", "ln");
            string? dash = (string?)Child(ln, "prstDash")?.Attribute("val");
            double width = ln?.Attribute("w") is { } w && double.TryParse(w.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var emu) && emu > 0 ? emu / 12700 : 2.25;
            return new ChartTrendline(type, SolidFill(ln, theme) ?? seriesColor, width, dash is not null and not "solid",
                (int)Math.Clamp(Number(Child(t, "order"), "val", 2), 2, 6), (int)Math.Clamp(Number(Child(t, "period"), "val", 2), 2, 100),
                Number(Child(t, "forward"), "val"), Number(Child(t, "backward"), "val"));
        }

        IReadOnlyList<ChartErrorBars>? ReadErrorBars(XElement ser, List<double?> values, List<double?>? xValues, ChartKind seriesKind)
        {
            var result = new List<ChartErrorBars>();
            foreach (var bars in ser.Elements().Where(e => e.Name.LocalName == "errBars"))
            {
                bool horizontal = (string?)Child(bars, "errDir")?.Attribute("val") == "x" && xValues is not null;
                var data = horizontal ? xValues! : values;
                var type = (string?)Child(bars, "errBarType")?.Attribute("val") ?? "both";
                var valueType = (string?)Child(bars, "errValType")?.Attribute("val") ?? "fixedVal";
                double amount = Number(Child(bars, "val"), "val", 1);
                var present = data.Where(v => v.HasValue).Select(v => v!.Value).ToList();
                double mean = present.Count > 0 ? present.Average() : 0;
                double deviation = present.Count > 1 ? Math.Sqrt(present.Sum(v => (v - mean) * (v - mean)) / (present.Count - 1)) : 0;
                var plus = new double[data.Count];
                var minus = new double[data.Count];
                var customPlus = valueType == "cust" ? Numbers(Child(bars, "plus")).Values : null;
                var customMinus = valueType == "cust" ? Numbers(Child(bars, "minus")).Values : null;
                for (int i = 0; i < data.Count; i++)
                {
                    if (data[i] is not { } v) continue;
                    double up = valueType switch
                    {
                        "percentage" => Math.Abs(v) * amount / 100,
                        "stdDev" => amount * deviation,
                        "stdErr" => present.Count > 0 ? deviation / Math.Sqrt(present.Count) : 0,
                        "cust" => customPlus is not null && customPlus.Count > 0 ? customPlus[Math.Min(i, customPlus.Count - 1)] ?? 0 : 0,
                        _ => amount,
                    };
                    double down = valueType == "cust" ? (customMinus is not null && customMinus.Count > 0 ? customMinus[Math.Min(i, customMinus.Count - 1)] ?? 0 : 0) : up;
                    plus[i] = type == "minus" ? 0 : Math.Abs(up);
                    minus[i] = type == "plus" ? 0 : Math.Abs(down);
                }
                var line = Path(bars, "spPr", "ln");
                double lineWidth = line?.Attribute("w") is { } w && double.TryParse(w.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var emu) && emu > 0 ? emu / 12700 : 1;
                result.Add(new ChartErrorBars(horizontal, plus, minus, !Flag(Child(bars, "noEndCap"), "val"), SolidFill(line, theme), lineWidth));
            }
            return result.Count > 0 ? result : null;
        }

        var kind = KindOf(group);
        bool slices = kind is ChartKind.Pie or ChartKind.Doughnut, xy = kind is ChartKind.Scatter or ChartKind.Bubble;
        bool filled = kind == ChartKind.Radar && (string?)Child(group, "radarStyle")?.Attribute("val") == "filled";
        var grouping = GroupingOf(group, kind);
        var series = ReadSeries(group, kind, 0);
        var extraSeries = new List<(XElement Group, ChartKind Kind, List<ChartSeries> Series)>();
        int nextIndex = series.Count;
        foreach (var extra in extras)
        {
            var extraKind = KindOf(extra);
            var read = ReadSeries(extra, extraKind, nextIndex);
            nextIndex += read.Count;
            extraSeries.Add((extra, extraKind, read));
        }
        int points = Math.Max(categories.Count, series.Concat(extraSeries.SelectMany(e => e.Series)).Select(s => s.Values.Count).DefaultIfEmpty(0).Max());
        while (categories.Count < points) categories.Add((categories.Count + 1).ToString(CultureInfo.InvariantCulture));

        // ---- axes and scale
        var valueAxes = plot.Elements().Where(e => e.Name.LocalName == "valAx").ToList();
        var groupAxisIds = group.Elements().Where(e => e.Name.LocalName == "axId").Select(e => (string?)e.Attribute("val")).ToList();
        var valueAxis = valueAxes.FirstOrDefault(a => groupAxisIds.Contains((string?)Child(a, "axId")?.Attribute("val"))) ?? valueAxes.FirstOrDefault();
        var categoryAxis = plot.Elements().FirstOrDefault(e => e.Name.LocalName is "catAx" or "dateAx" && groupAxisIds.Contains((string?)Child(e, "axId")?.Attribute("val")))
                           ?? plot.Elements().FirstOrDefault(e => e.Name.LocalName is "catAx" or "dateAx");
        if (xy)
        {
            // A scatter or bubble chart has two value axes: the horizontal one (axPos b/t, else the first) and the vertical one.
            categoryAxis = valueAxes.FirstOrDefault(a => (string?)Child(a, "axPos")?.Attribute("val") is "b" or "t") ?? valueAxes.FirstOrDefault();
            valueAxis = valueAxes.FirstOrDefault(a => a != categoryAxis);
        }
        string? TextOf(XElement? holder) => holder is null ? null : string.Join(" ", holder.Descendants().Where(e => e.Name.LocalName == "p").Select(p => string.Concat(p.Descendants().Where(e => e.Name.LocalName == "t").Select(t => t.Value))).Where(s => s.Length > 0)) is { Length: > 0 } text ? text : null;
        // A value axis' scale: the data range's automatic scale, or the fixed limits, units, logarithmic base and direction of the axis.
        string? AxisId(XElement? e) => (string?)Child(e, "axId")?.Attribute("val");
        (double Min, double Max, List<ChartTick> Ticks, double LogBase, bool Reversed) ValueScale(XElement? axis, double lo, double hi, double loPositive, ChartGrouping grouping, string? fallbackCode)
        {
            var scaling = Child(axis, "scaling");
            double? fixedMin = Child(scaling, "min") is { } mn ? Number(mn, "val") : null, fixedMax = Child(scaling, "max") is { } mx ? Number(mx, "val") : null;
            double? fixedMajor = Child(axis, "majorUnit") is { } mu && Number(mu, "val") > 0 ? Number(mu, "val") : null;
            bool reversed = (string?)Child(scaling, "orientation")?.Attribute("val") == "maxMin";
            double logBase = Child(scaling, "logBase") is { } lb ? Number(lb, "val") : 0;
            if (logBase < 2 || logBase > 1000) logBase = 0;
            string? code = Child(axis, "numFmt") is { } nf && !Flag(nf, "sourceLinked") ? (string?)nf.Attribute("formatCode") : null;
            code ??= fallbackCode;
            if (lo > hi) (lo, hi) = (0, 1);
            if (logBase > 0 && grouping == ChartGrouping.Clustered)
            {
                // Logarithmic axes run from a power of the base below the smallest value to one above the largest.
                if (loPositive == double.MaxValue || hi <= 0) (loPositive, hi) = (1, logBase);
                double lMin = fixedMin is > 0 ? fixedMin.Value : Math.Pow(logBase, Math.Floor(Math.Log(loPositive, logBase) + 1e-9));
                double lMax = fixedMax is { } fm && fm > lMin ? fm : Math.Pow(logBase, Math.Ceiling(Math.Log(hi, logBase) - 1e-9));
                if (lMax <= lMin) lMax = lMin * logBase;
                var logTicks = new List<ChartTick>();
                for (double v = Math.Pow(logBase, Math.Ceiling(Math.Log(lMin, logBase) - 1e-9)); v <= lMax * (1 + 1e-9) && logTicks.Count < 40; v *= logBase)
                    logTicks.Add(new ChartTick(v, Format(v, code)));
                return (lMin, lMax, logTicks, logBase, reversed);
            }
            if (grouping == ChartGrouping.PercentStacked)
            {
                double pMin = fixedMin ?? 0, pMax = fixedMax ?? 1;
                code ??= "0%";
                if (fallbackCode is not null && code == fallbackCode) code = "0%";
                return (pMin, pMax, Ticks(pMin, pMax, fixedMajor ?? 0.1, v => Format(v, code)), 0, reversed);
            }
            var scale = AutoScale(lo, hi, fixedMin, fixedMax, fixedMajor);
            return (scale.Min, scale.Max, Ticks(scale.Min, scale.Max, scale.Step, v => Format(v, code)), 0, reversed);
        }

        // The range of the data of some groups on one axis: plain values, or per-category sums when stacked.
        (double Lo, double Hi, double LoPositive) Range(IEnumerable<(ChartGrouping Grouping, IReadOnlyList<ChartSeries> Series)> groups)
        {
            double lo = double.MaxValue, hi = double.MinValue, loPositive = double.MaxValue;
            foreach (var (g, list) in groups)
            {
                foreach (var v in list.SelectMany(s => s.Values))
                    if (v is { } x) { (lo, hi) = (Math.Min(lo, x), Math.Max(hi, x)); if (x > 0) loPositive = Math.Min(loPositive, x); }
                // Error bars extend the automatic scale.
                foreach (var sr in list)
                    foreach (var bars in sr.ErrorBars?.Where(b => !b.Horizontal) ?? [])
                        for (int p = 0; p < sr.Values.Count && p < bars.Plus.Count; p++)
                            if (sr.Values[p] is { } x) (lo, hi) = (Math.Min(lo, x - bars.Minus[p]), Math.Max(hi, x + bars.Plus[p]));
                if (g == ChartGrouping.Stacked)
                    for (int p = 0; p < points; p++)
                    {
                        double positive = 0, negative = 0;
                        foreach (var s in list)
                            if (p < s.Values.Count && s.Values[p] is { } x) { if (x >= 0) positive += x; else negative += x; }
                        (lo, hi) = (Math.Min(lo, negative), Math.Max(hi, positive));
                    }
            }
            return (lo, hi, loPositive);
        }

        var ticks = new List<ChartTick>();
        double min = 0, max = 1, logBase = 0, secondaryMin = 0, secondaryMax = 1, secondaryLog = 0, xMin = 0, xMax = 1, xLog = 0;
        bool reverseValues = false, reverseX = false, reverseSecondary = false, secondaryShown = false;
        var xTicks = new List<ChartTick>();
        var secondaryTicks = new List<ChartTick>();
        string? secondaryTitle = null;
        var extraGroups = new List<ChartGroup>();
        if (xy)
        {
            double xlo = double.MaxValue, xhi = double.MinValue, xPositive = double.MaxValue;
            foreach (var v in series.SelectMany(sr => sr.XValues ?? []))
                if (v is { } x) { (xlo, xhi) = (Math.Min(xlo, x), Math.Max(xhi, x)); if (x > 0) xPositive = Math.Min(xPositive, x); }
            foreach (var sr in series)
                foreach (var bars in sr.ErrorBars?.Where(b => b.Horizontal) ?? [])
                    for (int p = 0; p < (sr.XValues?.Count ?? 0) && p < bars.Plus.Count; p++)
                        if (sr.XValues![p] is { } x) (xlo, xhi) = (Math.Min(xlo, x - bars.Minus[p]), Math.Max(xhi, x + bars.Plus[p]));
            string? xFallback = null;
            (xMin, xMax, xTicks, xLog, reverseX) = ValueScale(categoryAxis, xlo, xhi, xPositive, ChartGrouping.Clustered, xFallback);
        }
        if (!slices)
        {
            var primaryId = AxisId(valueAxis);
            var primary = new List<(ChartGrouping, IReadOnlyList<ChartSeries>)> { (grouping, series) };
            var secondary = new List<(ChartGrouping, IReadOnlyList<ChartSeries>)>();
            XElement? secondaryAxis = null;
            foreach (var (extra, extraKind, read) in extraSeries)
            {
                var extraAxis = valueAxes.FirstOrDefault(a => extra.Elements().Any(i => i.Name.LocalName == "axId" && (string?)i.Attribute("val") == AxisId(a)));
                bool onSecondary = extraAxis is not null && extraAxis != valueAxis;
                var extraGrouping = GroupingOf(extra, extraKind);
                (onSecondary ? secondary : primary).Add((extraGrouping, read));
                if (onSecondary) secondaryAxis = extraAxis;
                extraGroups.Add(new ChartGroup(extraKind, extraGrouping, read, onSecondary,
                    Child(extra, "gapWidth") is { } eg ? Number(eg, "val", 150) : 150,
                    Child(extra, "overlap") is { } eo ? Number(eo, "val") : extraGrouping == ChartGrouping.Clustered ? 0 : 100));
            }
            var (lo, hi, loPositive) = Range(primary);
            var primaryGrouping = primary.Any(g => g.Item1 == ChartGrouping.PercentStacked) ? ChartGrouping.PercentStacked : ChartGrouping.Clustered;
            (min, max, ticks, logBase, reverseValues) = ValueScale(valueAxis, lo, hi, loPositive, primaryGrouping, sourceFormat);
            if (secondaryAxis is not null)
            {
                var (slo, shi, sPositive) = Range(secondary);
                var secondaryGrouping = secondary.Any(g => g.Item1 == ChartGrouping.PercentStacked) ? ChartGrouping.PercentStacked : ChartGrouping.Clustered;
                (secondaryMin, secondaryMax, secondaryTicks, secondaryLog, reverseSecondary) = ValueScale(secondaryAxis, slo, shi, sPositive, secondaryGrouping, sourceFormat);
                secondaryShown = !Flag(Child(secondaryAxis, "delete"), "val");
                secondaryTitle = TextOf(Path(secondaryAxis, "title", "tx", "rich"));
            }
        }

        // ---- text and chrome
        var titleElement = Child(chart, "title");
        string? title = null;
        double titleSize = 14;
        if (titleElement is not null)
        {
            var rich = Path(titleElement, "tx", "rich");
            title = TextOf(rich);
            if (title is null && rich is null && series.Count == 1) title = series[0].Name; // an automatic title is the name of the only series
            if (Descendant(rich, "rPr", "defRPr")?.Attribute("sz") is { } sz && double.TryParse(sz.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) && size > 0) titleSize = size / 100;
        }
        else if (!Flag(Child(chart, "autoTitleDeleted"), "val") && series.Count == 1 && slices) title = series[0].Name;

        // Manual layouts count only when they are positions measured from the chart's edge, as Excel writes them.
        static (double X, double Y, double W, double H, bool Inner, bool HasSize)? ManualLayout(XElement? holder)
        {
            var m = Path(holder, "layout", "manualLayout");
            if (m is null || (string?)Child(m, "xMode")?.Attribute("val") != "edge" || (string?)Child(m, "yMode")?.Attribute("val") != "edge") return null;
            if (Child(m, "x") is null || Child(m, "y") is null) return null;
            double x = Number(Child(m, "x"), "val"), y = Number(Child(m, "y"), "val");
            double w = Number(Child(m, "w"), "val"), h = Number(Child(m, "h"), "val");
            if ((string?)Child(m, "wMode")?.Attribute("val") == "edge") w -= x;
            if ((string?)Child(m, "hMode")?.Attribute("val") == "edge") h -= y;
            return (x, y, w, h, (string?)Child(m, "layoutTarget")?.Attribute("val") == "inner", Child(m, "w") is not null && Child(m, "h") is not null);
        }
        ChartBox? plotBox = null;
        if (ManualLayout(plot) is { HasSize: true } pl && pl.X >= 0 && pl.Y >= 0 && pl.W > 0.05 && pl.H > 0.05 && pl.X + pl.W <= 1.001 && pl.Y + pl.H <= 1.001)
            plotBox = new ChartBox(pl.X, pl.Y, pl.W, pl.H, pl.Inner);
        (double X, double Y)? titlePosition = ManualLayout(titleElement) is { } tl && tl.X is >= 0 and < 1 && tl.Y is >= 0 and < 1 ? (tl.X, tl.Y) : null;

        var legendElement = Child(chart, "legend");
        var legend = legendElement is null ? LegendPosition.None : (string?)Child(legendElement, "legendPos")?.Attribute("val") switch
        {
            "l" => LegendPosition.Left, "t" => LegendPosition.Top, "b" => LegendPosition.Bottom, _ => LegendPosition.Right,
        };

        Color? background = Color.White, border = new Color(0xD9, 0xD9, 0xD9);
        if (Child(space, "spPr") is { } chartProperties)
        {
            background = SolidFill(chartProperties, theme, Color.White);
            var ln = Child(chartProperties, "ln");
            if (ln is not null && ln.Elements().Any(e => e.Name.LocalName is "noFill" or "solidFill")) border = SolidFill(ln, theme);
        }
        double fontSize = Descendant(Child(space, "txPr"), "defRPr")?.Attribute("sz") is { } baseSize && double.TryParse(baseSize.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var s0) && s0 > 0 ? s0 / 100 : 10;

        double gap = Child(group, "gapWidth") is { } g ? Number(g, "val", 150) : 150;
        double overlap = Child(group, "overlap") is { } o ? Number(o, "val") : grouping == ChartGrouping.Clustered ? 0 : 100;

        return template with
        {
            Kind = kind, Grouping = grouping, Title = title, TitleSize = titleSize, Categories = categories, Series = series, Legend = legend,
            Ticks = ticks, ValueMin = min, ValueMax = max,
            ShowValueAxis = valueAxis is not null && !Flag(Child(valueAxis, "delete"), "val"),
            ShowCategoryAxis = categoryAxis is not null && !Flag(Child(categoryAxis, "delete"), "val"),
            MajorGridlines = Child(valueAxis, "majorGridlines") is not null,
            ValueAxisTitle = TextOf(Path(valueAxis, "title", "tx", "rich")), CategoryAxisTitle = TextOf(Path(categoryAxis, "title", "tx", "rich")),
            GapWidth = gap, Overlap = Math.Clamp(overlap, -100, 100),
            ShowPercent = series.Any(s => s.Labels is not null) && slices,
            Background = background, Border = border, FontSize = fontSize,
            XTicks = xTicks, XMin = xMin, XMax = xMax, XGridlines = xy && Child(categoryAxis, "majorGridlines") is not null,
            CategoriesBetween = (string?)Child(valueAxis, "crossBetween")?.Attribute("val") is { } crossing ? crossing != "midCat" : kind != ChartKind.Area,
            HoleSize = Child(group, "holeSize") is { } hole ? Math.Clamp(Number(hole, "val", 50), 10, 90) / 100 : 0.5,
            FirstSliceAngle = Child(group, "firstSliceAng") is { } angle ? Number(angle, "val") : 0,
            BubbleScale = Child(group, "bubbleScale") is { } bubble ? Math.Clamp(Number(bubble, "val", 100), 0, 300) : 100,
            Filled = filled,
            LogBase = logBase, XLogBase = xLog, SecondaryLogBase = secondaryLog, ReverseValues = reverseValues, ReverseX = reverseX, ReverseSecondary = reverseSecondary,
            ReverseCategories = (string?)Child(Child(categoryAxis, "scaling"), "orientation")?.Attribute("val") == "maxMin" && !xy,
            ExtraGroups = extraGroups, SecondaryTicks = secondaryTicks, SecondaryMin = secondaryMin, SecondaryMax = secondaryMax,
            SecondaryAxisTitle = secondaryTitle, ShowSecondaryAxis = secondaryShown,
            PlotArea = plotBox, TitlePosition = titlePosition,
        };
    }

    private static XElement? Descendant(XElement? root, params string[] names) =>
        root?.Descendants().FirstOrDefault(e => names.Contains(e.Name.LocalName));

    /// <summary>The default color of the i-th series: the theme's accents, then darker and lighter variants of them.</summary>
    internal static Color PaletteColor(XlsxTheme theme, int index)
    {
        var accent = theme[4 + index % 6] ?? new Color(0x44, 0x72, 0xC4);
        return (index / 6) switch
        {
            0 => accent,
            1 => accent.Blend(Color.Black, 0.4),
            2 => accent.Blend(Color.White, 0.4),
            _ => accent.Blend(Color.Black, 0.2),
        };
    }

    private static List<ChartTick> Ticks(double min, double max, double step, Func<double, string> label)
    {
        var ticks = new List<ChartTick>();
        if (step <= 0 || !double.IsFinite(step)) step = 1;
        for (int i = 0; i <= 60; i++)
        {
            double value = Math.Round(min + i * step, 10);
            if (value > max + step * 1e-6) break;
            ticks.Add(new ChartTick(value, label(value)));
        }
        return ticks;
    }

    /// <summary>
    /// Excel's automatic value axis: the minimum is zero unless the data lies in a narrow band far from it, the maximum leaves 5 % of
    /// headroom, and the major unit is the smallest 1, 2 or 5 times a power of ten that gives at most ten intervals.
    /// </summary>
    internal static (double Min, double Max, double Step) AutoScale(double dataMin, double dataMax, double? fixedMin, double? fixedMax, double? fixedMajor)
    {
        double lo = dataMin, hi = dataMax;
        if (lo >= 0)
        {
            if (hi > 0 && (hi - lo) / hi <= 1.0 / 6) lo -= (hi - lo) / 2; else lo = 0;
            if (lo < 0 && dataMin >= 0) lo = Math.Max(lo, 0);
        }
        else if (hi <= 0)
        {
            if (lo < 0 && (hi - lo) / -lo <= 1.0 / 6) hi += (hi - lo) / 2; else hi = 0;
            if (hi > 0 && dataMax <= 0) hi = Math.Min(hi, 0);
        }
        if (hi > lo && hi > 0) hi += 0.05 * (hi - lo);
        if (lo < 0 && dataMin < 0) lo -= 0.05 * (hi - lo);
        if (fixedMin is { } fm) lo = fm;
        if (fixedMax is { } fx) hi = fx;
        if (hi <= lo) hi = lo + 1;

        double step;
        if (fixedMajor is { } fixedStep) step = fixedStep;
        else
        {
            double magnitude = Math.Pow(10, Math.Floor(Math.Log10((hi - lo) / 10)));
            step = magnitude;
            foreach (var factor in new[] { 1.0, 2, 5, 10 })
            {
                step = factor * magnitude;
                if ((hi - lo) / step <= 10 + 1e-9) break;
            }
        }
        if (fixedMin is null) lo = Math.Floor(lo / step + 1e-9) * step;
        if (fixedMax is null) hi = Math.Ceiling(hi / step - 1e-9) * step;
        if (hi <= lo) hi = lo + step;
        return (Math.Round(lo, 10), Math.Round(hi, 10), step);
    }
}
