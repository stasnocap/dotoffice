using System.Globalization;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Xlsx;

/// <summary>A conditional-format value object (<c>cfvo</c>): where a color scale, data bar or icon set threshold lies.</summary>
internal sealed record Cfvo(string Type, CfExpr? Value, bool GreaterOrEqual = true);

/// <summary>Settings a data bar takes from its Excel 2010 extension (<c>x14:dataBar</c>), matched by the rule's <c>x14:id</c>.</summary>
internal sealed record DataBarExtension(bool Gradient, double? MinLength, double? MaxLength, Color? NegativeColor, bool NegativeSameAsPositive);

/// <summary>One <c>cfRule</c> of the main namespace.</summary>
internal sealed class CfRule
{
    public required string Type { get; init; }
    public int Priority { get; init; }
    public bool StopIfTrue { get; init; }
    public int? DxfId { get; init; }
    public required List<CellRange> Ranges { get; init; }
    public int OriginRow { get; init; }
    public int OriginColumn { get; init; }

    public string? Operator { get; init; }
    public string? Text { get; init; }
    public List<CfExpr> Formulas { get; } = [];

    public int Rank { get; init; } = 10;
    public bool Percent { get; init; }
    public bool Bottom { get; init; }
    public bool AboveAverage { get; init; } = true;
    public bool EqualAverage { get; init; }
    public int StdDev { get; init; }

    public List<Cfvo> Cfvos { get; } = [];
    public List<Color> Colors { get; } = [];
    public bool ShowValue { get; init; } = true;
    public double MinLength { get; init; } = 10;
    public double MaxLength { get; init; } = 90;
    public string? X14Id { get; init; }

    /// <summary>The Excel 2010 settings of a data bar, once matched by <see cref="X14Id"/>.</summary>
    public DataBarExtension? Extension { get; set; }
    public string? IconSet { get; init; }
    public bool Reverse { get; init; }

    public bool Contains(int row, int column)
    {
        foreach (var range in Ranges) if (range.Contains(row, column)) return true;
        return false;
    }
}

/// <summary>The conditional formatting collected from one worksheet, plus what could not be understood.</summary>
internal sealed class CfSheet
{
    public List<CfRule> Rules { get; } = [];
    public Dictionary<string, DataBarExtension> Extensions { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Descriptions of unsupported rules, reported together as one warning.</summary>
    public SortedSet<string> Unsupported { get; } = new(StringComparer.Ordinal);

    public bool IsEmpty => Rules.Count == 0 && Unsupported.Count == 0;
}

internal static class ConditionalFormatting
{
    private static readonly HashSet<string> TextTypes = new(StringComparer.Ordinal)
    {
        "containsText", "notContainsText", "beginsWith", "endsWith", "containsBlanks", "notContainsBlanks", "containsErrors", "notContainsErrors",
    };

    private static readonly Dictionary<string, CellIcon[]> IconSets = new(StringComparer.Ordinal)
    {
        ["3Arrows"] = [new(IconShape.ArrowDown, Red), new(IconShape.ArrowSideways, Yellow), new(IconShape.ArrowUp, Green)],
        ["3TrafficLights1"] = [new(IconShape.Circle, Red), new(IconShape.Circle, Yellow), new(IconShape.Circle, Green)],
        ["3Symbols"] = [new(IconShape.Cross, Red), new(IconShape.Exclamation, Yellow), new(IconShape.Check, Green)],
    };

    private static Color Red => new(0xE0, 0x30, 0x30);
    private static Color Yellow => new(0xFF, 0xC0, 0x00);
    private static Color Green => new(0x00, 0xB0, 0x50);

    // ------------------------------------------------------------------ reading

    /// <summary>Reads a main-namespace <c>conditionalFormatting</c> element into rules; unsupported rules are noted in <paramref name="sheet"/>.</summary>
    public static void ReadBlock(XElement block, CfSheet sheet, Func<XElement?, Color?> resolve)
    {
        var ranges = ParseRanges((string?)block.Attribute("sqref"));
        foreach (var rule in block.Elements().Where(e => e.Name.LocalName == "cfRule"))
        {
            var type = (string?)rule.Attribute("type") ?? "";
            try
            {
                if (ranges.Count == 0) throw new NotSupportedException("no valid cell range");
                sheet.Rules.Add(ReadRule(rule, type, ranges, resolve));
            }
            catch (NotSupportedException e)
            {
                sheet.Unsupported.Add(type.Length == 0 ? $"a rule without a type ({e.Message})" : $"{type} ({e.Message})");
            }
        }
    }

    /// <summary>Reads the Excel 2010 data-bar extensions of a worksheet <c>extLst</c> (the x14 elements are not part of the main namespace).</summary>
    public static void ReadExtensions(XElement extList, CfSheet sheet, Func<XElement?, Color?> resolve)
    {
        foreach (var rule in extList.Descendants().Where(e => e.Name.LocalName == "cfRule" && (string?)e.Attribute("type") == "dataBar"))
        {
            if ((string?)rule.Attribute("id") is not { } id) continue;
            var bar = rule.Elements().FirstOrDefault(e => e.Name.LocalName == "dataBar");
            if (bar is null) continue;
            sheet.Extensions[id] = new DataBarExtension(
                Gradient: (string?)bar.Attribute("gradient") is not ("0" or "false"),
                MinLength: Number(bar, "minLength"),
                MaxLength: Number(bar, "maxLength"),
                NegativeColor: resolve(bar.Elements().FirstOrDefault(e => e.Name.LocalName == "negativeFillColor")),
                NegativeSameAsPositive: (string?)bar.Attribute("negativeBarColorSameAsPositive") is "1" or "true");
        }
    }

    private static double? Number(XElement e, string name) =>
        double.TryParse((string?)e.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) ? v : null;

    private static List<CellRange> ParseRanges(string? sqref)
    {
        var ranges = new List<CellRange>();
        if (string.IsNullOrWhiteSpace(sqref)) return ranges;
        foreach (var area in CellReference.ParseAreas(string.Join(',', sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries))))
        {
            if (area.Range is { } range) ranges.Add(range);
            else if (area.RowSpan is { } rows) ranges.Add(new CellRange(rows.From, 0, rows.To, CellReference.MaxColumns - 1));
            else if (area.ColumnSpan is { } columns) ranges.Add(new CellRange(0, columns.From, CellReference.MaxRows - 1, columns.To));
        }
        return ranges;
    }

    private static CfRule ReadRule(XElement rule, string type, List<CellRange> ranges, Func<XElement?, Color?> resolve)
    {
        int Int(string name, int fallback) => int.TryParse((string?)rule.Attribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : fallback;
        bool Flag(string name, bool fallback) => (string?)rule.Attribute(name) switch { "1" or "true" => true, "0" or "false" => false, _ => fallback };
        XElement? Child(XElement? parent, string name) => parent?.Elements().FirstOrDefault(e => e.Name.LocalName == name);

        bool known = type is "cellIs" or "expression" or "top10" or "aboveAverage" or "duplicateValues" or "uniqueValues"
                     or "colorScale" or "dataBar" or "iconSet" || TextTypes.Contains(type);
        if (!known) throw new NotSupportedException("rule type not supported");

        int dxf = Int("dxfId", -1);
        var bar = Child(rule, "dataBar");
        var icons = Child(rule, "iconSet");
        var result = new CfRule
        {
            Type = type,
            Priority = Int("priority", int.MaxValue),
            StopIfTrue = Flag("stopIfTrue", false),
            DxfId = dxf >= 0 ? dxf : null,
            Ranges = ranges,
            OriginRow = ranges.Min(r => r.FirstRow),
            OriginColumn = ranges.Min(r => r.FirstColumn),
            Operator = (string?)rule.Attribute("operator"),
            Text = (string?)rule.Attribute("text"),
            Rank = Math.Max(Int("rank", 10), 1),
            Percent = Flag("percent", false),
            Bottom = Flag("bottom", false),
            AboveAverage = Flag("aboveAverage", true),
            EqualAverage = Flag("equalAverage", false),
            StdDev = Math.Max(Int("stdDev", 0), 0),
            ShowValue = bar is not null ? ShowValueOf(bar) : icons is not null ? ShowValueOf(icons) : true,
            MinLength = bar is not null ? Number(bar, "minLength") ?? 10 : 10,
            MaxLength = bar is not null ? Number(bar, "maxLength") ?? 90 : 90,
            X14Id = rule.Descendants().FirstOrDefault(e => e.Name.LocalName == "id")?.Value.Trim(),
            IconSet = icons is null ? null : (string?)icons.Attribute("iconSet") ?? "3TrafficLights1",
            Reverse = icons is not null && (string?)icons.Attribute("reverse") is "1" or "true",
        };

        foreach (var formula in rule.Elements().Where(e => e.Name.LocalName == "formula"))
            result.Formulas.Add(Parse(formula.Value, "formula"));

        switch (type)
        {
            case "cellIs":
            {
                int needed = result.Operator is "between" or "notBetween" ? 2 : 1;
                if (result.Operator is not ("lessThan" or "lessThanOrEqual" or "equal" or "notEqual" or "greaterThanOrEqual" or "greaterThan" or "between" or "notBetween"))
                    throw new NotSupportedException("missing or unknown operator");
                if (result.Formulas.Count < needed) throw new NotSupportedException("missing operand");
                break;
            }
            case "expression":
                if (result.Formulas.Count < 1) throw new NotSupportedException("missing formula");
                break;
            case "containsText" or "notContainsText" or "beginsWith" or "endsWith":
                if (result.Text is null) throw new NotSupportedException("missing text");
                break;
            case "colorScale":
            {
                var scale = Child(rule, "colorScale") ?? throw new NotSupportedException("missing color scale");
                ReadCfvos(scale, result, resolve);
                foreach (var color in scale.Elements().Where(e => e.Name.LocalName == "color"))
                    result.Colors.Add(resolve(color) ?? throw new NotSupportedException("unresolvable color"));
                if (result.Cfvos.Count is not (2 or 3) || result.Colors.Count != result.Cfvos.Count)
                    throw new NotSupportedException("only 2- and 3-color scales are supported");
                break;
            }
            case "dataBar":
            {
                if (bar is null) throw new NotSupportedException("missing data bar");
                ReadCfvos(bar, result, resolve);
                if (result.Cfvos.Count != 2) throw new NotSupportedException("a data bar needs two limits");
                result.Colors.Add(resolve(Child(bar, "color")) ?? new Color(0x63, 0x8E, 0xC6));
                break;
            }
            case "iconSet":
            {
                if (icons is null) throw new NotSupportedException("missing icon set");
                if (!IconSets.TryGetValue(result.IconSet!, out var set)) throw new NotSupportedException($"icon set {result.IconSet}");
                ReadCfvos(icons, result, resolve);
                if (result.Cfvos.Count != set.Length) throw new NotSupportedException("unexpected number of thresholds");
                break;
            }
        }
        return result;
    }

    private static bool ShowValueOf(XElement e) => (string?)e.Attribute("showValue") is not ("0" or "false");

    private static void ReadCfvos(XElement parent, CfRule rule, Func<XElement?, Color?> resolve)
    {
        foreach (var cfvo in parent.Elements().Where(e => e.Name.LocalName == "cfvo"))
        {
            var kind = (string?)cfvo.Attribute("type") ?? "";
            if (kind is not ("min" or "max" or "num" or "percent" or "percentile" or "formula" or "autoMin" or "autoMax"))
                throw new NotSupportedException($"threshold type '{kind}'");
            var value = (string?)cfvo.Attribute("val");
            CfExpr? expr = kind is "min" or "max" or "autoMin" or "autoMax" || string.IsNullOrWhiteSpace(value) ? null : Parse(value, "threshold");
            if (expr is null && kind is "num" or "percent" or "percentile" or "formula") throw new NotSupportedException("threshold without a value");
            rule.Cfvos.Add(new Cfvo(kind, expr, (string?)cfvo.Attribute("gte") is not ("0" or "false")));
        }
    }

    private static CfExpr Parse(string formula, string what)
    {
        try
        {
            return CfParser.Parse(formula);
        }
        catch (NotSupportedException e)
        {
            throw new NotSupportedException($"{what}: {e.Message}");
        }
    }

    // ------------------------------------------------------------------ applying

    /// <summary>Everything an evaluation needs besides the sheet.</summary>
    public sealed record Environment(IReadOnlyList<Dxf> Dxfs, bool Date1904);

    /// <summary>What the rules decided for one cell; the first rule (highest priority) to set a property keeps it.</summary>
    private sealed class Outcome
    {
        public bool? Bold, Italic, Strike;
        public UnderlineStyle? Underline;
        public Color? FontColor, Fill;
        public Border? Left, Right, Top, Bottom;
        public string? NumberFormat;
        public CellDataBar? Bar;
        public CellIcon? Icon;
        public bool HideValue;
        public bool Any;

        public void Add(Dxf d)
        {
            Any = true;
            Bold ??= d.Bold;
            Italic ??= d.Italic;
            Strike ??= d.Strike;
            Underline ??= d.Underline;
            FontColor ??= d.FontColor;
            Fill ??= d.Fill;
            Left ??= d.Left;
            Right ??= d.Right;
            Top ??= d.Top;
            Bottom ??= d.Bottom;
            NumberFormat ??= d.NumberFormat;
        }
    }

    /// <summary>Sorted numbers of a rule's range with the statistics the rule types need.</summary>
    private sealed class Stats
    {
        public double[] Sorted = [];
        public double Min, Max, Average, StdDev;
        public Dictionary<string, int>? Counts;
    }

    public static void Apply(Worksheet sheet, CfSheet cf, Environment env, Action<string> warn)
    {
        if (cf.Unsupported.Count > 0)
            warn($"uses conditional formatting that is not supported ({string.Join(", ", cf.Unsupported)}); those rules are ignored.");
        if (cf.Rules.Count == 0) return;
        foreach (var rule in cf.Rules)
            if (rule.Type == "dataBar" && rule.X14Id is { } id && cf.Extensions.TryGetValue(id, out var ext))
                rule.Extension = ext;

        if (Bounds(sheet) is not { } bounds) return;
        var rules = cf.Rules.Select((r, i) => (Rule: r, Order: i)).OrderBy(x => x.Rule.Priority).ThenBy(x => x.Order).Select(x => x.Rule).ToList();
        var clamped = new Dictionary<CfRule, List<CellRange>>();
        foreach (var rule in rules)
        {
            var list = rule.Ranges.Select(r => Intersect(r, bounds)).OfType<CellRange>().ToList();
            if (list.Count > 0) clamped[rule] = list;
        }
        if (clamped.Count == 0) return;

        var positions = new HashSet<(int, int)>();
        foreach (var list in clamped.Values)
            foreach (var range in list)
                for (int r = range.FirstRow; r <= range.LastRow; r++)
                    for (int c = range.FirstColumn; c <= range.LastColumn; c++)
                        positions.Add((r, c));

        var cells = sheet.Cells;
        object ValueAt(int row, int column) => ValueOf(cells.GetValueOrDefault((row, column)));
        var stats = new Dictionary<CfRule, Stats>();
        Stats StatsOf(CfRule rule)
        {
            if (stats.TryGetValue(rule, out var existing)) return existing;
            var numbers = new List<double>();
            foreach (var ((r, c), cell) in cells)
                if (cell.Kind == CellValueKind.Number && cell.Value is { } v && clamped[rule].Any(range => range.Contains(r, c))) numbers.Add(v);
            numbers.Sort();
            var s = new Stats { Sorted = numbers.ToArray() };
            if (numbers.Count > 0)
            {
                s.Min = numbers[0];
                s.Max = numbers[^1];
                s.Average = numbers.Average();
                s.StdDev = numbers.Count > 1 ? Math.Sqrt(numbers.Sum(x => (x - s.Average) * (x - s.Average)) / (numbers.Count - 1)) : 0;
            }
            return stats[rule] = s;
        }
        Dictionary<string, int> CountsOf(CfRule rule)
        {
            var s = StatsOf(rule);
            if (s.Counts is not null) return s.Counts;
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var ((r, c), cell) in cells)
            {
                if (!clamped[rule].Any(range => range.Contains(r, c)) || KeyOf(ValueOf(cell)) is not { } key) continue;
                counts[key] = counts.GetValueOrDefault(key) + 1;
            }
            return s.Counts = counts;
        }

        var outcomes = new Dictionary<(int, int), Outcome>();
        foreach (var (row, column) in positions)
        {
            Outcome? outcome = null;
            foreach (var rule in rules)
            {
                if (!clamped.TryGetValue(rule, out var ranges) || !ranges.Any(range => range.Contains(row, column))) continue;
                var cell = cells.GetValueOrDefault((row, column));
                var ctx = new CfContext(ValueAt, rule.OriginRow, rule.OriginColumn, row, column, bounds);
                bool scaleLike = rule.Type is "colorScale" or "dataBar" or "iconSet";
                if (scaleLike)
                {
                    if (cell is not { Kind: CellValueKind.Number, Value: { } value }) continue;
                    outcome ??= new Outcome();
                    ApplyScale(rule, value, StatsOf(rule), ctx, outcome);
                    continue;
                }
                if (!Matches(rule, cell, ctx, StatsOf, CountsOf)) continue;
                outcome ??= new Outcome();
                if (rule.DxfId is { } id && id >= 0 && id < env.Dxfs.Count) outcome.Add(env.Dxfs[id]);
                if (rule.StopIfTrue) break;
            }
            if (outcome is not null) outcomes[(row, column)] = outcome;
        }

        foreach (var ((row, column), outcome) in outcomes)
        {
            if (cells.TryGetValue((row, column), out var cell)) cells[(row, column)] = Bake(cell, outcome, env);
            else if (outcome.Fill is not null || outcome.Left is not null || outcome.Right is not null || outcome.Top is not null || outcome.Bottom is not null)
                cells[(row, column)] = Bake(new Cell(), outcome, env);
        }
    }

    /// <summary>The cells conditional formatting can touch: the used range and the print area, without runaway whole-column areas.</summary>
    private static CellRange? Bounds(Worksheet sheet)
    {
        int r0 = int.MaxValue, c0 = int.MaxValue, r1 = -1, c1 = -1;
        foreach (var (row, column) in sheet.Cells.Keys)
        {
            r0 = Math.Min(r0, row);
            c0 = Math.Min(c0, column);
            r1 = Math.Max(r1, row);
            c1 = Math.Max(c1, column);
        }
        var used = r1 < 0 ? (CellRange?)null : new CellRange(r0, c0, r1, c1);
        var result = used;
        foreach (var printArea in sheet.PageSetup.PrintAreas)
        {
            var a = printArea;
            int lastRow = a.RowCount > 100_000 && used is { } u1 ? Math.Min(a.LastRow, u1.LastRow) : a.LastRow;
            int lastColumn = a.ColumnCount > 1_000 && used is { } u2 ? Math.Min(a.LastColumn, u2.LastColumn) : a.LastColumn;
            a = new CellRange(a.FirstRow, a.FirstColumn, Math.Max(lastRow, a.FirstRow), Math.Max(lastColumn, a.FirstColumn));
            var union = result is { } u ? new CellRange(Math.Min(u.FirstRow, a.FirstRow), Math.Min(u.FirstColumn, a.FirstColumn), Math.Max(u.LastRow, a.LastRow), Math.Max(u.LastColumn, a.LastColumn)) : a;
            if ((long)union.RowCount * union.ColumnCount <= 4_000_000) result = union;
        }
        return result;
    }

    private static CellRange? Intersect(CellRange a, CellRange b)
    {
        int r0 = Math.Max(a.FirstRow, b.FirstRow), r1 = Math.Min(a.LastRow, b.LastRow);
        int c0 = Math.Max(a.FirstColumn, b.FirstColumn), c1 = Math.Min(a.LastColumn, b.LastColumn);
        return r0 > r1 || c0 > c1 ? null : new CellRange(r0, c0, r1, c1);
    }

    /// <summary>A cell as a formula value.</summary>
    internal static object ValueOf(Cell? cell)
    {
        if (cell is null) return CfBlank.Instance;
        switch (cell.Kind)
        {
            case CellValueKind.Number:
                if (cell.Value is { } v) return v;
                return double.TryParse(cell.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed : cell.Text;
            case CellValueKind.Bool: return cell.Text == "TRUE";
            case CellValueKind.Error: return new CfError(cell.Text);
            default: return cell.Text.Length == 0 ? CfBlank.Instance : cell.Text;
        }
    }

    private static string? KeyOf(object value) => value switch
    {
        double d => "n" + d.ToString("R", CultureInfo.InvariantCulture),
        string s => "s" + s.ToUpperInvariant(),
        bool b => "b" + b,
        _ => null,
    };

    private static bool Matches(CfRule rule, Cell? cell, CfContext ctx, Func<CfRule, Stats> statsOf, Func<CfRule, Dictionary<string, int>> countsOf)
    {
        var value = ValueOf(cell);
        switch (rule.Type)
        {
            case "expression":
                return CfEvaluator.IsTrue(CfEvaluator.Evaluate(rule.Formulas[0], ctx));
            case "cellIs":
            {
                if (value is CfError) return false;
                var a = CfEvaluator.Evaluate(rule.Formulas[0], ctx);
                if (a is CfError) return false;
                switch (rule.Operator)
                {
                    case "between" or "notBetween":
                    {
                        var b = CfEvaluator.Evaluate(rule.Formulas[1], ctx);
                        if (b is CfError) return false;
                        bool inside = CfEvaluator.Compare(value, a) >= 0 && CfEvaluator.Compare(value, b) <= 0;
                        return rule.Operator == "between" ? inside : !inside;
                    }
                    default:
                    {
                        int cmp = CfEvaluator.Compare(value, a);
                        return rule.Operator switch
                        {
                            "lessThan" => cmp < 0,
                            "lessThanOrEqual" => cmp <= 0,
                            "equal" => cmp == 0,
                            "notEqual" => cmp != 0,
                            "greaterThanOrEqual" => cmp >= 0,
                            "greaterThan" => cmp > 0,
                            _ => false,
                        };
                    }
                }
            }
            case "top10":
            {
                if (value is not double v) return false;
                var s = statsOf(rule);
                int n = s.Sorted.Length;
                int k = rule.Percent ? Math.Max(1, (int)Math.Floor(n * rule.Rank / 100.0)) : rule.Rank;
                k = Math.Min(k, n);
                return rule.Bottom ? v <= s.Sorted[k - 1] : v >= s.Sorted[n - k];
            }
            case "aboveAverage":
            {
                if (value is not double v) return false;
                var s = statsOf(rule);
                double limit = s.Average + (rule.AboveAverage ? 1 : -1) * rule.StdDev * s.StdDev;
                if (rule.AboveAverage) return rule.EqualAverage && rule.StdDev == 0 ? v >= limit : v > limit;
                return rule.EqualAverage && rule.StdDev == 0 ? v <= limit : v < limit;
            }
            case "duplicateValues" or "uniqueValues":
            {
                if (KeyOf(value) is not { } key) return false;
                int count = countsOf(rule).GetValueOrDefault(key);
                return rule.Type == "duplicateValues" ? count > 1 : count == 1;
            }
            case "containsText" or "notContainsText" or "beginsWith" or "endsWith":
            {
                if (value is CfError) return false;
                var text = CfEvaluator.ToText(value);
                var needle = rule.Text!;
                return rule.Type switch
                {
                    "containsText" => text.Contains(needle, StringComparison.OrdinalIgnoreCase),
                    "notContainsText" => !text.Contains(needle, StringComparison.OrdinalIgnoreCase),
                    "beginsWith" => text.StartsWith(needle, StringComparison.OrdinalIgnoreCase),
                    _ => text.EndsWith(needle, StringComparison.OrdinalIgnoreCase),
                };
            }
            case "containsBlanks": return value is CfBlank || value is string s0 && s0.Trim().Length == 0;
            case "notContainsBlanks": return !(value is CfBlank || value is string s1 && s1.Trim().Length == 0);
            case "containsErrors": return value is CfError;
            case "notContainsErrors": return value is not CfError;
        }
        return false;
    }

    // ---- color scales, data bars, icon sets

    private static double Percentile(double[] sorted, double fraction)
    {
        if (sorted.Length == 0) return 0;
        double position = Math.Clamp(fraction, 0, 1) * (sorted.Length - 1);
        int lower = (int)Math.Floor(position);
        int upper = Math.Min(lower + 1, sorted.Length - 1);
        return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
    }

    private static double Threshold(Cfvo cfvo, Stats stats, CfContext ctx)
    {
        double Constant() => cfvo.Value is { } e && CfEvaluator.ToNumber(CfEvaluator.Evaluate(e, ctx)) is { } n ? n : stats.Min;
        return cfvo.Type switch
        {
            "min" or "autoMin" => stats.Min,
            "max" or "autoMax" => stats.Max,
            "percent" => stats.Min + (stats.Max - stats.Min) * Constant() / 100,
            "percentile" => Percentile(stats.Sorted, Constant() / 100),
            _ => Constant(),
        };
    }

    private static void ApplyScale(CfRule rule, double value, Stats stats, CfContext ctx, Outcome outcome)
    {
        switch (rule.Type)
        {
            case "colorScale":
            {
                if (outcome.Fill is not null) return;
                var limits = rule.Cfvos.Select(c => Threshold(c, stats, ctx)).ToList();
                var colors = rule.Colors;
                outcome.Any = true;
                if (value <= limits[0]) { outcome.Fill = colors[0]; return; }
                for (int i = 1; i < limits.Count; i++)
                {
                    if (value > limits[i]) continue;
                    double span = limits[i] - limits[i - 1];
                    outcome.Fill = span <= 1e-12 ? colors[i] : colors[i - 1].Blend(colors[i], (value - limits[i - 1]) / span);
                    return;
                }
                outcome.Fill = colors[^1];
                return;
            }
            case "dataBar":
            {
                if (outcome.Bar is not null) return;
                var ext = rule.Extension;
                double min = Threshold(rule.Cfvos[0], stats, ctx), max = Threshold(rule.Cfvos[1], stats, ctx);
                var color = rule.Colors[0];
                bool gradient = ext?.Gradient ?? true;
                outcome.Any = true;
                if (ext is not null && min < 0 && max > min)
                {
                    // Excel 2010 bars with negative values grow from an axis placed in proportion to the range.
                    double axis = Math.Clamp(-min / (max - min), 0, 1);
                    if (value < 0)
                    {
                        double length = Math.Clamp(value / min, 0, 1) * axis;
                        outcome.Bar = new CellDataBar(axis - length, axis, ext.NegativeSameAsPositive ? color : ext.NegativeColor ?? new Color(0xFF, 0, 0), gradient, Negative: true);
                    }
                    else
                    {
                        double length = max > 0 ? Math.Clamp(value / max, 0, 1) * (1 - axis) : 0;
                        outcome.Bar = new CellDataBar(axis, axis + length, color, gradient);
                    }
                }
                else
                {
                    double minLength = (ext?.MinLength ?? rule.MinLength) / 100, maxLength = (ext?.MaxLength ?? rule.MaxLength) / 100;
                    double position = max - min <= 1e-12 ? 1 : Math.Clamp((value - min) / (max - min), 0, 1);
                    outcome.Bar = new CellDataBar(0, minLength + position * (maxLength - minLength), color, gradient);
                }
                if (!rule.ShowValue) outcome.HideValue = true;
                return;
            }
            default:
            {
                if (outcome.Icon is not null) return;
                var set = IconSets[rule.IconSet!];
                int index = 0;
                for (int i = 1; i < rule.Cfvos.Count; i++)
                {
                    double limit = Threshold(rule.Cfvos[i], stats, ctx);
                    if (rule.Cfvos[i].GreaterOrEqual ? value >= limit : value > limit) index = i;
                }
                outcome.Any = true;
                outcome.Icon = set[rule.Reverse ? set.Length - 1 - index : index];
                if (!rule.ShowValue) outcome.HideValue = true;
                return;
            }
        }
    }

    // ---- writing the outcome into the cell

    private static Cell Bake(Cell cell, Outcome o, Environment env)
    {
        RunProperties Font(RunProperties f) => f with
        {
            Bold = o.Bold ?? f.Bold,
            Italic = o.Italic ?? f.Italic,
            Strike = o.Strike ?? f.Strike,
            Underline = o.Underline ?? f.Underline,
            Color = o.FontColor ?? f.Color,
        };

        var style = cell.Style;
        var text = cell.Text;
        if (o.NumberFormat is { } code && cell.Kind == CellValueKind.Number && cell.Value is { } number)
        {
            var shown = NumberFormat.Parse(code).FormatNumber(number, env.Date1904);
            text = shown.Text;
            if (shown.Color is { } c && o.FontColor is null) style = style with { Font = style.Font with { Color = c } };
        }
        style = style with
        {
            Font = Font(style.Font),
            Fill = o.Fill ?? style.Fill,
            Gradient = o.Fill is not null ? null : style.Gradient,
            Borders = style.Borders with
            {
                Left = o.Left ?? style.Borders.Left,
                Right = o.Right ?? style.Borders.Right,
                Top = o.Top ?? style.Borders.Top,
                Bottom = o.Bottom ?? style.Borders.Bottom,
            },
        };
        var runs = text == cell.Text ? cell.Runs?.Select(r => r with { Font = Font(r.Font) }).ToList() : null;
        if (o.HideValue) (text, runs) = ("", null);
        return cell with { Style = style, Text = text, Runs = runs, DataBar = o.Bar ?? cell.DataBar, Icon = o.Icon ?? cell.Icon };
    }
}
