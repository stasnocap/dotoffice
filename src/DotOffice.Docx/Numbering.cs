using System.Text;
using System.Xml;
using DotOffice.Core;

namespace DotOffice.Docx;

internal sealed record NumberingLevel
{
    public int Start { get; init; } = 1;
    public string Format { get; init; } = "decimal";
    public string? Text { get; init; }
    public ParagraphAlignment Alignment { get; init; }
    public ListLabelSuffix Suffix { get; init; }

    /// <summary>1-based level after which this level restarts; 0 = never; null = after any higher level.</summary>
    public int? Restart { get; init; }

    public bool IsLegal { get; init; }
    public string? ParagraphStyle { get; init; }
    public ParaFormat? Para { get; init; }
    public RunFormat? Run { get; init; }
}

internal sealed class AbstractNumbering
{
    public NumberingLevel?[] Levels { get; } = new NumberingLevel?[9];
    public string? NumStyleLink { get; init; }
}

internal sealed class NumberingInstance
{
    public int AbstractId { get; init; }
    public Dictionary<int, int> StartOverrides { get; } = [];
    public Dictionary<int, NumberingLevel> LevelOverrides { get; } = [];
}

/// <summary>A list label computed for one paragraph.</summary>
internal sealed record NumberingResult(string Text, NumberingLevel Level);

/// <summary>
/// numbering.xml definitions plus the running list counters. Counters are shared per abstract
/// definition, except for num instances that override a start value, which count on their own.
/// </summary>
internal sealed class Numbering
{
    private readonly Dictionary<int, AbstractNumbering> _abstracts = [];
    private readonly Dictionary<int, NumberingInstance> _nums = [];
    private readonly Dictionary<string, int?[]> _counters = new(StringComparer.Ordinal);

    public static Numbering Read(XmlReader r)
    {
        var numbering = new Numbering();
        r.MoveToContent();
        int depth = r.Depth;
        if (!r.Enter()) return numbering;
        while (r.NextChild(depth))
        {
            if (!r.IsW()) { r.Skip(); continue; }
            if (r.LocalName == "abstractNum" && r.IntAttr("abstractNumId") is { } aid) numbering._abstracts[aid] = ReadAbstract(r);
            else if (r.LocalName == "num" && r.IntAttr("numId") is { } nid) numbering._nums[nid] = ReadNum(r);
            else r.Skip();
        }
        return numbering;
    }

    private static AbstractNumbering ReadAbstract(XmlReader r)
    {
        var levels = new List<(int, NumberingLevel)>();
        string? styleLink = null;
        int depth = r.Depth;
        if (r.Enter())
        {
            while (r.NextChild(depth))
            {
                if (r.IsW() && r.LocalName == "lvl" && r.IntAttr("ilvl") is { } ilvl and >= 0 and <= 8)
                    levels.Add((ilvl, ReadLevel(r)));
                else if (r.IsW() && r.LocalName == "numStyleLink") { styleLink = r.WAttr("val"); r.Skip(); }
                else r.Skip();
            }
        }
        var result = new AbstractNumbering { NumStyleLink = styleLink };
        foreach (var (i, level) in levels) result.Levels[i] = level;
        return result;
    }

    private static NumberingInstance ReadNum(XmlReader r)
    {
        int abstractId = -1;
        var overrides = new List<(int Level, int? Start, NumberingLevel? Definition)>();
        int depth = r.Depth;
        if (r.Enter())
        {
            while (r.NextChild(depth))
            {
                if (!r.IsW()) { r.Skip(); continue; }
                if (r.LocalName == "abstractNumId") { abstractId = r.IntAttr("val") ?? -1; r.Skip(); }
                else if (r.LocalName == "lvlOverride" && r.IntAttr("ilvl") is { } ilvl and >= 0 and <= 8)
                {
                    int? start = null;
                    NumberingLevel? definition = null;
                    int inner = r.Depth;
                    if (r.Enter())
                    {
                        while (r.NextChild(inner))
                        {
                            if (r.IsW() && r.LocalName == "startOverride") { start = r.IntAttr("val"); r.Skip(); }
                            else if (r.IsW() && r.LocalName == "lvl") definition = ReadLevel(r);
                            else r.Skip();
                        }
                    }
                    overrides.Add((ilvl, start, definition));
                }
                else r.Skip();
            }
        }
        var num = new NumberingInstance { AbstractId = abstractId };
        foreach (var (level, start, definition) in overrides)
        {
            if (start is { } s) num.StartOverrides[level] = s;
            if (definition is not null) num.LevelOverrides[level] = definition;
        }
        return num;
    }

    private static NumberingLevel ReadLevel(XmlReader r)
    {
        var level = new NumberingLevel();
        int depth = r.Depth;
        if (!r.Enter()) return level;
        while (r.NextChild(depth))
        {
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "pPr": level = level with { Para = ParaFormat.Parse(r) }; continue;
                case "rPr": level = level with { Run = RunFormat.Parse(r) }; continue;
                case "start": level = level with { Start = Math.Clamp(r.IntAttr("val") ?? 1, 0, 1_000_000) }; break;
                case "numFmt": level = level with { Format = r.WAttr("val") ?? "decimal" }; break;
                case "lvlText": level = level with { Text = r.WAttr("val") ?? "" }; break;
                case "lvlJc": level = level with { Alignment = ParaFormat.ParseAlignment(r.WAttr("val")) }; break;
                case "lvlRestart": level = level with { Restart = r.IntAttr("val") }; break;
                case "isLgl": level = level with { IsLegal = r.OnOff() }; break;
                case "pStyle": level = level with { ParagraphStyle = r.WAttr("val") }; break;
                case "suff":
                    level = level with
                    {
                        Suffix = r.WAttr("val") switch
                        {
                            "space" => ListLabelSuffix.Space,
                            "nothing" => ListLabelSuffix.Nothing,
                            _ => ListLabelSuffix.Tab,
                        },
                    };
                    break;
            }
            r.Skip();
        }
        return level;
    }

    /// <summary>Returns the level definition without touching counters.</summary>
    public NumberingLevel? GetLevel(int numId, int ilvl, StyleSheet styles)
    {
        if (!_nums.TryGetValue(numId, out var num)) return null;
        if (num.LevelOverrides.TryGetValue(ilvl, out var overridden)) return overridden;
        return GetAbstract(num, styles)?.Levels[ilvl];
    }

    /// <summary>Finds the level whose w:pStyle links it to <paramref name="styleId"/>.</summary>
    public int? FindLevelForStyle(int numId, string? styleId, StyleSheet styles)
    {
        if (styleId is null || !_nums.TryGetValue(numId, out var num) || GetAbstract(num, styles) is not { } abs) return null;
        for (int i = 0; i < 9; i++)
            if (abs.Levels[i]?.ParagraphStyle == styleId) return i;
        return null;
    }

    /// <summary>Advances the counters for a list paragraph and formats its label.</summary>
    public NumberingResult? Next(int numId, int ilvl, StyleSheet styles)
    {
        if (!_nums.TryGetValue(numId, out var num) || GetLevel(numId, ilvl, styles) is not { } level) return null;
        var key = num.StartOverrides.Count > 0 ? $"n{numId}" : $"a{ResolveAbstractId(num, styles)}";
        if (!_counters.TryGetValue(key, out var counters)) _counters[key] = counters = new int?[9];

        counters[ilvl] = counters[ilvl] is { } c ? c + 1 : StartOf(num, numId, ilvl, styles);
        for (int deeper = ilvl + 1; deeper < 9; deeper++)
        {
            int restartAfter = GetLevel(numId, deeper, styles)?.Restart ?? deeper;
            if (restartAfter != 0 && ilvl < restartAfter) counters[deeper] = null;
        }

        if (level.Format == "bullet") return new NumberingResult(level.Text ?? "•", level);
        var text = new StringBuilder();
        var template = level.Text ?? $"%{ilvl + 1}.";
        for (int i = 0; i < template.Length; i++)
        {
            if (template[i] == '%' && i + 1 < template.Length && template[i + 1] is >= '1' and <= '9')
            {
                int referenced = template[++i] - '1';
                var refLevel = GetLevel(numId, referenced, styles);
                int value = counters[referenced] ?? StartOf(num, numId, referenced, styles);
                var format = level.IsLegal && referenced < ilvl ? "decimal" : refLevel?.Format ?? "decimal";
                text.Append(FormatNumber(value, format));
            }
            else text.Append(template[i]);
        }
        return new NumberingResult(text.ToString(), level);
    }

    private int StartOf(NumberingInstance num, int numId, int ilvl, StyleSheet styles) =>
        num.StartOverrides.TryGetValue(ilvl, out var s) ? s : GetLevel(numId, ilvl, styles)?.Start ?? 1;

    private AbstractNumbering? GetAbstract(NumberingInstance num, StyleSheet styles)
    {
        int id = ResolveAbstractId(num, styles);
        return _abstracts.GetValueOrDefault(id);
    }

    /// <summary>Follows numStyleLink → numbering style → its num → abstract definition (bounded).</summary>
    private int ResolveAbstractId(NumberingInstance num, StyleSheet styles)
    {
        int id = num.AbstractId;
        for (int hops = 0; hops < 8; hops++)
        {
            if (!_abstracts.TryGetValue(id, out var abs) || abs.NumStyleLink is null) return id;
            var linkedNum = styles.Resolve(abs.NumStyleLink).Para.NumId;
            if (linkedNum is not { } n || !_nums.TryGetValue(n, out var target) || target.AbstractId == id) return id;
            id = target.AbstractId;
        }
        return id;
    }

    public static string FormatNumber(int value, string format) => format switch
    {
        "decimalZero" => value < 10 && value >= 0 ? "0" + value : value.ToString(System.Globalization.CultureInfo.InvariantCulture),
        "upperRoman" => Roman(value),
        "lowerRoman" => Roman(value).ToLowerInvariant(),
        "upperLetter" => Letters(value),
        "lowerLetter" => Letters(value).ToLowerInvariant(),
        "ordinal" => value + OrdinalSuffix(value),
        "none" => "",
        _ => value.ToString(System.Globalization.CultureInfo.InvariantCulture),
    };

    private static string Roman(int value)
    {
        if (value <= 0 || value >= 4000) return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        ReadOnlySpan<int> numbers = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
        string[] symbols = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
        var sb = new StringBuilder();
        for (int i = 0; i < numbers.Length; i++)
            while (value >= numbers[i]) { sb.Append(symbols[i]); value -= numbers[i]; }
        return sb.ToString();
    }

    /// <summary>Word letter numbering: A..Z, then AA..ZZ, AAA… (the letter repeats).</summary>
    private static string Letters(int value)
    {
        if (value <= 0) return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        int repeat = (value - 1) / 26 + 1;
        return new string((char)('A' + (value - 1) % 26), Math.Min(repeat, 50));
    }

    private static string OrdinalSuffix(int value) => (value % 100) switch
    {
        11 or 12 or 13 => "th",
        _ => (value % 10) switch { 1 => "st", 2 => "nd", 3 => "rd", _ => "th" },
    };
}