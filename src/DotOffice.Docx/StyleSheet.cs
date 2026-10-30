using System.Xml;
using DotOffice.Core;

namespace DotOffice.Docx;

internal enum StyleType { Paragraph, Character, Table, Numbering }

/// <summary>Formatting attached to one table-style region (w:tblStylePr), e.g. firstRow.</summary>
internal sealed record ConditionalFormat(ParaFormat? Para, RunFormat? Run, CellFormat? Cell, TableFormat? Table)
{
    public ConditionalFormat Merge(ConditionalFormat over) => new(
        Para?.Merge(over.Para) ?? over.Para, Run?.Merge(over.Run) ?? over.Run,
        Cell?.Merge(over.Cell) ?? over.Cell, Table?.Merge(over.Table) ?? over.Table);
}

internal sealed class Style
{
    public required string Id { get; init; }
    public StyleType Type { get; init; }
    public string? BasedOn { get; init; }
    public bool IsDefault { get; init; }
    public ParaFormat? Para { get; init; }
    public RunFormat? Run { get; init; }
    public TableFormat? Table { get; init; }
    public CellFormat? Cell { get; init; }
    public Dictionary<string, ConditionalFormat> Conditionals { get; } = new(StringComparer.Ordinal);
}

/// <summary>A style with its basedOn chain merged (base first, derived last).</summary>
internal sealed record ResolvedStyle(ParaFormat Para, RunFormat Run, TableFormat Table, CellFormat Cell,
    IReadOnlyDictionary<string, ConditionalFormat> Conditionals)
{
    public static ResolvedStyle Empty { get; } = new(ParaFormat.Empty, RunFormat.Empty, new TableFormat(), new CellFormat(),
        new Dictionary<string, ConditionalFormat>());
}

/// <summary>
/// styles.xml: docDefaults plus named styles, and the OOXML resolution rules on top of them.
/// Resolution order is docDefaults → table style → numbering → paragraph style → run style → direct
/// formatting; toggle properties (bold, italic, caps…) flip across style levels and are absolute in
/// direct formatting (ECMA-376 §17.7.3).
/// </summary>
internal sealed class StyleSheet
{
    private const int MaxChainLength = 32;

    private readonly Dictionary<string, Style> _styles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, ResolvedStyle> _resolved = new(StringComparer.Ordinal);

    public RunFormat DefaultRun { get; private set; } = RunFormat.Empty;
    public ParaFormat DefaultPara { get; private set; } = ParaFormat.Empty;
    public Theme? Theme { get; set; }

    public string? DefaultParagraphStyle { get; private set; }
    public string? DefaultCharacterStyle { get; private set; }
    public string? DefaultTableStyle { get; private set; }

    public Style? Get(string? id) => id is not null && _styles.TryGetValue(id, out var s) ? s : null;

    public static StyleSheet Read(XmlReader r)
    {
        var sheet = new StyleSheet();
        r.MoveToContent();
        int depth = r.Depth;
        if (!r.Enter()) return sheet;
        while (r.NextChild(depth))
        {
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "docDefaults": sheet.ReadDocDefaults(r); break;
                case "style":
                    if (ReadStyle(r) is { } style)
                    {
                        sheet._styles.TryAdd(style.Id, style);
                        if (style.IsDefault)
                        {
                            switch (style.Type)
                            {
                                case StyleType.Paragraph: sheet.DefaultParagraphStyle ??= style.Id; break;
                                case StyleType.Character: sheet.DefaultCharacterStyle ??= style.Id; break;
                                case StyleType.Table: sheet.DefaultTableStyle ??= style.Id; break;
                            }
                        }
                    }
                    break;
                default: r.Skip(); break;
            }
        }
        return sheet;
    }

    private void ReadDocDefaults(XmlReader r)
    {
        int depth = r.Depth;
        if (!r.Enter()) return;
        while (r.NextChild(depth))
        {
            if (r.IsW() && r.LocalName is "rPrDefault" or "pPrDefault")
            {
                bool run = r.LocalName == "rPrDefault";
                int inner = r.Depth;
                if (!r.Enter()) continue;
                while (r.NextChild(inner))
                {
                    if (run && r.IsW() && r.LocalName == "rPr") DefaultRun = RunFormat.Parse(r);
                    else if (!run && r.IsW() && r.LocalName == "pPr") DefaultPara = ParaFormat.Parse(r);
                    else r.Skip();
                }
            }
            else r.Skip();
        }
    }

    private static Style? ReadStyle(XmlReader r)
    {
        var id = r.WAttr("styleId");
        var type = r.WAttr("type") switch
        {
            "character" => StyleType.Character,
            "table" => StyleType.Table,
            "numbering" => StyleType.Numbering,
            _ => StyleType.Paragraph,
        };
        bool isDefault = r.OnOff("default") && r.WAttr("default") is not null;
        string? basedOn = null;
        ParaFormat? para = null;
        RunFormat? run = null;
        TableFormat? table = null;
        CellFormat? cell = null;
        var conditionals = new Dictionary<string, ConditionalFormat>(StringComparer.Ordinal);
        int depth = r.Depth;
        if (r.Enter())
        {
            while (r.NextChild(depth))
            {
                if (!r.IsW()) { r.Skip(); continue; }
                switch (r.LocalName)
                {
                    case "basedOn": basedOn = r.WAttr("val"); r.Skip(); break;
                    case "pPr": para = ParaFormat.Parse(r); break;
                    case "rPr": run = RunFormat.Parse(r); break;
                    case "tblPr": table = TableFormat.Parse(r); break;
                    case "tcPr": cell = CellFormat.Parse(r); break;
                    case "tblStylePr":
                        var kind = r.WAttr("type");
                        var conditional = ReadConditional(r);
                        if (kind is not null) conditionals[kind] = conditional;
                        break;
                    default: r.Skip(); break;
                }
            }
        }
        if (string.IsNullOrEmpty(id)) return null;
        var style = new Style
        {
            Id = id, Type = type, BasedOn = basedOn, IsDefault = isDefault,
            Para = para, Run = run, Table = table, Cell = cell,
        };
        foreach (var (k, v) in conditionals) style.Conditionals[k] = v;
        return style;
    }

    private static ConditionalFormat ReadConditional(XmlReader r)
    {
        ParaFormat? para = null;
        RunFormat? run = null;
        CellFormat? cell = null;
        TableFormat? table = null;
        int depth = r.Depth;
        if (!r.Enter()) return new ConditionalFormat(null, null, null, null);
        while (r.NextChild(depth))
        {
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "pPr": para = ParaFormat.Parse(r); break;
                case "rPr": run = RunFormat.Parse(r); break;
                case "tcPr": cell = CellFormat.Parse(r); break;
                case "tblPr": table = TableFormat.Parse(r); break;
                default: r.Skip(); break;
            }
        }
        return new ConditionalFormat(para, run, cell, table);
    }

    /// <summary>Merges a style with its basedOn ancestors (cycle- and depth-safe).</summary>
    public ResolvedStyle Resolve(string? id)
    {
        if (id is null || !_styles.ContainsKey(id)) return ResolvedStyle.Empty;
        if (_resolved.TryGetValue(id, out var cached)) return cached;

        var chain = new List<Style>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (var s = Get(id); s is not null && chain.Count < MaxChainLength && seen.Add(s.Id); s = Get(s.BasedOn))
            chain.Add(s);
        chain.Reverse();

        var para = ParaFormat.Empty;
        var run = RunFormat.Empty;
        var table = new TableFormat();
        var cell = new CellFormat();
        var conditionals = new Dictionary<string, ConditionalFormat>(StringComparer.Ordinal);
        foreach (var s in chain)
        {
            para = para.Merge(s.Para);
            run = run.Merge(s.Run);
            table = table.Merge(s.Table);
            cell = cell.Merge(s.Cell);
            foreach (var (kind, c) in s.Conditionals)
                conditionals[kind] = conditionals.TryGetValue(kind, out var existing) ? existing.Merge(c) : c;
        }
        // Style ids are not inherited; StyleId on the merged formats refers to linked styles only.
        var resolved = new ResolvedStyle(para with { StyleId = null }, run with { StyleId = null }, table, cell, conditionals);
        _resolved[id] = resolved;
        return resolved;
    }

    /// <summary>Resolves run formatting for a run inside a paragraph.</summary>
    /// <param name="tableStyle">Merged table style run formatting for the containing cell, if any.</param>
    /// <param name="paragraphStyleId">The paragraph's style (null = default paragraph style).</param>
    /// <param name="direct">Direct run formatting (its <c>StyleId</c> selects the character style).</param>
    public RunProperties ResolveRun(RunFormat? tableStyle, string? paragraphStyleId, RunFormat? direct)
    {
        var paraStyle = Resolve(paragraphStyleId ?? DefaultParagraphStyle).Run;
        var charStyleId = direct?.StyleId ?? DefaultCharacterStyle;
        var charStyle = Get(charStyleId)?.Type == StyleType.Character ? Resolve(charStyleId).Run : RunFormat.Empty;

        var merged = DefaultRun.Merge(tableStyle).Merge(paraStyle).Merge(charStyle).Merge(direct);

        bool Toggle(Func<RunFormat, bool?> get)
        {
            bool value = get(DefaultRun) ?? false;
            foreach (var level in (ReadOnlySpan<RunFormat?>)[tableStyle, paraStyle, charStyle])
                if (level is not null && get(level) == true) value = !value;
            return direct is not null && get(direct) is { } d ? d : value;
        }

        var ascii = FontName(merged.Ascii) ?? "Times New Roman";
        return new RunProperties
        {
            FontFamily = ascii,
            HighAnsiFontFamily = FontName(merged.HAnsi) is { } h && h != ascii ? h : null,
            EastAsianFontFamily = FontName(merged.EastAsia),
            FontSize = merged.Size ?? 10,
            Bold = Toggle(static f => f.Bold),
            Italic = Toggle(static f => f.Italic),
            Caps = Toggle(static f => f.Caps),
            SmallCaps = Toggle(static f => f.SmallCaps),
            Strike = Toggle(static f => f.Strike),
            DoubleStrike = Toggle(static f => f.DoubleStrike),
            Hidden = Toggle(static f => f.Vanish),
            Underline = merged.Underline ?? UnderlineStyle.None,
            Color = merged.Color?.Value,
            Highlight = merged.Highlight?.Value,
            Shading = merged.Shading?.Value,
            VerticalPosition = merged.VerticalAlign ?? VerticalPosition.Baseline,
            CharacterSpacing = Math.Clamp(merged.Spacing ?? 0, -100, 100),
            BaselineShift = Math.Clamp(merged.Position ?? 0, -500, 500),
        };
    }

    private string? FontName(FontSpec? spec)
    {
        if (spec is not { } s) return null;
        if (s.Theme is not null && Theme?.Resolve(s.Theme) is { } themed) return themed;
        return string.IsNullOrWhiteSpace(s.Name) ? null : s.Name;
    }

    /// <summary>
    /// Merges paragraph formatting in hierarchy order. Numbering indentation sits below the paragraph
    /// style when the numbering came from the style, and above it when applied directly (Word behavior).
    /// </summary>
    public ParaFormat ResolveParagraphFormat(ParaFormat? tableStyle, string? styleId, ParaFormat? numbering,
        bool numberingIsDirect, ParaFormat direct)
    {
        var f = DefaultPara.Merge(tableStyle);
        if (!numberingIsDirect) f = f.Merge(numbering);
        f = f.Merge(Resolve(styleId).Para);
        if (numberingIsDirect) f = f.Merge(numbering);
        return f.Merge(direct with { StyleId = styleId });
    }

    public static ParagraphProperties ToProperties(ParaFormat f)
    {
        var tabs = new List<TabStop>();
        foreach (var t in f.Tabs ?? [])
        {
            tabs.RemoveAll(x => Math.Abs(x.Position - t.Position) < 0.5);
            if (!t.Clear) tabs.Add(new TabStop(t.Position, t.Alignment, t.Leader));
        }
        tabs.Sort((a, b) => a.Position.CompareTo(b.Position));

        return new ParagraphProperties
        {
            StyleId = f.StyleId,
            Alignment = f.Alignment ?? ParagraphAlignment.Left,
            LeftIndent = f.Left ?? 0,
            RightIndent = f.Right ?? 0,
            FirstLineIndent = f.FirstLine ?? 0,
            SpaceBefore = f.BeforeAuto == true ? 14 : Math.Max(0, f.Before ?? 0),
            SpaceAfter = f.AfterAuto == true ? 14 : Math.Max(0, f.After ?? 0),
            LineSpacing = f.Line ?? LineSpacing.Single,
            ContextualSpacing = f.ContextualSpacing ?? false,
            KeepWithNext = f.KeepNext ?? false,
            KeepLinesTogether = f.KeepLines ?? false,
            PageBreakBefore = f.PageBreakBefore ?? false,
            WidowControl = f.WidowControl ?? true,
            TabStops = tabs,
            Borders = f.Top is null && f.Bottom is null && f.LeftBorder is null && f.RightBorder is null && f.Between is null
                ? BorderSet.Empty
                : new BorderSet
                {
                    Top = f.Top ?? Border.None,
                    Bottom = f.Bottom ?? Border.None,
                    Left = f.LeftBorder ?? Border.None,
                    Right = f.RightBorder ?? Border.None,
                    InsideHorizontal = f.Between ?? Border.None,
                },
            Shading = f.Shading?.Value,
        };
    }
}