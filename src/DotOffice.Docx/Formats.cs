using System.Xml;
using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>An explicit color setting; a null <see cref="Value"/> means "auto"/"none".</summary>
internal readonly record struct ColorSpec(Color? Value);

/// <summary>One rFonts slot: explicit name and/or theme reference (theme wins).</summary>
internal readonly record struct FontSpec(string? Name, string? Theme);

internal readonly record struct TabSpec(double Position, TabAlignment Alignment, TabLeader Leader, bool Clear);

/// <summary>Run properties as written at one level of the style hierarchy (unspecified = null).</summary>
internal sealed record RunFormat
{
    public static RunFormat Empty { get; } = new();

    public FontSpec? Ascii { get; init; }
    public FontSpec? HAnsi { get; init; }
    public FontSpec? EastAsia { get; init; }
    public FontSpec? ComplexScript { get; init; }
    public double? Size { get; init; }
    public double? SizeCs { get; init; }
    public bool? BoldCs { get; init; }
    public bool? ItalicCs { get; init; }
    public bool? RightToLeft { get; init; }
    public bool? UseComplexScript { get; init; }
    public double? Kerning { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? Caps { get; init; }
    public bool? SmallCaps { get; init; }
    public bool? Strike { get; init; }
    public bool? DoubleStrike { get; init; }
    public bool? Vanish { get; init; }
    public UnderlineStyle? Underline { get; init; }
    public ColorSpec? Color { get; init; }
    public ColorSpec? Highlight { get; init; }
    public ColorSpec? Shading { get; init; }
    public VerticalPosition? VerticalAlign { get; init; }
    public double? Spacing { get; init; }
    public double? Position { get; init; }
    public string? StyleId { get; init; }

    /// <summary>Plain override merge: values set in <paramref name="over"/> replace these.</summary>
    public RunFormat Merge(RunFormat? over) => over is null ? this : new RunFormat
    {
        Ascii = over.Ascii ?? Ascii,
        HAnsi = over.HAnsi ?? HAnsi,
        EastAsia = over.EastAsia ?? EastAsia,
        ComplexScript = over.ComplexScript ?? ComplexScript,
        Size = over.Size ?? Size,
        SizeCs = over.SizeCs ?? SizeCs,
        BoldCs = over.BoldCs ?? BoldCs,
        ItalicCs = over.ItalicCs ?? ItalicCs,
        RightToLeft = over.RightToLeft ?? RightToLeft,
        UseComplexScript = over.UseComplexScript ?? UseComplexScript,
        Kerning = over.Kerning ?? Kerning,
        Bold = over.Bold ?? Bold,
        Italic = over.Italic ?? Italic,
        Caps = over.Caps ?? Caps,
        SmallCaps = over.SmallCaps ?? SmallCaps,
        Strike = over.Strike ?? Strike,
        DoubleStrike = over.DoubleStrike ?? DoubleStrike,
        Vanish = over.Vanish ?? Vanish,
        Underline = over.Underline ?? Underline,
        Color = over.Color ?? Color,
        Highlight = over.Highlight ?? Highlight,
        Shading = over.Shading ?? Shading,
        VerticalAlign = over.VerticalAlign ?? VerticalAlign,
        Spacing = over.Spacing ?? Spacing,
        Position = over.Position ?? Position,
        StyleId = over.StyleId ?? StyleId,
    };

    public static RunFormat Parse(XmlReader r)
    {
        int depth = r.Depth;
        var f = new RunFormat();
        if (!r.Enter()) return f;
        while (r.NextChild(depth))
        {
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "rStyle": f = f with { StyleId = r.WAttr("val") }; break;
                case "rFonts":
                    f = f with
                    {
                        Ascii = Font(r, "ascii", "asciiTheme") ?? f.Ascii,
                        HAnsi = Font(r, "hAnsi", "hAnsiTheme") ?? f.HAnsi,
                        EastAsia = Font(r, "eastAsia", "eastAsiaTheme") ?? f.EastAsia,
                        ComplexScript = Font(r, "cs", "cstheme") ?? f.ComplexScript,
                    };
                    break;
                case "sz":
                    if (r.DoubleAttr("val") is { } sz && sz > 0) f = f with { Size = Math.Min(Units.FromHalfPoints(sz), 1638) };
                    break;
                case "szCs":
                    if (r.DoubleAttr("val") is { } szCs && szCs > 0) f = f with { SizeCs = Math.Min(Units.FromHalfPoints(szCs), 1638) };
                    break;
                case "kern":
                    if (r.DoubleAttr("val") is { } kern) f = f with { Kerning = Units.FromHalfPoints(Math.Max(0, kern)) };
                    break;
                case "b": f = f with { Bold = r.OnOff() }; break;
                case "i": f = f with { Italic = r.OnOff() }; break;
                case "bCs": f = f with { BoldCs = r.OnOff() }; break;
                case "iCs": f = f with { ItalicCs = r.OnOff() }; break;
                case "rtl": f = f with { RightToLeft = r.OnOff() }; break;
                case "cs": f = f with { UseComplexScript = r.OnOff() }; break;
                case "caps": f = f with { Caps = r.OnOff() }; break;
                case "smallCaps": f = f with { SmallCaps = r.OnOff() }; break;
                case "strike": f = f with { Strike = r.OnOff() }; break;
                case "dstrike": f = f with { DoubleStrike = r.OnOff() }; break;
                case "vanish": f = f with { Vanish = r.OnOff() }; break;
                case "u": f = f with { Underline = ParseUnderline(r.WAttr("val")) }; break;
                case "color":
                    f = f with { Color = r.WAttr("val") is "auto" or null ? new ColorSpec(null) : new ColorSpec(Core.Color.TryParseHex(r.WAttr("val"))) };
                    break;
                case "highlight": f = f with { Highlight = new ColorSpec(HighlightColor(r.WAttr("val"))) }; break;
                case "shd": f = f with { Shading = ParseShading(r) }; break;
                case "vertAlign":
                    f = f with
                    {
                        VerticalAlign = r.WAttr("val") switch
                        {
                            "superscript" => VerticalPosition.Superscript,
                            "subscript" => VerticalPosition.Subscript,
                            _ => VerticalPosition.Baseline,
                        },
                    };
                    break;
                case "spacing": f = f with { Spacing = r.TwipsAttr("val") }; break;
                case "position": f = f with { Position = r.DoubleAttr("val") is { } p ? Units.FromHalfPoints(p) : null }; break;
            }
            r.Skip();
        }
        return f;
    }

    private static FontSpec? Font(XmlReader r, string name, string theme)
    {
        var n = r.WAttr(name);
        var t = r.WAttr(theme);
        return n is null && t is null ? null : new FontSpec(n, t);
    }

    private static UnderlineStyle ParseUnderline(string? value) => value switch
    {
        null or "single" => UnderlineStyle.Single,
        "none" => UnderlineStyle.None,
        "words" => UnderlineStyle.Words,
        "double" => UnderlineStyle.Double,
        "thick" => UnderlineStyle.Thick,
        "dotted" or "dottedHeavy" => UnderlineStyle.Dotted,
        "dash" or "dashedHeavy" or "dashLong" or "dashLongHeavy" or "dotDash" or "dashDotHeavy" or "dotDotDash"
            or "dashDotDotHeavy" => UnderlineStyle.Dashed,
        "wave" or "wavyHeavy" or "wavyDouble" => UnderlineStyle.Wavy,
        _ => UnderlineStyle.Single,
    };

    internal static Color? HighlightColor(string? name) => name switch
    {
        "black" => new Color(0, 0, 0),
        "blue" => new Color(0, 0, 255),
        "cyan" => new Color(0, 255, 255),
        "green" => new Color(0, 255, 0),
        "magenta" => new Color(255, 0, 255),
        "red" => new Color(255, 0, 0),
        "yellow" => new Color(255, 255, 0),
        "white" => new Color(255, 255, 255),
        "darkBlue" => new Color(0, 0, 139),
        "darkCyan" => new Color(0, 139, 139),
        "darkGreen" => new Color(0, 100, 0),
        "darkMagenta" => new Color(139, 0, 139),
        "darkRed" => new Color(139, 0, 0),
        "darkYellow" => new Color(128, 128, 0),
        "darkGray" => new Color(169, 169, 169),
        "lightGray" => new Color(211, 211, 211),
        _ => null,
    };

    /// <summary>Resolves w:shd (pattern of <c>color</c> over <c>fill</c>) to a single background color.</summary>
    internal static ColorSpec ParseShading(XmlReader r)
    {
        var pattern = r.WAttr("val");
        var fill = r.WAttr("fill") is "auto" or null ? (Color?)null : Core.Color.TryParseHex(r.WAttr("fill"));
        var color = r.WAttr("color") is "auto" or null ? (Color?)null : Core.Color.TryParseHex(r.WAttr("color"));
        if (pattern is "nil" or null or "clear") return new ColorSpec(fill);
        if (pattern == "solid") return new ColorSpec(color ?? Core.Color.Black);
        if (pattern.StartsWith("pct", StringComparison.Ordinal) && int.TryParse(pattern.AsSpan(3), out var pct))
            return new ColorSpec((fill ?? Core.Color.White).Blend(color ?? Core.Color.Black, pct / 100.0));
        return new ColorSpec(fill);
    }
}

internal sealed record ParaFormat
{
    public static ParaFormat Empty { get; } = new();

    public string? StyleId { get; init; }
    public ParagraphAlignment? Alignment { get; init; }
    public double? Left { get; init; }
    public double? Right { get; init; }
    public double? FirstLine { get; init; }
    public double? Before { get; init; }
    public double? After { get; init; }
    public bool? BeforeAuto { get; init; }
    public bool? AfterAuto { get; init; }
    public LineSpacing? Line { get; init; }
    public bool? KeepNext { get; init; }
    public bool? KeepLines { get; init; }
    public bool? PageBreakBefore { get; init; }
    public bool? WidowControl { get; init; }
    public bool? ContextualSpacing { get; init; }
    public IReadOnlyList<TabSpec>? Tabs { get; init; }
    public Border? Top { get; init; }
    public Border? Bottom { get; init; }
    public Border? LeftBorder { get; init; }
    public Border? RightBorder { get; init; }
    public Border? Between { get; init; }
    public ColorSpec? Shading { get; init; }
    public int? NumId { get; init; }
    public int? Ilvl { get; init; }
    public bool? Bidi { get; init; }

    public ParaFormat Merge(ParaFormat? over) => over is null ? this : new ParaFormat
    {
        StyleId = over.StyleId ?? StyleId,
        Alignment = over.Alignment ?? Alignment,
        Left = over.Left ?? Left,
        Right = over.Right ?? Right,
        FirstLine = over.FirstLine ?? FirstLine,
        Before = over.Before ?? Before,
        After = over.After ?? After,
        BeforeAuto = over.BeforeAuto ?? BeforeAuto,
        AfterAuto = over.AfterAuto ?? AfterAuto,
        Line = over.Line ?? Line,
        KeepNext = over.KeepNext ?? KeepNext,
        KeepLines = over.KeepLines ?? KeepLines,
        PageBreakBefore = over.PageBreakBefore ?? PageBreakBefore,
        WidowControl = over.WidowControl ?? WidowControl,
        ContextualSpacing = over.ContextualSpacing ?? ContextualSpacing,
        Tabs = Tabs is null ? over.Tabs : over.Tabs is null ? Tabs : [.. Tabs, .. over.Tabs],
        Top = over.Top ?? Top,
        Bottom = over.Bottom ?? Bottom,
        LeftBorder = over.LeftBorder ?? LeftBorder,
        RightBorder = over.RightBorder ?? RightBorder,
        Between = over.Between ?? Between,
        Shading = over.Shading ?? Shading,
        NumId = over.NumId ?? NumId,
        Ilvl = over.Ilvl ?? Ilvl,
        Bidi = over.Bidi ?? Bidi,
    };

    /// <summary>Parses w:pPr. The paragraph mark's w:rPr and w:sectPr are handed to the callbacks.</summary>
    public static ParaFormat Parse(XmlReader r, Action<RunFormat>? onMarkRun = null, Action<XmlReader>? onSectPr = null)
    {
        int depth = r.Depth;
        var f = new ParaFormat();
        if (!r.Enter()) return f;
        while (r.NextChild(depth))
        {
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "rPr":
                    var mark = RunFormat.Parse(r);
                    onMarkRun?.Invoke(mark);
                    continue;
                case "sectPr":
                    if (onSectPr is null) r.Skip(); else onSectPr(r);
                    continue;
                case "tabs":
                    f = f with { Tabs = ParseTabs(r) };
                    continue;
                case "pBdr":
                    f = ParseBorders(r, f);
                    continue;
                case "numPr":
                    f = ParseNumPr(r, f);
                    continue;
                case "pStyle": f = f with { StyleId = r.WAttr("val") }; break;
                case "jc": f = f with { Alignment = ParseAlignment(r.WAttr("val")) }; break;
                case "ind":
                    var hanging = r.TwipsAttr("hanging");
                    var firstLine = r.TwipsAttr("firstLine");
                    f = f with
                    {
                        Left = r.TwipsAttr("left") ?? r.TwipsAttr("start") ?? f.Left,
                        Right = r.TwipsAttr("right") ?? r.TwipsAttr("end") ?? f.Right,
                        FirstLine = hanging is { } h ? -h : firstLine ?? f.FirstLine,
                    };
                    break;
                case "spacing":
                    var line = r.DoubleAttr("line");
                    LineSpacing? spacing = line is null ? null : r.WAttr("lineRule") switch
                    {
                        "exact" => new LineSpacing(Math.Abs(Units.FromTwips(line.Value)), LineSpacingRule.Exact),
                        "atLeast" => new LineSpacing(Units.FromTwips(line.Value), LineSpacingRule.AtLeast),
                        _ => new LineSpacing(Math.Clamp(line.Value / 240.0, 0.06, 132), LineSpacingRule.Auto),
                    };
                    f = f with
                    {
                        Before = r.TwipsAttr("before") ?? f.Before,
                        After = r.TwipsAttr("after") ?? f.After,
                        BeforeAuto = r.WAttr("beforeAutospacing") is null ? f.BeforeAuto : r.OnOff("beforeAutospacing"),
                        AfterAuto = r.WAttr("afterAutospacing") is null ? f.AfterAuto : r.OnOff("afterAutospacing"),
                        Line = spacing ?? f.Line,
                    };
                    break;
                case "keepNext": f = f with { KeepNext = r.OnOff() }; break;
                case "bidi": f = f with { Bidi = r.OnOff() }; break;
                case "keepLines": f = f with { KeepLines = r.OnOff() }; break;
                case "pageBreakBefore": f = f with { PageBreakBefore = r.OnOff() }; break;
                case "widowControl": f = f with { WidowControl = r.OnOff() }; break;
                case "contextualSpacing": f = f with { ContextualSpacing = r.OnOff() }; break;
                case "shd": f = f with { Shading = RunFormat.ParseShading(r) }; break;
            }
            r.Skip();
        }
        return f;
    }

    internal static ParagraphAlignment ParseAlignment(string? value) => value switch
    {
        "center" => ParagraphAlignment.Center,
        "right" or "end" => ParagraphAlignment.Right,
        "both" or "thaiDistribute" or "lowKashida" or "mediumKashida" or "highKashida" => ParagraphAlignment.Justify,
        "distribute" => ParagraphAlignment.Distribute,
        _ => ParagraphAlignment.Left,
    };

    private static List<TabSpec> ParseTabs(XmlReader r)
    {
        var tabs = new List<TabSpec>();
        int depth = r.Depth;
        if (!r.Enter()) return tabs;
        while (r.NextChild(depth))
        {
            if (r.IsW() && r.LocalName == "tab" && r.TwipsAttr("pos") is { } pos)
            {
                var val = r.WAttr("val");
                var alignment = val switch
                {
                    "center" => TabAlignment.Center,
                    "right" or "end" => TabAlignment.Right,
                    "decimal" => TabAlignment.Decimal,
                    "bar" => TabAlignment.Bar,
                    _ => TabAlignment.Left,
                };
                var leader = r.WAttr("leader") switch
                {
                    "dot" => TabLeader.Dot,
                    "hyphen" => TabLeader.Hyphen,
                    "underscore" or "heavy" => TabLeader.Underscore,
                    "middleDot" => TabLeader.MiddleDot,
                    _ => TabLeader.None,
                };
                tabs.Add(new TabSpec(pos, alignment, leader, val == "clear"));
            }
            r.Skip();
        }
        return tabs;
    }

    private static ParaFormat ParseBorders(XmlReader r, ParaFormat f)
    {
        int depth = r.Depth;
        if (!r.Enter()) return f;
        while (r.NextChild(depth))
        {
            if (r.IsW())
            {
                var border = Borders.Parse(r);
                f = r.LocalName switch
                {
                    "top" => f with { Top = border },
                    "bottom" => f with { Bottom = border },
                    "left" or "start" => f with { LeftBorder = border },
                    "right" or "end" => f with { RightBorder = border },
                    "between" => f with { Between = border },
                    _ => f,
                };
            }
            r.Skip();
        }
        return f;
    }

    private static ParaFormat ParseNumPr(XmlReader r, ParaFormat f)
    {
        int depth = r.Depth;
        if (!r.Enter()) return f;
        while (r.NextChild(depth))
        {
            if (r.IsW() && r.LocalName == "numId") f = f with { NumId = r.IntAttr("val") };
            else if (r.IsW() && r.LocalName == "ilvl") f = f with { Ilvl = Math.Clamp(r.IntAttr("val") ?? 0, 0, 8) };
            r.Skip();
        }
        return f;
    }
}

internal static class Borders
{
    /// <summary>Parses a border element (w:top, w:insideH…) at the reader's position without moving.</summary>
    public static Border Parse(XmlReader r)
    {
        var val = r.WAttr("val");
        var style = val switch
        {
            null or "nil" or "none" => BorderStyle.None,
            "single" or "wave" or "threeDEmboss" or "threeDEngrave" or "outset" or "inset" => BorderStyle.Single,
            "thick" => BorderStyle.Thick,
            "dotted" or "dotDash" => BorderStyle.Dotted,
            "dashed" or "dashSmallGap" or "dotDotDash" or "dashDotStroked" => BorderStyle.Dashed,
            _ when val.Contains("ouble", StringComparison.Ordinal) || val.Contains("Thin", StringComparison.Ordinal)
                   || val.Contains("riple", StringComparison.Ordinal) => BorderStyle.Double,
            _ => BorderStyle.Single,
        };
        if (style == BorderStyle.None) return Border.None;
        double width = Math.Clamp(Units.FromEighthPoints(r.DoubleAttr("sz") ?? 2), 0.25, 12);
        double space = Math.Clamp(r.DoubleAttr("space") ?? 0, 0, 31);
        var color = r.WAttr("color") is "auto" or null ? (Color?)null : Color.TryParseHex(r.WAttr("color"));
        return new Border(style, width, color, space);
    }
}

internal sealed record TableFormat
{
    public string? StyleId { get; init; }
    public TableWidth? Width { get; init; }
    public TableAlignment? Alignment { get; init; }
    public double? Indent { get; init; }
    public TableLayoutMode? Layout { get; init; }
    public EdgeBorders Borders { get; init; } = EdgeBorders.Empty;
    public EdgeLengths CellMargins { get; init; } = EdgeLengths.Empty;
    public ColorSpec? Shading { get; init; }
    public int? Look { get; init; }
    public int? RowBandSize { get; init; }
    public int? ColumnBandSize { get; init; }

    public TableFormat Merge(TableFormat? over) => over is null ? this : new TableFormat
    {
        StyleId = over.StyleId ?? StyleId,
        Width = over.Width ?? Width,
        Alignment = over.Alignment ?? Alignment,
        Indent = over.Indent ?? Indent,
        Layout = over.Layout ?? Layout,
        Borders = Borders.Merge(over.Borders),
        CellMargins = CellMargins.Merge(over.CellMargins),
        Shading = over.Shading ?? Shading,
        Look = over.Look ?? Look,
        RowBandSize = over.RowBandSize ?? RowBandSize,
        ColumnBandSize = over.ColumnBandSize ?? ColumnBandSize,
    };

    public static TableFormat Parse(XmlReader r)
    {
        int depth = r.Depth;
        var f = new TableFormat();
        if (!r.Enter()) return f;
        while (r.NextChild(depth))
        {
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "tblBorders": f = f with { Borders = EdgeBorders.Parse(r) }; continue;
                case "tblCellMar": f = f with { CellMargins = EdgeLengths.Parse(r) }; continue;
                case "tblStyle": f = f with { StyleId = r.WAttr("val") }; break;
                case "tblW": f = f with { Width = ParseWidth(r) }; break;
                case "jc":
                    f = f with
                    {
                        Alignment = r.WAttr("val") switch
                        {
                            "center" => TableAlignment.Center,
                            "right" or "end" => TableAlignment.Right,
                            _ => TableAlignment.Left,
                        },
                    };
                    break;
                case "tblInd": f = f with { Indent = ParseWidth(r) is { Unit: TableWidthUnit.Points } w ? w.Value : 0 }; break;
                case "tblLayout": f = f with { Layout = r.WAttr("type") == "fixed" ? TableLayoutMode.Fixed : TableLayoutMode.Auto }; break;
                case "shd": f = f with { Shading = RunFormat.ParseShading(r) }; break;
                case "tblStyleRowBandSize": f = f with { RowBandSize = Math.Max(1, r.IntAttr("val") ?? 1) }; break;
                case "tblStyleColBandSize": f = f with { ColumnBandSize = Math.Max(1, r.IntAttr("val") ?? 1) }; break;
                case "tblLook": f = f with { Look = ParseLook(r) }; break;
            }
            r.Skip();
        }
        return f;
    }

    public const int LookFirstRow = 0x20, LookLastRow = 0x40, LookFirstColumn = 0x80, LookLastColumn = 0x100,
        LookNoHBand = 0x200, LookNoVBand = 0x400;

    private static int ParseLook(XmlReader r)
    {
        int look = int.TryParse(r.WAttr("val"), System.Globalization.NumberStyles.HexNumber, null, out var v) ? v : 0;
        void Flag(string attr, int bit)
        {
            if (r.WAttr(attr) is null) return;
            look = r.OnOff(attr) ? look | bit : look & ~bit;
        }
        Flag("firstRow", LookFirstRow);
        Flag("lastRow", LookLastRow);
        Flag("firstColumn", LookFirstColumn);
        Flag("lastColumn", LookLastColumn);
        Flag("noHBand", LookNoHBand);
        Flag("noVBand", LookNoVBand);
        return look;
    }

    internal static TableWidth? ParseWidth(XmlReader r)
    {
        var raw = r.WAttr("w");
        switch (r.WAttr("type"))
        {
            case "auto":
            case "nil":
                return TableWidth.Auto;
            case "pct":
                if (raw is null) return null;
                if (raw.EndsWith('%'))
                    return double.TryParse(raw.AsSpan(0, raw.Length - 1), System.Globalization.CultureInfo.InvariantCulture, out var p)
                        ? new TableWidth(p, TableWidthUnit.Percent) : null;
                return double.TryParse(raw, System.Globalization.CultureInfo.InvariantCulture, out var fiftieths)
                    ? new TableWidth(fiftieths / 50, TableWidthUnit.Percent) : null;
            default:
                return WordXml.ParseTwips(raw) is { } pts ? new TableWidth(pts, TableWidthUnit.Points) : null;
        }
    }
}

internal sealed record EdgeBorders(Border? Top, Border? Bottom, Border? Left, Border? Right, Border? InsideH, Border? InsideV)
{
    public static EdgeBorders Empty { get; } = new(null, null, null, null, null, null);

    public EdgeBorders Merge(EdgeBorders over) => new(over.Top ?? Top, over.Bottom ?? Bottom, over.Left ?? Left,
        over.Right ?? Right, over.InsideH ?? InsideH, over.InsideV ?? InsideV);

    public static EdgeBorders Parse(XmlReader r)
    {
        var b = Empty;
        int depth = r.Depth;
        if (!r.Enter()) return b;
        while (r.NextChild(depth))
        {
            if (r.IsW())
            {
                var border = Borders.Parse(r);
                b = r.LocalName switch
                {
                    "top" => b with { Top = border },
                    "bottom" => b with { Bottom = border },
                    "left" or "start" => b with { Left = border },
                    "right" or "end" => b with { Right = border },
                    "insideH" => b with { InsideH = border },
                    "insideV" => b with { InsideV = border },
                    _ => b,
                };
            }
            r.Skip();
        }
        return b;
    }
}

internal sealed record EdgeLengths(double? Top, double? Bottom, double? Left, double? Right)
{
    public static EdgeLengths Empty { get; } = new(null, null, null, null);

    public EdgeLengths Merge(EdgeLengths over) =>
        new(over.Top ?? Top, over.Bottom ?? Bottom, over.Left ?? Left, over.Right ?? Right);

    public static EdgeLengths Parse(XmlReader r)
    {
        var m = Empty;
        int depth = r.Depth;
        if (!r.Enter()) return m;
        while (r.NextChild(depth))
        {
            if (r.IsW() && TableFormat.ParseWidth(r) is { Unit: TableWidthUnit.Points } w)
            {
                m = r.LocalName switch
                {
                    "top" => m with { Top = w.Value },
                    "bottom" => m with { Bottom = w.Value },
                    "left" or "start" => m with { Left = w.Value },
                    "right" or "end" => m with { Right = w.Value },
                    _ => m,
                };
            }
            r.Skip();
        }
        return m;
    }
}

internal sealed record RowFormat
{
    public double? Height { get; init; }
    public RowHeightRule? HeightRule { get; init; }
    public bool? IsHeader { get; init; }
    public bool? CantSplit { get; init; }
    public int? GridBefore { get; init; }

    public static RowFormat Parse(XmlReader r)
    {
        int depth = r.Depth;
        var f = new RowFormat();
        if (!r.Enter()) return f;
        while (r.NextChild(depth))
        {
            if (r.IsW())
            {
                switch (r.LocalName)
                {
                    case "trHeight":
                        f = f with
                        {
                            Height = r.TwipsAttr("val"),
                            HeightRule = r.WAttr("hRule") switch
                            {
                                "exact" => RowHeightRule.Exact,
                                "auto" => RowHeightRule.Auto,
                                _ => RowHeightRule.AtLeast,
                            },
                        };
                        break;
                    case "tblHeader": f = f with { IsHeader = r.OnOff() }; break;
                    case "cantSplit": f = f with { CantSplit = r.OnOff() }; break;
                    case "gridBefore": f = f with { GridBefore = Math.Clamp(r.IntAttr("val") ?? 0, 0, 63) }; break;
                }
            }
            r.Skip();
        }
        return f;
    }
}

internal sealed record CellFormat
{
    public TableWidth? Width { get; init; }
    public int? GridSpan { get; init; }
    public VerticalMerge? VerticalMerge { get; init; }
    public bool? HorizontalMergeContinue { get; init; }
    public EdgeBorders Borders { get; init; } = EdgeBorders.Empty;
    public ColorSpec? Shading { get; init; }
    public EdgeLengths Margins { get; init; } = EdgeLengths.Empty;
    public CellVerticalAlignment? VerticalAlignment { get; init; }

    public CellFormat Merge(CellFormat? over) => over is null ? this : new CellFormat
    {
        Width = over.Width ?? Width,
        GridSpan = over.GridSpan ?? GridSpan,
        VerticalMerge = over.VerticalMerge ?? VerticalMerge,
        HorizontalMergeContinue = over.HorizontalMergeContinue ?? HorizontalMergeContinue,
        Borders = Borders.Merge(over.Borders),
        Shading = over.Shading ?? Shading,
        Margins = Margins.Merge(over.Margins),
        VerticalAlignment = over.VerticalAlignment ?? VerticalAlignment,
    };

    public static CellFormat Parse(XmlReader r)
    {
        int depth = r.Depth;
        var f = new CellFormat();
        if (!r.Enter()) return f;
        while (r.NextChild(depth))
        {
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "tcBorders": f = f with { Borders = EdgeBorders.Parse(r) }; continue;
                case "tcMar": f = f with { Margins = EdgeLengths.Parse(r) }; continue;
                case "tcW": f = f with { Width = TableFormat.ParseWidth(r) }; break;
                case "gridSpan": f = f with { GridSpan = Math.Clamp(r.IntAttr("val") ?? 1, 1, 63) }; break;
                case "vMerge": f = f with { VerticalMerge = r.WAttr("val") == "restart" ? Core.VerticalMerge.Restart : Core.VerticalMerge.Continue }; break;
                case "hMerge": f = f with { HorizontalMergeContinue = r.WAttr("val") != "restart" }; break;
                case "shd": f = f with { Shading = RunFormat.ParseShading(r) }; break;
                case "vAlign":
                    f = f with
                    {
                        VerticalAlignment = r.WAttr("val") switch
                        {
                            "center" => CellVerticalAlignment.Center,
                            "bottom" => CellVerticalAlignment.Bottom,
                            _ => CellVerticalAlignment.Top,
                        },
                    };
                    break;
            }
            r.Skip();
        }
        return f;
    }
}