using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Xlsx;

/// <summary>A resolved cell format (<c>cellXfs</c> entry): the visual style plus the number format to apply to values.</summary>
internal sealed record CellFormat(CellStyle Style, int NumberFormatId, string? NumberFormatCode)
{
    public static CellFormat Default { get; } = new(CellStyle.Default, 0, null);
}

/// <summary>
/// A differential format (<c>dxf</c>) of conditional formatting. Every property is optional: only what the dxf mentions overrides the cell.
/// </summary>
internal sealed record Dxf
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public UnderlineStyle? Underline { get; init; }
    public bool? Strike { get; init; }
    public Color? FontColor { get; init; }
    public Color? Fill { get; init; }
    public Border? Left { get; init; }
    public Border? Right { get; init; }
    public Border? Top { get; init; }
    public Border? Bottom { get; init; }
    public string? NumberFormat { get; init; }
}

/// <summary>Colors of the workbook theme, indexed the way SpreadsheetML refers to them (0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2, 4–9 accents, 10–11 links).</summary>
internal sealed class XlsxTheme
{
    // clrScheme child names in SpreadsheetML theme-index order.
    private static readonly string[] SchemeOrder =
        ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

    // Office 2013+ defaults, used when the workbook has no theme part.
    private static readonly string[] DefaultColors =
        ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"];

    private readonly Color?[] _colors = new Color?[12];

    public static XlsxTheme Default { get; } = new();

    public XlsxTheme()
    {
        for (int i = 0; i < _colors.Length; i++) _colors[i] = Color.TryParseHex(DefaultColors[i]);
        ReadFormats(XElement.Parse(DefaultFormats));
    }

    public Color? this[int index] => index >= 0 && index < _colors.Length ? _colors[index] : null;

    // Office 2013+ format scheme: the fills and effects that a shape's style references (fillRef/effectRef idx 1-3) stand for.
    private const string DefaultFormats =
        "<fmtScheme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
        "<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>" +
        "<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs>" +
        "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></fillStyleLst>" +
        "<effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>" +
        "<a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></effectStyleLst></fmtScheme>";

    private List<XElement> _fillStyles = [], _effectStyles = [];

    /// <summary>The fill of style reference <paramref name="index"/> (1-3), or null when the theme has none.</summary>
    public XElement? FillStyle(int index) => index >= 1 && index <= _fillStyles.Count ? _fillStyles[index - 1] : null;

    /// <summary>The <c>effectLst</c> of effect style reference <paramref name="index"/> (1-3), or null when the theme has none.</summary>
    public XElement? EffectStyle(int index) => index >= 1 && index <= _effectStyles.Count ? _effectStyles[index - 1] : null;

    private void ReadFormats(XElement? scheme)
    {
        _fillStyles = scheme?.Elements().FirstOrDefault(e => e.Name.LocalName == "fillStyleLst")?.Elements().ToList() ?? [];
        _effectStyles = scheme?.Elements().FirstOrDefault(e => e.Name.LocalName == "effectStyleLst")?.Elements()
            .Select(style => style.Elements().FirstOrDefault(e => e.Name.LocalName == "effectLst") ?? new XElement("effectLst")).ToList() ?? [];
    }

    public static XlsxTheme Read(XmlReader reader)
    {
        var theme = new XlsxTheme();
        var doc = XDocument.Load(reader);
        var scheme = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "clrScheme");
        if (doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "fmtScheme") is { } formats) theme.ReadFormats(formats);
        if (scheme is null) return theme;
        for (int i = 0; i < SchemeOrder.Length; i++)
        {
            var element = scheme.Elements().FirstOrDefault(e => e.Name.LocalName == SchemeOrder[i])?.Elements().FirstOrDefault();
            if (element is null) continue;
            var hex = element.Name.LocalName switch
            {
                "srgbClr" => (string?)element.Attribute("val"),
                "sysClr" => (string?)element.Attribute("lastClr"),
                _ => null,
            };
            if (Color.TryParseHex(hex) is { } color) theme._colors[i] = color;
        }
        return theme;
    }
}

/// <summary>Reads <c>styles.xml</c> into a table of <see cref="CellFormat"/> indexed by a cell's <c>s</c> attribute.</summary>
internal static class XlsxStyles
{
    private static readonly string[] LegacyPalette =
    [
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
        "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
        "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
        "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
        "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
        "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
    ];

    /// <summary>An entry of Excel's default indexed palette (the one <c>[ColorN]</c> number formats use, as index N + 7).</summary>
    internal static Color DefaultIndexedColor(int index) => Color.TryParseHex(LegacyPalette[Math.Clamp(index, 0, LegacyPalette.Length - 1)])!.Value;

    /// <summary>The cell formats of a workbook plus the color resolver that rich-text runs share with them.</summary>
    public sealed record Result(List<CellFormat> Formats, Func<XElement?, Color?> Resolve, List<Dxf> Dxfs, IReadOnlyList<Color?> Palette, Func<XElement, Dxf> ReadInlineDxf);

    public static Result Read(XmlReader reader, XlsxTheme theme, IConversionWarningSink? warnings)
    {
        var doc = XDocument.Load(reader);
        var root = doc.Root;
        var formats = new List<CellFormat>();
        var palette = new Color?[LegacyPalette.Length];
        for (int i = 0; i < palette.Length; i++) palette[i] = Color.TryParseHex(LegacyPalette[i]);
        Color? Resolve(XElement? color) => ResolveColor(color, theme, palette);
        var dxfs = new List<Dxf>();
        if (root is null) return new Result(formats, Resolve, dxfs, palette, _ => new Dxf());

        ReadPalette(root, palette);

        var numberFormats = new Dictionary<int, string>();
        foreach (var f in Children(Child(root, "numFmts"), "numFmt"))
            if (Int(f, "numFmtId") is { } id && (string?)f.Attribute("formatCode") is { } code) numberFormats[id] = code;

        var fonts = Children(Child(root, "fonts"), "font").Select(f => ApplyFont(f, CellStyle.Default.Font, Resolve)).ToList();

        var fills = new List<(Color? Solid, GradientFill? Gradient)>();
        foreach (var fill in Children(Child(root, "fills"), "fill"))
        {
            if (Child(fill, "gradientFill") is { } gradient)
            {
                var (solid, parsed) = ReadGradient(gradient, Resolve);
                fills.Add((solid, parsed));
                continue;
            }
            var pattern = Child(fill, "patternFill");
            fills.Add(((string?)pattern?.Attribute("patternType") == "solid" ? Resolve(Child(pattern, "fgColor")) : null, null));
        }

        var borders = Children(Child(root, "borders"), "border").Select(b => ReadBorders(b, Resolve)).ToList();

        foreach (var xf in Children(Child(root, "cellXfs"), "xf"))
        {
            int numFmtId = Int(xf, "numFmtId") ?? 0;
            var style = new CellStyle
            {
                Font = Pick(fonts, Int(xf, "fontId")) ?? CellStyle.Default.Font,
                Fill = Int(xf, "fillId") is { } fillId && fillId >= 0 && fillId < fills.Count ? fills[fillId].Solid : null,
                Gradient = Int(xf, "fillId") is { } gradientId && gradientId >= 0 && gradientId < fills.Count ? fills[gradientId].Gradient : null,
                Borders = Pick(borders, Int(xf, "borderId")) ?? BorderSet.Empty,
            };
            if (Child(xf, "alignment") is { } alignment)
            {
                style = style with
                {
                    HorizontalAlignment = ((string?)alignment.Attribute("horizontal")) switch
                    {
                        "left" => HorizontalAlignment.Left,
                        "center" or "centerContinuous" => HorizontalAlignment.Center,
                        "right" => HorizontalAlignment.Right,
                        "fill" => HorizontalAlignment.Fill,
                        "justify" or "distributed" => HorizontalAlignment.Justify,
                        _ => HorizontalAlignment.General,
                    },
                    VerticalAlignment = ((string?)alignment.Attribute("vertical")) switch
                    {
                        "top" => CellVerticalAlignment.Top,
                        "center" or "distributed" or "justify" => CellVerticalAlignment.Center,
                        _ => CellVerticalAlignment.Bottom,
                    },
                    Wrap = Bool(alignment, "wrapText"),
                    Indent = Math.Clamp(Int(alignment, "indent") ?? 0, 0, 250),
                    Rotation = Rotation(Int(alignment, "textRotation") ?? 0),
                    Stacked = Int(alignment, "textRotation") == 255,
                };
            }
            string? code = numberFormats.GetValueOrDefault(numFmtId);
            formats.Add(new CellFormat(style, numFmtId, code));
        }
        foreach (var dxf in Children(Child(root, "dxfs"), "dxf")) dxfs.Add(ReadDxf(dxf, numberFormats, Resolve));
        return new Result(formats, Resolve, dxfs, palette, d => ReadDxf(d, numberFormats, Resolve));
    }

    private static Dxf ReadDxf(XElement dxf, Dictionary<int, string> numberFormats, Func<XElement?, Color?> resolve)
    {
        var result = new Dxf();
        if (Child(dxf, "font") is { } font)
        {
            // ApplyFont inherits what the element does not mention; the two probes tell which properties are really present.
            var off = new RunProperties { Bold = false, Italic = false, Strike = false, Underline = UnderlineStyle.None };
            var on = new RunProperties { Bold = true, Italic = true, Strike = true, Underline = UnderlineStyle.Double };
            var a = ApplyFont(font, off, resolve);
            var b = ApplyFont(font, on, resolve);
            result = result with
            {
                Bold = a.Bold == b.Bold ? a.Bold : null,
                Italic = a.Italic == b.Italic ? a.Italic : null,
                Strike = a.Strike == b.Strike ? a.Strike : null,
                Underline = a.Underline == b.Underline ? a.Underline : null,
                FontColor = Child(font, "color") is { } color ? resolve(color) : null,
            };
        }
        if (Child(dxf, "fill") is { } fill && Child(fill, "patternFill") is { } pattern)
        {
            // In a dxf the color of a solid fill is the background color (Excel writes it the other way round for plain patterns).
            var type = (string?)pattern.Attribute("patternType");
            if (type != "none") result = result with { Fill = resolve(Child(pattern, "bgColor")) ?? resolve(Child(pattern, "fgColor")) };
        }
        if (Child(dxf, "border") is { } border)
        {
            var edges = ReadBorders(border, resolve);
            result = result with
            {
                Left = edges.Left.IsVisible ? edges.Left : null,
                Right = edges.Right.IsVisible ? edges.Right : null,
                Top = edges.Top.IsVisible ? edges.Top : null,
                Bottom = edges.Bottom.IsVisible ? edges.Bottom : null,
            };
        }
        if (Child(dxf, "numFmt") is { } numFmt)
        {
            var code = (string?)numFmt.Attribute("formatCode");
            if (code is null && Int(numFmt, "numFmtId") is { } id) code = numberFormats.GetValueOrDefault(id) ?? NumberFormat.BuiltIn(id);
            result = result with { NumberFormat = code };
        }
        return result;
    }

    /// <summary>
    /// Reads a <c>gradientFill</c>. Stops whose color cannot be resolved are dropped; with fewer than two left the fill degrades to the
    /// remaining color (or nothing). Excel's path gradients are always rectangular, so there is no separate circle kind.
    /// </summary>
    private static (Color? Solid, GradientFill? Gradient) ReadGradient(XElement gradient, Func<XElement?, Color?> resolve)
    {
        var stops = new List<GradientStop>();
        foreach (var stop in Children(gradient, "stop"))
            if (Double(stop, "position") is { } position && resolve(Child(stop, "color")) is { } color)
                stops.Add(new GradientStop(Math.Clamp(position, 0, 1), color));
        stops = stops.OrderBy(x => x.Position).ToList();
        if (stops.Count < 2) return (stops.Count == 1 ? stops[0].Color : null, null);
        bool path = (string?)gradient.Attribute("type") == "path";
        return (null, new GradientFill
        {
            Kind = path ? GradientKind.Path : GradientKind.Linear,
            Degree = Double(gradient, "degree") ?? 0,
            Left = Math.Clamp(Double(gradient, "left") ?? 0, 0, 1),
            Right = Math.Clamp(Double(gradient, "right") ?? 0, 0, 1),
            Top = Math.Clamp(Double(gradient, "top") ?? 0, 0, 1),
            Bottom = Math.Clamp(Double(gradient, "bottom") ?? 0, 0, 1),
            Stops = stops,
        });
    }

    /// <summary>Converts the file's textRotation (0–90 counter-clockwise, 91–180 clockwise) to degrees counter-clockwise; 255 (stacked) is no rotation.</summary>
    private static double Rotation(int value) => value switch
    {
        >= 0 and <= 90 => value,
        > 90 and <= 180 => 90 - value,
        _ => 0,
    };

    /// <summary>
    /// Applies the properties found in a <c>font</c> or rich-text <c>rPr</c> element on top of <paramref name="props"/>.
    /// Properties the element does not mention are inherited (for a rich-text run: from the cell's font, which is what Excel shows).
    /// </summary>
    internal static RunProperties ApplyFont(XElement font, RunProperties props, Func<XElement?, Color?> resolve)
    {
        if ((Child(font, "name") ?? Child(font, "rFont")) is { } name && (string?)name.Attribute("val") is { Length: > 0 } family) props = props with { FontFamily = family };
        if (Child(font, "sz") is { } size && Double(size, "val") is { } points and > 0) props = props with { FontSize = Math.Min(points, 409) };
        if (Child(font, "b") is { } b) props = props with { Bold = Flag(b) };
        if (Child(font, "i") is { } i) props = props with { Italic = Flag(i) };
        if (Child(font, "strike") is { } strike) props = props with { Strike = Flag(strike) };
        if (Child(font, "u") is { } u)
            props = props with
            {
                Underline = !Flag(u, "none") ? UnderlineStyle.None
                    : ((string?)u.Attribute("val")) is "double" or "doubleAccounting" ? UnderlineStyle.Double : UnderlineStyle.Single,
            };
        if (Child(font, "vertAlign") is { } v)
            props = props with
            {
                VerticalPosition = ((string?)v.Attribute("val")) switch
                {
                    "superscript" => VerticalPosition.Superscript,
                    "subscript" => VerticalPosition.Subscript,
                    _ => VerticalPosition.Baseline,
                },
            };
        if (Child(font, "color") is { } color) props = props with { Color = resolve(color) };
        return props;
    }

    private static BorderSet ReadBorders(XElement border, Func<XElement?, Color?> resolve)
    {
        Border Edge(params string[] names)
        {
            foreach (var name in names)
            {
                var edge = Child(border, name);
                if (edge is null) continue;
                var (style, width) = ((string?)edge.Attribute("style")) switch
                {
                    "thin" => (BorderStyle.Single, 0.75),
                    "hair" => (BorderStyle.Single, 0.25),
                    "medium" => (BorderStyle.Single, 1.5),
                    "thick" => (BorderStyle.Thick, 2.25),
                    "double" => (BorderStyle.Double, 1.5),
                    "dotted" => (BorderStyle.Dotted, 0.75),
                    "dashed" or "dashDot" or "dashDotDot" or "slantDashDot" => (BorderStyle.Dashed, 0.75),
                    "mediumDashed" or "mediumDashDot" or "mediumDashDotDot" => (BorderStyle.Dashed, 1.5),
                    _ => (BorderStyle.None, 0.0),
                };
                return style == BorderStyle.None ? Border.None : new Border(style, width, resolve(Child(edge, "color")));
            }
            return Border.None;
        }
        return new BorderSet { Left = Edge("left", "start"), Right = Edge("right", "end"), Top = Edge("top"), Bottom = Edge("bottom") };
    }

    private static void ReadPalette(XElement root, Color?[] palette)
    {
        int index = 0;
        foreach (var rgb in Children(Child(Child(root, "colors"), "indexedColors"), "rgbColor"))
        {
            if (index < palette.Length && Color.TryParseHex(Argb((string?)rgb.Attribute("rgb"))) is { } color) palette[index] = color;
            index++;
        }
    }

    private static Color? ResolveColor(XElement? element, XlsxTheme theme, Color?[] palette)
    {
        if (element is null) return null;
        Color? color = null;
        if ((string?)element.Attribute("rgb") is { } rgb) color = Color.TryParseHex(Argb(rgb));
        else if (Int(element, "indexed") is { } indexed)
            color = indexed >= 0 && indexed < palette.Length ? palette[indexed] : indexed == 64 ? null : indexed == 65 ? Color.White : null;
        else if (Int(element, "theme") is { } themeIndex) color = theme[themeIndex];
        if (color is { } c && Double(element, "tint") is { } tint and not 0) color = ApplyTint(c, Math.Clamp(tint, -1, 1));
        return color;
    }

    /// <summary>Strips the alpha byte of an ARGB value.</summary>
    private static string? Argb(string? value) => value is { Length: 8 } ? value[2..] : value;

    /// <summary>Excel's tint: lighten (positive) or darken (negative) the luminance in HLS space.</summary>
    internal static Color ApplyTint(Color color, double tint)
    {
        var (h, l, s) = ToHls(color);
        l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
        return FromHls(h, Math.Clamp(l, 0, 1), s);
    }

    private static (double H, double L, double S) ToHls(Color color)
    {
        double r = color.R / 255.0, g = color.G / 255.0, b = color.B / 255.0;
        double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
        double l = (max + min) / 2, d = max - min, h = 0, s = 0;
        if (d > 1e-9)
        {
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            if (max == r) h = (g - b) / d + (g < b ? 6 : 0);
            else if (max == g) h = (b - r) / d + 2;
            else h = (r - g) / d + 4;
            h /= 6;
        }
        return (h, l, s);
    }

    private static Color FromHls(double h, double l, double s)
    {
        double r, g, b;
        if (s <= 1e-9)
        {
            r = g = b = l;
        }
        else
        {
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
            static double Channel(double p, double q, double t)
            {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1.0 / 6) return p + (q - p) * 6 * t;
                if (t < 0.5) return q;
                if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
                return p;
            }
            r = Channel(p, q, h + 1.0 / 3);
            g = Channel(p, q, h);
            b = Channel(p, q, h - 1.0 / 3);
        }
        return new Color((byte)Math.Round(r * 255), (byte)Math.Round(g * 255), (byte)Math.Round(b * 255));
    }

    // ---- small XElement helpers (namespace-agnostic: matches the strict and transitional namespaces)

    private static XElement? Child(XElement? parent, string name) => parent?.Elements().FirstOrDefault(e => e.Name.LocalName == name);

    private static IEnumerable<XElement> Children(XElement? parent, string name) =>
        parent?.Elements().Where(e => e.Name.LocalName == name) ?? [];

    private static int? Int(XElement e, string name) =>
        int.TryParse((string?)e.Attribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    private static double? Double(XElement e, string name) =>
        double.TryParse((string?)e.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) ? v : null;

    private static bool Bool(XElement e, string name) => (string?)e.Attribute(name) is "1" or "true";

    /// <summary>An on/off element such as <c>&lt;b/&gt;</c>: true unless val is "0"/"false" (or equals <paramref name="offValue"/>).</summary>
    private static bool Flag(XElement e, string? offValue = null) =>
        (string?)e.Attribute("val") is var val && val is not ("0" or "false") && (offValue is null || val != offValue);

    private static T? Pick<T>(List<T> list, int? index) where T : class =>
        index is { } i && i >= 0 && i < list.Count ? list[i] : null;
}
