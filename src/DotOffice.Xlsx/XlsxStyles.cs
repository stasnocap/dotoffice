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
    }

    public Color? this[int index] => index >= 0 && index < _colors.Length ? _colors[index] : null;

    public static XlsxTheme Read(XmlReader reader)
    {
        var theme = new XlsxTheme();
        var doc = XDocument.Load(reader);
        var scheme = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "clrScheme");
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

    /// <summary>The cell formats of a workbook plus the color resolver that rich-text runs share with them.</summary>
    public sealed record Result(List<CellFormat> Formats, Func<XElement?, Color?> Resolve);

    public static Result Read(XmlReader reader, XlsxTheme theme, IConversionWarningSink? warnings)
    {
        var doc = XDocument.Load(reader);
        var root = doc.Root;
        var formats = new List<CellFormat>();
        var palette = new Color?[LegacyPalette.Length];
        for (int i = 0; i < palette.Length; i++) palette[i] = Color.TryParseHex(LegacyPalette[i]);
        Color? Resolve(XElement? color) => ResolveColor(color, theme, palette);
        if (root is null) return new Result(formats, Resolve);

        ReadPalette(root, palette);

        var numberFormats = new Dictionary<int, string>();
        foreach (var f in Children(Child(root, "numFmts"), "numFmt"))
            if (Int(f, "numFmtId") is { } id && (string?)f.Attribute("formatCode") is { } code) numberFormats[id] = code;

        var fonts = Children(Child(root, "fonts"), "font").Select(f => ApplyFont(f, CellStyle.Default.Font, Resolve)).ToList();

        bool gradientWarned = false;
        var fills = new List<Color?>();
        foreach (var fill in Children(Child(root, "fills"), "fill"))
        {
            if (Child(fill, "gradientFill") is not null)
            {
                if (!gradientWarned)
                {
                    gradientWarned = true;
                    warnings.Warn("xlsx.gradient-fill", "Gradient cell fills are not supported and were left unfilled.");
                }
                fills.Add(null);
                continue;
            }
            var pattern = Child(fill, "patternFill");
            fills.Add((string?)pattern?.Attribute("patternType") == "solid" ? Resolve(Child(pattern, "fgColor")) : null);
        }

        var borders = Children(Child(root, "borders"), "border").Select(b => ReadBorders(b, Resolve)).ToList();

        foreach (var xf in Children(Child(root, "cellXfs"), "xf"))
        {
            int numFmtId = Int(xf, "numFmtId") ?? 0;
            var style = new CellStyle
            {
                Font = Pick(fonts, Int(xf, "fontId")) ?? CellStyle.Default.Font,
                Fill = Int(xf, "fillId") is { } fillId && fillId >= 0 && fillId < fills.Count ? fills[fillId] : null,
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
                };
            }
            string? code = numberFormats.GetValueOrDefault(numFmtId);
            formats.Add(new CellFormat(style, numFmtId, code));
        }
        return new Result(formats, Resolve);
    }

    /// <summary>Converts the file's textRotation (0–90 counter-clockwise, 91–180 clockwise, 255 stacked) to degrees counter-clockwise.</summary>
    private static double Rotation(int value) => value switch
    {
        >= 0 and <= 90 => value,
        > 90 and <= 180 => 90 - value,
        255 => 90,
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
