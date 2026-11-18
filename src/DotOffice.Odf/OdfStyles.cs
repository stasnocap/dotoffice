using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Odf;

internal static class Ns
{
    public static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    public static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    public static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    public static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    public static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    public static readonly XNamespace Number = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
    public static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
}

/// <summary>ODF length and colour attribute parsing.</summary>
internal static class OdfUnits
{
    /// <summary>Parses "1.5cm", "0.2in", "12pt", "3mm", "4px", "1pc" into points.</summary>
    public static double? Length(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var s = value.Trim();
        int i = s.Length;
        while (i > 0 && !(char.IsDigit(s[i - 1]) || s[i - 1] == '.')) i--;
        if (i == 0 || !double.TryParse(s.AsSpan(0, i), NumberStyles.Float, CultureInfo.InvariantCulture, out var n) || !double.IsFinite(n)) return null;
        return s[i..].ToLowerInvariant() switch
        {
            "cm" => n * 72 / 2.54,
            "mm" => n * 72 / 25.4,
            "in" => n * 72,
            "pt" or "" => n,
            "pc" => n * 12,
            "px" => n * 0.75,
            _ => null,
        };
    }

    public static Color? Hex(string? value) => value is null or "transparent" ? null : Color.TryParseHex(value);
}

/// <summary>Collected <c>style:style</c> and number-style elements of styles.xml and content.xml.</summary>
internal sealed class OdfStyles
{
    private readonly Dictionary<(string Family, string Name), XElement> _styles = [];
    private readonly Dictionary<string, XElement> _numberStyles = [];
    private readonly Dictionary<string, XElement> _pageLayouts = [];
    private readonly Dictionary<string, string> _masterLayouts = [];
    private readonly Dictionary<string, string> _codes = [];
    private readonly Dictionary<string, XElement> _defaults = [];
    private readonly Dictionary<string, XElement> _listStyles = [];
    private readonly Dictionary<string, string> _fonts = [];
    private readonly Dictionary<string, XElement> _masters = [];

    public XElement? ListStyle(string? name) => name is not null && _listStyles.TryGetValue(name, out var e) ? e : null;

    public XElement? MasterPage(string? name) => name is not null && _masters.TryGetValue(name, out var e) ? e : null;

    public IEnumerable<XElement> MasterPages => _masters.Values;

    /// <summary>The family a <c>style:font-name</c> stands for (from the font face declarations), or the name itself.</summary>
    public string FontFamily(string name) => _fonts.TryGetValue(name, out var f) ? f : name.Trim('\'', '"');

    public void Add(XElement container)
    {
        foreach (var e in container.Elements())
        {
            if (e.Name == Ns.Style + "default-style" && (string?)e.Attribute(Ns.Style + "family") is { } fam) _defaults[fam] = e;
            var name = (string?)e.Attribute(Ns.Style + "name");
            if (name is null) continue;
            if (e.Name == Ns.Style + "font-face" && (string?)e.Attribute(Ns.Svg + "font-family") is { } ff) _fonts[name] = ff.Trim('\'', '"');
            else if (e.Name == Ns.Text + "list-style") _listStyles[name] = e;
            else if (e.Name == Ns.Style + "style")
                _styles[((string?)e.Attribute(Ns.Style + "family") ?? "", name)] = e;
            else if (e.Name == Ns.Style + "page-layout") _pageLayouts[name] = e;
            else if (e.Name == Ns.Style + "master-page")
            {
                _masterLayouts[name] = (string?)e.Attribute(Ns.Style + "page-layout-name") ?? "";
                _masters[name] = e;
                DefaultMasterPage ??= name;
            }
            else if (e.Name.Namespace == Ns.Number && e.Name.LocalName.EndsWith("-style", StringComparison.Ordinal))
                _numberStyles[name] = e;
        }
    }

    public XElement? Find(string family, string? name) =>
        name is not null && _styles.TryGetValue((family, name), out var e) ? e : null;

    /// <summary>The merged properties of a style and its parents, keyed "element/attribute" by local names.</summary>
    public Dictionary<string, string> Properties(string family, string? name)
    {
        var chain = new List<XElement>();
        for (var e = Find(family, name); e is not null && chain.Count < 16; e = Find(family, (string?)e.Attribute(Ns.Style + "parent-style-name")))
            chain.Add(e);
        var result = new Dictionary<string, string>();
        if (_defaults.TryGetValue(family, out var def))
            foreach (var props in def.Elements())
                foreach (var a in props.Attributes())
                    if (!a.IsNamespaceDeclaration) result[props.Name.LocalName + "/" + a.Name.LocalName] = a.Value;
        for (int i = chain.Count - 1; i >= 0; i--)
            foreach (var props in chain[i].Elements())
                foreach (var a in props.Attributes())
                    if (!a.IsNamespaceDeclaration) result[props.Name.LocalName + "/" + a.Name.LocalName] = a.Value;
        // The data style is named on the style itself.
        for (int i = chain.Count - 1; i >= 0; i--)
            if ((string?)chain[i].Attribute(Ns.Style + "data-style-name") is { } ds) result["style/data-style-name"] = ds;
        return result;
    }

    public XElement? PageLayoutOf(string? masterPage) =>
        masterPage is not null && _masterLayouts.TryGetValue(masterPage, out var l) && _pageLayouts.TryGetValue(l, out var e) ? e : null;

    public string? DefaultMasterPage { get; set; }

    /// <summary>The spreadsheet number format code (Excel syntax) equivalent to the named ODF data style.</summary>
    public string? NumberCode(string? name, int depth = 0)
    {
        if (name is null) return null;
        if (_codes.TryGetValue(name, out var cached)) return cached;
        if (!_numberStyles.TryGetValue(name, out var style) || depth > 4) return null;
        var code = Build(style);
        // style:map "value()>=0" applies another style to non-negative values; this style keeps the rest.
        foreach (var map in style.Elements(Ns.Style + "map"))
        {
            var cond = ((string?)map.Attribute(Ns.Style + "condition") ?? "").Replace(" ", "");
            if (cond is "value()>=0" or "value()>0" && NumberCode((string?)map.Attribute(Ns.Style + "apply-style-name"), depth + 1) is { } pos)
            {
                code = pos + ";" + code;
                break;
            }
        }
        return _codes[name] = code;
    }

    private static string Build(XElement style)
    {
        var sb = new StringBuilder();
        bool hours = false;
        foreach (var e in style.Elements())
        {
            var long_ = (string?)e.Attribute(Ns.Number + "style") == "long";
            switch (e.Name.LocalName)
            {
                case "number":
                {
                    int dec = Int(e, "decimal-places") ?? 0, min = Math.Clamp(Int(e, "min-integer-digits") ?? 1, 0, 30);
                    bool group = (string?)e.Attribute(Ns.Number + "grouping") == "true";
                    var digits = new string('0', min);
                    if (group) digits = digits.Length >= 4 ? digits : "#,##" + digits.PadLeft(1, '0');
                    sb.Append(digits.Length == 0 ? "#" : digits);
                    if (dec > 0) sb.Append('.').Append('0', Math.Min(dec, 30));
                    break;
                }
                case "scientific-number":
                {
                    int dec = Int(e, "decimal-places") ?? 0;
                    sb.Append('0');
                    if (dec > 0) sb.Append('.').Append('0', Math.Min(dec, 30));
                    sb.Append("E+").Append('0', Math.Clamp(Int(e, "min-exponent-digits") ?? 2, 1, 5));
                    break;
                }
                case "fraction":
                    sb.Append("# ").Append('?', Math.Clamp(Int(e, "min-numerator-digits") ?? 1, 1, 5)).Append('/')
                      .Append('?', Math.Clamp(Int(e, "min-denominator-digits") ?? 1, 1, 5));
                    break;
                case "text": sb.Append(Literal(e.Value)); break;
                case "currency-symbol": sb.Append(Literal(e.Value)); break;
                case "year": sb.Append(long_ ? "yyyy" : "yy"); break;
                case "month":
                    sb.Append((string?)e.Attribute(Ns.Number + "textual") == "true" ? (long_ ? "mmmm" : "mmm") : (long_ ? "mm" : "m"));
                    break;
                case "day": sb.Append(long_ ? "dd" : "d"); break;
                case "day-of-week": sb.Append(long_ ? "dddd" : "ddd"); break;
                case "hours":
                    sb.Append((string?)style.Elements(Ns.Number + "hours").FirstOrDefault()?.Attribute(Ns.Number + "truncate-on-overflow") == "false"
                        ? "[h]" : long_ ? "hh" : "h");
                    hours = true;
                    break;
                case "minutes": sb.Append(long_ ? "mm" : "m"); break;
                case "seconds":
                    sb.Append(long_ ? "ss" : "s");
                    if (Int(e, "decimal-places") is > 0 and var d) sb.Append('.').Append('0', Math.Min(d, 9));
                    break;
                case "am-pm": sb.Append("AM/PM"); break;
                case "quarter": sb.Append("\"Q\"0"); break;
                case "text-content": sb.Append('@'); break;
            }
        }
        _ = hours;
        return sb.Length == 0 ? "General" : sb.ToString();
    }

    private static int? Int(XElement e, string name) =>
        int.TryParse((string?)e.Attribute(Ns.Number + name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    private static string Literal(string text) => text.Length == 0 ? "" : "\"" + text.Replace("\"", "") + "\"";
}
