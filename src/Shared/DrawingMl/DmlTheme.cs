using System.Xml;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Dml;

/// <summary>Colors of a DrawingML theme, indexed the way SpreadsheetML refers to them (0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2, 4–9 accents, 10–11 links).</summary>
internal sealed class DmlTheme
{
    // clrScheme child names in SpreadsheetML theme-index order.
    private static readonly string[] SchemeOrder =
        ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

    // Office 2013+ defaults, used when the package has no theme part.
    private static readonly string[] DefaultColors =
        ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"];

    private readonly Color?[] _colors = new Color?[12];

    public static DmlTheme Default { get; } = new();

    public DmlTheme()
    {
        for (int i = 0; i < _colors.Length; i++) _colors[i] = Color.TryParseHex(DefaultColors[i]);
        ReadFormats(XElement.Parse(DefaultFormats));
    }

    /// <summary>The latin typeface of the theme's minor (body) font; null when the package has no theme part or it names none.</summary>
    public string? MinorLatin { get; private set; }

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

    private List<XElement> _fillStyles = [], _effectStyles = [], _backgroundFillStyles = [];

    /// <summary>The fill of style reference <paramref name="index"/> (1-3), or null when the theme has none.</summary>
    public XElement? FillStyle(int index) => index >= 1 && index <= _fillStyles.Count ? _fillStyles[index - 1] : null;

    /// <summary>The background fill of style reference <paramref name="index"/> (1-3, a <c>bgRef</c> idx minus 1000), or null when the theme has none.</summary>
    public XElement? BackgroundFillStyle(int index) => index >= 1 && index <= _backgroundFillStyles.Count ? _backgroundFillStyles[index - 1] : null;

    /// <summary>The <c>effectLst</c> of effect style reference <paramref name="index"/> (1-3), or null when the theme has none.</summary>
    public XElement? EffectStyle(int index) => index >= 1 && index <= _effectStyles.Count ? _effectStyles[index - 1] : null;

    private void ReadFormats(XElement? scheme)
    {
        _fillStyles = scheme?.Elements().FirstOrDefault(e => e.Name.LocalName == "fillStyleLst")?.Elements().ToList() ?? [];
        _backgroundFillStyles = scheme?.Elements().FirstOrDefault(e => e.Name.LocalName == "bgFillStyleLst")?.Elements().ToList() ?? [];
        _effectStyles = scheme?.Elements().FirstOrDefault(e => e.Name.LocalName == "effectStyleLst")?.Elements()
            .Select(style => style.Elements().FirstOrDefault(e => e.Name.LocalName == "effectLst") ?? new XElement("effectLst")).ToList() ?? [];
    }

    public static DmlTheme Read(XmlReader reader)
    {
        var theme = new DmlTheme();
        var doc = XDocument.Load(reader);
        var scheme = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "clrScheme");
        if (doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "minorFont")?.Elements().FirstOrDefault(e => e.Name.LocalName == "latin")?.Attribute("typeface")?.Value is { Length: > 0 } minor) theme.MinorLatin = minor;
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
