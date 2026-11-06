using System.Xml;

namespace DotOffice.Docx;

/// <summary>Theme font scheme (major = headings, minor = body) from theme/theme1.xml.</summary>
internal sealed class Theme
{
    public string? MajorLatin { get; private set; }
    public string? MinorLatin { get; private set; }
    public string? MajorEastAsia { get; private set; }
    public string? MinorEastAsia { get; private set; }
    public string? MajorComplexScript { get; private set; }
    public string? MinorComplexScript { get; private set; }

    /// <summary>The color scheme by element name (dk1, lt1, dk2, lt2, accent1–6, hlink, folHlink).</summary>
    public Dictionary<string, Core.Color> Colors { get; } = [];

    public static Theme Read(XmlReader r)
    {
        var theme = new Theme();
        bool? major = null;
        while (r.Read())
        {
            if (r.NodeType == XmlNodeType.EndElement && r.IsA() && r.LocalName is "majorFont" or "minorFont")
                major = null;
            if (r.NodeType != XmlNodeType.Element || !r.IsA()) continue;
            if (r.LocalName is "dk1" or "lt1" or "dk2" or "lt2" or "accent1" or "accent2" or "accent3" or "accent4" or "accent5" or "accent6" or "hlink" or "folHlink")
            {
                string name = r.LocalName;
                int d = r.Depth;
                while (!r.IsEmptyElement && r.Read() && r.Depth > d)
                {
                    if (r.NodeType != XmlNodeType.Element) continue;
                    var c = r.LocalName == "srgbClr" ? Core.Color.TryParseHex(r.GetAttribute("val"))
                          : r.LocalName == "sysClr" ? Core.Color.TryParseHex(r.GetAttribute("lastClr")) ?? (r.GetAttribute("val") == "window" ? Core.Color.White : Core.Color.Black)
                          : null;
                    if (c is { } color) theme.Colors.TryAdd(name, color);
                }
                if (r.NodeType != XmlNodeType.Element && r.NodeType != XmlNodeType.EndElement) continue;
            }
            switch (r.LocalName)
            {
                case "majorFont": major = true; break;
                case "minorFont": major = false; break;
                case "latin" when major is { } m:
                    if (m) theme.MajorLatin ??= NonEmpty(r.GetAttribute("typeface"));
                    else theme.MinorLatin ??= NonEmpty(r.GetAttribute("typeface"));
                    break;
                case "cs" when major is { } m:
                    if (m) theme.MajorComplexScript ??= NonEmpty(r.GetAttribute("typeface"));
                    else theme.MinorComplexScript ??= NonEmpty(r.GetAttribute("typeface"));
                    break;
                case "ea" when major is { } m:
                    if (m) theme.MajorEastAsia ??= NonEmpty(r.GetAttribute("typeface"));
                    else theme.MinorEastAsia ??= NonEmpty(r.GetAttribute("typeface"));
                    break;
            }
        }
        return theme;
    }

    /// <summary>Maps an rFonts theme attribute value (e.g. "minorHAnsi") to a typeface.</summary>
    public string? Resolve(string theme) => theme switch
    {
        "majorAscii" or "majorHAnsi" => MajorLatin,
        "minorAscii" or "minorHAnsi" => MinorLatin,
        "majorBidi" => MajorComplexScript ?? MajorLatin,
        "minorBidi" => MinorComplexScript ?? MinorLatin,
        "majorEastAsia" => MajorEastAsia ?? MajorLatin,
        "minorEastAsia" => MinorEastAsia ?? MinorLatin,
        _ => null,
    };

    private static string? NonEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;
}