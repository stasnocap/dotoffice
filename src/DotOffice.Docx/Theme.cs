using System.Xml;

namespace DotOffice.Docx;

/// <summary>Theme font scheme (major = headings, minor = body) from theme/theme1.xml.</summary>
internal sealed class Theme
{
    public string? MajorLatin { get; private set; }
    public string? MinorLatin { get; private set; }
    public string? MajorEastAsia { get; private set; }
    public string? MinorEastAsia { get; private set; }

    public static Theme Read(XmlReader r)
    {
        var theme = new Theme();
        bool? major = null;
        while (r.Read())
        {
            if (r.NodeType == XmlNodeType.EndElement && r.IsA() && r.LocalName is "majorFont" or "minorFont")
                major = null;
            if (r.NodeType != XmlNodeType.Element || !r.IsA()) continue;
            switch (r.LocalName)
            {
                case "majorFont": major = true; break;
                case "minorFont": major = false; break;
                case "latin" when major is { } m:
                    if (m) theme.MajorLatin ??= NonEmpty(r.GetAttribute("typeface"));
                    else theme.MinorLatin ??= NonEmpty(r.GetAttribute("typeface"));
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
        "majorAscii" or "majorHAnsi" or "majorBidi" => MajorLatin,
        "minorAscii" or "minorHAnsi" or "minorBidi" => MinorLatin,
        "majorEastAsia" => MajorEastAsia ?? MajorLatin,
        "minorEastAsia" => MinorEastAsia ?? MinorLatin,
        _ => null,
    };

    private static string? NonEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;
}