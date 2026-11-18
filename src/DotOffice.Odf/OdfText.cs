using System.Globalization;
using DotOffice.Core;

namespace DotOffice.Odf;

/// <summary>Style-to-model conversions shared by the ODF readers.</summary>
internal static class OdfText
{
    /// <summary>Character formatting from the <c>text-properties</c> of a style, on top of <paramref name="inherited"/>.</summary>
    public static RunProperties Run(OdfStyles styles, Dictionary<string, string> p, RunProperties inherited)
    {
        string? V(string key) => p.GetValueOrDefault("text-properties/" + key);
        var r = inherited;
        if ((V("font-family") ?? V("font-name")) is { } family) r = r with { FontFamily = styles.FontFamily(family) };
        if ((V("font-name-asian") ?? V("font-family-asian")) is { } asian) r = r with { EastAsianFontFamily = styles.FontFamily(asian) };
        if ((V("font-name-complex") ?? V("font-family-complex")) is { } complex) r = r with { ComplexScriptFontFamily = styles.FontFamily(complex) };
        if (V("font-size") is { } fs)
        {
            if (fs.EndsWith('%') && double.TryParse(fs.AsSpan(0, fs.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var pct) && pct > 0)
                r = r with { FontSize = inherited.FontSize * pct / 100 };
            else if (OdfUnits.Length(fs) is { } size and > 0) r = r with { FontSize = size };
        }
        if (OdfUnits.Length(V("font-size-complex")) is { } cs and > 0) r = r with { ComplexScriptFontSize = cs };
        if (V("font-weight") is { } w) r = r with { Bold = w == "bold" || int.TryParse(w, out var n) && n >= 600 };
        if (V("font-style") is { } st) r = r with { Italic = st is "italic" or "oblique" };
        if (V("text-underline-style") is { } us) r = r with { Underline = us == "none" ? UnderlineStyle.None : V("text-underline-type") == "double" ? UnderlineStyle.Double : UnderlineStyle.Single };
        if (V("text-line-through-style") is { } ls) r = r with { Strike = ls != "none" };
        if (V("color") is { } color) r = r with { Color = OdfUnits.Hex(color) };
        if (V("background-color") is { } bg) r = r with { Highlight = OdfUnits.Hex(bg) };
        if (V("font-variant") is { } fv) r = r with { SmallCaps = fv == "small-caps" };
        if (V("text-transform") is { } tt) r = r with { Caps = tt == "uppercase" };
        if (V("text-display") is { } td) r = r with { Hidden = td == "none" };
        if (V("language") is { } lang) r = r with { Language = V("country") is { } country && country != "none" ? lang + "-" + country : lang };
        if (OdfUnits.Length(V("letter-spacing")) is { } ls2) r = r with { CharacterSpacing = ls2 };
        if (V("text-position") is { } pos)
        {
            var first = pos.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
            r = r with
            {
                VerticalPosition = first.StartsWith("sub", StringComparison.Ordinal) || first.StartsWith('-') ? VerticalPosition.Subscript
                    : first == "0%" ? VerticalPosition.Baseline : VerticalPosition.Superscript,
            };
            if (pos.Split(' ', StringSplitOptions.RemoveEmptyEntries) is { Length: > 1 } parts && parts[1].EndsWith('%')
                && double.TryParse(parts[1].AsSpan(0, parts[1].Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var scale) && scale > 0 && scale < 100)
                r = r with { FontSize = r.FontSize * scale / 100 };
        }
        return r;
    }
}
