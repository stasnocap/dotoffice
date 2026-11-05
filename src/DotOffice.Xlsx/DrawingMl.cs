using System.Globalization;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Xlsx;

/// <summary>Helpers for DrawingML (the markup of shapes and charts): colors with their modifiers and numbers.</summary>
internal static class DrawingMl
{
    public static XElement? Child(XElement? parent, string name) => parent?.Elements().FirstOrDefault(e => e.Name.LocalName == name);

    public static XElement? Path(XElement? parent, params string[] names)
    {
        foreach (var name in names) parent = Child(parent, name);
        return parent;
    }

    public static double Number(XElement? e, string attribute, double fallback = 0) =>
        double.TryParse((string?)e?.Attribute(attribute), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) ? v : fallback;

    public static bool Flag(XElement? e, string attribute) => (string?)e?.Attribute(attribute) is "1" or "true";

    /// <summary>
    /// The color a DrawingML color element describes (<c>srgbClr</c>, <c>schemeClr</c>, <c>sysClr</c>, <c>prstClr</c>), with luminance, tint and
    /// shade modifiers applied; <paramref name="placeholder"/> is what <c>phClr</c> stands for.
    /// </summary>
    public static Color? Resolve(XElement? holder, XlsxTheme theme, Color? placeholder = null)
    {
        var element = holder?.Elements().FirstOrDefault(e => e.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr" or "scrgbClr");
        if (element is null) return null;
        Color? color = element.Name.LocalName switch
        {
            "srgbClr" => Color.TryParseHex((string?)element.Attribute("val")),
            "sysClr" => Color.TryParseHex((string?)element.Attribute("lastClr")) ?? ((string?)element.Attribute("val") == "window" ? Color.White : Color.Black),
            "prstClr" => Preset((string?)element.Attribute("val")),
            "schemeClr" => (string?)element.Attribute("val") switch
            {
                "bg1" or "lt1" => theme[0],
                "tx1" or "dk1" => theme[1],
                "bg2" or "lt2" => theme[2],
                "tx2" or "dk2" => theme[3],
                "accent1" => theme[4], "accent2" => theme[5], "accent3" => theme[6], "accent4" => theme[7], "accent5" => theme[8], "accent6" => theme[9],
                "hlink" => theme[10], "folHlink" => theme[11],
                "phClr" => placeholder,
                _ => null,
            },
            _ => null,
        };
        if (color is not { } c) return null;
        foreach (var modifier in element.Elements())
        {
            double value = Number(modifier, "val") / 100_000.0;
            switch (modifier.Name.LocalName)
            {
                case "lumMod": c = WithHsl(c, (h, sat, l) => (h, sat, l * value)); break;
                case "lumOff": c = WithHsl(c, (h, sat, l) => (h, sat, l + value)); break;
                case "satMod": c = WithHsl(c, (h, sat, l) => (h, sat * value, l)); break;
                case "satOff": c = WithHsl(c, (h, sat, l) => (h, sat + value, l)); break;
                case "tint": c = c.Blend(Color.White, 1 - value); break;
                case "shade": c = c.Blend(Color.Black, 1 - value); break;
            }
        }
        return c;
    }

    /// <summary>The opacity (0–1) of the color element inside <paramref name="holder"/> (a <c>solidFill</c>, shadow or similar): its <c>alpha</c> modifier, 1 when absent.</summary>
    public static double ColorOpacity(XElement? holder)
    {
        var color = holder?.Elements().FirstOrDefault(e => e.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr" or "scrgbClr");
        return color is not null && Child(color, "alpha") is { } alpha ? Math.Clamp(Number(alpha, "val", 100_000) / 100_000.0, 0, 1) : 1;
    }

    /// <summary>The opacity (0–1) of the first solid fill under <paramref name="holder"/>: its color's <c>alpha</c> modifier, 1 when absent.</summary>
    public static double Opacity(XElement? holder)
    {
        var color = Child(holder, "solidFill")?.Elements().FirstOrDefault();
        return color is not null && Child(color, "alpha") is { } alpha ? Math.Clamp(Number(alpha, "val", 100_000) / 100_000.0, 0, 1) : 1;
    }

    private static Color? Preset(string? name) => name switch
    {
        "black" => Color.Black, "white" => Color.White, "red" => new Color(255, 0, 0), "green" => new Color(0, 128, 0), "blue" => new Color(0, 0, 255),
        "yellow" => new Color(255, 255, 0), "gray" or "grey" => new Color(128, 128, 128), "orange" => new Color(255, 165, 0), _ => null,
    };

    private static Color WithHsl(Color color, Func<double, double, double, (double H, double S, double L)> change)
    {
        double r = color.R / 255.0, g = color.G / 255.0, b = color.B / 255.0;
        double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b)), l = (max + min) / 2, h = 0, s = 0;
        if (max - min > 1e-9)
        {
            double d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            h = max == r ? (g - b) / d + (g < b ? 6 : 0) : max == g ? (b - r) / d + 2 : (r - g) / d + 4;
            h /= 6;
        }
        (h, s, l) = change(h, s, l);
        l = Math.Clamp(l, 0, 1);
        s = Math.Clamp(s, 0, 1);
        if (s < 1e-9) return new Color((byte)Math.Round(l * 255), (byte)Math.Round(l * 255), (byte)Math.Round(l * 255));
        double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
        static double Channel(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            return t < 1.0 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2.0 / 3 ? p + (q - p) * (2.0 / 3 - t) * 6 : p;
        }
        return new Color((byte)Math.Round(Channel(p, q, h + 1.0 / 3) * 255), (byte)Math.Round(Channel(p, q, h) * 255), (byte)Math.Round(Channel(p, q, h - 1.0 / 3) * 255));
    }

    /// <summary>
    /// The solid color of a fill property holder (<c>spPr</c> and the like): the <c>solidFill</c> color, null for <c>noFill</c>, and
    /// <paramref name="whenAbsent"/> when no fill is specified. Gradient fills take the color of their first stop.
    /// </summary>
    public static Color? SolidFill(XElement? holder, XlsxTheme theme, Color? whenAbsent = null, Color? placeholder = null)
    {
        if (holder is null) return whenAbsent;
        foreach (var child in holder.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "noFill": return null;
                case "solidFill": return Resolve(child, theme, placeholder) ?? whenAbsent;
                case "gradFill":
                    return Resolve(Child(Child(child, "gsLst"), "gs"), theme, placeholder) ?? whenAbsent;
                case "pattFill": return Resolve(Child(child, "fgClr"), theme, placeholder) ?? whenAbsent;
            }
        }
        return whenAbsent;
    }
}
