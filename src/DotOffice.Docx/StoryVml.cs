using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>Legacy VML text boxes and rectangles (<c>w:pict</c> / <c>v:shape</c>), as Word 2007 and the VML fallbacks write them.</summary>
internal sealed partial class StoryReader
{
    private const string Vml = "urn:schemas-microsoft-com:vml";

    private void ReadPict(XmlReader r, ParagraphState state, RunProperties props)
    {
        bool drawn = false;
        int depth = r.Depth;
        if (r.Enter())
        while (r.NextChild(depth))
        {
            if (r.NamespaceURI == Vml && r.LocalName is "shape" or "rect" or "roundrect" or "oval")
            {
                drawn |= ReadVmlShape(r, state, props);
                continue;
            }
            r.Skip();
        }
        if (!drawn) context.WarnOnce("docx.vml", "VML pictures and embedded objects are not supported and were skipped.");
    }

    private bool ReadVmlShape(XmlReader r, ParagraphState state, RunProperties props)
    {
        string kind = r.LocalName;
        var style = VmlStyle(r.GetAttribute("style"));
        var data = new ShapeData
        {
            Geometry = kind switch { "roundrect" => "roundRect", "oval" => "ellipse", _ => "rect" },
            Insets = new Thickness(3.6, 7.2, 3.6, 7.2),
        };
        bool filled = r.GetAttribute("filled") is not ("f" or "false");
        bool stroked = r.GetAttribute("stroked") is not ("f" or "false");
        var fill = VmlColor(r.GetAttribute("fillcolor")) ?? new Color(255, 255, 255);
        var line = VmlColor(r.GetAttribute("strokecolor")) ?? new Color(0, 0, 0);
        data.OutlineWidth = VmlLength(r.GetAttribute("strokeweight")) ?? 0.75;
        var wrap = TextWrap.None;
        bool hasBox = false;
        int depth = r.Depth;
        if (r.Enter())
        {
            while (r.NextChild(depth))
            {
                if (r.NamespaceURI == Vml && r.LocalName == "textbox")
                {
                    hasBox = true;
                    if (r.GetAttribute("inset") is { } inset)
                    {
                        var parts = inset.Split(',').Select(p => p.Trim()).ToArray();
                        double Part(int i, double fallback) => i < parts.Length && VmlLength(parts[i]) is { } v ? v : fallback;
                        data.Insets = new Thickness(Part(1, 3.6), Part(2, 7.2), Part(3, 3.6), Part(0, 7.2));
                    }
                    ReadContainer(r, "txbxContent", () => ReadBlocks(r, data.Blocks, null, 0, null));
                }
                else if (r.NamespaceURI == Vml && r.LocalName == "fill" && VmlColor(r.GetAttribute("color")) is { } fc) { fill = fc; r.Skip(); }
                else if (r.NamespaceURI == Vml && r.LocalName == "stroke" && VmlColor(r.GetAttribute("color")) is { } sc) { line = sc; r.Skip(); }
                else if (r.LocalName == "wrap" && r.NamespaceURI == "urn:schemas-microsoft-com:office:word")
                {
                    wrap = r.GetAttribute("type") switch { "square" or "tight" or "through" => TextWrap.Square, "topAndBottom" => TextWrap.TopAndBottom, _ => TextWrap.None };
                    r.Skip();
                }
                else r.Skip();
            }
        }
        if (filled) data.Fill = fill;
        if (stroked) data.Outline = line;
        if (!hasBox && !filled && !stroked) return false;
        double width = VmlLength(style.GetValueOrDefault("width")) ?? 0, height = VmlLength(style.GetValueOrDefault("height")) ?? 0;
        if (width <= 0 || height <= 0) return false;

        bool absolute = style.GetValueOrDefault("position") == "absolute";
        bool behind = style.TryGetValue("z-index", out var z) && int.TryParse(z, out var zi) && zi < 0;
        if (!absolute)
        {
            var origin = new FloatPosition(FloatReference.Column, FloatAlignment.Offset, 0);
            if (CreateShape(data, width, height, props, origin, origin, TextWrap.None, false, default) is { } inline) AddInline(state, new InlineShape(inline));
            return true;
        }
        var horizontal = VmlPosition(style, "horizontal", "margin-left", "column");
        var vertical = VmlPosition(style, "vertical", "margin-top", "paragraph");
        var distance = new Thickness(VmlLength(style.GetValueOrDefault("mso-wrap-distance-top")) ?? 0, VmlLength(style.GetValueOrDefault("mso-wrap-distance-right")) ?? 9,
            VmlLength(style.GetValueOrDefault("mso-wrap-distance-bottom")) ?? 0, VmlLength(style.GetValueOrDefault("mso-wrap-distance-left")) ?? 9);
        if (CreateShape(data, width, height, props, horizontal, vertical, wrap, behind, distance) is { } box) AddInline(state, box);
        return true;
    }

    private static FloatPosition VmlPosition(Dictionary<string, string> style, string axis, string margin, string defaultReference)
    {
        bool h = axis == "horizontal";
        var reference = ParseReference(style.GetValueOrDefault($"mso-position-{axis}-relative") ?? defaultReference, h);
        var alignment = style.GetValueOrDefault($"mso-position-{axis}") switch
        {
            "left" or "top" => FloatAlignment.Start,
            "center" => FloatAlignment.Center,
            "right" or "bottom" => FloatAlignment.End,
            "inside" => FloatAlignment.Inside,
            "outside" => FloatAlignment.Outside,
            _ => FloatAlignment.Offset,
        };
        double offset = VmlLength(style.GetValueOrDefault(margin)) ?? 0;
        return new FloatPosition(reference, alignment, Math.Clamp(offset, -1584, 1584));
    }

    private static Dictionary<string, string> VmlStyle(string? style)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in (style ?? "").Split(';'))
        {
            int colon = part.IndexOf(':');
            if (colon > 0) result[part[..colon].Trim()] = part[(colon + 1)..].Trim();
        }
        return result;
    }

    /// <summary>A VML length (pt, in, cm, mm, px, pc; bare numbers are points) in points.</summary>
    private static double? VmlLength(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        text = text.Trim();
        int i = text.Length;
        while (i > 0 && char.IsLetter(text[i - 1])) i--;
        if (!double.TryParse(text[..i], NumberStyles.Float, CultureInfo.InvariantCulture, out var value)) return null;
        return text[i..].ToLowerInvariant() switch
        {
            "" or "pt" => value,
            "in" => value * 72,
            "cm" => value * 72 / 2.54,
            "mm" => value * 72 / 25.4,
            "px" => value * 0.75,
            "pc" => value * 12,
            _ => null,
        };
    }

    private static Color? VmlColor(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        text = text.Trim();
        int space = text.IndexOf(' ');
        if (space > 0) text = text[..space];
        if (text.StartsWith('#') && text.Length == 7 && int.TryParse(text[1..], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
            return new Color((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);
        return text.ToLowerInvariant() switch
        {
            "white" => new Color(255, 255, 255), "black" => new Color(0, 0, 0), "red" => new Color(255, 0, 0), "green" => new Color(0, 128, 0),
            "blue" => new Color(0, 0, 255), "yellow" => new Color(255, 255, 0), "silver" => new Color(192, 192, 192), "gray" or "grey" => new Color(128, 128, 128),
            _ => null,
        };
    }
}
