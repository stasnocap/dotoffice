using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>Word drawing shapes (<c>wps:wsp</c>): text boxes and plain filled or outlined rectangles.</summary>
internal sealed partial class StoryReader
{
    private const string WordprocessingShape = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    private sealed class ShapeData
    {
        public Color? Fill { get; set; }
        public Color? Outline { get; set; }
        public double OutlineWidth { get; set; } = 0.75;
        public LineStyleKind Dash { get; set; }
        public Thickness Insets { get; set; } = new(3.6, 7.2, 3.6, 7.2);
        public TextBoxAnchor Anchor { get; set; }
        public double Rotation { get; set; }
        public string Geometry { get; set; } = "rect";
        public double CornerRadius { get; set; } = 1.0 / 6;
        public double? Adjust1 { get; set; }
        public double? Adjust2 { get; set; }
        public GradientFill? Gradient { get; set; }
        public TextBoxDirection Direction { get; set; }
        public bool UnsupportedVertical { get; set; }
        public List<Block> Blocks { get; } = [];
    }

    /// <summary>Reads the <c>a:graphicData</c> of a wps shape (the reader sits on it) and consumes it.</summary>
    private ShapeData ReadShapeData(XmlReader r)
    {
        var data = new ShapeData();
        int depth = r.Depth;
        if (!r.Enter()) return data;
        while (r.NextChild(depth))
        {
            if (r.NamespaceURI == WordprocessingShape && r.LocalName == "wsp") ReadShape(r, data);
            else r.Skip();
        }
        return data;
    }

    private void ReadShape(XmlReader r, ShapeData data)
    {
        XElement? properties = null, style = null;
        int depth = r.Depth;
        if (!r.Enter()) return;
        while (r.NextChild(depth))
        {
            if (r.NamespaceURI == WordprocessingShape)
            {
                switch (r.LocalName)
                {
                    case "txbx":
                        ReadContainer(r, "txbxContent", () => ReadBlocks(r, data.Blocks, null, 0, null));
                        continue;
                    case "bodyPr":
                        ReadBodyProperties(r, data);
                        r.Skip();
                        continue;
                    case "spPr":
                        properties = (XElement)XNode.ReadFrom(r);
                        continue;
                    case "style":
                        style = (XElement)XNode.ReadFrom(r);
                        continue;
                }
            }
            r.Skip();
        }
        ApplyShapeProperties(data, properties, style);
    }

    private static void ReadBodyProperties(XmlReader r, ShapeData data)
    {
        static double Inset(XmlReader r, string name, double fallback) =>
            long.TryParse(r.GetAttribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? Math.Clamp(Units.FromEmu(v), 0, 1584) : fallback;
        var d = data.Insets;
        data.Insets = new Thickness(Inset(r, "tIns", d.Top), Inset(r, "rIns", d.Right), Inset(r, "bIns", d.Bottom), Inset(r, "lIns", d.Left));
        data.Anchor = r.GetAttribute("anchor") switch { "ctr" => TextBoxAnchor.Center, "b" => TextBoxAnchor.Bottom, _ => TextBoxAnchor.Top };
        data.Direction = r.GetAttribute("vert") switch { "vert" or "eaVert" => TextBoxDirection.Clockwise, "vert270" => TextBoxDirection.CounterClockwise, _ => TextBoxDirection.Horizontal };
        data.UnsupportedVertical = r.GetAttribute("vert") is "wordArtVert" or "wordArtVertRtl" or "mongolianVert";
    }

    private void ApplyShapeProperties(ShapeData data, XElement? properties, XElement? style)
    {
        static XElement? Child(XElement? e, string name) => e?.Elements().FirstOrDefault(c => c.Name.LocalName == name);

        if (properties is not null)
        {
            if (Child(properties, "xfrm") is { } xfrm && double.TryParse((string?)xfrm.Attribute("rot"), NumberStyles.Float, CultureInfo.InvariantCulture, out var rot))
                data.Rotation = rot / 60000.0;
            if (Child(properties, "prstGeom") is { } geometry)
            {
                data.Geometry = (string?)geometry.Attribute("prst") ?? "rect";
                data.CornerRadius = data.Geometry == "flowChartTerminator" ? 0.5 : AdjustOf(geometry, "adj") ?? 1.0 / 6;
                data.CornerRadius = Math.Clamp(data.CornerRadius, 0, 0.5);
                data.Adjust1 = AdjustOf(geometry, "adj1") ?? AdjustOf(geometry, "adj");
                data.Adjust2 = AdjustOf(geometry, "adj2");
            }
            else if (Child(properties, "custGeom") is not null) data.Geometry = "custom";
        }

        // A fill or line the shape properties do not give comes from the style's references (index 0 means none).
        Color? Referenced(string name) =>
            Child(style, name) is { } reference && (string?)reference.Attribute("idx") is { } idx && idx != "0" ? ResolveDrawingColor(reference) : null;

        var fill = properties?.Elements().FirstOrDefault(e => e.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "blipFill" or "pattFill" or "grpFill");
        switch (fill?.Name.LocalName)
        {
            case null: data.Fill = Referenced("fillRef"); break;
            case "noFill": break;
            case "solidFill": data.Fill = ResolveDrawingColor(fill); break;
            case "gradFill":
                data.Gradient = ReadGradient(fill);
                data.Fill = data.Gradient?.Stops[0].Color ?? ResolveDrawingColor(fill.Elements().FirstOrDefault(e => e.Name.LocalName == "gsLst")?.Elements().FirstOrDefault());
                break;
            default:
                context.WarnOnce("docx.shape.fill", "Picture and pattern fills of shapes are drawn as a solid color or not at all.");
                break;
        }

        var line = Child(properties, "ln");
        if (line is null) data.Outline = Referenced("lnRef");
        else if (line.Elements().Any(e => e.Name.LocalName == "noFill")) data.Outline = null;
        else
        {
            data.Outline = Child(line, "solidFill") is { } lineFill ? ResolveDrawingColor(lineFill) : Referenced("lnRef");
            if (long.TryParse((string?)line.Attribute("w"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var w))
                data.OutlineWidth = Math.Clamp(Units.FromEmu(w), 0.25, 100);
            data.Dash = (string?)Child(line, "prstDash")?.Attribute("val") switch
            {
                "dot" or "sysDot" => LineStyleKind.Dotted,
                "dash" or "sysDash" or "lgDash" or "dashDot" or "lgDashDot" or "lgDashDotDot" or "sysDashDot" or "sysDashDotDot" => LineStyleKind.Dashed,
                _ => LineStyleKind.Solid,
            };
        }
    }

    /// <summary>A preset's adjust value (<c>val N</c> over 100000) from its <c>avLst</c>, or null when it states none.</summary>
    private static double? AdjustOf(XElement geometry, string name)
    {
        var gd = geometry.Elements().FirstOrDefault(e => e.Name.LocalName == "avLst")?.Elements().FirstOrDefault(e => (string?)e.Attribute("name") == name);
        return gd?.Attribute("fmla")?.Value is { } f && f.StartsWith("val ", StringComparison.Ordinal)
               && double.TryParse(f.AsSpan(4), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 100_000.0 : null;
    }

    private static ShapeGeometry? PresetGeometry(string preset) => preset switch
    {
        "rect" or "flowChartProcess" or "textBox" => ShapeGeometry.Rectangle,
        "roundRect" or "flowChartAlternateProcess" or "flowChartTerminator" => ShapeGeometry.RoundRectangle,
        "ellipse" or "flowChartConnector" => ShapeGeometry.Ellipse,
        "triangle" => ShapeGeometry.Triangle,
        "rtTriangle" => ShapeGeometry.RightTriangle,
        "diamond" or "flowChartDecision" => ShapeGeometry.Diamond,
        "parallelogram" or "flowChartInputOutput" => ShapeGeometry.Parallelogram,
        "trapezoid" => ShapeGeometry.Trapezoid,
        "hexagon" or "flowChartPreparation" => ShapeGeometry.Hexagon,
        "pentagon" => ShapeGeometry.Pentagon,
        "octagon" => ShapeGeometry.Octagon,
        "plus" => ShapeGeometry.Plus,
        "chevron" => ShapeGeometry.Chevron,
        "homePlate" => ShapeGeometry.HomePlate,
        "star5" => ShapeGeometry.Star5,
        "rightArrow" => ShapeGeometry.RightArrow,
        "leftArrow" => ShapeGeometry.LeftArrow,
        "upArrow" => ShapeGeometry.UpArrow,
        "downArrow" => ShapeGeometry.DownArrow,
        _ => null,
    };

    /// <summary>A linear or path gradient; null with fewer than two usable stops.</summary>
    private GradientFill? ReadGradient(XElement fill)
    {
        var stops = new List<GradientStop>();
        foreach (var gs in fill.Elements().FirstOrDefault(e => e.Name.LocalName == "gsLst")?.Elements().Where(e => e.Name.LocalName == "gs") ?? [])
            if (ResolveDrawingColor(gs) is { } color
                && double.TryParse((string?)gs.Attribute("pos"), NumberStyles.Float, CultureInfo.InvariantCulture, out var pos))
                stops.Add(new GradientStop(Math.Clamp(pos / 100_000.0, 0, 1), color));
        if (stops.Count < 2) return null;
        stops = [.. stops.OrderBy(s => s.Position)];
        static double Number(XElement? e, string name) =>
            double.TryParse((string?)e?.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : double.NaN;
        var path = fill.Elements().FirstOrDefault(e => e.Name.LocalName == "path");
        if (path is not null)
        {
            var focus = path.Elements().FirstOrDefault(e => e.Name.LocalName == "fillToRect");
            double Edge(string name) => double.IsNaN(Number(focus, name)) ? 0.5 : Number(focus, name) / 100_000.0;
            return new GradientFill
            {
                Kind = GradientKind.Path, Rectangular = (string?)path.Attribute("path") == "rect",
                Left = Edge("l"), Right = 1 - Edge("r"), Top = Edge("t"), Bottom = 1 - Edge("b"), Stops = stops,
            };
        }
        double angle = Number(fill.Elements().FirstOrDefault(e => e.Name.LocalName == "lin"), "ang");
        return new GradientFill { Kind = GradientKind.Linear, Degree = double.IsNaN(angle) ? 0 : angle / 60_000.0, Stops = stops };
    }

    private static readonly Dictionary<string, Color> OfficeColors = new()
    {
        ["dk1"] = new(0, 0, 0), ["lt1"] = new(255, 255, 255), ["dk2"] = new(0x44, 0x54, 0x6A), ["lt2"] = new(0xE7, 0xE6, 0xE6),
        ["accent1"] = new(0x44, 0x72, 0xC4), ["accent2"] = new(0xED, 0x7D, 0x31), ["accent3"] = new(0xA5, 0xA5, 0xA5),
        ["accent4"] = new(0xFF, 0xC0, 0x00), ["accent5"] = new(0x5B, 0x9B, 0xD5), ["accent6"] = new(0x70, 0xAD, 0x47),
        ["hlink"] = new(0x05, 0x63, 0xC1), ["folHlink"] = new(0x95, 0x4F, 0x72),
    };

    /// <summary>The color of the first color element under <paramref name="holder"/> (or of <paramref name="holder"/> itself), with its modifiers applied.</summary>
    private Color? ResolveDrawingColor(XElement? holder)
    {
        if (holder is null) return null;
        var element = holder.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr"
            ? holder
            : holder.Elements().FirstOrDefault(e => e.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr");
        if (element is null) return null;
        string? val = (string?)element.Attribute("val");
        Color? color = element.Name.LocalName switch
        {
            "srgbClr" => Color.TryParseHex(val),
            "sysClr" => Color.TryParseHex((string?)element.Attribute("lastClr")) ?? (val == "window" ? Color.White : Color.Black),
            "prstClr" => val switch
            {
                "black" => Color.Black, "white" => Color.White, "red" => new Color(255, 0, 0), "green" => new Color(0, 128, 0),
                "blue" => new Color(0, 0, 255), "yellow" => new Color(255, 255, 0), "gray" or "grey" => new Color(128, 128, 128), _ => null,
            },
            _ => SchemeColor(val),
        };
        if (color is not { } c) return null;
        foreach (var modifier in element.Elements())
        {
            double v = double.TryParse((string?)modifier.Attribute("val"), NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n / 100_000.0 : 0;
            switch (modifier.Name.LocalName)
            {
                case "lumMod": c = c.WithHsl((h, s, l) => (h, s, l * v)); break;
                case "lumOff": c = c.WithHsl((h, s, l) => (h, s, l + v)); break;
                case "satMod": c = c.WithHsl((h, s, l) => (h, s * v, l)); break;
                case "tint": c = c.Blend(Color.White, 1 - v); break;
                case "shade": c = c.Blend(Color.Black, 1 - v); break;
            }
        }
        return c;
    }

    private Color? SchemeColor(string? name)
    {
        string? key = name switch { "bg1" => "lt1", "tx1" => "dk1", "bg2" => "lt2", "tx2" => "dk2", _ => name };
        if (key is null) return null;
        if (Styles.Theme?.Colors.TryGetValue(key, out var themed) == true) return themed;
        return OfficeColors.TryGetValue(key, out var office) ? office : null;
    }

    /// <summary>Makes the floating box for a shape; null (with warnings) for shapes with nothing this converter draws.</summary>
    private FloatingTextBox? CreateShape(ShapeData data, double width, double height, RunProperties props, FloatPosition horizontal, FloatPosition vertical,
        TextWrap wrap, bool behind, Thickness distance)
    {
        if (width <= 0 || height <= 0) return null;
        width = Math.Min(width, context.Options.MaxDimension);
        height = Math.Min(height, context.Options.MaxDimension);
        var geometry = PresetGeometry(data.Geometry);
        bool hasText = data.Blocks.Count > 0;
        if (data.UnsupportedVertical) context.WarnOnce("docx.shape.vert", "Stacked and Mongolian vertical text in shapes is laid out horizontally.");
        if (geometry is null)
        {
            context.WarnOnce("docx.shape.geometry", $"Shapes other than the basic presets (rectangle, rounded rectangle, ellipse, triangles, diamond, hexagon, arrows, …) are not drawn (only their text, if any); first seen: '{data.Geometry}'.");
            if (!hasText) return null;
        }
        if (!hasText && data.Fill is null && data.Outline is null) return null;
        var box = new FloatingTextBox(width, height, props)
        {
            Fill = geometry is null ? null : data.Fill,
            FillGradient = geometry is null ? null : data.Gradient,
            Outline = geometry is null ? null : data.Outline,
            Geometry = geometry ?? ShapeGeometry.Rectangle,
            CornerRadius = data.CornerRadius, Adjust1 = data.Adjust1, Adjust2 = data.Adjust2,
            OutlineWidth = data.OutlineWidth,
            OutlineStyle = data.Dash,
            Insets = data.Insets,
            VerticalAnchor = data.Anchor,
            TextDirection = data.Direction,
            Rotation = data.Rotation,
            Horizontal = horizontal, Vertical = vertical, Wrap = behind ? TextWrap.None : wrap, BehindText = behind, Distance = distance,
        };
        box.Blocks.AddRange(data.Blocks);
        return box;
    }
}
