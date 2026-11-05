using System.Globalization;
using System.Xml.Linq;
using DotOffice.Core;
using static DotOffice.Xlsx.DrawingMl;

namespace DotOffice.Xlsx;

/// <summary>Reads the pictures and shapes of a worksheet's drawing part (SpreadsheetML Drawing, <c>xdr:wsDr</c>).</summary>
internal static class XlsxDrawings
{
    private const double EmuPerPoint = 12700;

    /// <summary>What a drawing contained that is not rendered, and what is rendered only approximately.</summary>
    internal sealed record Skipped(int Charts, int Shapes, int Other, IReadOnlyList<string> Notes, IReadOnlyList<string>? ChartKinds = null)
    {
        public bool Any => Charts + Shapes + Other > 0;
    }

    /// <summary>
    /// Adds the pictures and shapes of the drawing part to <paramref name="sheet"/> and reports what was left out. Objects that are hidden or
    /// excluded from printing (<c>fPrintsWithSheet="0"</c>, as Excel's "Print object" option writes it) are not printed.
    /// </summary>
    public static Skipped Read(XlsxContext context, Worksheet sheet, string drawingPart)
    {
        XDocument doc;
        using (var reader = context.Package.OpenXml(drawingPart)) doc = XDocument.Load(reader);
        if (doc.Root is not { } root || root.Name.LocalName != "wsDr") return new Skipped(0, 0, 1, []);
        var state = new State(context, sheet, context.Package.GetRelationships(drawingPart));

        foreach (var anchor in root.Elements())
        {
            var kind = anchor.Name.LocalName switch
            {
                "twoCellAnchor" => ImageAnchorKind.TwoCell,
                "oneCellAnchor" => ImageAnchorKind.OneCell,
                "absoluteAnchor" => ImageAnchorKind.Absolute,
                _ => (ImageAnchorKind?)null,
            };
            if (kind is null) continue;
            if ((string?)Child(anchor, "clientData")?.Attribute("fPrintsWithSheet") is "0" or "false") continue;

            var frame = new AnchorFrame(kind.Value, Point(Child(anchor, "from")), Point(Child(anchor, "to")), Number(Child(anchor, "pos"), "x") / EmuPerPoint,
                Number(Child(anchor, "pos"), "y") / EmuPerPoint, Number(Child(anchor, "ext"), "cx") / EmuPerPoint, Number(Child(anchor, "ext"), "cy") / EmuPerPoint,
                Child(anchor, "ext") is not null);
            foreach (var content in anchor.Elements())
            {
                switch (content.Name.LocalName)
                {
                    case "pic" or "sp" or "cxnSp" or "grpSp":
                        ReadObject(state, frame, FrameFraction.Whole, content);
                        break;
                    case "graphicFrame": ReadChart(state, frame, FrameFraction.Whole, content); break;
                    case "from" or "to" or "pos" or "ext" or "clientData": break;
                    default: state.Other++; break; // AlternateContent (slicers, …), contentPart
                }
            }
        }
        return new Skipped(state.Charts, state.Shapes, state.Other, state.Notes, state.ChartKinds);
    }

    /// <summary>The anchor of a top-level object: where its frame is, and whether the anchor itself states the size (<c>ext</c>).</summary>
    private sealed record AnchorFrame(ImageAnchorKind Kind, AnchorPoint From, AnchorPoint To, double X, double Y, double Width, double Height, bool HasExtent);

    private sealed class State(XlsxContext context, Worksheet sheet, IReadOnlyDictionary<string, Opc.Relationship> rels)
    {
        public XlsxContext Context { get; } = context;
        public Worksheet Sheet { get; } = sheet;
        public IReadOnlyDictionary<string, Opc.Relationship> Relationships { get; } = rels;
        public int Charts, Shapes, Other, ZOrder;
        public List<string> Notes { get; } = [];
        public List<string> ChartKinds { get; } = [];

        public void Note(string note)
        {
            if (!Notes.Contains(note)) Notes.Add(note);
        }
    }

    private static void ReadObject(State state, AnchorFrame anchor, FrameFraction frame, XElement element)
    {
        switch (element.Name.LocalName)
        {
            case "pic": ReadPicture(state, anchor, frame, element); break;
            case "sp" or "cxnSp": ReadShape(state, anchor, frame, element); break;
            case "grpSp": ReadGroup(state, anchor, frame, element); break;
            case "graphicFrame": ReadChart(state, anchor, frame, element); break;
        }
    }

    /// <summary>Reads a chart frame; charts of kinds that are not drawn are counted and named for the warning.</summary>
    private static void ReadChart(State state, AnchorFrame anchor, FrameFraction frame, XElement graphicFrame)
    {
        var context = state.Context;
        if (Child(Child(graphicFrame, "nvGraphicFramePr"), "cNvPr") is { } p && (string?)p.Attribute("hidden") is "1" or "true") return;
        var chartElement = graphicFrame.Descendants().FirstOrDefault(e => e.Name.LocalName == "chart" && e.Attributes().Any(a => a.Name.LocalName == "id"));
        var id = chartElement?.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value;
        if (id is null || !state.Relationships.TryGetValue(id, out var rel) || rel.IsExternal || !context.Package.Exists(rel.Target))
        {
            // Not a chart (a table or a diagram) or a chart whose part is missing.
            state.Charts++;
            return;
        }
        XDocument doc;
        try
        {
            using var reader = context.Package.OpenXml(rel.Target);
            doc = XDocument.Load(reader);
        }
        catch (System.Xml.XmlException)
        {
            state.Charts++;
            return;
        }
        var (kind, from, to, x, y, width, height) = Placement(anchor, Child(graphicFrame, "xfrm"), frame);
        var template = new SheetChart { Anchor = kind, From = from, To = to, X = x, Y = y, Width = width, Height = height, Frame = frame, ZOrder = state.ZOrder++ };
        var (unsupported, build) = XlsxCharts.Prepare(context, state.Sheet, doc, template, state.Notes);
        if (build is null)
        {
            state.Charts++;
            state.ChartKinds.Add(unsupported ?? "unknown");
            return;
        }
        context.DeferChart(state.Sheet, build);
    }

    /// <summary>
    /// Reads a group by placing its children inside the group's frame: a child's own coordinates are mapped from the group's child space
    /// (<c>chOff</c>, <c>chExt</c>) onto the group's frame. A rotated group's children are drawn unrotated.
    /// </summary>
    private static void ReadGroup(State state, AnchorFrame anchor, FrameFraction frame, XElement group)
    {
        if (Child(Child(group, "nvGrpSpPr"), "cNvPr") is { } p && (string?)p.Attribute("hidden") is "1" or "true") return;
        var xfrm = Child(Child(group, "grpSpPr"), "xfrm");
        if (Number(xfrm, "rot") != 0) state.Note("a rotated group is drawn unrotated");
        var childOffset = Child(xfrm, "chOff");
        var childExtent = Child(xfrm, "chExt");
        double cx = Number(childExtent, "cx"), cy = Number(childExtent, "cy");
        foreach (var child in group.Elements())
        {
            if (child.Name.LocalName is not ("pic" or "sp" or "cxnSp" or "grpSp")) continue;
            var childXfrm = child.Name.LocalName switch
            {
                "pic" => Child(Child(child, "spPr"), "xfrm"),
                "grpSp" => Child(Child(child, "grpSpPr"), "xfrm"),
                _ => Child(Child(child, "spPr"), "xfrm"),
            };
            var fraction = cx > 0 && cy > 0 && Child(childXfrm, "off") is { } off && Child(childXfrm, "ext") is { } ext
                ? new FrameFraction((Number(off, "x") - Number(childOffset, "x")) / cx, (Number(off, "y") - Number(childOffset, "y")) / cy, Number(ext, "cx") / cx, Number(ext, "cy") / cy)
                : FrameFraction.Whole;
            ReadObject(state, anchor, frame.Within(fraction), child);
        }
    }

    /// <summary>The size and position fields every object takes from its anchor.</summary>
    private static (ImageAnchorKind Kind, AnchorPoint From, AnchorPoint To, double X, double Y, double Width, double Height) Placement(AnchorFrame anchor, XElement? xfrm, FrameFraction frame)
    {
        // A top-level one-cell or absolute anchor without its own size takes it from the object's xfrm; inside groups the anchor decides.
        double width = anchor.HasExtent ? anchor.Width : Number(Child(xfrm, "ext"), "cx") / EmuPerPoint;
        double height = anchor.HasExtent ? anchor.Height : Number(Child(xfrm, "ext"), "cy") / EmuPerPoint;
        return (anchor.Kind, anchor.From, anchor.To, anchor.X, anchor.Y, width, height);
    }

    private static void ReadPicture(State state, AnchorFrame anchor, FrameFraction frame, XElement pic)
    {
        var context = state.Context;
        var sheet = state.Sheet;
        var properties = Child(Child(pic, "nvPicPr"), "cNvPr");
        if ((string?)properties?.Attribute("hidden") is "1" or "true") return;

        var blip = Child(Child(pic, "blipFill"), "blip");
        var embed = blip?.Attributes().FirstOrDefault(a => a.Name.LocalName == "embed")?.Value;
        if (embed is null || !state.Relationships.TryGetValue(embed, out var rel) || rel.IsExternal || !context.Package.Exists(rel.Target))
        {
            context.Options.Warnings.Warn("xlsx.image-missing", $"Sheet '{sheet.Name}' has a picture whose image is missing or linked from outside the file; it was skipped.");
            return;
        }
        ImageData image;
        try { image = context.GetImage(rel.Target); }
        catch (InvalidDataException e)
        {
            context.Options.Warnings.Warn("xlsx.image-unreadable", $"Sheet '{sheet.Name}' has a picture that could not be read: {e.Message}");
            return;
        }

        var src = Child(Child(pic, "blipFill"), "srcRect");
        double Crop(string name) => Math.Clamp(Number(src, name) / 100_000.0, 0, 1);
        var xfrm = Child(Child(pic, "spPr"), "xfrm");
        bool flipH = Flag(xfrm, "flipH"), flipV = Flag(xfrm, "flipV");
        double rotation = -Number(xfrm, "rot") / 60_000.0; // clockwise in the file, counter-clockwise in the model

        var (kind, from, to, x, y, width, height) = Placement(anchor, xfrm, frame);
        sheet.Images.Add(new SheetImage
        {
            Image = image,
            Anchor = kind, From = from, To = to, X = x, Y = y, Width = width, Height = height, Frame = frame,
            CropLeft = Crop("l"), CropTop = Crop("t"), CropRight = Crop("r"), CropBottom = Crop("b"),
            Rotation = rotation,
            FlipH = flipH, FlipV = flipV,
            ZOrder = state.ZOrder++,
        });
    }

    // ------------------------------------------------------------------ shapes

    private static readonly double[] StyleLineWidths = [0, 0.5, 1, 1.5];

    private static ShapeGeometry? GeometryOf(string? preset) => preset switch
    {
        null or "rect" or "flowChartProcess" or "snip1Rect" or "snip2SameRect" => ShapeGeometry.Rectangle,
        "roundRect" => ShapeGeometry.RoundRectangle,
        "ellipse" or "flowChartConnector" => ShapeGeometry.Ellipse,
        "triangle" => ShapeGeometry.Triangle,
        "rtTriangle" => ShapeGeometry.RightTriangle,
        "diamond" or "flowChartDecision" => ShapeGeometry.Diamond,
        "parallelogram" => ShapeGeometry.Parallelogram,
        "trapezoid" => ShapeGeometry.Trapezoid,
        "hexagon" => ShapeGeometry.Hexagon,
        "rightArrow" => ShapeGeometry.RightArrow,
        "leftArrow" => ShapeGeometry.LeftArrow,
        "upArrow" => ShapeGeometry.UpArrow,
        "downArrow" => ShapeGeometry.DownArrow,
        "line" or "straightConnector1" => ShapeGeometry.Line,
        _ => null,
    };

    private static void ReadShape(State state, AnchorFrame anchor, FrameFraction frame, XElement shape)
    {
        var context = state.Context;
        bool connector = shape.Name.LocalName == "cxnSp";
        var properties = Child(Child(shape, connector ? "nvCxnSpPr" : "nvSpPr"), "cNvPr");
        if ((string?)properties?.Attribute("hidden") is "1" or "true") return;
        var spPr = Child(shape, "spPr");
        var xfrm = Child(spPr, "xfrm");

        IReadOnlyList<ShapePath>? paths = null;
        ShapeGeometry? geometry;
        var preset = (string?)Child(spPr, "prstGeom")?.Attribute("prst");
        if (Child(spPr, "custGeom") is { } custom)
        {
            paths = ReadCustomGeometry(custom, Number(Child(xfrm, "ext"), "cx"), Number(Child(xfrm, "ext"), "cy"));
            if (paths is null)
            {
                state.Shapes++; // paths that use guide formulas are not evaluated
                return;
            }
            geometry = ShapeGeometry.Custom;
        }
        else if (preset is not null && ConnectorPath(preset, Child(spPr, "prstGeom")) is { } route)
        {
            paths = [new ShapePath(route, false, Fill: false)];
            geometry = ShapeGeometry.Line;
            if (preset is "curvedConnector4" or "curvedConnector5") state.Note("curved connectors with four or five segments are drawn approximately");
        }
        else
        {
            geometry = GeometryOf(preset);
            if (geometry is null)
            {
                state.Note($"shapes of the geometry '{preset}' are drawn as rectangles");
                geometry = connector ? ShapeGeometry.Line : ShapeGeometry.Rectangle;
            }
            else if (connector) geometry = ShapeGeometry.Line;
        }

        // The theme's style references stand in for what the shape does not state itself.
        var style = Child(shape, "style");
        var theme = context.Theme;
        Color? StyleColor(string name) => Resolve(Child(style, name), theme);
        var fillRef = Child(style, "fillRef");
        var lineRef = Child(style, "lnRef");
        bool lineOnly = geometry == ShapeGeometry.Line;
        Color? fill = lineOnly ? null : SolidFill(spPr, theme, Number(fillRef, "idx") > 0 ? StyleColor("fillRef") : null);
        GradientFill? gradient = lineOnly ? null : ReadGradientFill(spPr, theme, state);

        var ln = Child(spPr, "ln");
        Color? lineColor;
        double lineWidth;
        StrokeStyle lineStyle = StrokeStyle.Solid;
        if (ln is not null && ln.Elements().Any(e => e.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "pattFill"))
            lineColor = SolidFill(ln, theme);
        else lineColor = Number(lineRef, "idx") > 0 ? StyleColor("lnRef") : null;
        int lineIdx = (int)Number(lineRef, "idx");
        lineWidth = ln?.Attribute("w") is { } w && double.TryParse(w.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var emu) ? emu / EmuPerPoint
            : lineIdx is >= 1 and <= 3 ? StyleLineWidths[lineIdx] : 0.75;
        if ((string?)Child(ln, "prstDash")?.Attribute("val") is { } dash && dash != "solid")
            lineStyle = dash is "sysDot" or "dot" or "sysDash" ? StrokeStyle.Dotted : StrokeStyle.Dashed;
        if (geometry == ShapeGeometry.Line && lineColor is null && ln is null) lineColor = Color.Black;
        LineEnd End(string name)
        {
            var e = Child(ln, name);
            var kind = (string?)e?.Attribute("type") switch { "triangle" => LineEndKind.Triangle, "stealth" => LineEndKind.Stealth, "arrow" => LineEndKind.Arrow, "oval" => LineEndKind.Oval, "diamond" => LineEndKind.Diamond, _ => LineEndKind.None };
            double Size(string attribute) => (string?)e?.Attribute(attribute) switch { "sm" => 2, "lg" => 5, _ => 3 };
            return new LineEnd(kind, Size("w"), Size("len"));
        }

        // Text: the default color is the style's font color (white on the theme's default shapes).
        var fontColor = Resolve(Child(style, "fontRef"), theme);
        var paragraphs = ReadText(Child(shape, "txBody"), theme, fontColor, out var body);

        if (fill is null && lineColor is null && paragraphs.Count == 0) return;
        var (kind, from, to, x, y, width, height) = Placement(anchor, xfrm, frame);
        double Inset(string name, double fallback) => (string?)body?.Attribute(name) is { } v && double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var e) ? e / EmuPerPoint : fallback;
        state.Sheet.Shapes.Add(new SheetShape
        {
            Anchor = kind, From = from, To = to, X = x, Y = y, Width = width, Height = height, Frame = frame,
            Rotation = -Number(xfrm, "rot") / 60_000.0,
            FlipH = Flag(xfrm, "flipH"), FlipV = Flag(xfrm, "flipV"),
            ZOrder = state.ZOrder++,
            Geometry = geometry.Value,
            Fill = fill,
            FillGradient = gradient,
            Paths = paths,
            HeadEnd = End("headEnd"), TailEnd = End("tailEnd"),
            LineColor = lineColor,
            LineWidth = lineWidth,
            LineStyle = lineStyle,
            Paragraphs = paragraphs,
            TextAnchor = (string?)body?.Attribute("anchor") switch { "ctr" => ShapeTextAnchor.Middle, "b" => ShapeTextAnchor.Bottom, _ => ShapeTextAnchor.Top },
            TextInsets = new Thickness(Inset("tIns", 3.6), Inset("rIns", 7.2), Inset("bIns", 3.6), Inset("lIns", 7.2)),
            WrapText = (string?)body?.Attribute("wrap") != "none",
        });
    }

    /// <summary>
    /// A gradient fill of a shape: the stops, and either the angle of a linear gradient or the focus of a path gradient (which is drawn
    /// radially). Null when the fill is not a gradient.
    /// </summary>
    private static GradientFill? ReadGradientFill(XElement? spPr, XlsxTheme theme, State state)
    {
        var first = spPr?.Elements().FirstOrDefault(e => e.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "pattFill" or "blipFill");
        if (first is null || first.Name.LocalName != "gradFill") return null;
        var stops = new List<GradientStop>();
        foreach (var gs in Child(first, "gsLst")?.Elements().Where(e => e.Name.LocalName == "gs") ?? [])
            if (Resolve(gs, theme) is { } color) stops.Add(new GradientStop(Math.Clamp(Number(gs, "pos") / 100_000.0, 0, 1), color));
        if (stops.Count < 2) return null;
        stops = stops.OrderBy(g => g.Position).ToList();
        var path = Child(first, "path");
        if (path is not null)
        {
            if ((string?)path.Attribute("path") == "rect") state.Note("rectangular path gradients are drawn as radial ones");
            var focus = Child(path, "fillToRect");
            double Edge(string name, double fallback) => focus?.Attribute(name) is null ? fallback : Number(focus, name) / 100_000.0;
            double l = Edge("l", 0.5), t = Edge("t", 0.5), r = Edge("r", 0.5), b = Edge("b", 0.5);
            return new GradientFill { Kind = GradientKind.Path, Left = l, Right = 1 - r, Top = t, Bottom = 1 - b, Stops = stops };
        }
        double degree = Child(first, "lin")?.Attribute("ang") is not null ? Number(Child(first, "lin"), "ang") / 60_000.0 : 0;
        return new GradientFill { Kind = GradientKind.Linear, Degree = degree, Stops = stops };
    }

    /// <summary>The route of a bent or curved connector as points in fractions of its frame; null for any other geometry.</summary>
    private static List<(double X, double Y)>? ConnectorPath(string preset, XElement? prstGeom)
    {
        double Adjust(string name, double fallback)
        {
            var gd = prstGeom?.Elements().FirstOrDefault(e => e.Name.LocalName == "avLst")?.Elements().FirstOrDefault(e => (string?)e.Attribute("name") == name);
            return gd?.Attribute("fmla")?.Value is { } f && f.StartsWith("val ", StringComparison.Ordinal) && double.TryParse(f.AsSpan(4), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 100_000.0 : fallback;
        }
        static void Curve(List<(double X, double Y)> points, (double X, double Y) p0, (double X, double Y) c1, (double X, double Y) c2, (double X, double Y) p3)
        {
            for (int k = 1; k <= 16; k++)
            {
                double t = k / 16.0, u = 1 - t;
                points.Add((u * u * u * p0.X + 3 * u * u * t * c1.X + 3 * u * t * t * c2.X + t * t * t * p3.X,
                            u * u * u * p0.Y + 3 * u * u * t * c1.Y + 3 * u * t * t * c2.Y + t * t * t * p3.Y));
            }
        }
        switch (preset)
        {
            case "bentConnector2": return [(0, 0), (1, 0), (1, 1)];
            case "bentConnector3": { double x1 = Adjust("adj1", 0.5); return [(0, 0), (x1, 0), (x1, 1), (1, 1)]; }
            case "bentConnector4": { double x1 = Adjust("adj1", 0.5), y2 = Adjust("adj2", 0.5); return [(0, 0), (x1, 0), (x1, y2), (1, y2), (1, 1)]; }
            case "bentConnector5": { double x1 = Adjust("adj1", 0.5), y2 = Adjust("adj2", 0.5), x3 = Adjust("adj3", 0.5); return [(0, 0), (x1, 0), (x1, y2), (x3, y2), (x3, 1), (1, 1)]; }
            case "curvedConnector2":
            {
                var points = new List<(double X, double Y)> { (0, 0) };
                Curve(points, (0, 0), (0.5, 0), (1, 0.5), (1, 1));
                return points;
            }
            case "curvedConnector3":
            {
                double x2 = Adjust("adj1", 0.5);
                var points = new List<(double X, double Y)> { (0, 0) };
                Curve(points, (0, 0), (x2 / 2, 0), (x2, 0.25), (x2, 0.5));
                Curve(points, (x2, 0.5), (x2, 0.75), ((1 + x2) / 2, 1), (1, 1));
                return points;
            }
            case "curvedConnector4" or "curvedConnector5":
            {
                // Rounded through the corners of the matching bent connector.
                var corners = ConnectorPath(preset.Replace("curved", "bent"), prstGeom)!;
                var points = new List<(double X, double Y)> { corners[0] };
                for (int i = 0; i + 1 < corners.Count; i++)
                {
                    var p0 = corners[Math.Max(i - 1, 0)]; var p1 = corners[i]; var p2 = corners[i + 1]; var p3 = corners[Math.Min(i + 2, corners.Count - 1)];
                    Curve(points, p1, (p1.X + (p2.X - p0.X) / 6, p1.Y + (p2.Y - p0.Y) / 6), (p2.X - (p3.X - p1.X) / 6, p2.Y - (p3.Y - p1.Y) / 6), p2);
                }
                return points;
            }
            default: return null;
        }
    }

    /// <summary>
    /// The paths of a custom geometry as polylines in fractions of the shape's frame. Null when the geometry uses guide formulas rather
    /// than plain numbers, or has nothing to draw.
    /// </summary>
    private static List<ShapePath>? ReadCustomGeometry(XElement custGeom, double widthEmu, double heightEmu)
    {
        var result = new List<ShapePath>();
        foreach (var path in Child(custGeom, "pathLst")?.Elements().Where(e => e.Name.LocalName == "path") ?? [])
        {
            double w = Number(path, "w"), h = Number(path, "h");
            if (w <= 0) w = widthEmu;
            if (h <= 0) h = heightEmu;
            if (w <= 0 || h <= 0) return null;
            bool valid = true;
            (double X, double Y) Pt(XElement? pt)
            {
                if (!double.TryParse((string?)pt?.Attribute("x"), NumberStyles.Float, CultureInfo.InvariantCulture, out var x) ||
                    !double.TryParse((string?)pt?.Attribute("y"), NumberStyles.Float, CultureInfo.InvariantCulture, out var y)) { valid = false; return (0, 0); }
                return (x / w, y / h);
            }
            bool fill = (string?)path.Attribute("fill") != "none", stroke = (string?)path.Attribute("stroke") is not ("0" or "false");
            var points = new List<(double X, double Y)>();
            void Flush(bool closed)
            {
                if (points.Count >= 2) result.Add(new ShapePath(points.ToList(), closed, fill, stroke));
                points.Clear();
            }
            foreach (var command in path.Elements())
            {
                var pts = command.Elements().Where(e => e.Name.LocalName == "pt").ToList();
                switch (command.Name.LocalName)
                {
                    case "moveTo": Flush(false); points.Add(Pt(pts.FirstOrDefault())); break;
                    case "lnTo": points.Add(Pt(pts.FirstOrDefault())); break;
                    case "cubicBezTo" when pts.Count == 3 && points.Count > 0:
                    {
                        var (p0, c1, c2, p3) = (points[^1], Pt(pts[0]), Pt(pts[1]), Pt(pts[2]));
                        for (int k = 1; k <= 16; k++)
                        {
                            double t = k / 16.0, u = 1 - t;
                            points.Add((u * u * u * p0.X + 3 * u * u * t * c1.X + 3 * u * t * t * c2.X + t * t * t * p3.X, u * u * u * p0.Y + 3 * u * u * t * c1.Y + 3 * u * t * t * c2.Y + t * t * t * p3.Y));
                        }
                        break;
                    }
                    case "quadBezTo" when pts.Count == 2 && points.Count > 0:
                    {
                        var (p0, c, p2) = (points[^1], Pt(pts[0]), Pt(pts[1]));
                        for (int k = 1; k <= 12; k++)
                        {
                            double t = k / 12.0, u = 1 - t;
                            points.Add((u * u * p0.X + 2 * u * t * c.X + t * t * p2.X, u * u * p0.Y + 2 * u * t * c.Y + t * t * p2.Y));
                        }
                        break;
                    }
                    case "arcTo" when points.Count > 0:
                    {
                        // An arc of an ellipse with radii wR and hR from the current point, starting at stAng and sweeping swAng.
                        double wr = Number(command, "wR") / w, hr = Number(command, "hR") / h, start = Number(command, "stAng") / 60_000.0 * Math.PI / 180, sweep = Number(command, "swAng") / 60_000.0 * Math.PI / 180;
                        if (command.Attribute("wR") is null || command.Attribute("hR") is null || command.Attribute("stAng") is null || command.Attribute("swAng") is null) { valid = false; break; }
                        var current = points[^1];
                        double cx = current.X - wr * Math.Cos(start), cy = current.Y - hr * Math.Sin(start);
                        int steps = Math.Max(4, (int)Math.Ceiling(Math.Abs(sweep) * 180 / Math.PI / 6));
                        for (int k = 1; k <= steps; k++)
                        {
                            double angle = start + sweep * k / steps;
                            points.Add((cx + wr * Math.Cos(angle), cy + hr * Math.Sin(angle)));
                        }
                        break;
                    }
                    case "close": Flush(true); break;
                    default: valid = false; break;
                }
                if (!valid) return null;
            }
            Flush(false);
        }
        return result.Count > 0 ? result : null;
    }

    private static List<ShapeParagraph> ReadText(XElement? txBody, XlsxTheme theme, Color? defaultColor, out XElement? bodyProperties)
    {
        bodyProperties = Child(txBody, "bodyPr");
        var result = new List<ShapeParagraph>();
        if (txBody is null) return result;
        var baseFont = new RunProperties { FontFamily = "Calibri", FontSize = 11, Color = defaultColor };
        foreach (var p in txBody.Elements().Where(e => e.Name.LocalName == "p"))
        {
            RunProperties Font(XElement? rPr)
            {
                var font = baseFont;
                if (rPr is null) return font;
                if (rPr.Attribute("sz") is { } sz && double.TryParse(sz.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) && size > 0) font = font with { FontSize = size / 100 };
                if (rPr.Attribute("b") is { } b) font = font with { Bold = b.Value is "1" or "true" };
                if (rPr.Attribute("i") is { } i) font = font with { Italic = i.Value is "1" or "true" };
                if ((string?)rPr.Attribute("u") is { } u) font = font with { Underline = u == "none" ? UnderlineStyle.None : u.StartsWith("dbl", StringComparison.Ordinal) ? UnderlineStyle.Double : UnderlineStyle.Single };
                if ((string?)rPr.Attribute("strike") is { } strike) font = font with { Strike = strike != "noStrike" };
                if ((string?)Child(rPr, "latin")?.Attribute("typeface") is { Length: > 0 } face && !face.StartsWith('+')) font = font with { FontFamily = face };
                if (Child(rPr, "solidFill") is { } fillElement && Resolve(fillElement, theme) is { } color) font = font with { Color = color };
                return font;
            }

            var runs = new List<CellTextRun>();
            foreach (var child in p.Elements())
            {
                switch (child.Name.LocalName)
                {
                    case "r": runs.Add(new CellTextRun(Child(child, "t")?.Value ?? "", Font(Child(child, "rPr")))); break;
                    case "br": runs.Add(new CellTextRun("\n", Font(Child(child, "rPr")))); break;
                    case "fld": runs.Add(new CellTextRun(Child(child, "t")?.Value ?? "", Font(Child(child, "rPr")))); break;
                }
            }
            var alignment = (string?)Child(p, "pPr")?.Attribute("algn") switch
            {
                "ctr" => HorizontalAlignment.Center, "r" => HorizontalAlignment.Right, "just" or "dist" => HorizontalAlignment.Justify, _ => HorizontalAlignment.Left,
            };
            double emptySize = runs.Count > 0 ? runs.Max(r => r.Font.FontSize) : Font(Child(p, "endParaRPr")).FontSize;
            result.Add(new ShapeParagraph(alignment, runs, emptySize));
        }
        // A text body of empty paragraphs only carries no text.
        if (result.All(p => p.Runs.Count == 0 || p.Runs.All(r => r.Text.Length == 0))) result.Clear();
        return result;
    }

    private static AnchorPoint Point(XElement? e) => e is null
        ? default
        : new AnchorPoint(
            (int)Math.Clamp(Number(Child(e, "col"), null), 0, CellReference.MaxColumns - 1), (int)Math.Clamp(Number(Child(e, "row"), null), 0, CellReference.MaxRows - 1),
            Number(Child(e, "colOff"), null) / EmuPerPoint, Number(Child(e, "rowOff"), null) / EmuPerPoint);

    /// <summary>An element's text as a number; 0 when absent or malformed.</summary>
    private static double Number(XElement? e, string? attribute)
    {
        var text = attribute is null ? e?.Value : (string?)e?.Attribute(attribute);
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) ? v : 0;
    }
}
