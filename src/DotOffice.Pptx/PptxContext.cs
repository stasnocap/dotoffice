using System.Globalization;
using System.Xml.Linq;
using DotOffice.Core;
using DotOffice.Dml;
using DotOffice.Opc;
using static DotOffice.Dml.DrawingMl;

namespace DotOffice.Pptx;

/// <summary>
/// Reads the slides of a presentation with what they inherit from their layout and master: placeholder position, fill, outline, body
/// properties and text styles, the color map, the background, and the shapes of the layout and master that are not placeholders.
/// </summary>
internal sealed partial class PptxContext(OpcPackage package, PptxReaderOptions options)
{
    private const double EmuPerPoint = 12700;

    private static readonly string[] FillKinds = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
    private static readonly string[] BulletKinds = ["buNone", "buChar", "buAutoNum", "buBlip"];
    private static readonly string[] SymbolFonts = ["Wingdings", "Wingdings 2", "Wingdings 3", "Symbol", "Webdings", "Marlett"];

    private readonly Dictionary<string, ImageData> _images = [];
    private readonly Dictionary<string, Master> _masters = [];
    private readonly Dictionary<string, Layout> _layouts = [];
    private readonly List<string> _notes = [];
    private readonly SortedDictionary<string, int> _skipped = new(StringComparer.Ordinal);
    private XElement? _defaultTextStyle;

    private sealed record Master(XElement Root, Host Host, string MajorFont, string MinorFont, IReadOnlyDictionary<string, string> ColorMap);

    private sealed record Layout(XElement Root, Host Host, Master? Master);

    /// <summary>The part an object comes from (for its pictures) and what it inherits from.</summary>
    private sealed record Scope(Host Host, Layout? Layout, Master? Master, IReadOnlyDictionary<string, string> ColorMap, bool IsSlide);

    /// <summary>Where an object goes: the frame of its outermost group (or its own frame), the part of it the object takes, and the turned groups.</summary>
    private sealed record Placement(double X, double Y, double Width, double Height, FrameFraction Frame, GroupTurn[] Turns);

    private sealed class Host(PptxContext context, IReadOnlyDictionary<string, Relationship> relationships, DmlTheme theme) : IDmlHost
    {
        public DmlTheme Theme { get; } = theme;

        /// <summary>The part a relationship of this part points to inside the package; null when it is missing or external.</summary>
        public string? Part(string relationshipId) =>
            relationships.TryGetValue(relationshipId, out var rel) && !rel.IsExternal && context._package.Exists(rel.Target) ? rel.Target : null;

        public ImageData? Image(string relationshipId, out bool unreadable)
        {
            unreadable = false;
            if (!relationships.TryGetValue(relationshipId, out var rel) || rel.IsExternal || !context._package.Exists(rel.Target)) return null;
            try
            {
                if (!context._images.TryGetValue(rel.Target, out var image))
                    context._images[rel.Target] = image = new ImageData(context._package.ReadAllBytes(rel.Target), rel.Target);
                return image;
            }
            catch (InvalidDataException)
            {
                unreadable = true;
                return null;
            }
        }

        public void Note(string note) => context.Note(note);
    }

    private readonly OpcPackage _package = package;

    private void Note(string note)
    {
        if (!_notes.Contains(note)) _notes.Add(note);
    }

    private void Skip(string what) => _skipped[what] = _skipped.GetValueOrDefault(what) + 1;

    private XElement Load(string part)
    {
        using var reader = _package.OpenXml(part);
        return XDocument.Load(reader).Root ?? throw new InvalidDataException($"Part '{part}' is empty.");
    }

    public Presentation Read(string part)
    {
        var root = Load(part);
        var size = Child(root, "sldSz");
        double width = Number(size, "cx", 9_144_000) / EmuPerPoint, height = Number(size, "cy", 6_858_000) / EmuPerPoint;
        _defaultTextStyle = Child(root, "defaultTextStyle");
        if (_package.FindRelationship(part, "tableStyles") is { IsExternal: false } styles && _package.Exists(styles.Target))
            foreach (var style in Load(styles.Target).Elements().Where(e => e.Name.LocalName == "tblStyle"))
                if ((string?)style.Attribute("styleId") is { } styleId) _tableStyles[styleId] = style;
        var relationships = _package.GetRelationships(part);
        var presentation = new Presentation();
        foreach (var id in Child(root, "sldIdLst")?.Elements() ?? [])
        {
            if (RelationshipId(id) is not { } rid || !relationships.TryGetValue(rid, out var rel) || rel.IsExternal || !_package.Exists(rel.Target))
            {
                options.Warnings.Warn("pptx.slide-missing", "A slide listed in the presentation is missing from the file; it was skipped.");
                continue;
            }
            if (ReadSlide(rel.Target, width, height) is { } slide) presentation.Slides.Add(slide);
        }

        foreach (var (what, count) in _skipped)
            options.Warnings.Warn("pptx.unsupported", $"The presentation has {count} {what} that {(count == 1 ? "is" : "are")} not drawn.");
        foreach (var note in _notes) options.Warnings.Warn("pptx.approximate", $"In this presentation {note}.");
        return presentation;
    }

    // ------------------------------------------------------------------ parts

    private Master? MasterOf(string layoutPart)
    {
        var rel = _package.FindRelationship(layoutPart, "slideMaster");
        if (rel is null || rel.IsExternal || !_package.Exists(rel.Target)) return null;
        if (_masters.TryGetValue(rel.Target, out var master)) return master;
        var root = Load(rel.Target);
        var theme = DmlTheme.Default;
        string major = "Calibri Light", minor = "Calibri";
        if (_package.FindRelationship(rel.Target, "theme") is { IsExternal: false } themeRel && _package.Exists(themeRel.Target))
        {
            var themeRoot = Load(themeRel.Target);
            theme = DmlTheme.Read(themeRoot.CreateReader());
            string? Face(string name) => (string?)Child(themeRoot.Descendants().FirstOrDefault(e => e.Name.LocalName == name), "latin")?.Attribute("typeface");
            if (Face("majorFont") is { Length: > 0 } m) major = m;
            if (Face("minorFont") is { Length: > 0 } n) minor = n;
        }
        var map = Child(root, "clrMap")?.Attributes().ToDictionary(a => a.Name.LocalName, a => a.Value) ?? new Dictionary<string, string>();
        return _masters[rel.Target] = new Master(root, new Host(this, _package.GetRelationships(rel.Target), theme), major, minor, map);
    }

    private Layout? LayoutOf(string slidePart)
    {
        var rel = _package.FindRelationship(slidePart, "slideLayout");
        if (rel is null || rel.IsExternal || !_package.Exists(rel.Target)) return null;
        if (_layouts.TryGetValue(rel.Target, out var layout)) return layout;
        var master = MasterOf(rel.Target);
        var root = Load(rel.Target);
        return _layouts[rel.Target] = new Layout(root, new Host(this, _package.GetRelationships(rel.Target), master?.Host.Theme ?? DmlTheme.Default), master);
    }

    private static IReadOnlyDictionary<string, string>? OverrideMap(XElement? root) =>
        Path(root, "clrMapOvr", "overrideClrMapping")?.Attributes().ToDictionary(a => a.Name.LocalName, a => a.Value);

    private static bool ShowsMasterShapes(XElement? root) => (string?)root?.Attribute("showMasterSp") is not ("0" or "false");

    private Slide? ReadSlide(string part, double width, double height)
    {
        var root = Load(part);
        if ((string?)root.Attribute("show") is "0" or "false") return null;
        var layout = LayoutOf(part);
        var master = layout?.Master;
        var theme = master?.Host.Theme ?? DmlTheme.Default;
        var map = OverrideMap(root) ?? OverrideMap(layout?.Root) ?? master?.ColorMap ?? new Dictionary<string, string>();
        var host = new Host(this, _package.GetRelationships(part), theme);
        var slide = new Slide(width, height);

        ReadBackground(slide, [(root, host), (layout?.Root, layout?.Host), (master?.Root, master?.Host)], map, theme);

        int z = 0;
        bool layoutShapes = layout is not null && ShowsMasterShapes(root);
        if (master is not null && layoutShapes && ShowsMasterShapes(layout!.Root))
            ReadTree(slide, Path(master.Root, "cSld", "spTree"), new Scope(master.Host, null, master, map, false), ref z);
        if (layoutShapes)
            ReadTree(slide, Path(layout!.Root, "cSld", "spTree"), new Scope(layout.Host, layout, master, map, false), ref z);
        ReadTree(slide, Path(root, "cSld", "spTree"), new Scope(host, layout, master, map, true), ref z);
        return slide;
    }

    // ------------------------------------------------------------------ background

    private static void ReadBackground(Slide slide, (XElement? Root, Host? Host)[] chain, IReadOnlyDictionary<string, string> map, DmlTheme theme)
    {
        foreach (var (root, host) in chain)
        {
            if (Path(root, "cSld", "bg") is not { } original || host is null) continue;
            var bg = MapColors(new XElement(original), map);
            if (Child(bg, "bgPr") is { } properties)
            {
                var fill = properties.Elements().FirstOrDefault(e => FillKinds.Contains(e.Name.LocalName));
                ApplyBackgroundFill(slide, fill, theme, host, null);
            }
            else if (Child(bg, "bgRef") is { } reference)
            {
                int idx = (int)Number(reference, "idx");
                var color = Resolve(reference, theme);
                var style = idx >= 1001 ? theme.BackgroundFillStyle(idx - 1000) : theme.FillStyle(idx);
                if (style is null) slide.Background = color;
                else ApplyBackgroundFill(slide, MapColors(new XElement(style), map), theme, host, color);
            }
            return;
        }
    }

    private static void ApplyBackgroundFill(Slide slide, XElement? fill, DmlTheme theme, Host host, Color? placeholder)
    {
        switch (fill?.Name.LocalName)
        {
            case "solidFill": slide.Background = Resolve(fill, theme, placeholder); break;
            case "gradFill":
                slide.BackgroundGradient = DmlShapes.ReadGradientFill(fill, theme, placeholder);
                if (slide.BackgroundGradient is null) slide.Background = placeholder;
                break;
            case "blipFill":
                var embed = Child(fill, "blip")?.Attributes().FirstOrDefault(a => a.Name.LocalName == "embed")?.Value;
                if (embed is not null && host.Image(embed, out _) is { } image)
                {
                    slide.BackgroundImage = image;
                    if (Child(fill, "tile") is not null) host.Note("tiled picture backgrounds are stretched over the slide");
                }
                else host.Note("a picture background whose image is missing is left out");
                break;
            case "pattFill":
                slide.Background = Resolve(Child(fill, "fgClr"), theme, placeholder);
                host.Note("pattern backgrounds are drawn in their foreground color");
                break;
            default: slide.Background = placeholder; break;
        }
    }

    // ------------------------------------------------------------------ shapes

    private void ReadTree(Slide slide, XElement? tree, Scope scope, ref int z)
    {
        foreach (var element in Content(tree))
        {
            // The layout's and master's placeholders only lend their properties to the slide's; they are not drawn themselves.
            if (!scope.IsSlide && PlaceholderOf(element) is not null) continue;
            ReadObject(slide, element, scope, null, ref z);
        }
    }

    /// <summary>The drawable children of a shape tree or group, with <c>mc:AlternateContent</c> resolved to its fallback.</summary>
    private static IEnumerable<XElement> Content(XElement? parent)
    {
        foreach (var child in parent?.Elements() ?? [])
        {
            if (child.Name.LocalName == "AlternateContent")
            {
                var chosen = Child(child, "Fallback") ?? Child(child, "Choice");
                foreach (var inner in chosen?.Elements() ?? []) yield return inner;
            }
            else yield return child;
        }
    }

    private static XElement? PlaceholderOf(XElement element)
    {
        var nv = element.Elements().FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvCxnSpPr" or "nvGraphicFramePr");
        return Path(nv, "nvPr", "ph");
    }

    private void ReadObject(Slide slide, XElement element, Scope scope, Placement? placement, ref int z)
    {
        switch (element.Name.LocalName)
        {
            case "sp" or "cxnSp":
                {
                    var effective = Effective(element, scope);
                    var shape = DmlShapes.ReadShape(scope.Host, effective, out bool unsupported);
                    if (unsupported) Skip("shapes with a custom geometry this reader cannot evaluate");
                    if (shape is null) return;
                    var place = placement ?? Own(Child(Child(effective, "spPr"), "xfrm"));
                    slide.Objects.Add(shape with
                    {
                        Anchor = ImageAnchorKind.Absolute, X = place.X, Y = place.Y, Width = place.Width, Height = place.Height, Frame = place.Frame,
                        Rotation = shape.Rotation + place.Turns.Sum(t => t.Rotation), GroupTurns = place.Turns,
                        ZOrder = z++,
                    });
                    break;
                }
            case "pic": ReadPicture(slide, element, scope, placement, ref z); break;
            case "grpSp": ReadGroup(slide, element, scope, placement, ref z); break;
            case "graphicFrame":
                {
                    if ((string?)Path(element, "nvGraphicFramePr", "cNvPr")?.Attribute("hidden") is "1" or "true") return;
                    var uri = (string?)Path(element, "graphic", "graphicData")?.Attribute("uri") ?? "";
                    if (uri.EndsWith("/table", StringComparison.Ordinal) && Path(element, "graphic", "graphicData", "tbl") is { } table)
                    {
                        ReadTable(slide, table, scope, placement ?? Own(Child(element, "xfrm")), ref z);
                        break;
                    }
                    if (uri.EndsWith("/chart", StringComparison.Ordinal) && Path(element, "graphic", "graphicData", "chart") is { } chart)
                    {
                        ReadChart(slide, chart, scope, placement ?? Own(Child(element, "xfrm")), ref z);
                        break;
                    }
                    if (uri.EndsWith("/diagram", StringComparison.Ordinal) && Path(element, "graphic", "graphicData", "relIds") is { } relIds)
                    {
                        ReadDiagram(slide, element, relIds, scope, placement, ref z);
                        break;
                    }
                    // An embedded or linked OLE object is drawn from the preview picture saved with it; the object itself is not opened.
                    if (uri.EndsWith("/ole", StringComparison.Ordinal) && Path(element, "graphic", "graphicData", "oleObj") is { } ole)
                    {
                        if (Content(ole).FirstOrDefault(e => e.Name.LocalName == "pic") is { } preview)
                            ReadPicture(slide, preview, scope, placement ?? Own(Child(element, "xfrm")), ref z);
                        else Skip("embedded objects without a preview picture");
                        break;
                    }
                    Skip(uri.EndsWith("/chart", StringComparison.Ordinal) ? "charts"
                        : uri.EndsWith("/diagram", StringComparison.Ordinal) ? "SmartArt diagrams"
                        : "embedded objects");
                    break;
                }
            case "contentPart": Skip("ink drawings"); break;
        }
    }

    /// <summary>
    /// A chart part, read by the DrawingML chart reader shared with XLSX from its cached values; the embedded workbook is not opened, so a chart
    /// without caches is drawn empty. Scheme colors follow the slide's color map.
    /// </summary>
    private void ReadChart(Slide slide, XElement chart, Scope scope, Placement place, ref int z)
    {
        if (RelationshipId(chart) is not { } rid || scope.Host.Part(rid) is not { } part)
        {
            options.Warnings.Warn("pptx.chart-missing", "A chart's part is missing from the file; it was skipped.");
            return;
        }
        var doc = new XDocument(MapColors(Load(part), scope.ColorMap));
        var template = new SheetChart
        {
            Anchor = ImageAnchorKind.Absolute, X = place.X, Y = place.Y, Width = place.Width, Height = place.Height, Frame = place.Frame,
            GroupTurns = place.Turns, Rotation = place.Turns.Sum(t => t.Rotation), ZOrder = z++,
        };
        var notes = new List<string>();
        var (unsupported, build) = DmlCharts.Prepare(doc, template, notes);
        foreach (var note in notes) Note(note);
        if (build is null)
        {
            Skip(unsupported is null or "empty" ? "empty charts" : $"{unsupported} charts");
            return;
        }
        // Text without a color of its own takes the theme's text color through the color map, as other slide text does.
        var textColor = Resolve(new XElement("fill", new XElement("schemeClr", new XAttribute("val", scope.ColorMap.GetValueOrDefault("tx1", "dk1")))), scope.Host.Theme);
        slide.Objects.Add(build(new CachedChartData(scope.Host.Theme, textColor)));
    }

    /// <summary>A chart's data as it reads in a presentation: only its cached values, no cells.</summary>
    private sealed class CachedChartData(DmlTheme theme, Color? textColor) : IChartData
    {
        public DmlTheme Theme { get; } = theme;

        public Color? TextColor { get; } = textColor;

        public bool PresentationTitle => true;

        public bool Date1904 => false;

        public bool FramedChartArea => false;

        public bool FitAxisToHeight => true;

        public IEnumerable<Cell?> Cells(string formula) => [];
    }

    /// <summary>
    /// A SmartArt diagram, drawn from the drawing PowerPoint saves with it (<c>dsp:drawing</c>, which the diagram's data part names through
    /// <c>dataModelExt relId</c>): its shapes are read as a group filling the graphic frame. A shape whose text box (<c>txXfrm</c>) differs from
    /// its frame becomes the shape and an unfilled text box over it. A diagram without that drawing is not drawn.
    /// </summary>
    private void ReadDiagram(Slide slide, XElement frame, XElement relIds, Scope scope, Placement? placement, ref int z)
    {
        var dataId = relIds.Attributes().FirstOrDefault(a => a.Name.LocalName == "dm")?.Value;
        var drawingId = dataId is not null && scope.Host.Part(dataId) is { } dataPart
            ? (string?)Load(dataPart).Descendants().FirstOrDefault(e => e.Name.LocalName == "dataModelExt")?.Attribute("relId")
            : null;
        if (drawingId is null || scope.Host.Part(drawingId) is not { } drawingPart || Child(Load(drawingPart), "spTree") is not { } tree)
        {
            Skip("SmartArt diagrams without a saved drawing");
            return;
        }
        XNamespace a = A, p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var xfrm = Child(frame, "xfrm");
        var ext = Child(xfrm, "ext");
        var group = new XElement(p + "grpSp",
            new XElement(p + "grpSpPr",
                new XElement(a + "xfrm", new XElement(Child(xfrm, "off") ?? new XElement(a + "off", new XAttribute("x", 0), new XAttribute("y", 0))) { Name = a + "off" },
                    new XElement(a + "ext", new XAttribute("cx", Number(ext, "cx")), new XAttribute("cy", Number(ext, "cy"))),
                    new XElement(a + "chOff", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(a + "chExt", new XAttribute("cx", Number(ext, "cx")), new XAttribute("cy", Number(ext, "cy"))))));
        foreach (var child in Content(tree)) group.Add(SplitTextBox(child));
        var host = new Host(this, _package.GetRelationships(drawingPart), scope.Host.Theme);
        ReadGroup(slide, group, scope with { Host = host, IsSlide = false }, placement, ref z);
    }

    /// <summary>A diagram shape, or for one with its own text box (<c>txXfrm</c>) the shape without its text followed by that text box.</summary>
    private static IEnumerable<XElement> SplitTextBox(XElement element)
    {
        if (element.Name.LocalName == "grpSp")
        {
            var group = new XElement(element);
            foreach (var inner in group.Elements().Where(e => e.Name.LocalName is "sp" or "grpSp").ToList()) inner.ReplaceWith(SplitTextBox(inner));
            return [group];
        }
        var textFrame = Child(element, "txXfrm");
        var body = Child(element, "txBody");
        var frame = Path(element, "spPr", "xfrm");
        if (element.Name.LocalName != "sp" || textFrame is null || body is null || frame is null) return [element];
        string Key(XElement x) => string.Join(",", Number(Child(x, "off"), "x"), Number(Child(x, "off"), "y"), Number(Child(x, "ext"), "cx"), Number(Child(x, "ext"), "cy"));
        if (Key(textFrame) == Key(frame) && Number(textFrame, "rot") == Number(frame, "rot")) return [element];

        var shape = new XElement(element);
        Child(shape, "txBody")?.Remove();
        Child(shape, "txXfrm")?.Remove();
        XNamespace a = A;
        var text = new XElement(element);
        Child(text, "txXfrm")?.Remove();
        var spPr = Child(text, "spPr")!;
        spPr.RemoveNodes();
        spPr.Add(new XElement(textFrame) { Name = a + "xfrm" }, new XElement(a + "prstGeom", new XAttribute("prst", "rect")), new XElement(a + "noFill"),
            new XElement(a + "ln", new XElement(a + "noFill")));
        // The style's fill, line and effects belong to the shape; its font color stays with the text.
        if (Child(text, "style") is { } style)
            foreach (var reference in style.Elements().Where(e => e.Name.LocalName is "lnRef" or "fillRef" or "effectRef")) reference.SetAttributeValue("idx", "0");
        return [shape, text];
    }

    private static Placement Own(XElement? xfrm) => new(
        Number(Child(xfrm, "off"), "x") / EmuPerPoint, Number(Child(xfrm, "off"), "y") / EmuPerPoint,
        Number(Child(xfrm, "ext"), "cx") / EmuPerPoint, Number(Child(xfrm, "ext"), "cy") / EmuPerPoint, FrameFraction.Whole, []);

    /// <summary>
    /// A group places its children inside its frame: a child's coordinates are mapped from the group's child space (<c>chOff</c>, <c>chExt</c>).
    /// A rotated group swings its children about its center and turns them with it.
    /// </summary>
    private void ReadGroup(Slide slide, XElement group, Scope scope, Placement? placement, ref int z)
    {
        if ((string?)Path(group, "nvGrpSpPr", "cNvPr")?.Attribute("hidden") is "1" or "true") return;
        var xfrm = Path(group, "grpSpPr", "xfrm");
        var place = placement ?? Own(xfrm);
        double turn = -Number(xfrm, "rot") / 60_000.0;
        var turns = turn != 0 ? [.. place.Turns, new GroupTurn(place.Frame, turn)] : place.Turns;
        var childOffset = Child(xfrm, "chOff");
        double cx = Number(Child(xfrm, "chExt"), "cx"), cy = Number(Child(xfrm, "chExt"), "cy");
        foreach (var child in Content(group))
        {
            var childXfrm = child.Name.LocalName switch
            {
                "grpSp" => Path(child, "grpSpPr", "xfrm"),
                "graphicFrame" => Child(child, "xfrm"),
                _ => Path(child, "spPr", "xfrm"),
            };
            var fraction = cx > 0 && cy > 0 && Child(childXfrm, "off") is { } off && Child(childXfrm, "ext") is { } ext
                ? new FrameFraction((Number(off, "x") - Number(childOffset, "x")) / cx, (Number(off, "y") - Number(childOffset, "y")) / cy, Number(ext, "cx") / cx, Number(ext, "cy") / cy)
                : FrameFraction.Whole;
            ReadObject(slide, child, scope, place with { Frame = place.Frame.Within(fraction), Turns = turns }, ref z);
        }
    }

    private void ReadPicture(Slide slide, XElement pic, Scope scope, Placement? placement, ref int z)
    {
        if ((string?)Path(pic, "nvPicPr", "cNvPr")?.Attribute("hidden") is "1" or "true") return;
        var effective = Effective(pic, scope);
        var blip = Path(effective, "blipFill", "blip");
        var embed = blip?.Attributes().FirstOrDefault(a => a.Name.LocalName == "embed")?.Value;
        bool unreadable = false;
        if ((embed is null ? null : scope.Host.Image(embed, out unreadable)) is not { } image)
        {
            options.Warnings.Warn(unreadable ? "pptx.image-unreadable" : "pptx.image-missing",
                unreadable ? "A picture could not be read; it was skipped." : "A picture's image is missing or linked from outside the file; it was skipped.");
            return;
        }
        var src = Path(effective, "blipFill", "srcRect");
        double Crop(string name) => Math.Clamp(Number(src, name) / 100_000.0, 0, 1);
        var xfrm = Path(effective, "spPr", "xfrm");
        var place = placement ?? Own(xfrm);
        slide.Objects.Add(new SheetImage
        {
            Image = image,
            Anchor = ImageAnchorKind.Absolute, X = place.X, Y = place.Y, Width = place.Width, Height = place.Height, Frame = place.Frame,
            CropLeft = Crop("l"), CropTop = Crop("t"), CropRight = Crop("r"), CropBottom = Crop("b"),
            Rotation = -Number(xfrm, "rot") / 60_000.0 + place.Turns.Sum(t => t.Rotation), GroupTurns = place.Turns,
            FlipH = Flag(xfrm, "flipH"), FlipV = Flag(xfrm, "flipV"),
            ZOrder = z++,
        });
    }

    // ------------------------------------------------------------------ inheritance

    /// <summary>The placeholder type the master is searched for: titles stand for centered titles, the body for subtitles and content.</summary>
    private static string MasterType(string type) => type switch
    {
        "title" or "ctrTitle" => "title",
        "dt" or "ftr" or "sldNum" or "hdr" => type,
        _ => "body",
    };

    /// <summary>The placeholder of <paramref name="root"/>'s shape tree that a placeholder of this type and index inherits from: same index first, then same type.</summary>
    private static XElement? FindPlaceholder(XElement? root, string type, string? idx, bool byMasterType)
    {
        var shapes = Path(root, "cSld", "spTree")?.Descendants().Where(e => e.Name.LocalName is "sp" or "pic").ToList() ?? [];
        string Type(XElement ph) => (string?)ph.Attribute("type") ?? "obj";
        if (idx is not null && !byMasterType)
            foreach (var shape in shapes)
                if (PlaceholderOf(shape) is { } ph && (string?)ph.Attribute("idx") == idx) return shape;
        foreach (var shape in shapes)
            if (PlaceholderOf(shape) is { } ph && (byMasterType ? MasterType(Type(ph)) == MasterType(type) : Type(ph) == type)) return shape;
        return null;
    }

    /// <summary>
    /// A copy of a slide shape with everything it inherits written into it, so that the shared DrawingML reader sees one complete shape: the
    /// frame, geometry, fill, outline and style of its layout and master placeholders, their body properties, the text styles of every level,
    /// bullets as leading text, theme fonts by name and the color map applied.
    /// </summary>
    private XElement Effective(XElement element, Scope scope)
    {
        var shape = new XElement(element);
        XElement? layoutPh = null, masterPh = null;
        string? type = null;
        if (scope.IsSlide && PlaceholderOf(shape) is { } ph)
        {
            type = (string?)ph.Attribute("type") ?? "obj";
            layoutPh = FindPlaceholder(scope.Layout?.Root, type, (string?)ph.Attribute("idx"), false);
            // A slide placeholder with only an index takes its type from the layout placeholder it matches.
            if (ph.Attribute("type") is null && layoutPh is not null && PlaceholderOf(layoutPh) is { } layoutType) type = (string?)layoutType.Attribute("type") ?? type;
            masterPh = FindPlaceholder(scope.Master?.Root, type, null, true);
        }
        XElement?[] ancestors = [layoutPh, masterPh]; // nearest first

        var ns = shape.Name.Namespace;
        var spPr = Child(shape, "spPr");
        if (spPr is null) shape.Add(spPr = new XElement(ns + "spPr"));
        foreach (var ancestor in ancestors)
        {
            var from = Child(ancestor, "spPr");
            if (from is null) continue;
            if (Child(Child(spPr, "xfrm"), "ext") is null && Child(from, "xfrm") is { } xfrm)
            {
                Child(spPr, "xfrm")?.Remove();
                spPr.AddFirst(new XElement(xfrm));
            }
            if (Child(spPr, "prstGeom") is null && Child(spPr, "custGeom") is null && from.Elements().FirstOrDefault(e => e.Name.LocalName is "prstGeom" or "custGeom") is { } geometry)
                spPr.Add(new XElement(geometry));
            if (!spPr.Elements().Any(e => FillKinds.Contains(e.Name.LocalName)) && from.Elements().FirstOrDefault(e => FillKinds.Contains(e.Name.LocalName)) is { } fill)
                spPr.Add(new XElement(fill));
            if (Child(spPr, "ln") is null && Child(from, "ln") is { } ln) spPr.Add(new XElement(ln));
            if (Child(spPr, "effectLst") is null && Child(from, "effectLst") is { } effects) spPr.Add(new XElement(effects));
            if (Child(shape, "style") is null && Child(ancestor, "style") is { } style) shape.Add(new XElement(style));
        }

        if (Child(shape, "txBody") is { } body) ExpandText(body, scope, type, ancestors, Path(shape, "style", "fontRef")?.Elements().Any() == true);
        return MapColors(shape, scope.ColorMap);
    }

    /// <summary>
    /// Writes the inherited body properties and the text styles of each paragraph's level into the text body, with bullets as leading text.
    /// When the shape's style names a font color (<c>fontRef</c>, <paramref name="styleColor"/>), it outranks the presentation's and master's
    /// text colors.
    /// </summary>
    private void ExpandText(XElement body, Scope scope, string? type, XElement?[] ancestors, bool styleColor)
    {
        var master = scope.Master;
        var ns = body.Name.Namespace;
        var a = Child(body, "p")?.Name.Namespace ?? (XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";

        // Body properties: the master's, then the layout's, then the shape's own attributes; the autofit of the nearest that states one.
        var bodyPr = Child(body, "bodyPr");
        if (bodyPr is null) body.AddFirst(bodyPr = new XElement(a + "bodyPr"));
        foreach (var ancestor in ancestors)
        {
            if (Path(ancestor, "txBody", "bodyPr") is not { } inherited) continue;
            foreach (var attribute in inherited.Attributes())
                if (bodyPr.Attribute(attribute.Name) is null) bodyPr.SetAttributeValue(attribute.Name, attribute.Value);
            if (!bodyPr.Elements().Any(e => e.Name.LocalName is "noAutofit" or "normAutofit" or "spAutoFit")
                && inherited.Elements().FirstOrDefault(e => e.Name.LocalName is "noAutofit" or "normAutofit" or "spAutoFit") is { } fit)
                bodyPr.Add(new XElement(fit));
        }
        var normAutofit = Child(bodyPr, "normAutofit");
        double fontScale = normAutofit?.Attribute("fontScale") is null ? 1 : Number(normAutofit, "fontScale", 100_000) / 100_000.0;

        // Text styles from lowest to highest: the presentation's defaults, the master's style for this kind of placeholder, the list styles of
        // the master and layout placeholders, then the shape's own list style.
        var styles = new List<XElement?> { _defaultTextStyle };
        if (type is not null)
        {
            var txStyles = Child(master?.Root, "txStyles");
            styles.Add(MasterType(type) switch { "title" => Child(txStyles, "titleStyle"), "body" => Child(txStyles, "bodyStyle"), _ => Child(txStyles, "otherStyle") });
        }
        int baseStyles = styles.Count;
        for (int i = ancestors.Length - 1; i >= 0; i--) styles.Add(Path(ancestors[i], "txBody", "lstStyle"));
        styles.Add(Child(body, "lstStyle"));

        var numbering = new Dictionary<int, int>();
        foreach (var p in body.Elements().Where(e => e.Name.LocalName == "p").ToList())
        {
            var ownPPr = Child(p, "pPr");
            int level = Math.Clamp((int)Number(ownPPr, "lvl"), 0, 8);
            var pPr = new XElement(a + "pPr");
            var defRPr = new XElement(a + "defRPr");
            for (int s = 0; s < styles.Count; s++)
            {
                if (s == baseStyles && styleColor)
                    foreach (var fill in defRPr.Elements().Where(e => FillKinds.Contains(e.Name.LocalName)).ToList()) fill.Remove();
                if (Child(styles[s], $"lvl{level + 1}pPr") is not { } levelStyle) continue;
                MergeParagraph(pPr, levelStyle);
                MergeRun(defRPr, Child(levelStyle, "defRPr"));
            }
            if (ownPPr is not null)
            {
                MergeParagraph(pPr, ownPPr);
                MergeRun(defRPr, Child(ownPPr, "defRPr"));
                ownPPr.Remove();
            }
            p.AddFirst(pPr);

            XElement Run(XElement? own)
            {
                var rPr = new XElement(defRPr) { Name = a + "rPr" };
                MergeRun(rPr, own);
                double size = Number(rPr, "sz", 1800) * fontScale;
                rPr.SetAttributeValue("sz", Math.Max(100, Math.Round(size / 50) * 50).ToString(CultureInfo.InvariantCulture));
                var latin = Child(rPr, "latin");
                var face = (string?)latin?.Attribute("typeface");
                face = face switch
                {
                    null or "" or "+mn-lt" => master?.MinorFont ?? "Calibri",
                    "+mj-lt" => master?.MajorFont ?? "Calibri Light",
                    _ => face,
                };
                if (latin is null) rPr.Add(latin = new XElement(a + "latin"));
                latin.SetAttributeValue("typeface", face);
                return rPr;
            }

            bool hasText = false;
            XElement? firstRun = null;
            foreach (var run in p.Elements().Where(e => e.Name.LocalName is "r" or "fld" or "br").ToList())
            {
                var rPr = Run(Child(run, "rPr"));
                Child(run, "rPr")?.Remove();
                run.AddFirst(rPr);
                if (run.Name.LocalName != "br" && (Child(run, "t")?.Value.Length ?? 0) > 0)
                {
                    hasText = true;
                    firstRun ??= run;
                }
            }
            if (Child(p, "endParaRPr") is { } end) end.ReplaceWith(new XElement(Run(end)) { Name = a + "endParaRPr" });
            else p.Add(new XElement(Run(null)) { Name = a + "endParaRPr" });

            if (Bullet(pPr, level, hasText, numbering) is { } bullet && firstRun is not null)
            {
                var rPr = new XElement(Child(firstRun, "rPr")!);
                if (Child(pPr, "buClr") is { } color && color.Elements().FirstOrDefault() is { } colorElement)
                {
                    foreach (var f in rPr.Elements().Where(e => FillKinds.Contains(e.Name.LocalName)).ToList()) f.Remove();
                    rPr.Add(new XElement(a + "solidFill", new XElement(colorElement)));
                }
                if (Child(pPr, "buSzPct") is { } percent)
                    rPr.SetAttributeValue("sz", Math.Max(100, Math.Round(Number(rPr, "sz", 1800) * Number(percent, "val", 100_000) / 100_000.0 / 50) * 50).ToString(CultureInfo.InvariantCulture));
                firstRun.AddBeforeSelf(new XElement(a + "r", rPr, new XElement(a + "t", bullet)));
            }
        }
    }

    /// <summary>The text a bullet stands for, followed by a space; null when the paragraph has none. Numbered bullets count per level.</summary>
    private string? Bullet(XElement pPr, int level, bool hasText, Dictionary<int, int> numbering)
    {
        var kind = pPr.Elements().FirstOrDefault(e => BulletKinds.Contains(e.Name.LocalName));
        if (kind?.Name.LocalName != "buAutoNum")
            foreach (var deeper in numbering.Keys.Where(k => k >= level).ToList()) numbering.Remove(deeper);
        if (!hasText) return null;
        // With a hanging indent the text after the bullet starts at the margin, as a tab to it; otherwise a space follows.
        string gap = Number(pPr, "indent") < 0 ? "\t" : " ";
        switch (kind?.Name.LocalName)
        {
            case "buChar":
                var ch = (string?)kind.Attribute("char");
                if (string.IsNullOrEmpty(ch)) return null;
                // Symbol fonts put their pictures at letter code points; without the font the letter would show, so a plain bullet stands in.
                if ((string?)Child(pPr, "buFont")?.Attribute("typeface") is { } font && SymbolFonts.Contains(font, StringComparer.OrdinalIgnoreCase))
                {
                    Note("bullets in symbol fonts are drawn as round bullets");
                    ch = "•";
                }
                return ch + gap;
            case "buAutoNum":
                foreach (var deeper in numbering.Keys.Where(k => k > level).ToList()) numbering.Remove(deeper);
                int start = kind.Attribute("startAt") is null ? 1 : (int)Number(kind, "startAt", 1);
                int n = numbering[level] = numbering.TryGetValue(level, out var previous) ? previous + 1 : start;
                return AutoNumber((string?)kind.Attribute("type") ?? "arabicPeriod", n) + gap;
            case "buBlip":
                Note("picture bullets are drawn as round bullets");
                return "•" + gap;
            default: return null;
        }
    }

    private string AutoNumber(string scheme, int n)
    {
        string number;
        if (scheme.StartsWith("arabic", StringComparison.Ordinal)) number = n.ToString(CultureInfo.InvariantCulture);
        else if (scheme.StartsWith("alphaLc", StringComparison.Ordinal)) number = Letters(n);
        else if (scheme.StartsWith("alphaUc", StringComparison.Ordinal)) number = Letters(n).ToUpperInvariant();
        else if (scheme.StartsWith("romanLc", StringComparison.Ordinal)) number = Roman(n).ToLowerInvariant();
        else if (scheme.StartsWith("romanUc", StringComparison.Ordinal)) number = Roman(n);
        else
        {
            Note($"numbered bullets of the scheme '{scheme}' are numbered 1, 2, 3");
            number = n.ToString(CultureInfo.InvariantCulture);
        }
        if (scheme.EndsWith("ParenBoth", StringComparison.Ordinal)) return $"({number})";
        if (scheme.EndsWith("ParenR", StringComparison.Ordinal)) return number + ")";
        if (scheme.EndsWith("Plain", StringComparison.Ordinal)) return number;
        return number + ".";
    }

    private static string Letters(int n)
    {
        var s = "";
        for (n = Math.Max(n, 1); n > 0; n = (n - 1) / 26) s = (char)('a' + (n - 1) % 26) + s;
        return s;
    }

    private static string Roman(int n)
    {
        if (n is < 1 or > 3999) return n.ToString(CultureInfo.InvariantCulture);
        (int Value, string Text)[] digits = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
        var s = "";
        foreach (var (value, text) in digits)
            for (; n >= value; n -= value) s += text;
        return s;
    }

    /// <summary>Overlays paragraph properties: attributes replace, the bullet kind replaces as one choice, other children replace by name; run defaults are merged separately.</summary>
    private static void MergeParagraph(XElement target, XElement source)
    {
        foreach (var attribute in source.Attributes()) target.SetAttributeValue(attribute.Name, attribute.Value);
        foreach (var child in source.Elements())
        {
            var name = child.Name.LocalName;
            if (name == "defRPr") continue;
            if (BulletKinds.Contains(name)) foreach (var old in target.Elements().Where(e => BulletKinds.Contains(e.Name.LocalName)).ToList()) old.Remove();
            else Child(target, name)?.Remove();
            target.Add(new XElement(child));
        }
    }

    /// <summary>Overlays run properties: attributes replace, the fill replaces as one choice, other children (fonts, outline) replace by name.</summary>
    private static void MergeRun(XElement target, XElement? source)
    {
        if (source is null) return;
        foreach (var attribute in source.Attributes()) target.SetAttributeValue(attribute.Name, attribute.Value);
        foreach (var child in source.Elements())
        {
            var name = child.Name.LocalName;
            if (FillKinds.Contains(name)) foreach (var old in target.Elements().Where(e => FillKinds.Contains(e.Name.LocalName)).ToList()) old.Remove();
            else Child(target, name)?.Remove();
            target.Add(new XElement(child));
        }
    }

    /// <summary>Rewrites the scheme colors that the color map remaps (<c>bg1</c>, <c>tx1</c>, <c>bg2</c>, <c>tx2</c>) to the theme colors they stand for.</summary>
    private static XElement MapColors(XElement element, IReadOnlyDictionary<string, string> map)
    {
        if (map.Count == 0) return element;
        foreach (var color in element.DescendantsAndSelf().Where(e => e.Name.LocalName == "schemeClr"))
            if ((string?)color.Attribute("val") is "bg1" or "tx1" or "bg2" or "tx2" && map.TryGetValue(color.Attribute("val")!.Value, out var target))
                color.SetAttributeValue("val", target);
        return element;
    }

    private static string? RelationshipId(XElement e) =>
        e.Attributes().FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.Namespace != XNamespace.None)?.Value;
}
