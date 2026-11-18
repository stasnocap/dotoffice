using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DotOffice.Core;
using DotOffice.Opc;

namespace DotOffice.Odf;

/// <summary>Reads OpenDocument presentations (.odp, .otp) into the <see cref="Presentation"/> model.</summary>
public static class OdpReader
{
    /// <summary>Reads an ODP package. Malformed, encrypted or oversized input throws <see cref="InvalidDataException"/>.</summary>
    public static Presentation Read(Stream stream, OdfReaderOptions? options = null)
    {
        options ??= new OdfReaderOptions();
        try
        {
            using var package = OdfPackage.Open(stream, options, "application/vnd.oasis.opendocument.presentation", out var styles);
            XElement root;
            using (var r = package.OpenXml("content.xml")) root = XElement.Load(r);
            foreach (var part in root.Elements().Where(e => e.Name.LocalName is "automatic-styles" or "font-face-decls")) styles.Add(part);
            var body = root.Element(Ns.Office + "body")?.Element(Ns.Office + "presentation")
                       ?? throw new InvalidDataException("The document is not a presentation.");
            return new OdpContext(package, styles, options).Read(body);
        }
        catch (System.Xml.XmlException e)
        {
            throw new InvalidDataException($"The document contains malformed XML: {e.Message}", e);
        }
    }
}

internal sealed partial class OdpContext(OpcPackage package, OdfStyles styles, OdfReaderOptions options)
{
    private const int MaxPages = 5_000, MaxObjects = 200_000, MaxDepth = 32, MaxTableCells = 100_000;

    private readonly Dictionary<string, ImageData?> _images = [];
    private readonly HashSet<string> _warned = [];
    private int _objects, _tables;
    private XElement _body = null!;

    // Per slide.
    private Slide _slide = null!;
    private int _z, _pageNumber;
    private string _footer = "", _date = "";

    private void Warn(string code, string message)
    {
        if (_warned.Add(code)) options.Warnings.Warn(code, message);
    }

    public Presentation Read(XElement body)
    {
        _body = body;
        var presentation = new Presentation();
        int number = 0;
        foreach (var page in body.Elements(Ns.Draw + "page"))
        {
            if (++number > MaxPages) throw new InvalidDataException("The presentation has too many slides.");
            var pageProps = styles.Properties("drawing-page", (string?)page.Attribute(Ns.Draw + "style-name"), false);
            if (pageProps.GetValueOrDefault("drawing-page-properties/visibility") == "hidden") continue;
            presentation.Slides.Add(ReadPage(page, pageProps, number));
        }
        if (presentation.Slides.Count == 0) throw new InvalidDataException("The presentation has no slides.");
        return presentation;
    }

    // ------------------------------------------------------------------ pages

    private Slide ReadPage(XElement page, Dictionary<string, string> pageProps, int number)
    {
        var masterName = (string?)page.Attribute(Ns.Draw + "master-page-name");
        var master = styles.MasterPage(masterName) ?? styles.MasterPage(styles.DefaultMasterPage);
        var layout = styles.PageLayoutOf((string?)master?.Attribute(Ns.Style + "name"))?.Element(Ns.Style + "page-layout-properties");
        double w = OdfUnits.Length((string?)layout?.Attribute(Ns.Fo + "page-width")) ?? 792, h = OdfUnits.Length((string?)layout?.Attribute(Ns.Fo + "page-height")) ?? 612;
        if (w < 1 || h < 1 || w > 14_400 || h > 14_400) throw new InvalidDataException("The slide size is not valid.");

        _slide = new Slide(w, h);
        _z = 0;
        _pageNumber = number;
        _footer = FooterText(page, "footer-decl", "use-footer-name");
        _date = FooterText(page, "date-time-decl", "use-date-time-name");

        var masterProps = styles.Properties("drawing-page", (string?)master?.Attribute(Ns.Draw + "style-name"), false);
        bool Flag(string key, bool def) => pageProps.GetValueOrDefault("drawing-page-properties/" + key) is { } v ? v == "true" : def;
        if (Flag("background-visible", true)) ReadBackground(pageProps.ContainsKey("drawing-page-properties/fill") ? pageProps : masterProps);
        else ReadBackground(pageProps);

        if (master is not null && Flag("background-objects-visible", true))
        {
            bool footer = Flag("display-footer", false), date = Flag("display-date-time", false), pageNum = Flag("display-page-number", false);
            foreach (var shape in master.Elements())
            {
                if (shape.Attribute(Ns.Presentation + "placeholder") is { Value: "true" }) continue;
                switch ((string?)shape.Attribute(Ns.Presentation + "class"))
                {
                    case "title" or "outline" or "subtitle" or "notes" or "page" or "header": continue;
                    case "footer" when !footer: continue;
                    case "date-time" when !date: continue;
                    case "page-number" when !pageNum: continue;
                }
                ReadObject(shape, null, 0);
            }
        }
        foreach (var shape in page.Elements()) ReadObject(shape, master, 0);
        return _slide;
    }

    private string FooterText(XElement page, string decl, string useAttribute)
    {
        var name = (string?)page.Attribute(Ns.Presentation + useAttribute);
        if (name is null) return "";
        var e = _body.Elements(Ns.Presentation + decl).FirstOrDefault(d => (string?)d.Attribute(Ns.Presentation + "name") == name);
        return e?.Value ?? "";
    }

    private void ReadBackground(Dictionary<string, string> props)
    {
        string? P(string key) => props.GetValueOrDefault("drawing-page-properties/" + key);
        switch (P("fill"))
        {
            case "solid": _slide.Background = OdfUnits.Hex(P("fill-color")); break;
            case "gradient": _slide.BackgroundGradient = Gradient(P("fill-gradient-name")); break;
            case "bitmap": _slide.BackgroundImage = FillImage(P("fill-image-name")); break;
        }
    }

    // ------------------------------------------------------------------ fills

    private GradientFill? Gradient(string? name)
    {
        if (styles.Resource("gradient", name) is not { } g) return null;
        var start = OdfUnits.Hex((string?)g.Attribute(Ns.Draw + "start-color")) ?? Color.White;
        var end = OdfUnits.Hex((string?)g.Attribute(Ns.Draw + "end-color")) ?? Color.White;
        double angle = double.TryParse((string?)g.Attribute(Ns.Draw + "angle"), NumberStyles.Float, CultureInfo.InvariantCulture, out var a) ? a / 10 : 0;
        double Percent(string attr, double d) => double.TryParse(((string?)g.Attribute(Ns.Draw + attr))?.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 100 : d;
        double border = Math.Clamp(Percent("border", 0), 0, 0.95);
        switch ((string?)g.Attribute(Ns.Draw + "style"))
        {
            case "radial" or "ellipsoid" or "square" or "rectangular":
            {
                double cx = Percent("cx", 0.5), cy = Percent("cy", 0.5);
                return new GradientFill
                {
                    Kind = GradientKind.Path, Left = cx, Right = cx, Top = cy, Bottom = cy,
                    Rectangular = (string?)g.Attribute(Ns.Draw + "style") is "square" or "rectangular",
                    Stops = [new GradientStop(0, end), new GradientStop(1 - border, start), new GradientStop(1, start)],
                };
            }
            case "axial":
                return new GradientFill
                {
                    Kind = GradientKind.Linear, Degree = 90 - angle,
                    Stops = [new GradientStop(0, start), new GradientStop(0.5 * (1 - border), end), new GradientStop(0.5 + 0.5 * border, end), new GradientStop(1, start)],
                };
            default:
                return new GradientFill
                {
                    Kind = GradientKind.Linear, Degree = 90 - angle,
                    Stops = [new GradientStop(0, start), new GradientStop(border, start), new GradientStop(1, end)],
                };
        }
    }

    private ImageData? FillImage(string? name) =>
        styles.Resource("fill-image", name) is { } e ? LoadImage((string?)e.Attribute(Ns.XLink + "href"), e) : null;

    private ShapePattern? Hatch(string? name, Color? background)
    {
        if (styles.Resource("hatch", name) is not { } h) return null;
        var color = OdfUnits.Hex((string?)h.Attribute(Ns.Draw + "color")) ?? Color.Black;
        double spacing = OdfUnits.Length((string?)h.Attribute(Ns.Draw + "distance")) ?? 3;
        double rotation = double.TryParse((string?)h.Attribute(Ns.Draw + "rotation"), NumberStyles.Float, CultureInfo.InvariantCulture, out var r) ? r / 10 : 0;
        rotation = ((rotation % 180) + 180) % 180;
        bool diagonal = rotation is > 20 and < 70 or > 110 and < 160;
        bool single = (string?)h.Attribute(Ns.Draw + "style") == "single";
        var kind = single
            ? diagonal ? (rotation < 90 ? HatchKind.UpDiagonal : HatchKind.DownDiagonal) : rotation is >= 70 and <= 110 ? HatchKind.Vertical : HatchKind.Horizontal
            : diagonal ? HatchKind.DiagonalCross : HatchKind.Cross;
        return new ShapePattern(kind, color, background ?? Color.White, Math.Max(2, spacing), 0.75);
    }

    private ImageData? LoadImage(string? href, XElement image)
    {
        if (href is not null && !href.Contains("://") && !href.StartsWith('/') && !href.Contains("..") && package.Exists(href))
        {
            if (_images.TryGetValue(href, out var cached)) return cached;
            ImageData? data = null;
            try { data = new ImageData(package.ReadAllBytes(href), href); }
            catch (InvalidDataException) { }
            if (data is { Format: ImageFormat.Unknown })
            {
                Warn("odp.image-format", "A picture in a format that is not supported (SVG, SVM, ...) is not drawn.");
                data = null;
            }
            return _images[href] = data;
        }
        if (image.Element(Ns.Office + "binary-data") is { } b)
        {
            try { return new ImageData(Convert.FromBase64String(b.Value)); }
            catch (FormatException) { return null; }
        }
        Warn("odp.image-missing", $"The image '{href}' is not in the package or is external; it is not drawn.");
        return null;
    }

    // ------------------------------------------------------------------ objects

    private static Dictionary<string, string> Merge(params Dictionary<string, string>[] layers)
    {
        var result = new Dictionary<string, string>();
        foreach (var layer in layers)
            foreach (var (k, v) in layer) result[k] = v;
        return result;
    }

    private static double Attr(XElement e, XName name) => OdfUnits.Length((string?)e.Attribute(name)) ?? 0;

    /// <summary>The merged graphic and presentation style properties of an object.</summary>
    private Dictionary<string, string> GraphicProps(XElement e, bool defaultStandard)
    {
        var drawStyle = (string?)e.Attribute(Ns.Draw + "style-name");
        var props = styles.Properties("graphic", drawStyle ?? (defaultStandard ? "standard" : null));
        if ((string?)e.Attribute(Ns.Presentation + "style-name") is { } ps)
            props = Merge(props, styles.Properties("presentation", ps, false));
        return props;
    }

    private void ReadObject(XElement e, XElement? master, int depth)
    {
        if (e.Name.Namespace != Ns.Draw) return;
        if (depth > MaxDepth) throw new InvalidDataException("The presentation is nested too deeply.");
        if (++_objects > MaxObjects) throw new InvalidDataException("The presentation has too many objects.");
        switch (e.Name.LocalName)
        {
            case "frame": ReadFrame(e, master, depth); break;
            case "g":
                foreach (var child in e.Elements()) ReadObject(child, master, depth + 1);
                break;
            case "custom-shape" or "rect" or "ellipse" or "circle" or "caption" or "polygon" or "polyline" or "path" or "line" or "connector" or "measure":
                ReadShape(e);
                break;
            case "page-thumbnail" or "plugin" or "control": break;
            default: Warn("odp.object-skipped:" + e.Name.LocalName, $"A '{e.Name.LocalName}' object on a slide is not drawn."); break;
        }
    }

    private (double X, double Y, double W, double H, double Rotation) Box(XElement e, XElement? master)
    {
        double x = Attr(e, Ns.Svg + "x"), y = Attr(e, Ns.Svg + "y"), w = Attr(e, Ns.Svg + "width"), h = Attr(e, Ns.Svg + "height");
        if (e.Name.LocalName is "line" or "connector" or "measure")
        {
            double x1 = Attr(e, Ns.Svg + "x1"), y1 = Attr(e, Ns.Svg + "y1"), x2 = Attr(e, Ns.Svg + "x2"), y2 = Attr(e, Ns.Svg + "y2");
            x = Math.Min(x1, x2); y = Math.Min(y1, y2); w = Math.Abs(x2 - x1); h = Math.Abs(y2 - y1);
        }
        else if (e.Name.LocalName == "caption" || e.Attribute(Ns.Svg + "width") is null && e.Attribute(Ns.Svg + "x") is null)
        {
            // Placeholders without a position take that of the master's placeholder of the same class.
            var cls = (string?)e.Attribute(Ns.Presentation + "class");
            if (cls is not null && master?.Elements().FirstOrDefault(m => (string?)m.Attribute(Ns.Presentation + "class") == cls && m.Attribute(Ns.Svg + "width") is not null) is { } m)
                (x, y, w, h) = (Attr(m, Ns.Svg + "x"), Attr(m, Ns.Svg + "y"), Attr(m, Ns.Svg + "width"), Attr(m, Ns.Svg + "height"));
        }
        double rotation = 0;
        if ((string?)e.Attribute(Ns.Draw + "transform") is { } transform) (x, y, rotation) = Transform(transform, x, y, w, h);
        return (x, y, w, h, rotation);
    }

    private static readonly Regex TransformPart = new(@"(rotate|translate|scale|skewX|skewY|matrix)\s*\(([^)]*)\)", RegexOptions.Compiled);

    /// <summary>The top-left corner and rotation (degrees, counter-clockwise) of a shape of the size <paramref name="w"/> x <paramref name="h"/> after a draw:transform.</summary>
    private static (double X, double Y, double Rotation) Transform(string transform, double x, double y, double w, double h)
    {
        double angle = 0, tx = 0, ty = 0;
        foreach (Match m in TransformPart.Matches(transform))
        {
            var tokens = m.Groups[2].Value.Split([' ', ',', '\t'], StringSplitOptions.RemoveEmptyEntries);
            double Arg(int i) => i < tokens.Length ? OdfUnits.Length(tokens[i]) ?? 0 : 0;
            switch (m.Groups[1].Value)
            {
                case "rotate" when tokens.Length > 0:
                    angle = double.TryParse(tokens[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var rad) ? rad : 0;
                    break;
                case "translate": (tx, ty) = (Arg(0), Arg(1)); break;
            }
        }
        // The shape is turned about its own top-left corner, then moved: its center lies at translate + R * (w/2, h/2).
        double cos = Math.Cos(angle), sin = Math.Sin(angle);
        double hx = w / 2, hy = h / 2;
        double cx = tx + x + hx * cos + hy * sin, cy = ty + y - hx * sin + hy * cos;
        return (cx - hx, cy - hy, angle * 180 / Math.PI);
    }

    private void ReadFrame(XElement frame, XElement? master, int depth)
    {
        var (x, y, w, h, rotation) = Box(frame, master);
        var cls = (string?)frame.Attribute(Ns.Presentation + "class");
        if (cls is "page" or "handout" or "header") return;
        var g = GraphicProps(frame, false);
        if (frame.Element(Ns.Table + "table") is { } table)
        {
            ReadTable(table, x, y, w, h);
        }
        else if (frame.Element(Ns.Draw + "image") is { } image)
        {
            var data = LoadImage((string?)image.Attribute(Ns.XLink + "href"), image);
            if (data is null || w <= 0 || h <= 0) return;
            double cl = 0, ct = 0, cr = 0, cb = 0;
            if (g.GetValueOrDefault("graphic-properties/clip") is { } clip && clip.StartsWith("rect(", StringComparison.Ordinal) && data.PixelSize is { } px)
            {
                var parts = clip[5..].TrimEnd(')').Split([',', ' '], StringSplitOptions.RemoveEmptyEntries).Select(OdfUnits.Length).ToArray();
                double nw = px.Width * 72 / px.Dpi, nh = px.Height * 72 / px.Dpi;
                if (parts.Length == 4 && nw > 0 && nh > 0)
                    (ct, cr, cb, cl) = ((parts[0] ?? 0) / nh, (parts[1] ?? 0) / nw, (parts[2] ?? 0) / nh, (parts[3] ?? 0) / nw);
            }
            _slide.Objects.Add(new SheetImage
            {
                Image = data, Anchor = ImageAnchorKind.Absolute, X = x, Y = y, Width = w, Height = h, Rotation = rotation, ZOrder = _z++,
                CropLeft = cl, CropTop = ct, CropRight = cr, CropBottom = cb,
            });
        }
        else if (frame.Element(Ns.Draw + "text-box") is { } box)
        {
            var paragraphs = ReadText(box, frame, g, cls);
            bool hasFill = g.GetValueOrDefault("graphic-properties/fill") is "solid" or "gradient" or "bitmap" or "hatch";
            bool hasLine = g.GetValueOrDefault("graphic-properties/stroke") is "solid" or "dash";
            if (paragraphs.Count == 0 && !hasFill && !hasLine) return;
            AddShape(frame, new SheetShape { Geometry = ShapeGeometry.Rectangle }, g, x, y, w, h, rotation, paragraphs, wrap: true);
        }
        else if (frame.Element(Ns.Draw + "object") is not null || frame.Element(Ns.Draw + "object-ole") is not null)
        {
            Warn("odp.object-skipped", "A chart or embedded object on a slide is not drawn.");
        }
    }

    private void ReadShape(XElement e)
    {
        var (x, y, w, h, rotation) = Box(e, null);
        var g = GraphicProps(e, true);
        var shape = new SheetShape { Geometry = ShapeGeometry.Rectangle };
        bool wrap = g.GetValueOrDefault("graphic-properties/wrap-option") != "no-wrap";
        switch (e.Name.LocalName)
        {
            case "ellipse" or "circle": shape = shape with { Geometry = ShapeGeometry.Ellipse }; break;
            case "rect":
                if (Attr(e, Ns.Svg + "rx") is > 0 and var rx || Attr(e, Ns.Draw + "corner-radius") is > 0 and var rx2 && (rx = rx2) > 0)
                    shape = shape with { Geometry = ShapeGeometry.RoundRectangle, CornerRadius = Math.Clamp(rx / Math.Max(1, Math.Min(w, h)), 0, 0.5) };
                break;
            case "line" or "connector" or "measure":
            {
                double x1 = Attr(e, Ns.Svg + "x1"), y1 = Attr(e, Ns.Svg + "y1"), x2 = Attr(e, Ns.Svg + "x2"), y2 = Attr(e, Ns.Svg + "y2");
                shape = shape with { Geometry = ShapeGeometry.Line, FlipH = x2 < x1, FlipV = y2 < y1 };
                break;
            }
            case "polygon" or "polyline":
            {
                var path = PointsPath(e);
                if (path is null) return;
                shape = shape with { Geometry = ShapeGeometry.Custom, Paths = [path with { Closed = e.Name.LocalName == "polygon", Fill = e.Name.LocalName == "polygon" }] };
                break;
            }
            case "path":
            {
                var paths = SvgPath(e, (string?)e.Attribute(Ns.Svg + "d"), (string?)e.Attribute(Ns.Svg + "viewBox"));
                if (paths is null) return;
                shape = shape with { Geometry = ShapeGeometry.Custom, Paths = paths };
                break;
            }
            case "custom-shape":
                shape = CustomShape(e, shape);
                break;
        }
        var paragraphs = e.Name.LocalName is "line" or "connector" or "measure" ? [] : ReadText(e, e, g, null);
        AddShape(e, shape, g, x, y, w, h, rotation, paragraphs, wrap);
    }

    private SheetShape CustomShape(XElement e, SheetShape shape)
    {
        var geometry = e.Element(Ns.Draw + "enhanced-geometry");
        var type = (string?)geometry?.Attribute(Ns.Draw + "type") ?? "rectangle";
        ShapeGeometry? mapped = type switch
        {
            "rectangle" => ShapeGeometry.Rectangle,
            "round-rectangle" or "round-quadrangle" => ShapeGeometry.RoundRectangle,
            "ellipse" => ShapeGeometry.Ellipse,
            "diamond" => ShapeGeometry.Diamond,
            "isosceles-triangle" or "triangle" => ShapeGeometry.Triangle,
            "right-triangle" => ShapeGeometry.RightTriangle,
            "parallelogram" => ShapeGeometry.Parallelogram,
            "trapezoid" => ShapeGeometry.Trapezoid,
            "hexagon" => ShapeGeometry.Hexagon,
            "octagon" => ShapeGeometry.Octagon,
            "cross" or "mso-spt11" => ShapeGeometry.Plus,
            "right-arrow" => ShapeGeometry.RightArrow,
            "left-arrow" => ShapeGeometry.LeftArrow,
            "up-arrow" => ShapeGeometry.UpArrow,
            "down-arrow" => ShapeGeometry.DownArrow,
            "pentagon" or "regular-pentagon" or "mso-spt56" => ShapeGeometry.Pentagon,
            "pentagon-right" => ShapeGeometry.HomePlate,
            "chevron" => ShapeGeometry.Chevron,
            "star5" => ShapeGeometry.Star5,
            _ => null,
        };
        shape = shape with { FlipH = (string?)geometry?.Attribute(Ns.Draw + "mirror-horizontal") == "true", FlipV = (string?)geometry?.Attribute(Ns.Draw + "mirror-vertical") == "true" };
        if (mapped is { } m) return shape with { Geometry = m, FlipV = shape.FlipV ^ (m == ShapeGeometry.Trapezoid) };
        // Other geometry: the preset's drawn path when it is a plain one, else the bounding rectangle.
        var path = geometry is null ? null : EnhancedPath(geometry);
        if (path is not null) return shape with { Geometry = ShapeGeometry.Custom, Paths = path };
        Warn("odp.geometry:" + type, $"The shape geometry '{type}' is not supported; it is drawn as a rectangle.");
        return shape;
    }

    private static readonly Regex Number_ = new(@"-?\d+(\.\d+)?([eE][-+]?\d+)?", RegexOptions.Compiled);

    /// <summary>A custom shape's <c>draw:enhanced-path</c> of lines and curves in literal numbers (no formulas), as fractions of its viewBox.</summary>
    private static IReadOnlyList<ShapePath>? EnhancedPath(XElement geometry)
    {
        var d = (string?)geometry.Attribute(Ns.Draw + "enhanced-path");
        var vb = (string?)geometry.Attribute(Ns.Svg + "viewBox");
        if (d is null || vb is null || d.Contains('?') || d.Contains('$')) return null;
        return SvgPathCore(d, vb, enhanced: true);
    }

    private IReadOnlyList<ShapePath>? SvgPath(XElement e, string? d, string? viewBox) =>
        d is null || viewBox is null ? null : SvgPathCore(d, viewBox, enhanced: false);

    private static IReadOnlyList<ShapePath>? SvgPathCore(string d, string viewBox, bool enhanced)
    {
        var vb = viewBox.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries).Select(s => double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : double.NaN).ToArray();
        if (vb.Length != 4 || vb.Any(double.IsNaN) || vb[2] <= 0 || vb[3] <= 0) return null;
        var subpaths = new List<(List<(double, double)> Points, bool Closed)>();
        List<(double, double)>? current = null;
        double cx = 0, cy = 0;
        int i = 0;
        // Tokens: command letters and numbers.
        var tokens = Regex.Matches(d, @"[A-Za-z]|-?\d+(\.\d+)?([eE][-+]?\d+)?").Select(m => m.Value).ToList();
        if (tokens.Count > 200_000) return null;
        char cmd = ' ';
        double Next() => i < tokens.Count && double.TryParse(tokens[i++], NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0;
        bool Numeric() => i < tokens.Count && !char.IsLetter(tokens[i][0]);
        void Start(double x, double y) { current = [(x, y)]; subpaths.Add((current, false)); cx = x; cy = y; }
        void LineTo(double x, double y) { if (current is null) Start(cx, cy); current!.Add((x, y)); cx = x; cy = y; }
        while (i < tokens.Count)
        {
            if (char.IsLetter(tokens[i][0])) cmd = tokens[i++][0];
            else if (cmd is 'M') cmd = 'L';
            switch (cmd)
            {
                case 'M': { double x = Next(), y = Next(); Start(x, y); cmd = 'L'; break; }
                case 'L': { double x = Next(), y = Next(); LineTo(x, y); break; }
                case 'H': LineTo(Next(), cy); break;
                case 'V': LineTo(cx, Next()); break;
                case 'C':
                {
                    double x1 = Next(), y1 = Next(), x2 = Next(), y2 = Next(), x = Next(), y = Next();
                    if (current is null) Start(cx, cy);
                    double x0 = cx, y0 = cy;
                    for (int s = 1; s <= 12; s++)
                    {
                        double t = s / 12.0, u = 1 - t;
                        current!.Add((u * u * u * x0 + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x, u * u * u * y0 + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y));
                    }
                    (cx, cy) = (x, y);
                    break;
                }
                case 'Z' or 'z':
                    if (subpaths.Count > 0) subpaths[^1] = (subpaths[^1].Points, true);
                    current = null;
                    break;
                case 'N' or 'S' or 'F' or 'U' or 'W' or 'V' when enhanced: break; // end of path, no-stroke, no-fill markers
                default: return null; // arcs and other commands are not supported
            }
            if (cmd is 'Z' or 'z' or 'N' or 'S' or 'F') { if (!Numeric()) continue; }
        }
        var result = new List<ShapePath>();
        var rings = new List<IReadOnlyList<(double X, double Y)>>();
        (double X, double Y) Norm((double X, double Y) p) => ((p.X - vb[0]) / vb[2], (p.Y - vb[1]) / vb[3]);
        var normalized = subpaths.Where(s => s.Points.Count >= 2).Select(s => (Points: (IReadOnlyList<(double X, double Y)>)s.Points.Select(Norm).ToList(), s.Closed)).ToList();
        if (normalized.Count == 0) return null;
        var first = normalized[0];
        var more = normalized.Skip(1).Where(s => s.Closed).Select(s => s.Points).ToList();
        result.Add(new ShapePath(first.Points, first.Closed) { Rings = more.Count > 0 ? more : null });
        foreach (var open in normalized.Skip(1).Where(s => !s.Closed)) result.Add(new ShapePath(open.Points, false, Fill: false));
        return result;
    }

    private ShapePath? PointsPath(XElement e)
    {
        var vb = (string?)e.Attribute(Ns.Svg + "viewBox");
        var pts = (string?)e.Attribute(Ns.Draw + "points");
        if (vb is null || pts is null) return null;
        var v = vb.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries).Select(s => double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : double.NaN).ToArray();
        if (v.Length != 4 || v.Any(double.IsNaN) || v[2] <= 0 || v[3] <= 0) return null;
        var points = new List<(double, double)>();
        foreach (var pair in pts.Split([' '], StringSplitOptions.RemoveEmptyEntries))
        {
            var xy = pair.Split(',');
            if (xy.Length == 2 && double.TryParse(xy[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var px) && double.TryParse(xy[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var py))
                points.Add(((px - v[0]) / v[2], (py - v[1]) / v[3]));
            if (points.Count > 100_000) break;
        }
        return points.Count >= 2 ? new ShapePath(points, true) : null;
    }

    private void AddShape(XElement e, SheetShape shape, Dictionary<string, string> g, double x, double y, double w, double h, double rotation,
        List<ShapeParagraph> paragraphs, bool wrap)
    {
        string? P(string key) => g.GetValueOrDefault("graphic-properties/" + key);
        Color? fill = null;
        GradientFill? gradient = null;
        ImageData? fillImage = null;
        ShapePattern? pattern = null;
        var solid = OdfUnits.Hex(P("fill-color"));
        switch (P("fill"))
        {
            case "solid": fill = solid; break;
            case "gradient": gradient = Gradient(P("fill-gradient-name")); fill = gradient?.Stops.FirstOrDefault().Color; break;
            case "bitmap": fillImage = FillImage(P("fill-image-name")); break;
            case "hatch":
                pattern = Hatch(P("fill-hatch-name"), P("fill-hatch-solid") == "true" ? solid : null);
                fill = pattern?.Foreground;
                break;
        }
        if (fill is not null && P("opacity") is { } op && double.TryParse(op.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var opacity) && opacity < 100)
            fill = Color.White.Blend(fill.Value, Math.Clamp(opacity / 100, 0, 1)); // a translucent fill over the usual white page

        Color? line = null;
        double lineWidth = 1;
        var style = StrokeStyle.Solid;
        if (P("stroke") is "solid" or "dash")
        {
            line = OdfUnits.Hex(P("stroke-color")) ?? Color.Black;
            lineWidth = OdfUnits.Length(P("stroke-width")) is { } sw and > 0 ? sw : 0.75;
            if (P("stroke") == "dash")
            {
                var name = P("stroke-dash") ?? "";
                style = name.Contains("ot", StringComparison.OrdinalIgnoreCase) && !name.Contains("ash", StringComparison.OrdinalIgnoreCase) ? StrokeStyle.Dotted : StrokeStyle.Dashed;
            }
        }
        LineEnd End(string key)
        {
            if (P(key) is not { Length: > 0 } marker) return default;
            double width = OdfUnits.Length(P(key + "-width")) ?? lineWidth * 3;
            double scale = Math.Clamp(width / Math.Max(0.5, lineWidth), 2, 8);
            return new LineEnd(marker.Contains("ircle", StringComparison.Ordinal) || marker.Contains("val", StringComparison.Ordinal) ? LineEndKind.Oval
                : marker.Contains("iamond", StringComparison.Ordinal) ? LineEndKind.Diamond : marker.Contains("tealth", StringComparison.Ordinal) ? LineEndKind.Stealth : LineEndKind.Triangle, scale, scale);
        }
        ShapeShadow? shadow = null;
        if (P("shadow") == "visible")
        {
            double shadowOpacity = double.TryParse(P("shadow-opacity")?.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var so) ? so / 100 : 1;
            shadow = new ShapeShadow(OdfUnits.Hex(P("shadow-color")) ?? Color.Black, Math.Clamp(shadowOpacity, 0, 1), OdfUnits.Length(P("shadow-offset-x")) ?? 3, OdfUnits.Length(P("shadow-offset-y")) ?? 3);
        }
        double Pad(string side, double d) => OdfUnits.Length(P("padding-" + side)) ?? d;
        var anchor = P("textarea-vertical-align") switch { "middle" => ShapeTextAnchor.Middle, "bottom" => ShapeTextAnchor.Bottom, _ => ShapeTextAnchor.Top };
        // Shapes' text is centered in the shape unless the style says otherwise.
        if (P("textarea-vertical-align") is null && e.Name.LocalName is "custom-shape" or "ellipse" or "circle" && paragraphs.Count > 0) anchor = ShapeTextAnchor.Middle;
        bool vertical = g.GetValueOrDefault("paragraph-properties/writing-mode") is "tb-rl" or "tb";

        _slide.Objects.Add(shape with
        {
            Anchor = ImageAnchorKind.Absolute, X = x, Y = y, Width = w, Height = h, Rotation = rotation, ZOrder = _z++,
            Fill = fill, FillGradient = gradient, FillImage = fillImage, FillPattern = pattern,
            LineColor = line, LineWidth = lineWidth, LineStyle = style,
            HeadEnd = End("marker-start"), TailEnd = End("marker-end"), Shadow = shadow,
            Paragraphs = paragraphs, TextAnchor = anchor, WrapText = wrap,
            TextInsets = new Thickness(Pad("top", 3.6), Pad("right", 7.2), Pad("bottom", 3.6), Pad("left", 7.2)),
            TextRotation = vertical ? -90 : 0,
        });
    }
}
