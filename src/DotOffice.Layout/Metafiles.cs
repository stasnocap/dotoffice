using System.Buffers.Binary;
using System.Text;
using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>
/// Draws Windows metafiles (EMF and WMF pictures) as layout items: their GDI records are played into paths, text runs and bitmaps
/// fitted to the picture's box, so the renderer only sees primitives it already draws.
/// </summary>
internal static class Metafiles
{
    /// <summary>Replaces every metafile picture on the pages by its drawing; pictures that cannot be played stay (the renderer draws a placeholder).</summary>
    public static void Expand(LayoutDocument document, FontSelector fonts, IConversionWarningSink? warnings)
    {
        Dictionary<ImageData, Drawing?> cache = new(ReferenceEqualityComparer.Instance);
        foreach (var page in document.Pages)
        {
            var items = page.Items;
            for (int i = 0; i < items.Count; i++)
            {
                if (items[i] is not ImageItem { Image.Format: ImageFormat.Emf or ImageFormat.Wmf } image || image.Width <= 0 || image.Height <= 0) continue;
                if (!cache.TryGetValue(image.Image, out var drawing))
                {
                    drawing = Play(image.Image, fonts, warnings);
                    cache[image.Image] = drawing;
                }
                if (drawing is null) continue;
                var expanded = new List<LayoutItem>(drawing.Items.Count + 2)
                {
                    new PushItem((image.X, image.Y, image.Width, image.Height)) { ClipOutline = image.Clip },
                };
                // The drawing is played into a unit box; scale it to this picture's box.
                foreach (var item in drawing.Items) expanded.Add(Fit(item, image.X, image.Y, image.Width / UnitBox, image.Height / drawing.Height));
                expanded.Add(new PopItem());
                items.RemoveAt(i);
                items.InsertRange(i, expanded);
                // The drawing is scanned in turn, so pictures nested in it (EMF+ metafile images) are drawn as well.
            }
        }
    }

    /// <summary>A played metafile: items in a box <see cref="UnitBox"/> wide and <see cref="Height"/> high (the picture's own proportions).</summary>
    private sealed record Drawing(List<LayoutItem> Items, double Height);

    /// <summary>The width of the box metafiles are played into; items are scaled from it to the picture's box.</summary>
    public const double UnitBox = 1000;

    private static LayoutItem Fit(LayoutItem item, double x, double y, double sx, double sy)
    {
        double s = Math.Sqrt(sx * sy);
        (double, double) P(double px, double py) => (x + px * sx, y + py * sy);
        switch (item)
        {
            case PathItem path:
                return path with
                {
                    Points = path.Points.Select(p => { var (fx, fy) = P(p.X, p.Y); return p with { X = fx, Y = fy }; }).ToList(),
                    LineWidth = path.LineWidth * s,
                };
            case TextItem text:
            {
                var (fx, fy) = P(text.X, text.Baseline);
                return text with { X = fx, Baseline = fy, FontSize = text.FontSize * sy, CharacterSpacing = text.CharacterSpacing * sx };
            }
            case ImageItem picture:
            {
                var (fx, fy) = P(picture.X, picture.Y);
                return picture with { X = fx, Y = fy, Width = picture.Width * sx, Height = picture.Height * sy };
            }
            case LineItem line:
            {
                var (x1, y1) = P(line.X1, line.Y1);
                var (x2, y2) = P(line.X2, line.Y2);
                return line with { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Width = line.Width * s };
            }
            case PushItem push:
            {
                var (cx, cy) = P(push.CenterX, push.CenterY);
                return push with
                {
                    CenterX = cx, CenterY = cy,
                    ClipOutline = push.ClipOutline?.Select(q => double.IsNaN(q.X) ? q : P(q.X, q.Y)).ToList(),
                    Clip = push.Clip is { } c ? (x + c.X * sx, y + c.Y * sy, c.Width * sx, c.Height * sy) : null,
                };
            }
            case GradientItem gradient:
            {
                // The blend direction is a normal to the lines of equal color, so it scales by the inverse of the box's stretch.
                double radians = gradient.Degree * Math.PI / 180;
                var (fx, fy) = P(gradient.X, gradient.Y);
                return gradient with
                {
                    X = fx, Y = fy, Width = gradient.Width * sx, Height = gradient.Height * sy,
                    Degree = gradient.Radial ? gradient.Degree : Math.Atan2(Math.Sin(radians) / sy, Math.Cos(radians) / sx) * 180 / Math.PI,
                    Clip = gradient.Clip?.Select(q => P(q.X, q.Y)).ToList(),
                };
            }
            default:
                return item;
        }
    }

    private static Drawing? Play(ImageData image, FontSelector fonts, IConversionWarningSink? warnings)
    {
        var player = new GdiPlayer(fonts);
        try
        {
            // A dual EMF draws its GDI records; one that draws only through EMF+ records (or whose GDI records draw nothing) its EMF+ records.
            bool played = image.Format == ImageFormat.Emf
                ? EmfReader.Play(image.Data, player, EmfPlusPlayer.PlusOnly(image.Data) ? new EmfPlusPlayer(player, fonts) : null)
                : WmfReader.Play(image.Data, player);
            if (!played)
            {
                player = new GdiPlayer(fonts);
                EmfReader.Play(image.Data, player, new EmfPlusPlayer(player, fonts));
            }
            player.Finish();
        }
        catch (Exception e) when (e is ArgumentOutOfRangeException or IndexOutOfRangeException or InvalidDataException or OverflowException)
        {
            warnings.Warn("metafile.invalid", $"Metafile '{image.Name}' could not be read: {e.Message}");
            return null;
        }
        return new Drawing(player.Items, player.BoxHeight);
    }
}

/// <summary>Drawing items with nested clip groups that follow the clip each item is drawn with.</summary>
internal sealed class ClipGroups
{
    public List<LayoutItem> Items { get; } = [];

    /// <summary>The clip groups currently open in <see cref="Items"/>, outermost first.</summary>
    private readonly List<IReadOnlyList<(double X, double Y)>> _open = [];

    /// <summary>Adds a drawing item, first opening or closing clip groups so the open ones are the outlines of <paramref name="clip"/>.</summary>
    public void Add(LayoutItem item, IReadOnlyList<IReadOnlyList<(double X, double Y)>> clip)
    {
        int common = 0;
        while (common < _open.Count && common < clip.Count && ReferenceEquals(_open[common], clip[common])) common++;
        for (int k = _open.Count; k > common; k--) Items.Add(new PopItem());
        _open.RemoveRange(common, _open.Count - common);
        for (int k = common; k < clip.Count; k++)
        {
            Items.Add(new PushItem(null) { ClipOutline = clip[k] });
            _open.Add(clip[k]);
        }
        Items.Add(item);
    }

    /// <summary>An outline for clipping from path points: curves flattened, each figure a ring (rings separated by NaN points).</summary>
    public static List<(double X, double Y)> Flatten(IReadOnlyList<PathPoint> path)
    {
        var outline = new List<(double X, double Y)>();
        (double X, double Y) last = (0, 0);
        for (int k = 0; k < path.Count; k++)
        {
            var p = path[k];
            switch (p.Kind)
            {
                case PathPointKind.Close: continue;
                case PathPointKind.Move:
                    if (outline.Count > 0) outline.Add((double.NaN, double.NaN));
                    outline.Add(last = (p.X, p.Y));
                    break;
                case PathPointKind.Curve when k + 2 < path.Count:
                {
                    var (c1, c2, e) = (path[k], path[k + 1], path[k + 2]);
                    for (int step = 1; step <= 16; step++)
                    {
                        double u = step / 16.0, v = 1 - u;
                        outline.Add((v * v * v * last.X + 3 * v * v * u * c1.X + 3 * v * u * u * c2.X + u * u * u * e.X,
                                     v * v * v * last.Y + 3 * v * v * u * c1.Y + 3 * v * u * u * c2.Y + u * u * u * e.Y));
                    }
                    last = (e.X, e.Y);
                    k += 2;
                    break;
                }
                default:
                    outline.Add(last = (p.X, p.Y));
                    break;
            }
        }
        return outline;
    }

    /// <summary>Closes the clip groups still open.</summary>
    public void Finish()
    {
        foreach (var _ in _open) Items.Add(new PopItem());
        _open.Clear();
    }
}

/// <summary>A 2D affine transform: x' = A x + C y + E, y' = B x + D y + F.</summary>
internal readonly record struct Affine(double A, double B, double C, double D, double E, double F)
{
    public static Affine Identity { get; } = new(1, 0, 0, 1, 0, 0);

    public (double X, double Y) Apply(double x, double y) => (A * x + C * y + E, B * x + D * y + F);

    /// <summary>This transform followed by <paramref name="next"/>.</summary>
    public Affine Then(Affine next) => new(
        next.A * A + next.C * B, next.B * A + next.D * B,
        next.A * C + next.C * D, next.B * C + next.D * D,
        next.A * E + next.C * F + next.E, next.B * E + next.D * F + next.F);

    /// <summary>The mean length a unit is scaled to.</summary>
    public double Scale => Math.Sqrt(Math.Abs(A * D - B * C));
}

internal sealed record GdiPen(int Style, double Width, Color Color, bool Cosmetic)
{
    public bool IsNull => (Style & 0xF) == 5;
}

internal sealed record GdiBrush(Color? Color);

internal sealed record GdiFont(int Height, int Escapement, int Weight, bool Italic, bool Underline, bool Strikeout, string Face);

/// <summary>GDI's drawing state and primitives, shared by the EMF and WMF readers; output is in a box 1000 wide with the picture's proportions.</summary>
internal sealed class GdiPlayer(FontSelector fonts)
{
    public List<LayoutItem> Items => Groups.Items;

    /// <summary>The output and its open clip groups; an EMF+ player drawing into the same picture shares them.</summary>
    public ClipGroups Groups { get; } = new();

    private void Add(LayoutItem item) => Groups.Add(item, _state.Clip);

    /// <summary>Closes the clip groups still open; call once the metafile has been played.</summary>
    public void Finish() => Groups.Finish();

    /// <summary>The height of the output box, which is <see cref="Metafiles.UnitBox"/> wide and has the picture's proportions.</summary>
    public double BoxHeight { get; set; } = Metafiles.UnitBox;

    private sealed class State
    {
        public GdiPen Pen = new(0, 0, Color.Black, true);
        public GdiBrush Brush = new(Color.White);
        public GdiFont Font = new(-12, 0, 400, false, false, false, "Arial");
        public Color TextColor = Color.Black, BackColor = Color.White;
        public bool Opaque = true, Winding, Clockwise;
        public int TextAlign, MapMode = 1;
        public (double X, double Y) WindowOrg, ViewportOrg, WindowExt = (1, 1), ViewportExt = (1, 1), Position;
        public Affine World = Affine.Identity;
        /// <summary>The clip: the intersection of these outlines (box coordinates; rings split by NaN points are even-odd).</summary>
        public IReadOnlyList<(double X, double Y)>[] Clip = [];

        public State Copy() => (State)MemberwiseClone();
    }

    private State _state = new();
    private readonly Stack<State> _saved = new();
    private readonly Dictionary<int, object> _objects = [];
    private List<PathPoint>? _path;
    private List<PathPoint>? _finishedPath;

    /// <summary>Device pixels to the output box; set by the reader from the header.</summary>
    public Affine DeviceToBox { get; set; } = Affine.Identity;

    /// <summary>The width of a device pixel in the output box (cosmetic pens draw one pixel wide).</summary>
    public double PixelWidth { get; set; } = 1;

    /// <summary>Device pixels per millimetre, for the fixed mapping modes.</summary>
    public (double X, double Y) PixelsPerMm { get; set; } = (96 / 25.4, 96 / 25.4);

    /// <summary>
    /// For WMF, which has no device: the picture's bounds in logical units (from the placeable header). Once the file sets its window
    /// extent the window maps to the box instead.
    /// </summary>
    public (double L, double T, double W, double H)? WmfFrame { get; set; }

    private bool _windowSet;

    private Affine PageToDevice()
    {
        var s = _state;
        if (WmfFrame is { } frame)
        {
            var (l, t, w, h) = _windowSet ? (s.WindowOrg.X, s.WindowOrg.Y, s.WindowExt.X, s.WindowExt.Y) : frame;
            if (w == 0) w = 1;
            if (h == 0) h = 1;
            return new Affine(1000 / w, 0, 0, BoxHeight / h, -l * 1000 / w, -t * BoxHeight / h);
        }
        double sx, sy;
        switch (s.MapMode)
        {
            case 7 or 8:
                sx = s.WindowExt.X == 0 ? 1 : s.ViewportExt.X / s.WindowExt.X;
                sy = s.WindowExt.Y == 0 ? 1 : s.ViewportExt.Y / s.WindowExt.Y;
                if (s.MapMode == 7)
                {
                    double m = Math.Min(Math.Abs(sx), Math.Abs(sy));
                    sx = Math.CopySign(m, sx);
                    sy = Math.CopySign(m, sy);
                }
                break;
            case >= 2 and <= 6:
                double mm = s.MapMode switch { 2 => 0.1, 3 => 0.01, 4 => 0.254, 5 => 0.0254, _ => 25.4 / 1440 };
                (sx, sy) = (mm * PixelsPerMm.X, -mm * PixelsPerMm.Y);
                break;
            default:
                (sx, sy) = (1, 1);
                break;
        }
        return new Affine(sx, 0, 0, sy, s.ViewportOrg.X - s.WindowOrg.X * sx, s.ViewportOrg.Y - s.WindowOrg.Y * sy);
    }

    private Affine Transform => _state.World.Then(PageToDevice()).Then(DeviceToBox);

    // ---- State ----

    public void SetMapMode(int mode) => _state.MapMode = mode;
    public void SetWindowOrg(double x, double y) => _state.WindowOrg = (x, y);
    public void SetWindowExt(double x, double y)
    {
        if (x == 0 || y == 0) return;
        _state.WindowExt = (x, y);
        _windowSet = true;
    }
    public void SetViewportOrg(double x, double y) => _state.ViewportOrg = (x, y);
    public void SetViewportExt(double x, double y) { if (x != 0 && y != 0) _state.ViewportExt = (x, y); }
    public void OffsetWindowOrg(double x, double y) => _state.WindowOrg = (_state.WindowOrg.X + x, _state.WindowOrg.Y + y);
    public void OffsetViewportOrg(double x, double y) => _state.ViewportOrg = (_state.ViewportOrg.X + x, _state.ViewportOrg.Y + y);
    public void ScaleWindowExt(double xNum, double xDenom, double yNum, double yDenom)
    {
        if (xDenom != 0 && yDenom != 0) SetWindowExt(_state.WindowExt.X * xNum / xDenom, _state.WindowExt.Y * yNum / yDenom);
    }
    public void ScaleViewportExt(double xNum, double xDenom, double yNum, double yDenom)
    {
        if (xDenom != 0 && yDenom != 0) SetViewportExt(_state.ViewportExt.X * xNum / xDenom, _state.ViewportExt.Y * yNum / yDenom);
    }
    public void SetWorld(Affine world) => _state.World = world;
    public void ModifyWorld(Affine xform, int mode) => _state.World = mode switch
    {
        1 => Affine.Identity,
        2 => xform.Then(_state.World),
        3 => _state.World.Then(xform),
        4 => xform,
        _ => _state.World,
    };
    public void SetTextColor(Color c) => _state.TextColor = c;
    public void SetBackColor(Color c) => _state.BackColor = c;
    public void SetBackMode(int mode) => _state.Opaque = mode == 2;
    public void SetTextAlign(int align) => _state.TextAlign = align;
    public void SetPolyFillMode(int mode) => _state.Winding = mode == 2;
    public void SetArcDirection(int direction) => _state.Clockwise = direction == 2;
    public void MoveTo(double x, double y) => _state.Position = (x, y);

    public void Save() => _saved.Push(_state.Copy());

    public void Restore(int relative)
    {
        // Negative: that many levels back; positive: an absolute level (1 is the first saved state).
        int pops = relative < 0 ? -relative : _saved.Count - relative + 1;
        for (int k = 0; k < pops && _saved.Count > 0; k++) _state = _saved.Pop();
    }

    // ---- Clipping ----

    /// <summary>IntersectClipRect: the clip is narrowed to a logical rectangle.</summary>
    public void IntersectClip(double l, double t, double r, double b)
    {
        var m = Transform;
        IntersectClip([m.Apply(l, t), m.Apply(r, t), m.Apply(r, b), m.Apply(l, b)]);
    }

    private void IntersectClip(IReadOnlyList<(double X, double Y)> outline) => _state.Clip = [.. _state.Clip, outline];

    /// <summary>Combines a region of device rectangles with the clip (RGN_AND 1, RGN_COPY 5); no rectangles with RGN_COPY resets it.</summary>
    public void SelectClipRegion(IReadOnlyList<(double L, double T, double R, double B)>? rects, int mode)
    {
        if (mode == 5 && rects is not { Count: > 0 }) { _state.Clip = []; return; }
        if (rects is not { Count: > 0 } || mode is not (1 or 5)) return;
        // A region's rectangles do not overlap, so they clip as the even-odd rings of one outline.
        var outline = new List<(double X, double Y)>();
        foreach (var (l, t, r, b) in rects)
        {
            if (outline.Count > 0) outline.Add((double.NaN, double.NaN));
            outline.AddRange([DeviceToBox.Apply(l, t), DeviceToBox.Apply(r, t), DeviceToBox.Apply(r, b), DeviceToBox.Apply(l, b)]);
        }
        if (mode == 5) _state.Clip = [];
        IntersectClip(outline);
    }

    /// <summary>SelectClipPath: the finished path (flattened) combines with the clip; only RGN_AND and RGN_COPY are played.</summary>
    public void SelectClipPath(int mode)
    {
        var path = _finishedPath;
        _finishedPath = null;
        if (path is not { Count: >= 3 } || mode is not (1 or 5)) return;
        var m = Transform;
        var outline = ClipGroups.Flatten(path.Select(p =>
        {
            if (p.Kind == PathPointKind.Close) return p;
            var (x, y) = m.Apply(p.X, p.Y);
            return p with { X = x, Y = y };
        }).ToList());
        if (mode == 5) _state.Clip = [];
        IntersectClip(outline);
    }

    // ---- Objects ----

    public void Create(int handle, object gdiObject) => _objects[handle] = gdiObject;
    public void Delete(int handle) => _objects.Remove(handle);

    public void Select(int handle)
    {
        if (_objects.TryGetValue(handle, out var o)) Select(o);
    }

    public void Select(object o)
    {
        switch (o)
        {
            case GdiPen pen: _state.Pen = pen; break;
            case GdiBrush brush: _state.Brush = brush; break;
            case GdiFont font: _state.Font = font; break;
        }
    }

    /// <summary>The stock object with this index (WHITE_BRUSH 0 … DC_PEN 19).</summary>
    public static object? Stock(int index) => index switch
    {
        0 => new GdiBrush(Color.White),
        1 => new GdiBrush(new Color(0xC0, 0xC0, 0xC0)),
        2 => new GdiBrush(new Color(0x80, 0x80, 0x80)),
        3 => new GdiBrush(new Color(0x40, 0x40, 0x40)),
        4 or 18 => new GdiBrush(Color.Black),
        5 => new GdiBrush(null),
        6 => new GdiPen(0, 0, Color.White, true),
        7 or 19 => new GdiPen(0, 0, Color.Black, true),
        8 => new GdiPen(5, 0, Color.Black, true),
        >= 10 and <= 17 => new GdiFont(-12, 0, 400, false, false, false, index is 11 or 16 ? "Courier New" : "Arial"),
        _ => null,
    };

    // ---- Paths ----

    public void BeginPath() => _path = [];
    public void EndPath() { _finishedPath = _path; _path = null; }
    public void AbortPath() { _path = null; _finishedPath = null; }

    public void CloseFigure()
    {
        if (_path is { Count: > 0 }) _path.Add(new PathPoint(0, 0, PathPointKind.Close));
    }

    public void FillPath(bool fill, bool stroke)
    {
        var path = _finishedPath ?? _path;
        _finishedPath = null;
        if (path is { Count: >= 2 }) Emit(path, fill, stroke);
    }

    // ---- Drawing ----

    /// <summary>Appends a figure (logical coordinates) to the open path, or draws it with the current pen and, when closed, brush.</summary>
    private void Figure(IReadOnlyList<PathPoint> logical, bool closed)
    {
        if (logical.Count < 2) return;
        if (_path is not null)
        {
            _path.AddRange(logical);
            if (closed) _path.Add(new PathPoint(0, 0, PathPointKind.Close));
            return;
        }
        var points = logical.ToList();
        if (closed) points.Add(new PathPoint(0, 0, PathPointKind.Close));
        Emit(points, closed, true);
    }

    private void Emit(List<PathPoint> logical, bool fill, bool stroke)
    {
        var m = Transform;
        var points = new List<PathPoint>(logical.Count);
        foreach (var p in logical)
        {
            if (p.Kind == PathPointKind.Close) { points.Add(p); continue; }
            var (x, y) = m.Apply(p.X, p.Y);
            points.Add(new PathPoint(x, y, p.Kind));
        }
        Color? fillColor = fill ? _state.Brush.Color : null;
        var pen = _state.Pen;
        Color? strokeColor = stroke && !pen.IsNull ? pen.Color : null;
        if (fillColor is null && strokeColor is null) return;
        // Cosmetic pens are a device pixel wide whatever the mapping.
        double width = pen.Cosmetic || pen.Width <= 0 ? PixelWidth : pen.Width * m.Scale;
        var style = (pen.Style & 0xF) switch { 1 or 3 or 4 => LineStyle.Dashed, 2 => LineStyle.Dotted, _ => LineStyle.Solid };
        Add(new PathItem(points, fillColor, strokeColor, Math.Max(width, 0.05), style, !_state.Winding));
    }

    private static PathPoint Pt(double x, double y, PathPointKind kind = PathPointKind.Line) => new(x, y, kind);

    public void Polygon(IReadOnlyList<(double X, double Y)> points) => Polyline(points, true);

    public void Polyline(IReadOnlyList<(double X, double Y)> points, bool closed = false)
    {
        if (points.Count < 2) return;
        Figure(points.Select((p, k) => Pt(p.X, p.Y, k == 0 ? PathPointKind.Move : PathPointKind.Line)).ToList(), closed);
    }

    /// <summary>Several polygons drawn as one shape (holes by the fill mode).</summary>
    public void PolyPolygon(IReadOnlyList<IReadOnlyList<(double X, double Y)>> polygons)
    {
        var all = new List<PathPoint>();
        foreach (var polygon in polygons)
        {
            if (polygon.Count < 2) continue;
            for (int k = 0; k < polygon.Count; k++) all.Add(Pt(polygon[k].X, polygon[k].Y, k == 0 ? PathPointKind.Move : PathPointKind.Line));
            all.Add(new PathPoint(0, 0, PathPointKind.Close));
        }
        if (all.Count == 0) return;
        if (_path is not null) { _path.AddRange(all); return; }
        Emit(all, true, true);
    }

    /// <summary>Lines from the current position through the points (LineTo, PolylineTo).</summary>
    public void PolylineTo(IReadOnlyList<(double X, double Y)> points)
    {
        if (points.Count == 0) return;
        var figure = new List<PathPoint>();
        if (_path is not { Count: > 0 }) figure.Add(Pt(_state.Position.X, _state.Position.Y, PathPointKind.Move));
        foreach (var p in points) figure.Add(Pt(p.X, p.Y));
        _state.Position = points[^1];
        if (_path is not null) { if (_path.Count == 0) _path.Add(Pt(figure[0].X, figure[0].Y, PathPointKind.Move)); _path.AddRange(figure.Where(f => f.Kind != PathPointKind.Move)); return; }
        Emit(figure, false, true);
    }

    /// <summary>Bézier segments; with <paramref name="fromCurrent"/> they start at the current position (PolyBezierTo).</summary>
    public void PolyBezier(IReadOnlyList<(double X, double Y)> points, bool fromCurrent)
    {
        var figure = new List<PathPoint>();
        int k = 0;
        if (fromCurrent)
        {
            if (_path is not { Count: > 0 }) figure.Add(Pt(_state.Position.X, _state.Position.Y, PathPointKind.Move));
        }
        else
        {
            if (points.Count < 4) return;
            figure.Add(Pt(points[0].X, points[0].Y, PathPointKind.Move));
            k = 1;
        }
        for (; k + 2 < points.Count; k += 3)
            for (int j = 0; j < 3; j++) figure.Add(Pt(points[k + j].X, points[k + j].Y, PathPointKind.Curve));
        if (points.Count > 0) _state.Position = points[^1];
        if (_path is not null)
        {
            if (fromCurrent && _path.Count == 0) _path.Add(Pt(_state.Position.X, _state.Position.Y, PathPointKind.Move));
            _path.AddRange(figure);
            return;
        }
        Emit(figure, false, true);
    }

    public void Rectangle(double left, double top, double right, double bottom) =>
        Figure([Pt(left, top, PathPointKind.Move), Pt(right, top), Pt(right, bottom), Pt(left, bottom)], true);

    public void RoundRect(double left, double top, double right, double bottom, double cornerWidth, double cornerHeight)
    {
        double rx = Math.Min(Math.Abs(cornerWidth) / 2, Math.Abs(right - left) / 2), ry = Math.Min(Math.Abs(cornerHeight) / 2, Math.Abs(bottom - top) / 2);
        if (rx <= 0 || ry <= 0) { Rectangle(left, top, right, bottom); return; }
        const double k = 0.5522847498;
        Figure([
            Pt(left + rx, top, PathPointKind.Move), Pt(right - rx, top),
            Pt(right - rx + k * rx, top, PathPointKind.Curve), Pt(right, top + ry - k * ry, PathPointKind.Curve), Pt(right, top + ry, PathPointKind.Curve),
            Pt(right, bottom - ry),
            Pt(right, bottom - ry + k * ry, PathPointKind.Curve), Pt(right - rx + k * rx, bottom, PathPointKind.Curve), Pt(right - rx, bottom, PathPointKind.Curve),
            Pt(left + rx, bottom),
            Pt(left + rx - k * rx, bottom, PathPointKind.Curve), Pt(left, bottom - ry + k * ry, PathPointKind.Curve), Pt(left, bottom - ry, PathPointKind.Curve),
            Pt(left, top + ry),
            Pt(left, top + ry - k * ry, PathPointKind.Curve), Pt(left + rx - k * rx, top, PathPointKind.Curve), Pt(left + rx, top, PathPointKind.Curve),
        ], true);
    }

    public void Ellipse(double left, double top, double right, double bottom)
    {
        double cx = (left + right) / 2, cy = (top + bottom) / 2, rx = (right - left) / 2, ry = (bottom - top) / 2;
        var figure = new List<PathPoint> { Pt(cx + rx, cy, PathPointKind.Move) };
        AppendArc(figure, cx, cy, rx, ry, 0, 2 * Math.PI);
        Figure(figure, true);
    }

    /// <summary>Kind 0 is an open arc, 1 a chord (closed by a straight line), 2 a pie (closed through the center).</summary>
    public void Arc(double left, double top, double right, double bottom, double x1, double y1, double x2, double y2, int kind)
    {
        double cx = (left + right) / 2, cy = (top + bottom) / 2, rx = Math.Abs(right - left) / 2, ry = Math.Abs(bottom - top) / 2;
        if (rx <= 0 || ry <= 0) return;
        // Angles in the usual counter-clockwise sense on screen (y grows downwards in logical space).
        double a1 = Math.Atan2(-(y1 - cy) / ry, (x1 - cx) / rx), a2 = Math.Atan2(-(y2 - cy) / ry, (x2 - cx) / rx);
        double sweep = _state.Clockwise ? -(((a1 - a2) % (2 * Math.PI) + 2 * Math.PI) % (2 * Math.PI)) : ((a2 - a1) % (2 * Math.PI) + 2 * Math.PI) % (2 * Math.PI);
        if (Math.Abs(sweep) < 1e-9) sweep = _state.Clockwise ? -2 * Math.PI : 2 * Math.PI;
        var figure = new List<PathPoint>();
        if (kind == 2) { figure.Add(Pt(cx, cy, PathPointKind.Move)); figure.Add(Pt(cx + rx * Math.Cos(a1), cy - ry * Math.Sin(a1))); }
        else figure.Add(Pt(cx + rx * Math.Cos(a1), cy - ry * Math.Sin(a1), PathPointKind.Move));
        AppendArc(figure, cx, cy, rx, ry, a1, sweep);
        Figure(figure, kind != 0);
    }

    /// <summary>Bézier segments of at most a quarter turn along an ellipse arc from <paramref name="start"/> through <paramref name="sweep"/> radians.</summary>
    private static void AppendArc(List<PathPoint> figure, double cx, double cy, double rx, double ry, double start, double sweep)
    {
        int segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(sweep) / (Math.PI / 2) - 1e-9));
        double step = sweep / segments, alpha = 4.0 / 3 * Math.Tan(step / 4);
        for (int s = 0; s < segments; s++)
        {
            double t1 = start + s * step, t2 = t1 + step;
            double c1 = Math.Cos(t1), s1 = Math.Sin(t1), c2 = Math.Cos(t2), s2 = Math.Sin(t2);
            figure.Add(Pt(cx + rx * (c1 - alpha * s1), cy - ry * (s1 + alpha * c1), PathPointKind.Curve));
            figure.Add(Pt(cx + rx * (c2 + alpha * s2), cy - ry * (s2 - alpha * c2), PathPointKind.Curve));
            figure.Add(Pt(cx + rx * c2, cy - ry * s2, PathPointKind.Curve));
        }
    }

    /// <summary>Fills rectangles (a region) with a brush, without outlines.</summary>
    public void FillRects(IEnumerable<(double L, double T, double R, double B)> rects, GdiBrush? brush = null)
    {
        var color = (brush ?? _state.Brush).Color;
        if (color is null) return;
        var m = Transform;
        foreach (var (l, t, r, b) in rects)
        {
            var points = new List<PathPoint>();
            foreach (var (x, y, k) in new[] { (l, t, PathPointKind.Move), (r, t, PathPointKind.Line), (r, b, PathPointKind.Line), (l, b, PathPointKind.Line) })
            {
                var (px, py) = m.Apply(x, y);
                points.Add(new PathPoint(px, py, k));
            }
            points.Add(new PathPoint(0, 0, PathPointKind.Close));
            Add(new PathItem(points, color, null));
        }
    }

    public void FillRegion(IEnumerable<(double L, double T, double R, double B)> rects, int brushHandle)
    {
        if (_objects.TryGetValue(brushHandle, out var o) && o is GdiBrush brush) FillRects(rects, brush);
    }

    public void PatBlt(double x, double y, double w, double h, int rop)
    {
        GdiBrush? brush = rop switch { 0x00F00021 => _state.Brush, 0x00000042 => new GdiBrush(Color.Black), 0x00FF0062 => new GdiBrush(Color.White), _ => null };
        if (brush is not null) FillRects([(x, y, x + w, y + h)], brush);
    }

    /// <summary>Draws a bitmap (a BMP, PNG or JPEG file) stretched over the logical rectangle.</summary>
    public void Bitmap(ImageData picture, double x, double y, double w, double h)
    {
        var m = Transform;
        var (x1, y1) = m.Apply(x, y);
        var (x2, y2) = m.Apply(x + w, y + h);
        bool flipH = x2 < x1, flipV = y2 < y1;
        double left = Math.Min(x1, x2), top = Math.Min(y1, y2), width = Math.Abs(x2 - x1), height = Math.Abs(y2 - y1);
        if (width <= 0 || height <= 0) return;
        if (flipH || flipV) Add(new PushItem(null, 0, left + width / 2, top + height / 2, flipH, flipV));
        Add(new ImageItem(left, top, width, height, picture));
        if (flipH || flipV) Add(new PopItem());
    }

    /// <summary>Draws a string at the logical reference point with the current font, text color and alignment.</summary>
    public void Text(double x, double y, string text, IReadOnlyList<double>? dx, bool opaque, (double L, double T, double R, double B)? background)
    {
        var s = _state;
        if (opaque && background is { } box) FillRects([box], new GdiBrush(s.BackColor));
        if (text.Length == 0) return;
        if ((s.TextAlign & 1) != 0) (x, y) = s.Position;

        var m = Transform;
        var (px, py) = m.Apply(x, y);
        double unitX = Math.Sqrt(m.A * m.A + m.B * m.B), unitY = Math.Sqrt(m.C * m.C + m.D * m.D);
        var font = s.Font;
        var resolved = fonts.Resolve(font.Face, font.Weight >= 600, font.Italic);
        var ttf = resolved.Font;
        double height = Math.Abs(font.Height == 0 ? 12 : font.Height) * unitY;
        // A positive height is the cell height (ascent plus descent); a negative one the em size.
        double size = font.Height > 0 && ttf.WinAscent + ttf.WinDescent > 0 ? height * ttf.UnitsPerEm / (ttf.WinAscent + ttf.WinDescent) : height;
        if (size <= 0) return;
        var metrics = FontMetrics.Of(ttf, size);

        var offsets = new List<double>(text.Length);
        double advance = 0;
        for (int k = 0; k < text.Length; k++)
        {
            offsets.Add(advance);
            advance += dx is not null && k < dx.Count ? dx[k] * unitX : ttf.MeasureText(text[k].ToString(), size);
        }
        double shiftX = (s.TextAlign & 6) switch { 6 => -advance / 2, 2 => -advance, _ => 0 };
        double baseline = (s.TextAlign & 24) switch { 24 => 0, 8 => -metrics.Descent, _ => metrics.Ascent };
        if ((s.TextAlign & 1) != 0) s.Position = (x + advance / unitX, y);

        // Escapement turns the text counter-clockwise; a rotation in the transform turns it as well (text stays upright when the y axis flips).
        double rotation = font.Escapement / 10.0 - Math.Atan2(m.B, m.A) * 180 / Math.PI;
        bool rotated = Math.Abs(rotation % 360) > 0.01;
        if (rotated) Add(new PushItem(null, rotation, px, py));
        double left = px + shiftX, line = py + baseline;
        if (dx is null) Add(new TextItem(left, line, ttf, size, text, s.TextColor, resolved.SyntheticBold, resolved.SyntheticItalic));
        else
            for (int k = 0; k < text.Length; k++)
                if (!char.IsWhiteSpace(text[k]))
                    Add(new TextItem(left + offsets[k], line, ttf, size, text[k].ToString(), s.TextColor, resolved.SyntheticBold, resolved.SyntheticItalic));
        double thickness = Math.Max(size * ttf.UnderlineThickness / ttf.UnitsPerEm, size / 20);
        if (font.Underline)
        {
            double uy = line - size * ttf.UnderlinePosition / ttf.UnitsPerEm;
            Add(new LineItem(left, uy, left + advance, uy, thickness, s.TextColor));
        }
        if (font.Strikeout)
        {
            double sy = line - size * (ttf.StrikeoutPosition != 0 ? ttf.StrikeoutPosition : ttf.UnitsPerEm / 4) / ttf.UnitsPerEm;
            Add(new LineItem(left, sy, left + advance, sy, thickness, s.TextColor));
        }
        if (rotated) Add(new PopItem());
    }

    /// <summary>A BMP file from a device-independent bitmap (its header and bits), or the embedded PNG or JPEG.</summary>
    public static ImageData? Dib(ReadOnlySpan<byte> header, ReadOnlySpan<byte> bits)
    {
        if (header.Length < 12 || bits.Length == 0) return null;
        if (header.Length >= 20)
        {
            int compression = BinaryPrimitives.ReadInt32LittleEndian(header[16..]);
            if (compression is 4 or 5) return new ImageData(bits.ToArray(), "metafile bitmap");
        }
        var file = new byte[14 + header.Length + bits.Length];
        file[0] = (byte)'B';
        file[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(file.AsSpan(2), file.Length);
        BinaryPrimitives.WriteInt32LittleEndian(file.AsSpan(10), 14 + header.Length);
        header.CopyTo(file.AsSpan(14));
        bits.CopyTo(file.AsSpan(14 + header.Length));
        return new ImageData(file, "metafile bitmap");
    }

    /// <summary>
    /// A 32-bit bitmap with premultiplied alpha (AlphaBlend with AC_SRC_ALPHA) as a BMP file with an alpha mask, its colors divided back out
    /// and its alpha scaled by the constant alpha.
    /// </summary>
    public static ImageData? AlphaDib(ReadOnlySpan<byte> header, ReadOnlySpan<byte> bits, int constantAlpha)
    {
        if (header.Length < 40 || BinaryPrimitives.ReadInt16LittleEndian(header[14..]) != 32) return Dib(header, bits);
        int compression = BinaryPrimitives.ReadInt32LittleEndian(header[16..]);
        if (compression is not (0 or 3)) return Dib(header, bits);
        const int v3 = 56;
        var info = new byte[v3];
        header[..40].CopyTo(info);
        BinaryPrimitives.WriteInt32LittleEndian(info, v3);
        BinaryPrimitives.WriteInt32LittleEndian(info.AsSpan(16), 3);
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(40), 0x00FF0000);
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(44), 0x0000FF00);
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(48), 0x000000FF);
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(52), 0xFF000000);
        var pixels = bits.ToArray();
        for (int k = 0; k + 3 < pixels.Length; k += 4)
        {
            int a = pixels[k + 3];
            if (a is > 0 and < 255)
                for (int c = 0; c < 3; c++) pixels[k + c] = (byte)Math.Min(255, pixels[k + c] * 255 / a);
            pixels[k + 3] = (byte)(a * constantAlpha / 255);
        }
        return Dib(info, pixels);
    }

    /// <summary>
    /// A 24- or 32-bit image combined with a one-bit mask of the same size (black opaque, white transparent) as a BMP file with alpha;
    /// null for other kinds of bitmap.
    /// </summary>
    public static ImageData? MaskedDib(ReadOnlySpan<byte> maskHeader, ReadOnlySpan<byte> maskBits, ReadOnlySpan<byte> header, ReadOnlySpan<byte> bits)
    {
        if (maskHeader.Length < 48 || header.Length < 40) return null;
        int width = BinaryPrimitives.ReadInt32LittleEndian(header[4..]), height = BinaryPrimitives.ReadInt32LittleEndian(header[8..]);
        int bpp = BinaryPrimitives.ReadInt16LittleEndian(header[14..]), compression = BinaryPrimitives.ReadInt32LittleEndian(header[16..]);
        int maskWidth = BinaryPrimitives.ReadInt32LittleEndian(maskHeader[4..]), maskHeight = BinaryPrimitives.ReadInt32LittleEndian(maskHeader[8..]);
        int maskHeaderSize = BinaryPrimitives.ReadInt32LittleEndian(maskHeader);
        if (bpp is not (24 or 32) || !(compression == 0 || compression == 3 && bpp == 32) || BinaryPrimitives.ReadInt16LittleEndian(maskHeader[14..]) != 1
            || maskWidth != width || Math.Abs(maskHeight) < Math.Abs(height) || width <= 0 || maskHeaderSize + 8 > maskHeader.Length) return null;
        int rows = Math.Abs(height), stride = (width * bpp + 31) / 32 * 4, maskStride = (width + 31) / 32 * 4;
        // A taller mask is cropped to its top rows (the end of a bottom-up bitmap).
        int skip = maskHeight > 0 ? maskHeight - rows : 0;
        if (bits.Length < stride * rows || maskBits.Length < maskStride * (rows + skip)) return null;
        // Which mask index is black (opaque), from the mask's two-color palette.
        var black = maskHeader.Slice(maskHeaderSize, 3);
        int opaqueBit = black[0] == 0 && black[1] == 0 && black[2] == 0 ? 0 : 1;
        bool flip = height > 0 != maskHeight > 0;

        var info = new byte[56];
        header[..40].CopyTo(info);
        BinaryPrimitives.WriteInt32LittleEndian(info, 56);
        BinaryPrimitives.WriteInt16LittleEndian(info.AsSpan(14), 32);
        BinaryPrimitives.WriteInt32LittleEndian(info.AsSpan(16), 3);
        BinaryPrimitives.WriteInt32LittleEndian(info.AsSpan(20), width * rows * 4);
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(40), 0x00FF0000);
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(44), 0x0000FF00);
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(48), 0x000000FF);
        BinaryPrimitives.WriteUInt32LittleEndian(info.AsSpan(52), 0xFF000000);
        var pixels = new byte[width * rows * 4];
        for (int row = 0; row < rows; row++)
        {
            int maskRow = (flip ? rows - 1 - row : row) + skip;
            for (int x = 0; x < width; x++)
            {
                int src = row * stride + x * (bpp / 8), dst = (row * width + x) * 4;
                pixels[dst] = bits[src];
                pixels[dst + 1] = bits[src + 1];
                pixels[dst + 2] = bits[src + 2];
                int bit = maskBits[maskRow * maskStride + x / 8] >> (7 - x % 8) & 1;
                pixels[dst + 3] = (byte)(bit == opaqueBit ? 255 : 0);
            }
        }
        return Dib(info, pixels);
    }

    public static Color ColorRef(ReadOnlySpan<byte> d) => new(d[0], d[1], d[2]);
}

/// <summary>Plays the records of an Enhanced Metafile (EMF, [MS-EMF]).</summary>
internal static class EmfReader
{
    /// <summary>
    /// Plays the GDI records, or with <paramref name="plus"/> the EMF+ records and the GDI records after their GetDC. False when GDI records
    /// alone were played and only EMF+ records draw.
    /// </summary>
    public static bool Play(byte[] data, GdiPlayer gdi, EmfPlusPlayer? plus = null)
    {
        var d = data.AsSpan();
        if (d.Length < 88 || I32(d, 0) != 1) throw new InvalidDataException("Not an EMF file.");
        int headerSize = I32(d, 4);
        var bounds = (L: I32(d, 8), T: I32(d, 12), R: I32(d, 16), B: I32(d, 20));
        var frame = (L: I32(d, 24), T: I32(d, 28), R: I32(d, 32), B: I32(d, 36));
        double devX = I32(d, 72), devY = I32(d, 76), mmX = I32(d, 80), mmY = I32(d, 84);
        if (headerSize >= 108 && d.Length >= 108 && I32(d, 100) > 0 && I32(d, 104) > 0) (mmX, mmY) = (I32(d, 100) / 1000.0, I32(d, 104) / 1000.0);
        if (devX <= 0 || devY <= 0 || mmX <= 0 || mmY <= 0) (devX, devY, mmX, mmY) = (96, 96, 25.4, 25.4);
        gdi.PixelsPerMm = (devX / mmX, devY / mmY);

        // The frame (hundredths of a millimetre) is the picture's box; without one, the bounds (device pixels) are.
        double l, t, w, h;
        if (frame.R > frame.L && frame.B > frame.T)
        {
            double px = 100 * mmX / devX, py = 100 * mmY / devY;
            (l, t, w, h) = (frame.L / px, frame.T / py, (frame.R - frame.L) / px, (frame.B - frame.T) / py);
        }
        else (l, t, w, h) = (bounds.L, bounds.T, Math.Max(bounds.R - bounds.L, 1), Math.Max(bounds.B - bounds.T, 1));
        // Frame and bounds are in square units (hundredths of a millimetre scaled per axis, or pixels), so the box keeps their proportions.
        double physicalHeight = frame.R > frame.L && frame.B > frame.T ? (frame.B - frame.T) / (double)(frame.R - frame.L) : h / w;
        gdi.BoxHeight = Math.Clamp(1000 * physicalHeight, 1, 1e6);
        gdi.DeviceToBox = new Affine(1000 / w, 0, 0, gdi.BoxHeight / h, -l * 1000 / w, -t * gdi.BoxHeight / h);
        gdi.PixelWidth = gdi.DeviceToBox.Scale;

        bool emfPlus = false, drew = false;
        (byte[] Header, byte[] Bits, int X, int Y, int W, int H)? mask = null;
        int pos = 0;
        while (pos + 8 <= d.Length)
        {
            int type = I32(d, pos), size = I32(d, pos + 4);
            if (size < 8 || pos + size > d.Length) break;
            var r = d.Slice(pos, size);
            pos += size;
            if (type == 14) break;
            if (type == 70 && r.Length >= 16 && r.Slice(12, 4).SequenceEqual("EMF+"u8))
            {
                emfPlus = true;
                plus?.Play(r[16..]);
                continue;
            }
            if (plus is { GdiActive: false }) continue;
            int before = gdi.Items.Count;
            Record(type, r, gdi, ref mask);
            drew |= gdi.Items.Count > before;
        }
        return drew || !emfPlus || plus is not null;
    }

    private static int I32(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadInt32LittleEndian(d[o..]);
    private static short I16(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadInt16LittleEndian(d[o..]);
    private static float F32(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadSingleLittleEndian(d[o..]);

    private static Affine Xform(ReadOnlySpan<byte> d, int o) => new(F32(d, o), F32(d, o + 4), F32(d, o + 8), F32(d, o + 12), F32(d, o + 16), F32(d, o + 20));

    private static List<(double X, double Y)> Points(ReadOnlySpan<byte> r, int offset, int count, bool small)
    {
        int step = small ? 4 : 8;
        if (count < 0 || offset + (long)count * step > r.Length) throw new InvalidDataException("A point list runs past its record.");
        var points = new List<(double X, double Y)>(count);
        for (int k = 0; k < count; k++)
        {
            int o = offset + k * step;
            points.Add(small ? (I16(r, o), I16(r, o + 2)) : (I32(r, o), I32(r, o + 4)));
        }
        return points;
    }

    private static List<IReadOnlyList<(double X, double Y)>> PolyPoints(ReadOnlySpan<byte> r, bool small)
    {
        int polys = I32(r, 24), total = I32(r, 28);
        if (polys < 0 || 32 + (long)polys * 4 > r.Length) throw new InvalidDataException("A polygon list runs past its record.");
        var all = Points(r, 32 + polys * 4, total, small);
        var result = new List<IReadOnlyList<(double X, double Y)>>(polys);
        int at = 0;
        for (int k = 0; k < polys; k++)
        {
            int n = I32(r, 32 + k * 4);
            if (n < 0 || at + n > all.Count) break;
            result.Add(all.GetRange(at, n));
            at += n;
        }
        return result;
    }

    private static void Record(int type, ReadOnlySpan<byte> r, GdiPlayer gdi, ref (byte[] Header, byte[] Bits, int X, int Y, int W, int H)? mask)
    {
        switch (type)
        {
            case 2 or 85: gdi.PolyBezier(Points(r, 28, I32(r, 24), type == 85), false); break;
            case 3 or 86: gdi.Polygon(Points(r, 28, I32(r, 24), type == 86)); break;
            case 4 or 87: gdi.Polyline(Points(r, 28, I32(r, 24), type == 87)); break;
            case 5 or 88: gdi.PolyBezier(Points(r, 28, I32(r, 24), type == 88), true); break;
            case 6 or 89: gdi.PolylineTo(Points(r, 28, I32(r, 24), type == 89)); break;
            case 7 or 90: foreach (var line in PolyPoints(r, type == 90)) gdi.Polyline(line); break;
            case 8 or 91: gdi.PolyPolygon(PolyPoints(r, type == 91)); break;
            case 9: gdi.SetWindowExt(I32(r, 8), I32(r, 12)); break;
            case 10: gdi.SetWindowOrg(I32(r, 8), I32(r, 12)); break;
            case 11: gdi.SetViewportExt(I32(r, 8), I32(r, 12)); break;
            case 12: gdi.SetViewportOrg(I32(r, 8), I32(r, 12)); break;
            case 17: gdi.SetMapMode(I32(r, 8)); break;
            case 18: gdi.SetBackMode(I32(r, 8)); break;
            case 19: gdi.SetPolyFillMode(I32(r, 8)); break;
            case 22: gdi.SetTextAlign(I32(r, 8)); break;
            case 24: gdi.SetTextColor(GdiPlayer.ColorRef(r[8..])); break;
            case 25: gdi.SetBackColor(GdiPlayer.ColorRef(r[8..])); break;
            case 27: gdi.MoveTo(I32(r, 8), I32(r, 12)); break;
            case 31: gdi.ScaleViewportExt(I32(r, 8), I32(r, 12), I32(r, 16), I32(r, 20)); break;
            case 32: gdi.ScaleWindowExt(I32(r, 8), I32(r, 12), I32(r, 16), I32(r, 20)); break;
            case 33: gdi.Save(); break;
            case 34: gdi.Restore(I32(r, 8)); break;
            case 35: gdi.SetWorld(Xform(r, 8)); break;
            case 36: gdi.ModifyWorld(Xform(r, 8), I32(r, 32)); break;
            case 37:
            {
                int handle = I32(r, 8);
                if ((handle & 0x80000000) != 0) { if (GdiPlayer.Stock(handle & 0x7FFFFFFF) is { } stock) gdi.Select(stock); }
                else gdi.Select(handle);
                break;
            }
            case 38: gdi.Create(I32(r, 8), new GdiPen(I32(r, 12), I32(r, 16), GdiPlayer.ColorRef(r[24..]), I32(r, 16) <= 1)); break;
            case 39:
            {
                int style = I32(r, 12);
                gdi.Create(I32(r, 8), new GdiBrush(style == 1 ? null : GdiPlayer.ColorRef(r[16..])));
                break;
            }
            case 40: gdi.Delete(I32(r, 8)); break;
            case 42: gdi.Ellipse(I32(r, 8), I32(r, 12), I32(r, 16), I32(r, 20)); break;
            case 43: gdi.Rectangle(I32(r, 8), I32(r, 12), I32(r, 16), I32(r, 20)); break;
            case 44: gdi.RoundRect(I32(r, 8), I32(r, 12), I32(r, 16), I32(r, 20), I32(r, 24), I32(r, 28)); break;
            case 45 or 46 or 47 or 55:
                gdi.Arc(I32(r, 8), I32(r, 12), I32(r, 16), I32(r, 20), I32(r, 24), I32(r, 28), I32(r, 32), I32(r, 36), type switch { 46 => 1, 47 => 2, _ => 0 });
                break;
            case 54: gdi.PolylineTo([(I32(r, 8), I32(r, 12))]); break;
            case 57: gdi.SetArcDirection(I32(r, 8)); break;
            case 59: gdi.BeginPath(); break;
            case 60: gdi.EndPath(); break;
            case 61: gdi.CloseFigure(); break;
            case 62: gdi.FillPath(true, false); break;
            case 63: gdi.FillPath(true, true); break;
            case 64: gdi.FillPath(false, true); break;
            case 68: gdi.AbortPath(); break;
            case 30: gdi.IntersectClip(I32(r, 8), I32(r, 12), I32(r, 16), I32(r, 20)); break;
            case 67: gdi.SelectClipPath(I32(r, 8)); break;
            case 75 when r.Length >= 16:
                gdi.SelectClipRegion(I32(r, 8) == 0 ? null : Region(r[16..]), I32(r, 12));
                break;
            case 71 when r.Length >= 32:
                // FillRgn: the brush handle, then the region's rectangles.
                if (Region(r[32..]) is { } fillRegion) gdi.FillRegion(fillRegion, I32(r, 28));
                break;
            case 74 when r.Length >= 28:
                if (Region(r[28..]) is { } paintRegion) gdi.FillRects(paintRegion);
                break;
            case 76 or 77 or 114 when r.Length >= 100:
            {
                // BitBlt, StretchBlt, AlphaBlend: without a source bitmap BitBlt is a pattern fill.
                int x = I32(r, 24), y = I32(r, 28), cx = I32(r, 32), cy = I32(r, 36), rop = I32(r, 40);
                int offBmi = I32(r, 84), cbBmi = I32(r, 88), offBits = I32(r, 92), cbBits = I32(r, 96);
                if (cbBmi == 0 || cbBits == 0) { if (type == 76) gdi.PatBlt(x, y, cx, cy, rop); break; }
                var bmi = Slice(r, offBmi, cbBmi);
                var bits = Slice(r, offBits, cbBits);
                if (type != 114 && rop == 0x008800C6)
                {
                    // SRCAND: a mask, ANDed onto the page before an image is XORed or ORed into the same place (an icon's transparency).
                    mask = (bmi.ToArray(), bits.ToArray(), x, y, cx, cy);
                    break;
                }
                if (type != 114 && !CopiesSource(rop)) { mask = null; break; }
                ImageData? blt;
                if (type == 114) blt = (rop >> 24 & 1) != 0 ? GdiPlayer.AlphaDib(bmi, bits, rop >> 16 & 0xFF) : GdiPlayer.Dib(bmi, bits);
                else if (mask is { } m && (m.X, m.Y, m.W, m.H) == (x, y, cx, cy) && rop is 0x00660046 or 0x00EE0086)
                    blt = GdiPlayer.MaskedDib(m.Header, m.Bits, bmi, bits) ?? GdiPlayer.Dib(bmi, bits);
                else blt = GdiPlayer.Dib(bmi, bits);
                mask = null;
                if (blt is not null) gdi.Bitmap(blt, x, y, cx, cy);
                break;
            }
            case 80 when r.Length >= 76:
            {
                // SetDIBitsToDevice: drawn at the source's size.
                if (GdiPlayer.Dib(Slice(r, I32(r, 48), I32(r, 52)), Slice(r, I32(r, 56), I32(r, 60))) is { } dib)
                    gdi.Bitmap(dib, I32(r, 24), I32(r, 28), I32(r, 40), I32(r, 44));
                break;
            }
            case 81 when r.Length >= 80:
            {
                if (!CopiesSource(I32(r, 68))) break;
                if (GdiPlayer.Dib(Slice(r, I32(r, 48), I32(r, 52)), Slice(r, I32(r, 56), I32(r, 60))) is { } dib)
                    gdi.Bitmap(dib, I32(r, 24), I32(r, 28), I32(r, 72), I32(r, 76));
                break;
            }
            case 82 when r.Length >= 104:
            {
                var face = Encoding.Unicode.GetString(r.Slice(40, 64)).Split('\0')[0];
                gdi.Create(I32(r, 8), new GdiFont(I32(r, 12), I32(r, 20), I32(r, 28), r[32] != 0, r[33] != 0, r[34] != 0, face.Length > 0 ? face : "Arial"));
                break;
            }
            case 83 or 84 when r.Length >= 76:
            {
                int count = I32(r, 44), offString = I32(r, 48), options = I32(r, 52), offDx = I32(r, 72);
                if (count < 0) break;
                string text = type == 84
                    ? Encoding.Unicode.GetString(Slice(r, offString, count * 2))
                    : Encoding.Latin1.GetString(Slice(r, offString, count));
                List<double>? dx = null;
                if (offDx > 0 && offDx + (long)count * 4 <= r.Length)
                {
                    int step = (options & 0x2000) != 0 ? 8 : 4;
                    dx = [];
                    for (int k = 0; k < count && offDx + k * step + 4 <= r.Length; k++) dx.Add(I32(r, offDx + k * step));
                }
                var rect = (I32(r, 56), I32(r, 60), I32(r, 64), I32(r, 68));
                gdi.Text(I32(r, 36), I32(r, 40), text, dx, (options & 2) != 0, rect);
                break;
            }
            case 95 when r.Length >= 44:
            {
                int style = I32(r, 28), brushStyle = I32(r, 36);
                // Geometric pens (PS_GEOMETRIC) have a width in logical units; cosmetic ones are a pixel wide.
                gdi.Create(I32(r, 8), new GdiPen(brushStyle == 1 ? 5 : style, I32(r, 32), GdiPlayer.ColorRef(r[40..]), (style & 0x00010000) == 0));
                break;
            }
        }
    }

    /// <summary>Raster operations that paint the source bitmap as it is.</summary>
    private static bool CopiesSource(int rop) => rop is 0x00CC0020 or 0x00EE0086 or 0x00660046 or 0x00C000CA;

    private static ReadOnlySpan<byte> Slice(ReadOnlySpan<byte> r, int offset, int length) =>
        offset < 0 || length < 0 || offset + (long)length > r.Length ? throw new InvalidDataException("A record's data runs past its end.") : r.Slice(offset, length);

    private static List<(double L, double T, double R, double B)>? Region(ReadOnlySpan<byte> d)
    {
        if (d.Length < 32) return null;
        int count = I32(d, 8);
        var rects = new List<(double, double, double, double)>();
        for (int k = 0; k < count && 32 + k * 16 + 16 <= d.Length; k++)
        {
            int o = 32 + k * 16;
            rects.Add((I32(d, o), I32(d, o + 4), I32(d, o + 8), I32(d, o + 12)));
        }
        return rects;
    }
}

/// <summary>Plays the records of a Windows Metafile (WMF, [MS-WMF]), with or without the placeable (Aldus) header.</summary>
internal static class WmfReader
{
    public static bool Play(byte[] data, GdiPlayer gdi)
    {
        var d = data.AsSpan();
        int pos = 0;
        (double L, double T, double W, double H) frame = (0, 0, 1000, 1000);
        bool placeable = d.Length >= 22 && BinaryPrimitives.ReadUInt32LittleEndian(d) == 0x9AC6CDD7;
        if (placeable)
        {
            int l = I16(d, 6), t = I16(d, 8), r = I16(d, 10), b = I16(d, 12);
            frame = (Math.Min(l, r), Math.Min(t, b), Math.Max(Math.Abs(r - l), 1), Math.Max(Math.Abs(b - t), 1));
            pos = 22;
        }
        if (d.Length < pos + 18 || U16(d, pos + 2) != 9) throw new InvalidDataException("Not a WMF file.");
        pos += 18;
        if (!placeable && FirstWindow(d[pos..]) is { } window) frame = window;
        gdi.WmfFrame = frame;
        gdi.BoxHeight = Math.Clamp(1000 * Math.Abs(frame.H / frame.W), 1, 1e6);
        // A pixel at 96 dpi, from the placeable header's logical units per inch.
        int inch = placeable ? U16(d, 14) : 0;
        gdi.PixelWidth = inch > 0 ? 1000.0 * inch / 96 / Math.Abs(frame.W) : 1.5;

        var slots = new List<bool>();
        int Allocate(object? gdiObject)
        {
            int index = slots.IndexOf(false);
            if (index < 0) { index = slots.Count; slots.Add(true); } else slots[index] = true;
            if (gdiObject is not null) gdi.Create(index, gdiObject); else gdi.Delete(index);
            return index;
        }

        while (pos + 6 <= d.Length)
        {
            long words = BinaryPrimitives.ReadUInt32LittleEndian(d[pos..]);
            int function = U16(d, pos + 4);
            if (words < 3 || pos + words * 2 > d.Length) break;
            var p = d.Slice(pos + 6, (int)(words * 2 - 6));
            pos += (int)(words * 2);
            if (function == 0) break;
            switch (function)
            {
                case 0x0201: gdi.SetBackColor(GdiPlayer.ColorRef(p)); break;
                case 0x0102: gdi.SetBackMode(U16(p, 0)); break;
                case 0x0103: gdi.SetMapMode(U16(p, 0)); break;
                case 0x0106: gdi.SetPolyFillMode(U16(p, 0)); break;
                case 0x0209: gdi.SetTextColor(GdiPlayer.ColorRef(p)); break;
                case 0x012E: gdi.SetTextAlign(U16(p, 0)); break;
                case 0x020B: gdi.SetWindowOrg(I16(p, 2), I16(p, 0)); break;
                case 0x020C: gdi.SetWindowExt(I16(p, 2), I16(p, 0)); break;
                case 0x020D: gdi.SetViewportOrg(I16(p, 2), I16(p, 0)); break;
                case 0x020E: gdi.SetViewportExt(I16(p, 2), I16(p, 0)); break;
                case 0x020F: gdi.OffsetWindowOrg(I16(p, 2), I16(p, 0)); break;
                case 0x0211: gdi.OffsetViewportOrg(I16(p, 2), I16(p, 0)); break;
                case 0x0410: gdi.ScaleWindowExt(I16(p, 6), I16(p, 4), I16(p, 2), I16(p, 0)); break;
                case 0x0412: gdi.ScaleViewportExt(I16(p, 6), I16(p, 4), I16(p, 2), I16(p, 0)); break;
                case 0x0214: gdi.MoveTo(I16(p, 2), I16(p, 0)); break;
                case 0x0213: gdi.PolylineTo([(I16(p, 2), I16(p, 0))]); break;
                case 0x0324: gdi.Polygon(Points(p, 2, I16(p, 0))); break;
                case 0x0325: gdi.Polyline(Points(p, 2, I16(p, 0))); break;
                case 0x0538:
                {
                    int polys = I16(p, 0), at = 2 + polys * 2;
                    var polygons = new List<IReadOnlyList<(double X, double Y)>>();
                    for (int k = 0; k < polys; k++)
                    {
                        int n = I16(p, 2 + k * 2);
                        polygons.Add(Points(p, at, n));
                        at += n * 4;
                    }
                    gdi.PolyPolygon(polygons);
                    break;
                }
                case 0x041B: gdi.Rectangle(I16(p, 6), I16(p, 4), I16(p, 2), I16(p, 0)); break;
                case 0x0418: gdi.Ellipse(I16(p, 6), I16(p, 4), I16(p, 2), I16(p, 0)); break;
                case 0x061C: gdi.RoundRect(I16(p, 10), I16(p, 8), I16(p, 6), I16(p, 4), I16(p, 2), I16(p, 0)); break;
                case 0x0817 or 0x081A or 0x0830:
                    gdi.Arc(I16(p, 14), I16(p, 12), I16(p, 10), I16(p, 8), I16(p, 6), I16(p, 4), I16(p, 2), I16(p, 0),
                        function switch { 0x0830 => 1, 0x081A => 2, _ => 0 });
                    break;
                case 0x001E: gdi.Save(); break;
                case 0x0416: gdi.IntersectClip(I16(p, 6), I16(p, 4), I16(p, 2), I16(p, 0)); break;
                case 0x0127: gdi.Restore(I16(p, 0)); break;
                case 0x012D: gdi.Select(U16(p, 0)); break;
                case 0x01F0:
                {
                    int index = U16(p, 0);
                    if (index < slots.Count) slots[index] = false;
                    gdi.Delete(index);
                    break;
                }
                case 0x02FA:
                {
                    int style = U16(p, 0), width = I16(p, 2);
                    Allocate(new GdiPen(style, width, GdiPlayer.ColorRef(p[6..]), width <= 1));
                    break;
                }
                case 0x02FC:
                    Allocate(new GdiBrush(U16(p, 0) == 1 ? null : GdiPlayer.ColorRef(p[2..])));
                    break;
                case 0x02FB when p.Length >= 18:
                {
                    var face = Encoding.Latin1.GetString(p.Slice(18, Math.Min(32, p.Length - 18))).Split('\0')[0];
                    Allocate(new GdiFont(I16(p, 0), I16(p, 4), I16(p, 8), p[10] != 0, p[11] != 0, p[12] != 0, face.Length > 0 ? face : "Arial"));
                    break;
                }
                case 0x00F7 or 0x01F9 or 0x0142 or 0x06FF:
                    // Palettes, pattern brushes and regions take an object slot; pattern brushes paint grey.
                    Allocate(function is 0x01F9 or 0x0142 ? new GdiBrush(new Color(0x80, 0x80, 0x80)) : null);
                    break;
                case 0x0521:
                {
                    int count = I16(p, 0), padded = (count + 1) & ~1;
                    string text = Encoding.Latin1.GetString(p.Slice(2, count));
                    gdi.Text(I16(p, 2 + padded + 2), I16(p, 2 + padded), text, null, false, null);
                    break;
                }
                case 0x0A32:
                {
                    int y = I16(p, 0), x = I16(p, 2), count = I16(p, 4), options = U16(p, 6), at = 8;
                    (double, double, double, double)? rect = null;
                    if ((options & 6) != 0 && p.Length >= 16) { rect = (I16(p, 8), I16(p, 10), I16(p, 12), I16(p, 14)); at = 16; }
                    string text = Encoding.Latin1.GetString(p.Slice(at, count));
                    at += (count + 1) & ~1;
                    List<double>? dx = null;
                    if (at + count * 2 <= p.Length)
                    {
                        dx = [];
                        for (int k = 0; k < count; k++) dx.Add(I16(p, at + k * 2));
                    }
                    gdi.Text(x, y, text, dx, (options & 2) != 0, rect);
                    break;
                }
                case 0x041D: gdi.PatBlt(I16(p, 10), I16(p, 8), I16(p, 6), I16(p, 4), I32(p, 0)); break;
                case 0x0B41 or 0x0940:
                {
                    // DIBStretchBlt / DIBBitBlt; the variants without a bitmap have an extra reserved word and paint the brush.
                    bool stretch = function == 0x0B41;
                    int rop = I32(p, 0);
                    if (p.Length <= (stretch ? 22 : 18))
                    {
                        int o = stretch ? 14 : 10;
                        gdi.PatBlt(I16(p, o + 6), I16(p, o + 4), I16(p, o + 2), I16(p, o), rop);
                        break;
                    }
                    int dest = stretch ? 12 : 8;
                    double h = I16(p, dest), w = I16(p, dest + 2), yd = I16(p, dest + 4), xd = I16(p, dest + 6);
                    if (Dib(p[(dest + 8)..]) is { } bitmap) gdi.Bitmap(bitmap, xd, yd, w, h);
                    break;
                }
                case 0x0F43:
                {
                    double h = I16(p, 14), w = I16(p, 16), yd = I16(p, 18), xd = I16(p, 20);
                    if (Dib(p[22..]) is { } bitmap) gdi.Bitmap(bitmap, xd, yd, w, h);
                    break;
                }
            }
        }
        return true;
    }

    /// <summary>Without a placeable header, the first window the file sets (origin and extent) stands for its bounds.</summary>
    private static (double L, double T, double W, double H)? FirstWindow(ReadOnlySpan<byte> d)
    {
        double l = 0, t = 0;
        for (int pos = 0; pos + 6 <= d.Length;)
        {
            long words = BinaryPrimitives.ReadUInt32LittleEndian(d[pos..]);
            int function = U16(d, pos + 4);
            if (words < 3 || pos + words * 2 > d.Length || function == 0) break;
            if (function == 0x020B && words >= 5) (l, t) = (I16(d, pos + 8), I16(d, pos + 6));
            if (function == 0x020C && words >= 5 && I16(d, pos + 6) != 0 && I16(d, pos + 8) != 0) return (l, t, I16(d, pos + 8), I16(d, pos + 6));
            pos += (int)(words * 2);
        }
        return null;
    }

    /// <summary>Splits a packed DIB (header, color table, bits) for <see cref="GdiPlayer.Dib"/>.</summary>
    private static ImageData? Dib(ReadOnlySpan<byte> d)
    {
        if (d.Length < 12) return null;
        int headerSize = I32(d, 0);
        if (headerSize < 12 || headerSize > d.Length) return null;
        int bpp = headerSize == 12 ? U16(d, 10) : U16(d, 14);
        int compression = headerSize >= 20 ? I32(d, 16) : 0, used = headerSize >= 36 ? I32(d, 32) : 0;
        int colors = used > 0 ? used : bpp <= 8 ? 1 << bpp : 0;
        int table = colors * (headerSize == 12 ? 3 : 4) + (compression == 3 && headerSize == 40 ? 12 : 0);
        int bitsAt = headerSize + table;
        if (bitsAt > d.Length) return null;
        return GdiPlayer.Dib(d[..bitsAt], d[bitsAt..]);
    }

    private static List<(double X, double Y)> Points(ReadOnlySpan<byte> p, int offset, int count)
    {
        if (count < 0 || offset + count * 4 > p.Length) throw new InvalidDataException("A point list runs past its record.");
        var points = new List<(double X, double Y)>(count);
        for (int k = 0; k < count; k++) points.Add((I16(p, offset + k * 4), I16(p, offset + k * 4 + 2)));
        return points;
    }

    private static int I32(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadInt32LittleEndian(d[o..]);
    private static short I16(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadInt16LittleEndian(d[o..]);
    private static int U16(ReadOnlySpan<byte> d, int o) => BinaryPrimitives.ReadUInt16LittleEndian(d[o..]);
}
