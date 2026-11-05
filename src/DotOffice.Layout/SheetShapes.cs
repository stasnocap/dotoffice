using DotOffice.Core;

namespace DotOffice.Layout;

public sealed partial class SheetLayoutEngine
{
    /// <summary>
    /// Draws a shape into its frame: the filled outline, the stroke, then the text. Rotation applies to everything, mirroring only to the
    /// outline (text is never mirrored).
    /// </summary>
    private void EmitShape(SheetShape shape, double x, double y, double width, double height,
        (double X, double Y, double Width, double Height) visible, List<LayoutItem> items)
    {
        int pushed = 1;
        items.Add(new PushItem(visible));
        double cx = x + width / 2, cy = y + height / 2;
        if (shape.Rotation != 0)
        {
            items.Add(new PushItem(null, shape.Rotation, cx, cy));
            pushed++;
        }
        bool flipped = shape.FlipH || shape.FlipV;
        if (shape.Glow is { } glow) EmitGlow(shape, glow, x, y, width, height, flipped, items);
        if (shape.Reflection is { } reflection && shape.Geometry != ShapeGeometry.Line) EmitReflection(shape, reflection, x, y, width, height, items);
        if (shape.Shadow is { } shadow)
        {
            // The shadow is the shape's outline in the shadow color, moved and translucent. A blur is a 3 x 3 grid of such copies spread
            // over the blur radius, each so faint that the middle, where they all overlap, ends up with the shadow's opacity.
            bool filled = shape.Fill is not null || shape.FillGradient is not null || shape.FillPattern is not null || shape.FillImage is not null;
            var dark = shape with
            {
                Fill = filled ? shadow.Color : null, FillGradient = null, FillPattern = null, FillImage = null, Shadow = null,
                LineColor = shape.LineColor is null ? null : shadow.Color, Paragraphs = [],
            };
            bool soft = shadow.Blur >= 0.5;
            double spread = shadow.Blur / 2, copyOpacity = soft ? 1 - Math.Pow(1 - Math.Min(shadow.Opacity, 0.999), 1.0 / 9) : shadow.Opacity;
            for (int gx = soft ? -1 : 0; gx <= (soft ? 1 : 0); gx++)
                for (int gy = soft ? -1 : 0; gy <= (soft ? 1 : 0); gy++)
                {
                    double dx = shadow.Dx + gx * spread, dy = shadow.Dy + gy * spread;
                    items.Add(new PushItem(null, 0, cx + dx, cy + dy, shape.FlipH, shape.FlipV, copyOpacity));
                    EmitOutline(dark, x + dx, y + dy, width, height, items);
                    items.Add(new PopItem());
                }
        }
        if (flipped) items.Add(new PushItem(null, 0, cx, cy, shape.FlipH, shape.FlipV));
        if (shape.SoftEdge >= 0.5 && shape.Geometry != ShapeGeometry.Line) EmitSoftEdged(shape, x, y, width, height, items);
        else EmitOutline(shape, x, y, width, height, items);
        if (flipped) items.Add(new PopItem());
        if (shape.Paragraphs.Count > 0 && shape.Geometry != ShapeGeometry.Line) EmitShapeText(shape, x, y, width, height, items);
        for (int i = 0; i < pushed; i++) items.Add(new PopItem());
    }

    /// <summary>
    /// A reflection: the shape (without its text and effects) mirrored about a line half the reflection's distance under its bottom edge,
    /// faded by layers. Layer k covers the reflected area from its top down to the end of the k-th band, so the opacity at a band is that of
    /// the layers reaching it (those with a band index at or after its own); each layer's opacity is picked so that they add up to the
    /// opacity the fade has there (from the start to the end opacity, down to where the reflection has faded out). Nested layers leave no seams.
    /// </summary>
    private static void EmitReflection(SheetShape shape, ShapeReflection reflection, double x, double y, double w, double h, List<LayoutItem> items)
    {
        const int Bands = 16;
        double cx = x + w / 2, bottom = y + h, top = bottom + reflection.Distance, axis = bottom + reflection.Distance / 2;
        var mirrored = shape with { Paragraphs = [], Glow = null, Shadow = null, SoftEdge = 0, Reflection = null };
        bool flipped = shape.FlipH || shape.FlipV;
        double span = h * reflection.EndPosition;
        double Target(int band) => band >= Bands ? 0 : Math.Clamp(reflection.StartOpacity + (reflection.EndOpacity - reflection.StartOpacity) * (band + 0.5) / Bands, 0, 0.999);
        for (int k = 0; k < Bands; k++)
        {
            double opacity = 1 - (1 - Target(k)) / (1 - Target(k + 1));
            if (opacity <= 0.002) continue;
            items.Add(new PushItem((x - w, top, w * 3, span * (k + 1) / Bands)));
            items.Add(new PushItem(null, 0, cx, axis, false, true, opacity));
            if (flipped) items.Add(new PushItem(null, 0, cx, y + h / 2, shape.FlipH, shape.FlipV));
            EmitOutline(mirrored, x, y, w, h, items);
            if (flipped) items.Add(new PopItem());
            items.Add(new PopItem());
            items.Add(new PopItem());
        }
    }

    /// <summary>
    /// A glow: the outline stroked four times, from thin to as wide as the glow's radius reaches, each translucent, under the shape. Near the
    /// outline all four overlap and the glow is at its strongest; it fades out to the radius.
    /// </summary>
    private static void EmitGlow(SheetShape shape, ShapeGlow glow, double x, double y, double w, double h, bool flipped, List<LayoutItem> items)
    {
        const int Layers = 4;
        double cx = x + w / 2, cy = y + h / 2;
        double layerOpacity = 1 - Math.Pow(1 - Math.Min(glow.Opacity, 0.999), 1.0 / Layers);
        var outlines = new List<(List<(double X, double Y)> Points, bool Closed)>();
        if (shape.Paths is { } paths)
            outlines.AddRange(paths.Where(p => p.Points.Count >= 2).Select(p => (p.Points.Select(q => (X: x + q.X * w, Y: y + q.Y * h)).ToList(), p.Closed)));
        else if (shape.Geometry == ShapeGeometry.Line) outlines.Add(([(x, y), (x + w, y + h)], false));
        else outlines.Add((Outline(shape, w, h).Select(p => (X: x + p.X, Y: y + p.Y)).ToList(), true));
        if (flipped) items.Add(new PushItem(null, 0, cx, cy, shape.FlipH, shape.FlipV));
        for (int k = Layers; k >= 1; k--)
        {
            items.Add(new PushItem(null, Opacity: layerOpacity));
            foreach (var (points, closed) in outlines)
                EmitStroke(points, closed, 2 * glow.Radius * k / Layers, glow.Color, LineStyle.Solid, default, default, items);
            items.Add(new PopItem());
        }
        if (flipped) items.Add(new PopItem());
    }

    /// <summary>
    /// A shape with soft edges: its fill drawn as copies of the outline moved inwards step by step (a real inset of the polygon, which
    /// follows notches and corners; not a scaled copy), each translucent enough that the opacity rises evenly from nothing at the outline to
    /// full one radius inside it. The outline stroke fades with it, so it is left out. The fill keeps its place in the frame and is cut at each
    /// inset outline.
    /// </summary>
    private static void EmitSoftEdged(SheetShape shape, double x, double y, double w, double h, List<LayoutItem> items)
    {
        const int Steps = 6;
        var plain = shape with { LineColor = null, SoftEdge = 0, Glow = null, Shadow = null };
        var outlines = new List<List<(double X, double Y)>>();
        if (shape.Paths is { } paths)
            outlines.AddRange(paths.Where(p => p.Fill && p.Points.Count >= 3 && (p.Closed || shape.Geometry == ShapeGeometry.Custom))
                .Select(p => p.Points.Select(q => (X: x + q.X * w, Y: y + q.Y * h)).ToList()));
        else outlines.Add(Outline(shape, w, h).Select(p => (X: x + p.X, Y: y + p.Y)).ToList());
        double radius = Math.Min(shape.SoftEdge, Math.Min(w, h) / 2 - 0.5);
        for (int j = 1; j <= Steps; j++)
        {
            double inset = radius * (j - 1) / Steps;
            var insets = outlines.Select(o => inset == 0 ? o : InsetPolygon(o, inset)).ToList();
            if (insets.Any(o => o is null)) break; // a notch or a thin part has closed up: nothing further inside
            items.Add(new PushItem(null, Opacity: 1.0 / (Steps - j + 1)));
            foreach (var polygon in insets) EmitFill(plain, polygon!, x, y, w, h, false, items, clip: true);
            items.Add(new PopItem());
        }
    }

    /// <summary>
    /// The polygon moved <paramref name="distance"/> inwards: every edge is moved along its inward normal and neighbouring edges are joined
    /// at their intersection (a mitre, limited at very sharp corners), so a reflex corner moves out of the way and a convex one in. Null when
    /// an edge would turn round (a feature narrower than twice the distance) or the polygon vanishes.
    /// </summary>
    internal static List<(double X, double Y)>? InsetPolygon(IReadOnlyList<(double X, double Y)> polygon, double distance)
    {
        var p = new List<(double X, double Y)>();
        foreach (var q in polygon)
            if (p.Count == 0 || Math.Abs(q.X - p[^1].X) > 1e-9 || Math.Abs(q.Y - p[^1].Y) > 1e-9) p.Add(q);
        if (p.Count > 1 && Math.Abs(p[0].X - p[^1].X) < 1e-9 && Math.Abs(p[0].Y - p[^1].Y) < 1e-9) p.RemoveAt(p.Count - 1);
        int n = p.Count;
        if (n < 3) return null;
        double area = 0;
        for (int i = 0; i < n; i++) area += p[i].X * p[(i + 1) % n].Y - p[(i + 1) % n].X * p[i].Y;
        if (Math.Abs(area) < 1e-9) return null;
        double side = area > 0 ? 1 : -1; // the interior lies to the left of an edge when the area is positive
        var normals = new (double X, double Y)[n];
        for (int i = 0; i < n; i++)
        {
            var (a, b) = (p[i], p[(i + 1) % n]);
            double dx = b.X - a.X, dy = b.Y - a.Y, length = Math.Sqrt(dx * dx + dy * dy);
            normals[i] = (-dy / length * side, dx / length * side);
        }
        var result = new List<(double X, double Y)>(n);
        for (int i = 0; i < n; i++)
        {
            var (n1, n2) = (normals[(i + n - 1) % n], normals[i]);
            double scale = distance / Math.Max(1 + n1.X * n2.X + n1.Y * n2.Y, 0.2);
            result.Add((p[i].X + (n1.X + n2.X) * scale, p[i].Y + (n1.Y + n2.Y) * scale));
        }
        double after = 0;
        for (int i = 0; i < n; i++)
        {
            var (a, b) = (result[i], result[(i + 1) % n]);
            after += a.X * b.Y - b.X * a.Y;
            var (oa, ob) = (p[i], p[(i + 1) % n]);
            if ((b.X - a.X) * (ob.X - oa.X) + (b.Y - a.Y) * (ob.Y - oa.Y) <= 0) return null;
        }
        return after * area > 0 ? result : null;
    }

    private static List<(double X, double Y)> Outline(SheetShape shape, double w, double h)
    {
        double m = Math.Min(w, h);
        switch (shape.Geometry)
        {
            case ShapeGeometry.RoundRectangle:
            {
                double r = Math.Clamp(m * shape.CornerRadius, 0, m / 2);
                var points = new List<(double, double)>();
                (double X, double Y, double Start)[] corners = [(w - r, r, -90), (w - r, h - r, 0), (r, h - r, 90), (r, r, 180)];
                foreach (var (ox, oy, start) in corners)
                    for (int k = 0; k <= 6; k++)
                    {
                        double angle = (start + k * 15) * Math.PI / 180;
                        points.Add((ox + r * Math.Cos(angle), oy + r * Math.Sin(angle)));
                    }
                return points;
            }
            case ShapeGeometry.Ellipse:
            {
                var points = new List<(double, double)>();
                for (int k = 0; k < 72; k++)
                {
                    double angle = k * 5 * Math.PI / 180;
                    points.Add((w / 2 + w / 2 * Math.Cos(angle), h / 2 + h / 2 * Math.Sin(angle)));
                }
                return points;
            }
            case ShapeGeometry.Triangle: return [(w / 2, 0), (w, h), (0, h)];
            case ShapeGeometry.RightTriangle: return [(0, 0), (0, h), (w, h)];
            case ShapeGeometry.Diamond: return [(w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2)];
            case ShapeGeometry.Parallelogram:
            {
                double o = Math.Min(m * 0.25, w / 2);
                return [(o, 0), (w, 0), (w - o, h), (0, h)];
            }
            case ShapeGeometry.Trapezoid:
            {
                double o = Math.Min(m * 0.25, w / 2);
                return [(o, 0), (w - o, 0), (w, h), (0, h)];
            }
            case ShapeGeometry.Hexagon:
            {
                double o = Math.Min(m * 0.25, w / 2);
                return [(o, 0), (w - o, 0), (w, h / 2), (w - o, h), (o, h), (0, h / 2)];
            }
            case ShapeGeometry.RightArrow or ShapeGeometry.LeftArrow:
            {
                double head = Math.Min(m * 0.5, w);
                var points = new List<(double X, double Y)>
                {
                    (0, h * 0.25), (w - head, h * 0.25), (w - head, 0), (w, h / 2), (w - head, h), (w - head, h * 0.75), (0, h * 0.75),
                };
                return shape.Geometry == ShapeGeometry.RightArrow ? points : points.Select(p => (w - p.X, p.Y)).ToList();
            }
            case ShapeGeometry.UpArrow or ShapeGeometry.DownArrow:
            {
                double head = Math.Min(m * 0.5, h);
                var points = new List<(double X, double Y)>
                {
                    (w * 0.25, h), (w * 0.25, head), (0, head), (w / 2, 0), (w, head), (w * 0.75, head), (w * 0.75, h),
                };
                return shape.Geometry == ShapeGeometry.UpArrow ? points : points.Select(p => (p.X, h - p.Y)).ToList();
            }
            default: return [(0, 0), (w, 0), (w, h), (0, h)];
        }
    }

    private static void EmitOutline(SheetShape shape, double x, double y, double w, double h, List<LayoutItem> items)
    {
        var dash = shape.LineStyle switch { StrokeStyle.Dotted => LineStyle.Dotted, StrokeStyle.Dashed => LineStyle.Dashed, _ => LineStyle.Solid };
        double strokeWidth = Math.Max(shape.LineWidth, 0.25);
        if (shape.Paths is { } paths)
        {
            foreach (var path in paths)
            {
                var points = path.Points.Select(p => (X: x + p.X * w, Y: y + p.Y * h)).ToList();
                if (points.Count < 2) continue;
                if (path.Fill && points.Count >= 3 && (path.Closed || shape.Geometry == ShapeGeometry.Custom)) EmitFill(shape, points, x, y, w, h, false, items);
                if (shape.LineColor is not { } color || shape.LineWidth <= 0 || !path.Stroke) continue;
                EmitStroke(points, path.Closed, strokeWidth, color, dash, path.Closed ? default : shape.HeadEnd, path.Closed ? default : shape.TailEnd, items);
            }
            return;
        }
        if (shape.Geometry == ShapeGeometry.Line)
        {
            if (shape.LineColor is { } lineColor)
                EmitStroke([(x, y), (x + w, y + h)], false, strokeWidth, lineColor, dash, shape.HeadEnd, shape.TailEnd, items);
            return;
        }
        var outline = Outline(shape, w, h).Select(p => (X: x + p.X, Y: y + p.Y)).ToList();
        EmitFill(shape, outline, x, y, w, h, shape.Geometry == ShapeGeometry.Ellipse, items);
        if (shape.LineColor is { } color2 && shape.LineWidth > 0)
            EmitStroke(outline, true, shape.LineWidth, color2, dash, default, default, items);
    }

    /// <summary>Fills an outline: with its gradient, as an ellipse when it is one, else as a polygon.</summary>
    private static void EmitFill(SheetShape shape, List<(double X, double Y)> points, double x, double y, double w, double h, bool ellipse, List<LayoutItem> items, bool clip = false)
    {
        if (shape.FillGradient is { } gradient && gradient.Stops.Count >= 2)
        {
            bool box = shape.Geometry == ShapeGeometry.Rectangle && !clip;
            items.Add(new GradientItem(x, y, w, h, gradient.Degree, gradient.Stops)
            {
                Clip = box ? null : points,
                Radial = gradient.Kind == GradientKind.Path,
                Rectangular = gradient.Kind == GradientKind.Path && gradient.Rectangular,
                FocusX = (gradient.Left + gradient.Right) / 2, FocusY = (gradient.Top + gradient.Bottom) / 2,
            });
        }
        else if (shape.FillPattern is { } pattern)
            EmitPattern(pattern, points, x, y, w, h, ellipse, items);
        else if (shape.FillImage is { } picture)
        {
            if (shape.FillImageAlpha < 1) items.Add(new PushItem(null, Opacity: shape.FillImageAlpha));
            if (shape.FillTile is not { } tile || !EmitTiles(picture, tile, points, x, y, w, h, items))
                items.Add(new ImageItem(x, y, w, h, picture) { Clip = shape.Geometry == ShapeGeometry.Rectangle && !clip ? null : points });
            if (shape.FillImageAlpha < 1) items.Add(new PopItem());
        }
        else if (shape.Fill is { } fill)
        {
            if (ellipse) items.Add(new EllipseItem(x, y, w, h, fill));
            else items.Add(new PolygonItem(points, fill));
        }
    }

    /// <summary>
    /// A picture fill repeated in tiles from the corner or edge the tile's alignment names, each clipped to the outline. False (nothing drawn)
    /// when the tiles would be too many to be worth drawing, so that the picture is stretched instead.
    /// </summary>
    private static bool EmitTiles(ImageData picture, ShapeTile tile, List<(double X, double Y)> outline, double x, double y, double w, double h, List<LayoutItem> items)
    {
        string align = tile.Alignment;
        double originX = align.EndsWith('r') ? x + w - tile.Width : align is "t" or "ctr" or "b" ? x + (w - tile.Width) / 2 : x;
        double originY = align.StartsWith('b') ? y + h - tile.Height : align is "l" or "ctr" or "r" ? y + (h - tile.Height) / 2 : y;
        originX += tile.OffsetX;
        originY += tile.OffsetY;
        int firstColumn = (int)Math.Floor((x - originX) / tile.Width), lastColumn = (int)Math.Ceiling((x + w - originX) / tile.Width) - 1;
        int firstRow = (int)Math.Floor((y - originY) / tile.Height), lastRow = (int)Math.Ceiling((y + h - originY) / tile.Height) - 1;
        if ((long)(lastColumn - firstColumn + 1) * (lastRow - firstRow + 1) > 2500) return false;
        for (int row = firstRow; row <= lastRow; row++)
            for (int column = firstColumn; column <= lastColumn; column++)
            {
                double tileX = originX + column * tile.Width, tileY = originY + row * tile.Height;
                // A mirrored fill flips every other tile (counted from the one at the origin) across x, y or both.
                bool flipH = tile.Flip is "x" or "xy" && ((column % 2) + 2) % 2 == 1, flipV = tile.Flip is "y" or "xy" && ((row % 2) + 2) % 2 == 1;
                if (!flipH && !flipV) { items.Add(new ImageItem(tileX, tileY, tile.Width, tile.Height, picture) { Clip = outline }); continue; }
                // The picture is mirrored about its tile's center; the clip is given mirrored too, so that it comes back to the outline.
                double cx = tileX + tile.Width / 2, cy = tileY + tile.Height / 2;
                var clip = outline.Select(p => (flipH ? 2 * cx - p.X : p.X, flipV ? 2 * cy - p.Y : p.Y)).ToList();
                items.Add(new PushItem(null, 0, cx, cy, flipH, flipV));
                items.Add(new ImageItem(tileX, tileY, tile.Width, tile.Height, picture) { Clip = clip });
                items.Add(new PopItem());
            }
        return true;
    }

    /// <summary>
    /// A pattern fill: the background, then the hatch lines of the pattern cut at the outline (anchored at the shape's top-left corner), or
    /// for tint patterns the colors blended by their coverage.
    /// </summary>
    private static void EmitPattern(ShapePattern pattern, List<(double X, double Y)> outline, double x, double y, double w, double h, bool ellipse, List<LayoutItem> items)
    {
        var back = pattern.Kind == HatchKind.Tint ? pattern.Background.Blend(pattern.Foreground, pattern.Coverage) : pattern.Background;
        if (ellipse) items.Add(new EllipseItem(x, y, w, h, back));
        else items.Add(new PolygonItem(outline, back));
        double[] angles = pattern.Kind switch
        {
            HatchKind.Horizontal => [0], HatchKind.Vertical => [90], HatchKind.UpDiagonal => [-45], HatchKind.DownDiagonal => [45],
            HatchKind.Cross => [0, 90], HatchKind.DiagonalCross => [-45, 45], _ => [],
        };
        foreach (double degrees in angles)
        {
            double radians = degrees * Math.PI / 180, dx = Math.Cos(radians), dy = Math.Sin(radians), nx = -dy, ny = dx; // direction and normal
            double lo = outline.Min(p => p.X * nx + p.Y * ny), hi = outline.Max(p => p.X * nx + p.Y * ny);
            double spacing = Math.Max(pattern.Spacing, (hi - lo) / 400), origin = x * nx + y * ny;
            for (double t = origin + Math.Ceiling((lo - origin) / spacing) * spacing; t <= hi; t += spacing)
            {
                var crossings = new List<double>();
                double probe = t + 1e-6; // keeps a line through a vertex from being counted twice
                for (int i = 0; i < outline.Count; i++)
                {
                    var (a, b) = (outline[i], outline[(i + 1) % outline.Count]);
                    double ta = a.X * nx + a.Y * ny, tb = b.X * nx + b.Y * ny;
                    if ((ta - probe) * (tb - probe) >= 0) continue;
                    double u = (probe - ta) / (tb - ta);
                    crossings.Add((a.X + (b.X - a.X) * u) * dx + (a.Y + (b.Y - a.Y) * u) * dy);
                }
                crossings.Sort();
                for (int k = 0; k + 1 < crossings.Count; k += 2)
                    items.Add(new LineItem(nx * t + dx * crossings[k], ny * t + dy * crossings[k], nx * t + dx * crossings[k + 1], ny * t + dy * crossings[k + 1], pattern.Weight, pattern.Foreground));
            }
        }
    }

    /// <summary>Strokes a polyline (closed when asked) and puts arrowheads on its open ends.</summary>
    private static void EmitStroke(List<(double X, double Y)> points, bool closed, double width, Color color, LineStyle dash, LineEnd head, LineEnd tail, List<LayoutItem> items)
    {
        points = new List<(double X, double Y)>(points);
        if (!closed && points.Count >= 2)
        {
            // The line stops short of a filled arrowhead so that its square end does not show beside the tip.
            if (head.Kind != LineEndKind.None) points[0] = ArrowHead(points[0], points[1], head, width, color, items);
            if (tail.Kind != LineEndKind.None) points[^1] = ArrowHead(points[^1], points[^2], tail, width, color, items);
        }
        int segments = closed ? points.Count : points.Count - 1;
        for (int i = 0; i < segments; i++)
        {
            var (a, b) = (points[i], points[(i + 1) % points.Count]);
            items.Add(new LineItem(a.X, a.Y, b.X, b.Y, width, color, dash));
        }
    }

    /// <summary>
    /// Draws an arrowhead at <paramref name="tip"/> pointing away from <paramref name="next"/> and returns where the line should end.
    /// </summary>
    private static (double X, double Y) ArrowHead((double X, double Y) tip, (double X, double Y) next, LineEnd end, double lineWidth, Color color, List<LayoutItem> items)
    {
        double dx = tip.X - next.X, dy = tip.Y - next.Y, length = Math.Sqrt(dx * dx + dy * dy);
        if (length < 1e-6) return tip;
        (dx, dy) = (dx / length, dy / length);
        double unit = Math.Max(lineWidth, 1), headWidth = end.Width * unit, headLength = Math.Min(end.Length * unit, length);
        double px = -dy, py = dx; // perpendicular
        (double X, double Y) At(double along, double across) => (tip.X - dx * along + px * across, tip.Y - dy * along + py * across);
        switch (end.Kind)
        {
            case LineEndKind.Triangle:
                items.Add(new PolygonItem([At(0, 0), At(headLength, headWidth / 2), At(headLength, -headWidth / 2)], color));
                return At(headLength * 0.9, 0);
            case LineEndKind.Stealth:
                items.Add(new PolygonItem([At(0, 0), At(headLength, headWidth / 2), At(headLength * 0.65, 0), At(headLength, -headWidth / 2)], color));
                return At(headLength * 0.65, 0);
            case LineEndKind.Arrow:
                items.Add(new LineItem(tip.X, tip.Y, At(headLength, headWidth / 2).X, At(headLength, headWidth / 2).Y, Math.Max(lineWidth, 0.5), color));
                items.Add(new LineItem(tip.X, tip.Y, At(headLength, -headWidth / 2).X, At(headLength, -headWidth / 2).Y, Math.Max(lineWidth, 0.5), color));
                return tip;
            case LineEndKind.Oval:
                items.Add(new EllipseItem(tip.X - headWidth / 2, tip.Y - headWidth / 2, headWidth, headWidth, color));
                return tip;
            case LineEndKind.Diamond:
                items.Add(new PolygonItem([At(-headLength / 2, 0), At(0, headWidth / 2), At(headLength / 2, 0), At(0, -headWidth / 2)], color));
                return tip;
            default: return tip;
        }
    }

    /// <summary>The text of a shape inside its insets, wrapped to the shape's width and placed by its vertical anchor.</summary>
    private void EmitShapeText(SheetShape shape, double x, double y, double width, double height, List<LayoutItem> items)
    {
        double left = x + shape.TextInsets.Left, right = x + width - shape.TextInsets.Right;
        double top = y + shape.TextInsets.Top, bottom = y + height - shape.TextInsets.Bottom;
        double areaWidth = Math.Max(right - left, 1);
        var blocks = new List<(TextBlock Block, HorizontalAlignment Alignment)>();
        foreach (var paragraph in shape.Paragraphs)
        {
            var text = string.Concat(paragraph.Runs.Select(r => r.Text));
            var font = paragraph.Runs.Count > 0 ? paragraph.Runs[0].Font : RunProperties.Default with { FontFamily = "Calibri", FontSize = paragraph.FontSize };
            var cell = new Cell
            {
                Text = text,
                Runs = paragraph.Runs.Count > 0 ? paragraph.Runs : null,
                Kind = CellValueKind.Text,
                Style = CellStyle.Default with { Wrap = true, HorizontalAlignment = shape.WrapText ? paragraph.Alignment : HorizontalAlignment.Left, Font = font },
            };
            blocks.Add((BuildBlock(cell, text, shape.WrapText ? areaWidth : 1e6), paragraph.Alignment));
        }
        double total = blocks.Sum(b => b.Block.Height);
        double lineTop = shape.TextAnchor switch
        {
            ShapeTextAnchor.Middle => (top + bottom) / 2 - total / 2,
            ShapeTextAnchor.Bottom => bottom - total,
            _ => top,
        };
        foreach (var (block, alignment) in blocks)
        {
            for (int i = 0; i < block.Lines.Count; i++)
            {
                var (pitch, ascent, descent) = block.Metrics[i];
                var line = block.Lines[i];
                if (line.Atoms.Count > 0)
                {
                    double x0 = left;
                    if (!shape.WrapText)
                    {
                        double w = LineWidth(line);
                        x0 = alignment switch { HorizontalAlignment.Right => right - w, HorizontalAlignment.Center => (left + right) / 2 - w / 2, _ => left };
                    }
                    EmitLine(line, x0, lineTop + (pitch - (ascent + descent)) / 2 + ascent, x - 1e6, x + 1e6, items);
                }
                lineTop += pitch;
            }
        }
    }
}
