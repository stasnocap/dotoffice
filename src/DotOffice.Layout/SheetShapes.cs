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
        EmitOutline(shape, x, y, width, height, items);
        if (flipped) items.Add(new PopItem());
        if (shape.Paragraphs.Count > 0 && shape.Geometry != ShapeGeometry.Line) EmitShapeText(shape, x, y, width, height, items);
        for (int i = 0; i < pushed; i++) items.Add(new PopItem());
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
    private static void EmitFill(SheetShape shape, List<(double X, double Y)> points, double x, double y, double w, double h, bool ellipse, List<LayoutItem> items)
    {
        if (shape.FillGradient is { } gradient && gradient.Stops.Count >= 2)
        {
            bool box = shape.Geometry == ShapeGeometry.Rectangle;
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
            items.Add(new ImageItem(x, y, w, h, picture) { Clip = shape.Geometry == ShapeGeometry.Rectangle ? null : points });
        else if (shape.Fill is { } fill)
        {
            if (ellipse) items.Add(new EllipseItem(x, y, w, h, fill));
            else items.Add(new PolygonItem(points, fill));
        }
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
