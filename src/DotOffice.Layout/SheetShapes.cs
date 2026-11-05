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
        if (shape.Geometry == ShapeGeometry.Line)
        {
            if (shape.LineColor is { } lineColor) items.Add(new LineItem(x, y, x + w, y + h, Math.Max(shape.LineWidth, 0.25), lineColor, dash));
            return;
        }
        var points = Outline(shape, w, h).Select(p => (X: x + p.X, Y: y + p.Y)).ToList();
        if (shape.Fill is { } fill)
        {
            if (shape.Geometry == ShapeGeometry.Ellipse) items.Add(new EllipseItem(x, y, w, h, fill));
            else items.Add(new PolygonItem(points, fill));
        }
        if (shape.LineColor is { } color && shape.LineWidth > 0)
            for (int i = 0; i < points.Count; i++)
            {
                var (a, b) = (points[i], points[(i + 1) % points.Count]);
                items.Add(new LineItem(a.X, a.Y, b.X, b.Y, shape.LineWidth, color, dash));
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
