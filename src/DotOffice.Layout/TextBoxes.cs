using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>Lays out the content of text boxes (floating shapes with <c>w:txbxContent</c>).</summary>
internal static class TextBoxes
{
    /// <summary>
    /// Fill, outline and content of <paramref name="box"/> with its top-left corner at (<paramref name="x"/>, <paramref name="y"/>).
    /// The content is clipped to the box; a rotated box turns as a whole about its center.
    /// </summary>
    public static List<LayoutItem> Realize(FloatingTextBox box, double x, double y, FlowContext context)
    {
        var items = new List<LayoutItem>();
        double w = box.Width, h = box.Height;
        bool plain = box.Geometry == ShapeGeometry.Rectangle && box.FillGradient is null;
        if (!plain)
        {
            // Other outlines and gradients are drawn by the same code as the shapes of spreadsheets.
            var shape = new SheetShape
            {
                Geometry = box.Geometry, CornerRadius = box.CornerRadius, Adjust1 = box.Adjust1, Adjust2 = box.Adjust2,
                Fill = box.Fill, FillGradient = box.FillGradient, LineColor = box.Outline, LineWidth = box.OutlineWidth,
                LineStyle = box.OutlineStyle switch { LineStyleKind.Dotted => StrokeStyle.Dotted, LineStyleKind.Dashed => StrokeStyle.Dashed, _ => StrokeStyle.Solid },
            };
            SheetLayoutEngine.EmitOutline(shape, x, y, w, h, items);
        }
        else if (box.Fill is { } fill) items.Add(new RectItem(x, y, w, h, fill));
        if (plain && box.Outline is { } outline && box.OutlineWidth > 0)
        {
            var style = box.OutlineStyle switch { LineStyleKind.Dotted => LineStyle.Dotted, LineStyleKind.Dashed => LineStyle.Dashed, _ => LineStyle.Solid };
            double lw = box.OutlineWidth;
            items.Add(new LineItem(x, y, x + w, y, lw, outline, style));
            items.Add(new LineItem(x + w, y, x + w, y + h, lw, outline, style));
            items.Add(new LineItem(x + w, y + h, x, y + h, lw, outline, style));
            items.Add(new LineItem(x, y + h, x, y, lw, outline, style));
        }

        var inset = box.Insets;
        // The text rectangle of an ellipse is the rectangle inscribed in it; of a rounded rectangle, the one inside its corners.
        if (box.Geometry == ShapeGeometry.Ellipse)
            inset = new Thickness(inset.Top + h * 0.14645, inset.Right + w * 0.14645, inset.Bottom + h * 0.14645, inset.Left + w * 0.14645);
        else if (box.Geometry == ShapeGeometry.RoundRectangle)
        {
            double corner = Math.Clamp(Math.Min(w, h) * box.CornerRadius, 0, Math.Min(w, h) / 2) * 0.29289;
            inset = new Thickness(inset.Top + corner, inset.Right + corner, inset.Bottom + corner, inset.Left + corner);
        }
        double bodyWidth = w - inset.Left - inset.Right;
        double bodyHeight = h - inset.Top - inset.Bottom;
        var direction = box.TextDirection;
        // Turned text is laid out horizontally in a frame as long as the body is high, and the frame is turned about the body's centre.
        if (direction != TextBoxDirection.Horizontal) (bodyWidth, bodyHeight) = (bodyHeight, bodyWidth);
        if (box.Blocks.Count > 0 && bodyWidth > 1 && context.Depth < 8)
        {
            var nested = context.Nested();
            var stacked = new List<(double Y, Slice Slice)>();
            double left = direction == TextBoxDirection.Horizontal ? x + inset.Left : x + inset.Left + (w - inset.Left - inset.Right) / 2 - bodyWidth / 2;
            double height = BlockFlow.Stack(BlockFlow.Build(box.Blocks, left, bodyWidth, nested), stacked);
            double free = Math.Max(0, bodyHeight - height);
            double by = y + inset.Top;
            double cx = x + inset.Left + (w - inset.Left - inset.Right) / 2, cy = by + (h - inset.Top - inset.Bottom) / 2;
            double top = (direction == TextBoxDirection.Horizontal ? by : cy - bodyHeight / 2) + box.VerticalAnchor switch { TextBoxAnchor.Center => free / 2, TextBoxAnchor.Bottom => free, _ => 0 };
            var content = new List<LayoutItem>();
            foreach (var (sy, slice) in stacked)
                foreach (var item in slice.Items)
                {
                    var moved = item.Offset(0, top + sy);
                    if (moved is TextItem { Field: { } field } text) moved = text with { Text = context.FieldText(field), Glyphs = null };
                    content.Add(moved);
                }
            items.Add(new PushItem((x, y, w, h)));
            // PushItem turns counter-clockwise; text turned clockwise needs a negative angle.
            if (direction != TextBoxDirection.Horizontal)
                items.Add(new PushItem(null, direction == TextBoxDirection.CounterClockwise ? 90 : -90, cx, cy));
            items.AddRange(direction == TextBoxDirection.Clockwise ? UprightText.Apply(content) : content);
            if (direction != TextBoxDirection.Horizontal) items.Add(new PopItem());
            items.Add(new PopItem());
        }

        if (Math.Abs(box.Rotation) > 0.01)
        {
            // PushItem turns counter-clockwise; DrawingML rotates clockwise.
            items.Insert(0, new PushItem(null, -box.Rotation, x + w / 2, y + h / 2));
            items.Add(new PopItem());
        }
        return items;
    }
}
