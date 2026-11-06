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
        if (box.Fill is { } fill) items.Add(new RectItem(x, y, w, h, fill));
        if (box.Outline is { } outline && box.OutlineWidth > 0)
        {
            var style = box.OutlineStyle switch { LineStyleKind.Dotted => LineStyle.Dotted, LineStyleKind.Dashed => LineStyle.Dashed, _ => LineStyle.Solid };
            double lw = box.OutlineWidth;
            items.Add(new LineItem(x, y, x + w, y, lw, outline, style));
            items.Add(new LineItem(x + w, y, x + w, y + h, lw, outline, style));
            items.Add(new LineItem(x + w, y + h, x, y + h, lw, outline, style));
            items.Add(new LineItem(x, y + h, x, y, lw, outline, style));
        }

        var inset = box.Insets;
        double bodyWidth = w - inset.Left - inset.Right;
        if (box.Blocks.Count > 0 && bodyWidth > 1 && context.Depth < 8)
        {
            var nested = context.Nested();
            var stacked = new List<(double Y, Slice Slice)>();
            double height = BlockFlow.Stack(BlockFlow.Build(box.Blocks, x + inset.Left, bodyWidth, nested), stacked);
            double free = Math.Max(0, h - inset.Top - inset.Bottom - height);
            double top = y + inset.Top + box.VerticalAnchor switch { TextBoxAnchor.Center => free / 2, TextBoxAnchor.Bottom => free, _ => 0 };
            var content = new List<LayoutItem>();
            foreach (var (sy, slice) in stacked)
                foreach (var item in slice.Items)
                {
                    var moved = item.Offset(0, top + sy);
                    if (moved is TextItem { Field: { } field } text) moved = text with { Text = context.FieldText(field), Glyphs = null };
                    content.Add(moved);
                }
            items.Add(new PushItem((x, y, w, h)));
            items.AddRange(content);
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
