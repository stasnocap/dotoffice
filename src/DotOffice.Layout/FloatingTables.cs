using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>A laid-out table (<c>w:tblpPr</c>) placed like a floating picture; its items are relative to its top-left corner.</summary>
internal sealed class FloatingTable(double width, double height, IReadOnlyList<LayoutItem> items)
    : FloatingObject(width, height, RunProperties.Default)
{
    public IReadOnlyList<LayoutItem> Items { get; } = items;

    /// <summary>
    /// Lays out paragraphs sharing one <c>w:framePr</c> as a floating box. Without a frame width the box is as wide as the
    /// widest paragraph's text (indents included), at most the column; the height is the text's unless the frame sets one.
    /// </summary>
    public static FloatingTable? BuildFrame(IReadOnlyList<Paragraph> paragraphs, FrameProperties frame, double columnWidth, double maxHeight, FlowContext ctx, Paragraph? following = null)
    {
        double width = frame.Width ?? Math.Min(columnWidth, paragraphs.Max(p =>
            ctx.Itemizer.Itemize(p, ctx.FieldText).Where(a => a.Kind != AtomKind.LineBreak).Sum(a => a.Width)
            + p.Properties.LeftIndent + Math.Max(0, p.Properties.FirstLineIndent) + p.Properties.RightIndent + 0.5));
        width = Math.Min(width, columnWidth);
        var slices = BlockFlow.Build(paragraphs.Cast<Block>().ToList(), 0, width, ctx);
        if (slices.Count == 0) return null;
        var stacked = new List<(double Y, Slice Slice)>();
        double height = BlockFlow.Stack(slices, stacked);
        // A drop cap spans exactly `lines` lines of the paragraph after it, whatever the letter's own height.
        if (frame.DropLines is { } n && frame.Height is null && following is not null
            && BlockFlow.Build([following], 0, columnWidth, ctx) is [{ } first, ..] && first.Height > 0)
            height = n * (first.Height - first.LeadingSpace);
        if (frame.Height is { } fixedHeight) height = frame.HeightExact ? fixedHeight : Math.Max(height, fixedHeight);
        if (height <= 0 || height > maxHeight) return null;
        var items = new List<LayoutItem>();
        foreach (var (y, slice) in stacked)
            foreach (var item in slice.Items) items.Add(item.Offset(0, y));
        return new FloatingTable(width, height, items)
        {
            Horizontal = frame.Horizontal, Vertical = frame.Vertical, Distance = frame.Distance, Wrap = frame.Wrap,
        };
    }

    public static FloatingTable? Build(Table table, double columnWidth, double maxHeight, FlowContext ctx)
    {
        var position = table.Properties.Position;
        if (position is null) return null;
        // A floating table sits at the position of its tblpPr: the table's own alignment and indent do not apply.
        var plain = new Table { Properties = table.Properties with { Alignment = TableAlignment.Left, Indent = 0, Position = null } };
        plain.Grid.AddRange(table.Grid);
        plain.Rows.AddRange(table.Rows);
        var slices = TableLayouter.Layout(plain, 0, columnWidth, ctx);
        if (slices.Count == 0 || slices.Any(s => s is not RowSlice)) return null;
        var stacked = new List<(double Y, Slice Slice)>();
        double height = BlockFlow.Stack(slices, stacked);
        if (height <= 0 || height > maxHeight) return null;
        double left = slices.Cast<RowSlice>().Min(s => s.Extent.Left), right = slices.Cast<RowSlice>().Max(s => s.Extent.Right);
        var items = new List<LayoutItem>();
        foreach (var (y, slice) in stacked)
            foreach (var item in slice.Items) items.Add(item.Offset(-left, y));
        // The position is carried by the FloatingObject members so the picture machinery places and wraps it.
        return new FloatingTable(right - left, height, items)
        {
            Horizontal = position.Horizontal, Vertical = position.Vertical, Distance = position.Distance, Wrap = TextWrap.Square,
        };
    }
}
