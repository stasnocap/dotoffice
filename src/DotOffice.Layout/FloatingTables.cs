using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>A laid-out table (<c>w:tblpPr</c>) placed like a floating picture; its items are relative to its top-left corner.</summary>
internal sealed class FloatingTable(double width, double height, IReadOnlyList<LayoutItem> items)
    : FloatingObject(width, height, RunProperties.Default)
{
    public IReadOnlyList<LayoutItem> Items { get; } = items;

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
