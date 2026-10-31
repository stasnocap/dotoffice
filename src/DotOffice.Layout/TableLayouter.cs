using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>A laid-out line of cell content; <see cref="Y"/> is relative to the row's top.</summary>
internal sealed record CellLine(double Y, double Height, IReadOnlyList<LayoutItem> Items, IReadOnlyList<Note> Footnotes,
    IReadOnlyList<FloatingImage> Floats);

/// <summary>The part of a (possibly vertically merged) cell that falls into one row.</summary>
internal sealed record CellPart(double X, double Width, Thickness Margins, Color? Shading, BorderSet Borders,
    bool DrawTop, bool DrawBottom, IReadOnlyList<CellLine> Lines);

/// <summary>A table row. Its items are regenerated from cell parts so the row can be split across pages.</summary>
internal sealed class RowSlice : Slice
{
    private readonly IReadOnlyList<CellPart> _cells;
    private readonly bool _cantSplit;

    public RowSlice(IReadOnlyList<CellPart> cells, double height, bool cantSplit)
    {
        _cells = cells;
        _cantSplit = cantSplit;
        Height = height;
        Footnotes = cells.SelectMany(c => c.Lines).SelectMany(l => l.Footnotes).ToList();
        Floats = cells.SelectMany(c => c.Lines).SelectMany(l => l.Floats).ToList();
        Render();
    }

    private void Render()
    {
        Items.Clear();
        foreach (var c in _cells)
            if (c.Shading is { } fill) Items.Add(new RectItem(c.X, 0, c.Width, Height, fill));
        foreach (var c in _cells)
            foreach (var line in c.Lines)
                foreach (var item in line.Items) Items.Add(item.Offset(0, line.Y));
        foreach (var c in _cells)
        {
            var b = c.Borders;
            if (c.DrawTop && b.Top.IsVisible) ParagraphLayouter.AddBorder(Items, b.Top, c.X, 0, c.X + c.Width, 0);
            if (c.DrawBottom && b.Bottom.IsVisible) ParagraphLayouter.AddBorder(Items, b.Bottom, c.X, Height, c.X + c.Width, Height);
            if (b.Left.IsVisible) ParagraphLayouter.AddBorder(Items, b.Left, c.X, 0, c.X, Height);
            if (b.Right.IsVisible) ParagraphLayouter.AddBorder(Items, b.Right, c.X + c.Width, 0, c.X + c.Width, Height);
        }
    }

    public override (Slice First, Slice Remainder)? Split(double available)
    {
        if (_cantSplit || available <= 0) return null;
        var first = new List<CellPart>(_cells.Count);
        var rest = new List<CellPart>(_cells.Count);
        bool anyFirst = false, anyRest = false;
        double firstHeight = 0, restHeight = 0;
        foreach (var c in _cells)
        {
            double limit = available - c.Margins.Bottom;
            var head = c.Lines.TakeWhile(l => l.Y + l.Height <= limit + 0.01).ToList();
            var tail = c.Lines.Skip(head.Count).ToList();
            double shift = tail.Count > 0 ? tail[0].Y - c.Margins.Top : 0;
            var moved = tail.Select(l => l with { Y = l.Y - shift }).ToList();
            anyFirst |= head.Count > 0;
            anyRest |= moved.Count > 0;
            if (head.Count > 0) firstHeight = Math.Max(firstHeight, head[^1].Y + head[^1].Height + c.Margins.Bottom);
            if (moved.Count > 0) restHeight = Math.Max(restHeight, moved[^1].Y + moved[^1].Height + c.Margins.Bottom);
            first.Add(c with { Lines = head, DrawBottom = false });
            rest.Add(c with { Lines = moved, DrawTop = false });
        }
        if (!anyFirst || !anyRest) return null;
        var firstSlice = new RowSlice(first, Math.Min(available, firstHeight), cantSplit: false)
        {
            KeepWithPrevious = KeepWithPrevious, BreakBefore = BreakBefore, RepeatHeader = RepeatHeader, IsHeaderRow = IsHeaderRow,
        };
        var restSlice = new RowSlice(rest, restHeight, cantSplit: false) { RepeatHeader = RepeatHeader, IsHeaderRow = IsHeaderRow };
        return (firstSlice, restSlice);
    }
}

internal static class TableLayouter
{
    private sealed record Placed(TableCell Cell, int Row, int Column, int Span);

    public static List<Slice> Layout(Table table, double x0, double available, FlowContext ctx)
    {
        var rows = table.Rows;
        if (rows.Count == 0) return [];
        int columns = Math.Max(table.Grid.Count, rows.Max(r => r.GridBefore + r.Cells.Sum(c => Math.Max(1, c.GridSpan))));
        if (columns == 0) return [];

        var placed = new List<List<Placed>>();
        for (int r = 0; r < rows.Count; r++)
        {
            int col = rows[r].GridBefore;
            var list = new List<Placed>();
            foreach (var cell in rows[r].Cells)
            {
                int span = Math.Clamp(cell.GridSpan, 1, columns - Math.Min(col, columns - 1));
                list.Add(new Placed(cell, r, col, span));
                col += span;
            }
            placed.Add(list);
        }

        var widths = ResolveColumns(table, placed, columns, available, ctx);
        double tableWidth = widths.Sum();
        double left = table.Properties.Alignment switch
        {
            TableAlignment.Center => x0 + (available - tableWidth) / 2,
            TableAlignment.Right => x0 + available - tableWidth,
            _ => x0 + table.Properties.Indent,
        };
        if (!ctx.Settings.TableEdgeAtIndent && table.Properties.Alignment == TableAlignment.Left && placed[0].Count > 0)
            left -= placed[0][0].Cell.Margins.Left;
        var colX = new double[columns + 1];
        colX[0] = left;
        for (int c = 0; c < columns; c++) colX[c + 1] = colX[c] + widths[c];

        // Vertical merge groups: restart cell → last row index of the group.
        var groupEnd = new Dictionary<Placed, int>();
        var continued = new HashSet<Placed>();
        for (int r = 0; r < rows.Count; r++)
        {
            foreach (var p in placed[r])
            {
                if (p.Cell.VerticalMerge != VerticalMerge.Restart) continue;
                int end = r;
                while (end + 1 < rows.Count && placed[end + 1].FirstOrDefault(q => q.Column == p.Column) is { Cell.VerticalMerge: VerticalMerge.Continue } next)
                {
                    continued.Add(next);
                    end++;
                }
                groupEnd[p] = end;
            }
        }

        // Lay out cell contents.
        var contents = new Dictionary<Placed, (List<(double Y, Slice Slice)> Lines, double Height)>();
        var nested = ctx.Nested();
        foreach (var p in placed.SelectMany(r => r))
        {
            double w = colX[p.Column + p.Span] - colX[p.Column];
            var m = p.Cell.Margins;
            double inner = Math.Max(1, w - m.Left - m.Right);
            var slices = BlockFlow.Build(p.Cell.Blocks, colX[p.Column] + m.Left, inner, nested);
            var lines = new List<(double, Slice)>();
            double h = BlockFlow.Stack(slices, lines);
            contents[p] = (lines, h);
        }

        // Row heights: single-row cells first, then grow the last row of each merge group to fit it.
        var heights = new double[rows.Count];
        for (int r = 0; r < rows.Count; r++)
        {
            double h = 0;
            foreach (var p in placed[r])
            {
                if (groupEnd.TryGetValue(p, out var end) && end > r) continue;
                var m = p.Cell.Margins;
                h = Math.Max(h, m.Top + contents[p].Height + m.Bottom);
            }
            heights[r] = rows[r].HeightRule switch
            {
                RowHeightRule.Exact when rows[r].Height > 0 => rows[r].Height,
                RowHeightRule.AtLeast => Math.Max(h, rows[r].Height),
                _ => h,
            };
        }
        foreach (var (p, end) in groupEnd)
        {
            if (end == p.Row) continue;
            double needed = p.Cell.Margins.Top + contents[p].Height + p.Cell.Margins.Bottom;
            double have = 0;
            for (int r = p.Row; r <= end; r++) have += heights[r];
            if (needed > have && rows[end].HeightRule != RowHeightRule.Exact) heights[end] += needed - have;
        }

        // Build cell parts per row.
        var parts = Enumerable.Range(0, rows.Count).Select(_ => new List<CellPart>()).ToArray();
        foreach (var p in placed.SelectMany(r => r))
        {
            if (continued.Contains(p)) continue;
            int end = groupEnd.TryGetValue(p, out var e) ? e : p.Row;
            var m = p.Cell.Margins;
            double groupHeight = 0;
            var rowTops = new double[end - p.Row + 1];
            for (int r = p.Row; r <= end; r++)
            {
                rowTops[r - p.Row] = groupHeight;
                groupHeight += heights[r];
            }
            var (lines, contentHeight) = contents[p];
            double free = Math.Max(0, groupHeight - m.Top - m.Bottom - contentHeight);
            double offset = m.Top + p.Cell.VerticalAlignment switch
            {
                CellVerticalAlignment.Center => free / 2,
                CellVerticalAlignment.Bottom => free,
                _ => 0,
            };
            var perRow = Enumerable.Range(0, rowTops.Length).Select(_ => new List<CellLine>()).ToArray();
            foreach (var (y, slice) in lines)
            {
                // Assign each line to the row holding its bottom edge, so it is painted after every
                // row background it overlaps (rows paint in order).
                double top = offset + y, bottom = top + slice.Height;
                int k = rowTops.Length - 1;
                while (k > 0 && rowTops[k] >= bottom - 0.01) k--;
                perRow[k].Add(new CellLine(top - rowTops[k], slice.Height, slice.Items, slice.Footnotes, slice.Floats));
            }
            double x = colX[p.Column], w = colX[p.Column + p.Span] - x;
            for (int r = p.Row; r <= end; r++)
            {
                var own = placed[r].FirstOrDefault(q => q.Column == p.Column)?.Cell ?? p.Cell;
                var borders = p.Cell.Borders with
                {
                    Left = own.Borders.Left,
                    Right = own.Borders.Right,
                    Bottom = r == end ? own.Borders.Bottom : Border.None,
                };
                parts[r].Add(new CellPart(x, w, m, p.Cell.Shading, borders, r == p.Row, r == end, perRow[r - p.Row]));
            }
        }

        var result = new List<Slice>(rows.Count);
        var header = new List<Slice>();
        bool leadingHeaders = true;
        for (int r = 0; r < rows.Count; r++)
        {
            leadingHeaders &= rows[r].IsHeader;
            var slice = new RowSlice(parts[r], heights[r], rows[r].CantSplit || leadingHeaders) { IsHeaderRow = leadingHeaders };
            if (leadingHeaders) header.Add(slice);
            else if (header.Count > 0)
            {
                slice.RepeatHeader = header;
                if (r == header.Count) slice.KeepWithPrevious = true; // never leave headers alone at a page bottom
            }
            // Rows in a vertical merge group stay together unless the group is taller than a page.
            if (r > 0 && placed[r].Any(continued.Contains)) slice.KeepWithPrevious = true;
            result.Add(slice);
        }
        return result;
    }

    private static double[] ResolveColumns(Table table, List<List<Placed>> placed, int columns, double available, FlowContext ctx)
    {
        var props = table.Properties;
        double? target = props.Width.Unit switch
        {
            TableWidthUnit.Percent => available * props.Width.Value / 100,
            TableWidthUnit.Points when props.Width.Value > 0 => props.Width.Value,
            _ => null,
        };

        var grid = table.Grid;
        if (grid.Count >= columns && grid.Take(columns).Sum() > 0)
        {
            var w = grid.Take(columns).ToArray();
            double sum = w.Sum();
            // Word recomputes the grid for percentage widths; otherwise the saved grid is authoritative.
            if (props.Width.Unit == TableWidthUnit.Percent && target is { } t && Math.Abs(sum - t) > 0.5)
                for (int i = 0; i < w.Length; i++) w[i] *= t / sum;
            return w;
        }

        // Auto layout from content: min = longest word, max = unwrapped content (CSS-style distribution).
        var min = new double[columns];
        var max = new double[columns];
        var preferred = new double?[columns];
        foreach (var p in placed.SelectMany(r => r))
        {
            var (cmin, cmax) = Intrinsic(p.Cell, ctx);
            if (p.Cell.Width.Unit == TableWidthUnit.Points && p.Cell.Width.Value > 0) cmax = Math.Max(cmin, p.Cell.Width.Value);
            if (p.Cell.Width.Unit == TableWidthUnit.Percent && p.Cell.Width.Value > 0)
                cmax = Math.Max(cmin, available * p.Cell.Width.Value / 100);
            if (p.Span == 1)
            {
                min[p.Column] = Math.Max(min[p.Column], cmin);
                max[p.Column] = Math.Max(max[p.Column], cmax);
                if (p.Cell.Width.Unit == TableWidthUnit.Points && p.Cell.Width.Value > 0) preferred[p.Column] = p.Cell.Width.Value;
            }
        }
        foreach (var p in placed.SelectMany(r => r).Where(p => p.Span > 1))
        {
            var (cmin, cmax) = Intrinsic(p.Cell, ctx);
            double haveMin = 0, haveMax = 0;
            for (int c = p.Column; c < p.Column + p.Span; c++) { haveMin += min[c]; haveMax += max[c]; }
            for (int c = p.Column; c < p.Column + p.Span; c++)
            {
                if (cmin > haveMin) min[c] += (cmin - haveMin) / p.Span;
                if (cmax > haveMax) max[c] += (cmax - haveMax) / p.Span;
            }
        }
        for (int c = 0; c < columns; c++)
        {
            if (preferred[c] is { } pw) max[c] = Math.Max(min[c], pw);
            if (max[c] < min[c]) max[c] = min[c];
            if (max[c] <= 0) max[c] = min[c] = 12;
        }

        double sumMin = min.Sum(), sumMax = max.Sum();
        double width = target ?? Math.Min(available, sumMax);
        var result = new double[columns];
        if (sumMin >= width)
            for (int c = 0; c < columns; c++) result[c] = min[c] * width / sumMin;
        else if (sumMax <= width)
            for (int c = 0; c < columns; c++) result[c] = max[c] * width / sumMax;
        else
            for (int c = 0; c < columns; c++) result[c] = min[c] + (max[c] - min[c]) * (width - sumMin) / (sumMax - sumMin);
        return result;
    }

    internal static (double Min, double Max) Intrinsic(TableCell cell, FlowContext ctx)
    {
        double min = 0, max = 0;
        foreach (var block in cell.Blocks)
        {
            switch (block)
            {
                case Paragraph p:
                    var (pmin, pmax) = LineBuilder.Intrinsic(ctx.Itemizer.Itemize(p, ctx.FieldText), ctx.Settings.DefaultTabStop);
                    double indent = Math.Max(0, p.Properties.LeftIndent) + Math.Max(0, p.Properties.RightIndent);
                    min = Math.Max(min, pmin + indent);
                    max = Math.Max(max, pmax + indent + Math.Max(0, p.Properties.FirstLineIndent));
                    break;
                case Table t:
                    double grid = t.Grid.Sum();
                    double tmin = 0, tmax = 0;
                    if (grid > 0) tmin = tmax = grid;
                    else
                        foreach (var row in t.Rows)
                        {
                            double rmin = 0, rmax = 0;
                            foreach (var c in row.Cells)
                            {
                                var (a, b) = Intrinsic(c, ctx);
                                rmin += a;
                                rmax += b;
                            }
                            tmin = Math.Max(tmin, rmin);
                            tmax = Math.Max(tmax, rmax);
                        }
                    min = Math.Max(min, tmin);
                    max = Math.Max(max, tmax);
                    break;
            }
        }
        return (min + cell.Margins.Left + cell.Margins.Right, max + cell.Margins.Left + cell.Margins.Right);
    }
}