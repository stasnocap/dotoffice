using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>
/// A vertical unit of flow (a paragraph line, a table row…). Item x coordinates are absolute page
/// coordinates; y coordinates are relative to the slice's top. Pagination only assigns a y offset.
/// </summary>
internal class Slice
{
    public double Height { get; set; }

    /// <summary>Space at the top of the slice (paragraph spacing before), included in <see cref="Height"/>.
    /// Dropped when the slice starts a page after a natural (not forced) page break, as in Word.</summary>
    public double LeadingSpace { get; set; }

    /// <summary>Space below the slice (paragraph spacing after); may be cut at the bottom of a page.</summary>
    public double TrailingSpace { get; set; }

    public List<LayoutItem> Items { get; init; } = [];

    /// <summary>Forced page break before this slice.</summary>
    public bool BreakBefore { get; set; }

    /// <summary>Forced column break before this slice (a page break in single-column sections).</summary>
    public bool ColumnBreakBefore { get; set; }

    /// <summary>No page break allowed between the previous slice and this one (keep lines, widows, keep-with-next).</summary>
    public bool KeepWithPrevious { get; set; }

    /// <summary>Table header rows to repeat when this slice starts a new page.</summary>
    public IReadOnlyList<Slice>? RepeatHeader { get; set; }

    public bool IsHeaderRow { get; set; }

    /// <summary>Footnotes referenced in this slice; they go to the bottom of the page the slice lands on.</summary>
    public IReadOnlyList<Note> Footnotes { get; set; } = [];

    /// <summary>Floating images anchored in this slice; positioned when the slice is placed.</summary>
    public IReadOnlyList<FloatingImage> Floats { get; set; } = [];

    /// <summary>Splits the slice so the first part fits in <paramref name="available"/>; null if it cannot be split.</summary>
    public virtual (Slice First, Slice Remainder)? Split(double available) => null;
}

/// <summary>An area text must avoid (a wrapped floating image plus its distance), in page coordinates.</summary>
internal readonly record struct Exclusion(double X, double Y, double Width, double Height, bool FullWidth)
{
    public double Right => X + Width;
    public double Bottom => Y + Height;
}

/// <summary>Exclusions affecting one column, queried by the vertical extent of a line.</summary>
internal sealed class Obstacles(IReadOnlyList<Exclusion> exclusions, double columnX, double columnWidth, double startY)
{
    /// <summary>Narrowest line worth placing beside an obstacle; narrower gaps are skipped.</summary>
    private const double MinWidth = 36;

    /// <summary>Page y at which the block being laid out starts.</summary>
    public double StartY { get; } = startY;

    /// <summary>
    /// Space taken from the column's left and right edges by obstacles beside a line spanning
    /// [<paramref name="top"/>, <paramref name="bottom"/>), or the y to retry at when the line cannot go there.
    /// </summary>
    public (double Left, double Right, double? Below) At(double top, double bottom)
    {
        double left = 0, right = 0, below = double.MaxValue;
        bool blocked = false;
        foreach (var e in exclusions)
        {
            if (e.Y >= bottom || e.Bottom <= top || e.Right <= columnX || e.X >= columnX + columnWidth) continue;
            below = Math.Min(below, e.Bottom);
            if (e.FullWidth) { blocked = true; continue; }
            // Text goes on the wider side of the obstacle.
            if (e.X - columnX >= columnX + columnWidth - e.Right) right = Math.Max(right, columnX + columnWidth - e.X);
            else left = Math.Max(left, e.Right - columnX);
        }
        if (blocked || columnWidth - left - right < MinWidth) return (0, 0, below);
        return (left, right, null);
    }
}

internal sealed class FlowContext
{
    public required FontSelector Fonts { get; init; }
    public required Itemizer Itemizer { get; init; }
    public required DocumentSettings Settings { get; init; }
    public required Func<FieldRun, string> FieldText { get; init; }
    public IConversionWarningSink? Warnings { get; init; }
    public int Depth { get; init; }

    public FlowContext Nested() => new()
    {
        Fonts = Fonts, Itemizer = Itemizer, Settings = Settings, FieldText = FieldText, Warnings = Warnings, Depth = Depth + 1,
    };
}

internal static class BlockFlow
{
    private const int MaxDepth = 24;

    /// <summary>Lays out blocks into slices for a column starting at page x <paramref name="x0"/>.</summary>
    public static List<Slice> Build(IReadOnlyList<Block> blocks, double x0, double width, FlowContext ctx)
    {
        var slices = new List<Slice>();
        if (ctx.Depth > MaxDepth) return slices;
        bool keepWithNext = false;
        for (int i = 0; i < blocks.Count; i++) slices.AddRange(BuildBlock(blocks, i, x0, width, ctx, ref keepWithNext));
        return slices;
    }

    /// <summary>
    /// Lays out one block of <paramref name="blocks"/> (neighbors matter for spacing and borders).
    /// <paramref name="keepWithNext"/> carries the previous block's keep-with-next into this one.
    /// </summary>
    public static List<Slice> BuildBlock(IReadOnlyList<Block> blocks, int i, double x0, double width, FlowContext ctx, ref bool keepWithNext,
        Obstacles? obstacles = null)
    {
        List<Slice> produced;
        switch (blocks[i])
        {
            case Paragraph p:
                var prev = i > 0 ? blocks[i - 1] as Paragraph : null;
                var next = i + 1 < blocks.Count ? blocks[i + 1] as Paragraph : null;
                produced = ParagraphLayouter.Layout(p, prev, next, x0, width, ctx, obstacles);
                break;
            case Table t:
                produced = TableLayouter.Layout(t, x0, width, ctx);
                break;
            default:
                return [];
        }
        if (produced.Count == 0) return produced;
        if (keepWithNext) produced[0].KeepWithPrevious = true;
        keepWithNext = blocks[i] is Paragraph { Properties.KeepWithNext: true };
        return produced;
    }

    /// <summary>Stacks slices without pagination (cells, headers); returns total height.</summary>
    public static double Stack(List<Slice> slices, List<(double Y, Slice Slice)> output)
    {
        double y = 0;
        for (int i = 0; i < slices.Count; i++)
        {
            output.Add((y, slices[i]));
            y += slices[i].Height + slices[i].TrailingSpace;
        }
        return y;
    }
}

internal static class ParagraphLayouter
{
    /// <summary>Line geometry around obstacles: tracks the page y of each line while the paragraph is broken.</summary>
    private sealed class WrapSpace(Obstacles obstacles, LineSpacing spacing, FontMetrics mark, double y) : ILineSpace
    {
        public (LineFrame Frame, double Gap) Next(LineFrame frame, bool first)
        {
            double estimate = LineHeight(spacing, mark.Ascent, mark.Descent).Height;
            double gap = 0;
            for (int guard = 0; guard < 100; guard++)
            {
                var (left, right, below) = obstacles.At(y + gap, y + gap + estimate);
                if (below is { } retry)
                {
                    gap = retry - y;
                    continue;
                }
                double start = frame.LeftIndent + (first ? frame.FirstLineIndent : 0);
                if (left > start) frame = frame with { LeftIndent = left - (first ? frame.FirstLineIndent : 0) };
                if (right > frame.RightIndent) frame = frame with { RightIndent = right };
                break;
            }
            y += gap;
            return (frame, gap);
        }

        public void Done(Line line)
        {
            double ascent = line.Ascent, descent = line.Descent;
            if (ascent + descent <= 0) (ascent, descent) = (mark.Ascent, mark.Descent);
            y += LineHeight(spacing, ascent, descent).Height;
        }
    }

    public static List<Slice> Layout(Paragraph p, Paragraph? prev, Paragraph? next, double x0, double width, FlowContext ctx,
        Obstacles? obstacles = null)
    {
        var props = p.Properties;
        var atoms = ctx.Itemizer.Itemize(p, ctx.FieldText);
        // Right/center-aligned list labels end at (or center on) the first-line position.
        double labelWidth = atoms.Where(a => a.IsLabel).Sum(a => a.Width);
        double labelShift = p.ListLabel?.Alignment switch
        {
            ParagraphAlignment.Right => labelWidth,
            ParagraphAlignment.Center => labelWidth / 2,
            _ => 0,
        };
        var frame = new LineFrame(width, props.LeftIndent, props.RightIndent, props.FirstLineIndent - labelShift, props.TabStops,
            ctx.Settings.DefaultTabStop);
        bool sameStylePrev = prev is not null && prev.Properties.StyleId == props.StyleId;
        bool sameStyleNext = next is not null && next.Properties.StyleId == props.StyleId;
        double spaceBefore = props.ContextualSpacing && sameStylePrev ? 0 : props.SpaceBefore;
        double spaceAfter = props.ContextualSpacing && sameStyleNext ? 0 : props.SpaceAfter;

        var borders = props.Borders;
        bool groupedPrev = borders.HasAny && prev is not null && SameBorderGroup(prev, p);
        bool groupedNext = borders.HasAny && next is not null && SameBorderGroup(p, next);
        var topBorder = groupedPrev ? borders.InsideHorizontal : borders.Top;
        var bottomBorder = groupedNext ? Border.None : borders.Bottom;
        double topExtra = topBorder.IsVisible ? topBorder.Space + topBorder.Width : 0;
        double bottomExtra = bottomBorder.IsVisible ? bottomBorder.Space + bottomBorder.Width : 0;

        var mark = FontMetrics.Of(ctx.Fonts.ForCharacter(p.MarkProperties, ' ').Font, p.MarkProperties.FontSize);
        var space = obstacles is null ? null : new WrapSpace(obstacles, props.LineSpacing, mark, obstacles.StartY + spaceBefore + topExtra);
        var lines = LineBuilder.Build(atoms, frame, props.Alignment, ctx.Settings.JustifyLinesEndingWithBreak, space);
        double areaLeft = x0 + props.LeftIndent - (borders.Left.IsVisible ? borders.Left.Space + borders.Left.Width : 0);
        double areaRight = x0 + width - props.RightIndent + (borders.Right.IsVisible ? borders.Right.Space + borders.Right.Width : 0);

        var slices = new List<Slice>(lines.Count);
        bool breakBefore = props.PageBreakBefore, columnBreakBefore = false;
        for (int i = 0; i < lines.Count; i++)
        {
            var line = lines[i];
            bool first = i == 0, last = i == lines.Count - 1;
            double ascent = line.Ascent, descent = line.Descent;
            if (ascent + descent <= 0) (ascent, descent) = (mark.Ascent, mark.Descent);
            var (height, baseline) = LineHeight(props.LineSpacing, ascent, descent);

            double y = first ? spaceBefore : 0;
            double areaTop = y;
            y += first ? topExtra : 0;
            y += line.Gap;
            double lineTop = y;
            y += height;
            double bottomBorderY = y + bottomBorder.Space + bottomBorder.Width / 2;
            if (last) y += bottomExtra;

            var slice = new Slice
            {
                Height = y,
                LeadingSpace = first ? spaceBefore : 0,
                TrailingSpace = last ? spaceAfter : 0,
                BreakBefore = breakBefore,
                ColumnBreakBefore = columnBreakBefore,
                KeepWithPrevious = i > 0 && !breakBefore && !columnBreakBefore && (props.KeepLinesTogether
                                   || (props.WidowControl && (i == 1 || i == lines.Count - 1))),
                Footnotes = line.Atoms.Where(a => a.Note is { Kind: NoteKind.Footnote }).Select(a => a.Note!).ToList(),
                Floats = first ? p.Inlines.OfType<FloatingImage>().Where(f => !f.Properties.Hidden).ToList() : [],
            };
            breakBefore = line.EndsWithPageBreak && !line.EndsWithColumnBreak;
            columnBreakBefore = line.EndsWithColumnBreak;

            if (props.Shading is { } shading)
                slice.Items.Add(new RectItem(areaLeft, areaTop, areaRight - areaLeft, y - areaTop, shading));
            DrawLine(line, x0, lineTop + baseline, slice.Items);
            AddLinks(line, x0, lineTop, lineTop + height, slice.Items);
            if (first && topBorder.IsVisible)
                AddBorder(slice.Items, topBorder, areaLeft, areaTop + topBorder.Width / 2, areaRight, areaTop + topBorder.Width / 2);
            if (last && bottomBorder.IsVisible)
                AddBorder(slice.Items, bottomBorder, areaLeft, bottomBorderY, areaRight, bottomBorderY);
            double sideTop = first ? areaTop : 0, sideBottom = y;
            if (borders.Left.IsVisible)
                AddBorder(slice.Items, borders.Left, areaLeft + borders.Left.Width / 2, sideTop, areaLeft + borders.Left.Width / 2, sideBottom);
            if (borders.Right.IsVisible)
                AddBorder(slice.Items, borders.Right, areaRight - borders.Right.Width / 2, sideTop, areaRight - borders.Right.Width / 2, sideBottom);
            slices.Add(slice);
        }
        return slices;
    }

    private static bool SameBorderGroup(Paragraph a, Paragraph b) =>
        a.Properties.Borders.HasAny && a.Properties.Borders == b.Properties.Borders
        && a.Properties.LeftIndent == b.Properties.LeftIndent && a.Properties.RightIndent == b.Properties.RightIndent;

    /// <summary>
    /// Applies OOXML line spacing. Extra space from multiple/at-least spacing goes above the text
    /// (the baseline keeps its distance to the line bottom), as in Word.
    /// </summary>
    internal static (double Height, double Baseline) LineHeight(LineSpacing spacing, double ascent, double descent)
    {
        double natural = ascent + descent;
        return spacing.Rule switch
        {
            LineSpacingRule.Exact => (spacing.Value, spacing.Value - descent * Math.Min(1, spacing.Value / Math.Max(natural, 0.01))),
            LineSpacingRule.AtLeast => (Math.Max(spacing.Value, natural), Math.Max(spacing.Value, natural) - descent),
            _ => (natural * spacing.Value, natural * spacing.Value - descent),
        };
    }

    internal static void AddBorder(List<LayoutItem> items, Border border, double x1, double y1, double x2, double y2)
    {
        var color = border.Color ?? Color.Black;
        var style = border.Style switch { BorderStyle.Dotted => LineStyle.Dotted, BorderStyle.Dashed => LineStyle.Dashed, _ => LineStyle.Solid };
        double width = border.Style == BorderStyle.Thick ? Math.Max(border.Width, 1.5) : border.Width;
        if (border.Style == BorderStyle.Double)
        {
            double w = Math.Max(0.25, width / 3);
            bool horizontal = Math.Abs(y2 - y1) < Math.Abs(x2 - x1);
            double d = w;
            items.Add(new LineItem(x1 - (horizontal ? 0 : d), y1 - (horizontal ? d : 0), x2 - (horizontal ? 0 : d), y2 - (horizontal ? d : 0), w, color));
            items.Add(new LineItem(x1 + (horizontal ? 0 : d), y1 + (horizontal ? d : 0), x2 + (horizontal ? 0 : d), y2 + (horizontal ? d : 0), w, color));
            return;
        }
        items.Add(new LineItem(x1, y1, x2, y2, width, color, style));
    }

    /// <summary>Emits backgrounds, glyph runs, images, leaders and decoration lines for one line.</summary>
    internal static void DrawLine(Line line, double x0, double baseline, List<LayoutItem> items)
    {
        var atoms = line.Atoms;
        int visible = line.VisibleCount;

        for (int k = 0; k < visible; k++)
        {
            var a = atoms[k];
            if (a.Kind is AtomKind.Image or AtomKind.LineBreak or AtomKind.PageBreak || a.Font is null) continue;
            var background = a.Props.Highlight ?? a.Props.Shading;
            if (background is null) continue;
            var m = a.Metrics;
            items.Add(new RectItem(x0 + a.X, baseline - a.Rise - m.Ascent, a.Advance, m.Height, background.Value));
        }

        for (int k = 0; k < visible; k++)
        {
            var a = atoms[k];
            var color = a.Props.Color ?? Color.Black;
            switch (a.Kind)
            {
                case AtomKind.Text:
                case AtomKind.Space:
                    items.Add(new TextItem(x0 + a.X, baseline - a.Rise, a.Font!.Font, a.Size, a.Text, color,
                        a.Font.SyntheticBold, a.Font.SyntheticItalic, a.Props.CharacterSpacing) { Field = a.Field });
                    break;
                case AtomKind.Image:
                    items.Add(new ImageItem(x0 + a.X, baseline - a.Image!.Height, a.Image.Width, a.Image.Height, a.Image.Image));
                    break;
                case AtomKind.Tab when a.Tab is { Leader: not TabLeader.None } stop && a.Advance > 4:
                    var ch = stop.Leader switch { TabLeader.Dot => ".", TabLeader.Hyphen => "-", TabLeader.Underscore => "_", _ => "·" };
                    double charWidth = a.Font!.Font.MeasureText(ch, a.Size);
                    if (charWidth <= 0) break;
                    int count = (int)((a.Advance - charWidth) / charWidth);
                    if (count <= 0) break;
                    double end = x0 + a.X + a.Advance - charWidth * 0.5;
                    items.Add(new TextItem(end - count * charWidth, baseline, a.Font.Font, a.Size, new string(ch[0], count), color,
                        a.Font.SyntheticBold, a.Font.SyntheticItalic));
                    break;
            }
        }

        DrawDecorations(atoms, visible, x0, baseline, items);
    }

    /// <summary>Emits bookmark anchors and one link rectangle per run of atoms sharing a hyperlink.</summary>
    internal static void AddLinks(Line line, double x0, double top, double bottom, List<LayoutItem> items)
    {
        LinkItem? pending = null;
        foreach (var a in line.Atoms)
        {
            if (a.Kind == AtomKind.Marker)
            {
                if (a.Bookmark is { } name) items.Add(new AnchorItem(name, x0 + a.X, top));
                continue;
            }
            if (a.Link is null || a.Kind is AtomKind.LineBreak or AtomKind.PageBreak) continue;
            double x = x0 + a.X;
            if (pending is not null && pending.Link == a.Link && Math.Abs(pending.X + pending.Width - x) < 0.5)
                pending = pending with { Width = x + a.Advance - pending.X };
            else
            {
                if (pending is not null) items.Add(pending);
                pending = new LinkItem(x, top, a.Advance, bottom - top, a.Link);
            }
        }
        if (pending is not null) items.Add(pending);
    }

    private static void DrawDecorations(List<Atom> atoms, int visible, double x0, double baseline, List<LayoutItem> items)
    {
        LineItem? pending = null;
        void Flush()
        {
            if (pending is not null) items.Add(pending);
            pending = null;
        }

        for (int k = 0; k < visible; k++)
        {
            var a = atoms[k];
            if (a.Font is null || a.Kind is AtomKind.Image or AtomKind.LineBreak or AtomKind.PageBreak) { Flush(); continue; }
            var p = a.Props;
            var font = a.Font.Font;
            double scale = a.Size / font.UnitsPerEm;
            var color = p.Color ?? Color.Black;
            double x1 = x0 + a.X, x2 = x1 + a.Advance;
            double y = baseline - a.Rise;

            if (p.Underline != UnderlineStyle.None && !(p.Underline == UnderlineStyle.Words && a.Kind != AtomKind.Text))
            {
                double thickness = Math.Max(0.4, font.UnderlineThickness * scale * (p.Underline == UnderlineStyle.Thick ? 2 : 1));
                double uy = y - font.UnderlinePosition * scale + thickness / 2;
                var style = p.Underline switch { UnderlineStyle.Dotted => LineStyle.Dotted, UnderlineStyle.Dashed => LineStyle.Dashed, _ => LineStyle.Solid };
                if (pending is not null && Math.Abs(pending.Y1 - uy) < 0.01 && Math.Abs(pending.X2 - x1) < 0.01
                    && pending.Color == color && pending.Style == style && Math.Abs(pending.Width - thickness) < 0.01)
                    pending = pending with { X2 = x2 };
                else
                {
                    Flush();
                    pending = new LineItem(x1, uy, x2, uy, thickness, color, style);
                }
                if (p.Underline == UnderlineStyle.Double)
                {
                    Flush();
                    items.Add(new LineItem(x1, uy + thickness * 2, x2, uy + thickness * 2, thickness, color));
                }
            }
            else Flush();

            if (p.Strike || p.DoubleStrike)
            {
                double thickness = Math.Max(0.4, font.StrikeoutSize * scale);
                double sy = y - font.StrikeoutPosition * scale;
                if (p.DoubleStrike)
                {
                    items.Add(new LineItem(x1, sy - thickness, x2, sy - thickness, thickness, color));
                    items.Add(new LineItem(x1, sy + thickness, x2, sy + thickness, thickness, color));
                }
                else items.Add(new LineItem(x1, sy, x2, sy, thickness, color));
            }
        }
        Flush();
    }
}