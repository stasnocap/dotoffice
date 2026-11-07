using DotOffice.Core;

namespace DotOffice.Layout;

public sealed partial class SheetLayoutEngine
{
    /// <summary>Lays out each slide of the presentation as one page of its size; an empty presentation yields one blank page.</summary>
    public LayoutDocument Layout(Presentation presentation, CancellationToken cancellationToken = default)
    {
        var document = new LayoutDocument { Title = presentation.Title, Author = presentation.Author };
        _fontLineHeight = true;
        try
        {
            LayoutSlides(presentation, document, cancellationToken);
        }
        finally
        {
            _fontLineHeight = false;
        }
        if (document.Pages.Count == 0) document.Pages.Add(new LayoutPage(720, 540));
        Metafiles.Expand(document, _fonts, _warnings);
        return document;
    }

    /// <summary>
    /// Whether shape text is single-spaced by the font's own line height (ascent plus descent), as PowerPoint and LibreOffice space slide
    /// text; spreadsheet shapes keep the row-height pitch they were tuned to.
    /// </summary>
    private bool _fontLineHeight;

    private void LayoutSlides(Presentation presentation, LayoutDocument document, CancellationToken cancellationToken)
    {
        foreach (var slide in presentation.Slides)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var page = new LayoutPage(slide.Width, slide.Height);
            var items = page.Items;
            if (slide.BackgroundImage is { } picture) items.Add(new ImageItem(0, 0, slide.Width, slide.Height, picture));
            else if (slide.BackgroundGradient is { } gradient) SheetPass.EmitGradient(gradient, 0, 0, slide.Width, slide.Height, items);
            else if (slide.Background is { } color) items.Add(new RectItem(0, 0, slide.Width, slide.Height, color));

            var whole = (0.0, 0.0, slide.Width, slide.Height);
            var placed = slide.Objects.OrderBy(o => o.ZOrder).Select(o => (Object: o, Rect: PlaceInFrame(o, o.X, o.Y, o.Width, o.Height))).ToList();
            FitText(placed);
            foreach (var (obj, (x, y, w, h)) in placed)
            {
                if (w <= 0 && h <= 0) continue;
                switch (obj)
                {
                    case SheetShape shape: EmitShape(shape, x, y, w, h, whole, items); break;
                    case SheetChart chart: EmitChart(chart, x, y, w, h, whole, items); break;
                    case SheetImage image: SheetPass.EmitImage(image, x, y, w, h, whole, items); break;
                }
            }
            document.Pages.Add(page);
        }
    }

    /// <summary>
    /// Grows frames to fit their text: shapes that fit their size to the text (<see cref="SheetShape.GrowToFit"/>) get taller, keeping their
    /// top; table rows get taller when a cell's text needs more than the row's height, moving the rows below down. Frames never shrink.
    /// </summary>
    private void FitText(List<(SheetObject Object, (double X, double Y, double Width, double Height) Rect)> placed)
    {
        for (int i = 0; i < placed.Count; i++)
            if (placed[i].Object is SheetShape { GrowToFit: true, TableRows: null, TextRotation: 0 } shape && shape.Paragraphs.Count > 0 && shape.Geometry != ShapeGeometry.Line)
            {
                var rect = placed[i].Rect;
                if (ShapeTextHeight(shape, rect.Width) is var need && need > rect.Height) placed[i] = (shape, rect with { Height = need });
            }

        foreach (var table in placed.Select((p, i) => (p, i)).Where(t => t.p.Object is SheetShape { TableRows: not null }).GroupBy(t => ((SheetShape)t.p.Object).TableRows!.Table))
        {
            var parts = table.ToList();
            var heights = ((SheetShape)parts[0].p.Object).TableRows!.RowHeights;
            double sum = heights.Sum();
            if (sum <= 0) continue;
            // The table's height on the page, from a cell (the rows keep their proportions there).
            double scale = 0;
            foreach (var (p, _) in parts)
            {
                var rows = ((SheetShape)p.Object).TableRows!;
                double span = heights.Skip(rows.First).Take(rows.Count).Sum();
                if (rows.Count > 0 && span > 0 && p.Rect.Height > 0) { scale = p.Rect.Height / span; break; }
            }
            if (scale <= 0) continue;
            var grow = new double[heights.Count];
            // Cells of one row first, then merged cells, which only grow their last row by what their rows still lack.
            // Cells of vertical text wrap at the row's height and leave it as it is.
            foreach (var (p, _) in parts.Where(t => t.p.Object is SheetShape { TableRows.Count: > 0, TextRotation: 0, Paragraphs.Count: > 0 })
                         .OrderBy(t => ((SheetShape)t.p.Object).TableRows!.Count))
            {
                var shape = (SheetShape)p.Object;
                var rows = shape.TableRows!;
                double have = 0;
                for (int r = rows.First; r < rows.First + rows.Count; r++) have += heights[r] * scale + grow[r];
                double lack = ShapeTextHeight(shape, p.Rect.Width) - have;
                if (lack > 0.01) grow[rows.First + rows.Count - 1] += lack;
            }
            var shift = new double[heights.Count + 1];
            for (int r = 0; r < heights.Count; r++) shift[r + 1] = shift[r] + grow[r];
            if (shift[^1] <= 0) continue;
            foreach (var (p, i) in parts)
            {
                var rows = ((SheetShape)p.Object).TableRows!;
                int first = Math.Clamp(rows.First, 0, heights.Count), end = Math.Clamp(rows.First + rows.Count, 0, heights.Count);
                placed[i] = (p.Object, p.Rect with { Y = p.Rect.Y + shift[first], Height = p.Rect.Height + shift[end] - shift[first] });
            }
        }
    }
}
