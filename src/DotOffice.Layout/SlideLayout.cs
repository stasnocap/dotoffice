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
            foreach (var obj in slide.Objects.OrderBy(o => o.ZOrder))
            {
                var (x, y, w, h) = PlaceInFrame(obj, obj.X, obj.Y, obj.Width, obj.Height);
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
}
