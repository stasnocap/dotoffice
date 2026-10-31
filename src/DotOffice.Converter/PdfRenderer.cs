using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Layout;
using DotOffice.Pdf;

namespace DotOffice.Converter;

/// <summary>Paints laid-out pages onto PDF canvases, converting top-down layout space to PDF space once.</summary>
internal sealed class PdfRenderer(bool embedFonts, IConversionWarningSink? warnings)
{
    private readonly Dictionary<TrueTypeFont, PdfFont> _fonts = new(ReferenceEqualityComparer.Instance);
    private readonly Dictionary<ImageData, PdfImage?> _images = new(ReferenceEqualityComparer.Instance);
    private PdfDocument _pdf = new();

    public void Render(LayoutDocument layout, Stream output)
    {
        _pdf = new PdfDocument { Title = layout.Title, Author = layout.Author };
        var pages = layout.Pages.Select(page => _pdf.AddPage(new PageSize(page.Width, page.Height))).ToList();
        for (int i = 0; i < pages.Count; i++) RenderPage(layout.Pages[i], pages[i]);
        AddLinks(layout, pages);
        if (layout.Pages.Count == 0) _pdf.AddPage(PageSize.Letter);
        _pdf.Save(output);
    }

    /// <summary>Turns link items into annotations; internal links resolve to the first anchor of their bookmark.</summary>
    private void AddLinks(LayoutDocument layout, List<PdfPage> pages)
    {
        var anchors = new Dictionary<string, (PdfPage Page, double Y)>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < pages.Count; i++)
            foreach (var anchor in layout.Pages[i].Items.OfType<AnchorItem>())
                anchors.TryAdd(anchor.Name, (pages[i], layout.Pages[i].Height - anchor.Y));

        for (int i = 0; i < pages.Count; i++)
        {
            double h = layout.Pages[i].Height;
            foreach (var link in layout.Pages[i].Items.OfType<LinkItem>())
            {
                double y = h - link.Y - link.Height;
                if (link.Link.Bookmark is { } name && anchors.TryGetValue(name, out var target))
                    pages[i].AddLink(link.X, y, link.Width, link.Height, target.Page, target.Y);
                else if (link.Link.Uri is { } uri && IsSafeUri(uri))
                    pages[i].AddLink(link.X, y, link.Width, link.Height, uri);
                else if (link.Link.Bookmark is { } missing)
                    warnings.Warn("pdf.link-target", $"Link target bookmark '{missing}' was not found.");
            }
        }
    }

    /// <summary>Only absolute web/mail links become URI actions (no file:, javascript: or relative paths).</summary>
    private static bool IsSafeUri(string uri) =>
        Uri.TryCreate(uri, UriKind.Absolute, out var parsed) && parsed.Scheme is "http" or "https" or "mailto" or "ftp";

    private void RenderPage(LayoutPage page, PdfPage pdfPage)
    {
        var canvas = pdfPage.Canvas;
        double h = page.Height;
        var items = page.Items;
        Color? fill = null, stroke = null;

        void Fill(Color c)
        {
            if (fill == c) return;
            canvas.SetFillColor(c.R / 255.0, c.G / 255.0, c.B / 255.0);
            fill = c;
        }

        void Stroke(Color c)
        {
            if (stroke == c) return;
            canvas.SetStrokeColor(c.R / 255.0, c.G / 255.0, c.B / 255.0);
            stroke = c;
        }

        for (int i = 0; i < items.Count; i++)
        {
            switch (items[i])
            {
                case TextItem text:
                    // Coalesce following runs that share font, size, style and baseline into one text object.
                    var segments = new List<TextSegment> { new(text.X, text.Text) };
                    while (i + 1 < items.Count && items[i + 1] is TextItem next && SameRun(text, next))
                    {
                        segments.Add(new TextSegment(next.X, next.Text));
                        i++;
                    }
                    Fill(text.Color);
                    if (text.SyntheticBold) Stroke(text.Color);
                    canvas.DrawTextSegments(Font(text.Font), text.FontSize, h - text.Baseline, segments,
                        new TextStyle(text.SyntheticBold, text.SyntheticItalic, text.CharacterSpacing));
                    break;
                case RectItem rect:
                    Fill(rect.Fill);
                    canvas.Rectangle(rect.X, h - rect.Y - rect.Height, rect.Width, rect.Height).Fill();
                    break;
                case LineItem line:
                    Stroke(line.Color);
                    canvas.SetLineWidth(line.Width);
                    if (line.Style != LineStyle.Solid)
                        canvas.SetDash(line.Style == LineStyle.Dotted ? [line.Width, line.Width * 2] : [line.Width * 4, line.Width * 2]);
                    canvas.MoveTo(line.X1, h - line.Y1).LineTo(line.X2, h - line.Y2).Stroke();
                    if (line.Style != LineStyle.Solid) canvas.SetDash();
                    break;
                case ImageItem image:
                    if (Image(image.Image) is { } pdfImage)
                        canvas.DrawImage(pdfImage, image.X, h - image.Y - image.Height, image.Width, image.Height);
                    else
                    {
                        // Unsupported format: keep the layout box visible.
                        canvas.SaveState().SetFillColor(0.92, 0.92, 0.92).SetStrokeColor(0.6, 0.6, 0.6).SetLineWidth(0.5)
                              .Rectangle(image.X, h - image.Y - image.Height, image.Width, image.Height).FillAndStroke().RestoreState();
                    }
                    break;
            }
        }
    }

    private static bool SameRun(TextItem a, TextItem b) =>
        ReferenceEquals(a.Font, b.Font) && a.FontSize == b.FontSize && a.Color == b.Color && Math.Abs(a.Baseline - b.Baseline) < 0.001
        && a.SyntheticBold == b.SyntheticBold && a.SyntheticItalic == b.SyntheticItalic && a.CharacterSpacing == b.CharacterSpacing
        && b.X >= a.X;

    private PdfFont Font(TrueTypeFont font)
    {
        if (_fonts.TryGetValue(font, out var pdfFont)) return pdfFont;
        bool embed = embedFonts;
        if (embed && font.IsEmbeddingRestricted)
        {
            warnings.Warn("pdf.font-not-embeddable", $"Font '{font.PostScriptName}' forbids embedding; it is referenced by name only.");
            embed = false;
        }
        return _fonts[font] = _pdf.AddFont(font, embed);
    }

    private PdfImage? Image(ImageData data)
    {
        if (_images.TryGetValue(data, out var cached)) return cached;
        PdfImage? image = null;
        try
        {
            image = data.Format switch
            {
                ImageFormat.Jpeg => PdfImage.FromJpeg(data.Data),
                ImageFormat.Png => PdfImage.FromPng(data.Data),
                _ => null,
            };
            if (image is null)
                warnings.Warn("pdf.image-format", $"Image format {data.Format} ('{data.Name}') is not supported yet; a placeholder was drawn.");
        }
        catch (Exception e) when (e is InvalidDataException or NotSupportedException)
        {
            warnings.Warn("pdf.image-invalid", $"Image '{data.Name}' could not be decoded: {e.Message}");
        }
        if (image is not null) _pdf.AddImage(image);
        return _images[data] = image;
    }
}