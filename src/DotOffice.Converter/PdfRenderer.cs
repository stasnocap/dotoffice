using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;
using DotOffice.Layout;
using DotOffice.Pdf;

namespace DotOffice.Converter;

/// <summary>Paints laid-out pages onto PDF canvases, converting top-down layout space to PDF space once.</summary>
internal sealed class PdfRenderer(bool embedFonts, IConversionWarningSink? warnings)
{
    /// <summary>
    /// Adds a polygon to the current path; points that are NaN separate its rings (a shape with holes). True when there are several rings,
    /// which are to be filled or clipped by the even-odd rule.
    /// </summary>
    private static bool PolygonPath(PdfCanvas canvas, IReadOnlyList<(double X, double Y)> points, double h)
    {
        bool rings = false, start = true;
        foreach (var (x, y) in points)
        {
            if (double.IsNaN(x)) { canvas.ClosePath(); start = true; rings = true; continue; }
            if (start) canvas.MoveTo(x, h - y); else canvas.LineTo(x, h - y);
            start = false;
        }
        canvas.ClosePath();
        return rings;
    }

    private static void ClipTo(PdfCanvas canvas, IReadOnlyList<(double X, double Y)> points, double h)
    {
        if (PolygonPath(canvas, points, h)) canvas.ClipEvenOdd(); else canvas.Clip();
    }

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
        var saved = new Stack<(Color? Fill, Color? Stroke)>();

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
                case TextItem { Glyphs: { } shaped } text:
                    Fill(text.Color);
                    if (text.SyntheticBold) Stroke(text.Color);
                    canvas.DrawGlyphs(Font(text.Font), text.FontSize, h - text.Baseline, Positions(text, shaped),
                        new TextStyle(text.SyntheticBold, text.SyntheticItalic));
                    break;
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
                case PushItem push:
                    // q/Q also save and restore the colors, so the cached ones must follow.
                    saved.Push((fill, stroke));
                    canvas.SaveState();
                    if (push.Opacity < 1) canvas.SetAlpha(push.Opacity);
                    if (push.Clip is { } clip) canvas.Rectangle(clip.X, h - clip.Y - clip.Height, clip.Width, clip.Height).Clip();
                    if (push.ClipOutline is { Count: >= 3 } outline) ClipTo(canvas, outline, h);
                    if (push.FlipH || push.FlipV)
                    {
                        // Mirror about the center first: the matrix is applied to the content before the rotation below.
                        double fcy = h - push.CenterY, sx = push.FlipH ? -1 : 1, sy = push.FlipV ? -1 : 1;
                        if (push.Rotation != 0)
                        {
                            double rad = push.Rotation * Math.PI / 180, c = Math.Cos(rad), s = Math.Sin(rad);
                            canvas.Transform(c, s, -s, c, push.CenterX - push.CenterX * c + fcy * s, fcy - push.CenterX * s - fcy * c);
                        }
                        canvas.Transform(sx, 0, 0, sy, push.FlipH ? 2 * push.CenterX : 0, push.FlipV ? 2 * fcy : 0);
                    }
                    else if (push.Rotation != 0)
                    {
                        double radians = push.Rotation * Math.PI / 180, cos = Math.Cos(radians), sin = Math.Sin(radians), cy = h - push.CenterY;
                        canvas.Transform(cos, sin, -sin, cos, push.CenterX - push.CenterX * cos + cy * sin, cy - push.CenterX * sin - cy * cos);
                    }
                    break;
                case PopItem when saved.Count > 0:
                    canvas.RestoreState();
                    (fill, stroke) = saved.Pop();
                    break;
                case GradientItem gradient:
                {
                    // The gradient line runs through the center, as long as the projection of the box on it (so the corners get the end colors).
                    double radians = gradient.Degree * Math.PI / 180, dx = Math.Cos(radians), dy = Math.Sin(radians);
                    double half = (Math.Abs(gradient.Width * dx) + Math.Abs(gradient.Height * dy)) / 2;
                    double cx = gradient.X + gradient.Width / 2, cy = gradient.Y + gradient.Height / 2;
                    var stops = gradient.Stops.Select(s => new ShadingStop(s.Position, s.Color.R / 255.0, s.Color.G / 255.0, s.Color.B / 255.0)).ToList();
                    PdfShading? shading = null;
                    if (gradient.Radial && gradient.Rectangular)
                    {
                        // Concentric rectangles from the focus out to the box's edges: stop 0 at the focus, the last stop at the edges.
                        double fx = gradient.X + gradient.Width * gradient.FocusX, fy = gradient.Y + gradient.Height * gradient.FocusY;
                        double toLeft = fx - gradient.X, toRight = gradient.X + gradient.Width - fx, toTop = fy - gradient.Y, toBottom = gradient.Y + gradient.Height - fy;
                        var fillStops = gradient.Stops;
                        Color At(double t)
                        {
                            if (t <= fillStops[0].Position) return fillStops[0].Color;
                            for (int k = 1; k < fillStops.Count; k++)
                                if (t <= fillStops[k].Position)
                                {
                                    double span = Math.Max(fillStops[k].Position - fillStops[k - 1].Position, 1e-9), u = (t - fillStops[k - 1].Position) / span;
                                    return fillStops[k - 1].Color.Blend(fillStops[k].Color, u);
                                }
                            return fillStops[^1].Color;
                        }
                        canvas.SaveState();
                        if (gradient.Clip is { Count: >= 3 } rectClip) ClipTo(canvas, rectClip, h);
                        else canvas.Rectangle(gradient.X, h - gradient.Y - gradient.Height, gradient.Width, gradient.Height).Clip();
                        const int bands = 48;
                        for (int band = bands; band >= 1; band--)
                        {
                            double t = (double)band / bands, color = (band - 0.5) / bands;
                            var c = At(color);
                            canvas.SetFillColor(c.R / 255.0, c.G / 255.0, c.B / 255.0)
                                  .Rectangle(fx - toLeft * t, h - (fy + toBottom * t), (toLeft + toRight) * t, (toTop + toBottom) * t).Fill();
                        }
                        canvas.RestoreState();
                        break;
                    }
                    if (gradient.Radial)
                    {
                        // Blends outwards from the focus to the farthest corner.
                        double fx = gradient.X + gradient.Width * gradient.FocusX, fy = gradient.Y + gradient.Height * gradient.FocusY;
                        double reachX = Math.Max(fx - gradient.X, gradient.X + gradient.Width - fx), reachY = Math.Max(fy - gradient.Y, gradient.Y + gradient.Height - fy);
                        shading = _pdf.AddShading(PdfShading.Radial(fx, h - fy, 0, fx, h - fy, Math.Max(Math.Sqrt(reachX * reachX + reachY * reachY), 0.01), stops));
                    }
                    else shading = _pdf.AddShading(PdfShading.Axial(cx - dx * half, h - (cy - dy * half), cx + dx * half, h - (cy + dy * half), stops));
                    canvas.SaveState();
                    if (gradient.Clip is { Count: >= 3 } clipPolygon) ClipTo(canvas, clipPolygon, h);
                    else canvas.Rectangle(gradient.X, h - gradient.Y - gradient.Height, gradient.Width, gradient.Height).Clip();
                    canvas.PaintShading(shading!).RestoreState();
                    break;
                }
                case ShadedTriangleItem { Corners.Count: >= 3 } shaded:
                    canvas.PaintShading(_pdf.AddShading(PdfShading.Triangles(shaded.Corners.Take(shaded.Corners.Count / 3 * 3).Select(c => (c.X, h - c.Y, c.Color.R / 255.0, c.Color.G / 255.0, c.Color.B / 255.0)).ToList())));
                    break;
                case RectItem rect:
                    Fill(rect.Fill);
                    canvas.Rectangle(rect.X, h - rect.Y - rect.Height, rect.Width, rect.Height).Fill();
                    break;
                case PolygonItem { Points.Count: >= 3 } polygon:
                    Fill(polygon.Fill);
                    if (PolygonPath(canvas, polygon.Points, h)) canvas.FillEvenOdd(); else canvas.Fill();
                    break;
                case EllipseItem ellipse:
                {
                    // Four Bézier arcs approximate the ellipse (kappa = 4/3 (√2 − 1)).
                    const double kappa = 0.5522847498;
                    double rx = ellipse.Width / 2, ry = ellipse.Height / 2, cx = ellipse.X + rx, cy = h - ellipse.Y - ry;
                    Fill(ellipse.Fill);
                    canvas.MoveTo(cx + rx, cy)
                          .CurveTo(cx + rx, cy + kappa * ry, cx + kappa * rx, cy + ry, cx, cy + ry)
                          .CurveTo(cx - kappa * rx, cy + ry, cx - rx, cy + kappa * ry, cx - rx, cy)
                          .CurveTo(cx - rx, cy - kappa * ry, cx - kappa * rx, cy - ry, cx, cy - ry)
                          .CurveTo(cx + kappa * rx, cy - ry, cx + rx, cy - kappa * ry, cx + rx, cy)
                          .ClosePath().Fill();
                    break;
                }
                case PathItem { Points.Count: >= 2 } path when path.Fill is not null || path.Stroke is not null:
                {
                    var points = path.Points;
                    for (int k = 0; k < points.Count; k++)
                    {
                        var p = points[k];
                        switch (p.Kind)
                        {
                            case PathPointKind.Move: canvas.MoveTo(p.X, h - p.Y); break;
                            case PathPointKind.Line: canvas.LineTo(p.X, h - p.Y); break;
                            case PathPointKind.Close: canvas.ClosePath(); break;
                            case PathPointKind.Curve when k + 2 < points.Count:
                                canvas.CurveTo(p.X, h - p.Y, points[k + 1].X, h - points[k + 1].Y, points[k + 2].X, h - points[k + 2].Y);
                                k += 2;
                                break;
                        }
                    }
                    if (path.Fill is { } pathFill) Fill(pathFill);
                    if (path.Stroke is { } pathStroke)
                    {
                        Stroke(pathStroke);
                        canvas.SetLineWidth(path.LineWidth);
                        if (path.Style != LineStyle.Solid)
                            canvas.SetDash(path.Style == LineStyle.Dotted ? [path.LineWidth, path.LineWidth * 2] : [path.LineWidth * 4, path.LineWidth * 2]);
                    }
                    if (path.Fill is null) canvas.Stroke();
                    else if (path.Stroke is null) { if (path.EvenOdd) canvas.FillEvenOdd(); else canvas.Fill(); }
                    else if (path.EvenOdd) canvas.FillAndStrokeEvenOdd();
                    else canvas.FillAndStroke();
                    if (path.Stroke is not null && path.Style != LineStyle.Solid) canvas.SetDash();
                    break;
                }
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
                    {
                        if (image.Clip is { Count: >= 3 } imageClip)
                        {
                            canvas.SaveState();
                            ClipTo(canvas, imageClip, h);
                        }
                        canvas.DrawImage(pdfImage, image.X, h - image.Y - image.Height, image.Width, image.Height);
                        if (image.Clip is { Count: >= 3 }) canvas.RestoreState();
                    }
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

    /// <summary>Absolute glyph positions in visual order: right-to-left runs are drawn last glyph first.</summary>
    private static List<PositionedGlyph> Positions(TextItem text, IReadOnlyList<ShapedGlyph> glyphs)
    {
        double scale = text.FontSize / text.Font.UnitsPerEm;
        var result = new List<PositionedGlyph>(glyphs.Count);
        double x = text.X;
        for (int k = 0; k < glyphs.Count; k++)
        {
            var g = glyphs[text.RightToLeft ? glyphs.Count - 1 - k : k];
            result.Add(new PositionedGlyph(g.GlyphId, g.Text, x + g.XOffset * scale, g.YOffset * scale));
            x += g.Advance * scale + (g.Advance != 0 ? text.CharacterSpacing : 0);
        }
        return result;
    }

    private static bool SameRun(TextItem a, TextItem b) =>
        ReferenceEquals(a.Font, b.Font) && a.FontSize == b.FontSize && a.Color == b.Color && Math.Abs(a.Baseline - b.Baseline) < 0.001
        && a.SyntheticBold == b.SyntheticBold && a.SyntheticItalic == b.SyntheticItalic && a.CharacterSpacing == b.CharacterSpacing
        && b.X >= a.X && b.Glyphs is null;

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
                ImageFormat.Bmp => PdfImage.FromBmp(data.Data),
                ImageFormat.Gif => PdfImage.FromGif(data.Data),
                ImageFormat.Tiff => PdfImage.FromTiff(data.Data),
                _ => null,
            };
            if (image is null)
                warnings.Warn("pdf.image-format", $"Image format {data.Format} ('{data.Name}') is not supported; a placeholder was drawn.");
        }
        catch (Exception e) when (e is InvalidDataException or NotSupportedException)
        {
            warnings.Warn("pdf.image-invalid", $"Image '{data.Name}' could not be decoded: {e.Message}");
        }
        if (image is not null) _pdf.AddImage(image);
        return _images[data] = image;
    }
}