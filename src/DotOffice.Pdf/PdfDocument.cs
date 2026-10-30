using System.Text;
using DotOffice.Fonts;

namespace DotOffice.Pdf;

public readonly record struct PageSize(double Width, double Height)
{
    public static PageSize A4 { get; } = new(595.276, 841.89);
    public static PageSize Letter { get; } = new(612, 792);
}

public sealed class PdfPage
{
    internal PdfPage(PageSize size) => Size = size;

    public PageSize Size { get; }
    public PdfCanvas Canvas { get; } = new();
}

public sealed class PdfDocument
{
    private readonly List<PdfPage> _pages = [];
    private readonly List<PdfFont> _fonts = [];
    private readonly List<PdfImage> _images = [];

    public string? Title { get; set; }
    public string? Author { get; set; }
    public string Producer { get; set; } = "DotOffice";

    public IReadOnlyList<PdfPage> Pages => _pages;

    public PdfPage AddPage(PageSize size)
    {
        var page = new PdfPage(size);
        _pages.Add(page);
        return page;
    }

    public PdfFont AddFont(TrueTypeFont font)
    {
        if (font.IsEmbeddingRestricted)
            throw new InvalidOperationException($"Font '{font.PostScriptName}' does not permit embedding.");
        var pdfFont = new PdfFont(font, $"F{_fonts.Count + 1}");
        _fonts.Add(pdfFont);
        return pdfFont;
    }

    public PdfImage AddImage(PdfImage image)
    {
        image.ResourceName = $"Im{_images.Count + 1}";
        _images.Add(image);
        return image;
    }

    public void Save(Stream output)
    {
        if (_pages.Count == 0) throw new InvalidOperationException("A PDF document needs at least one page.");

        var writer = new PdfWriter(output);
        int catalogId = writer.Allocate();
        int pagesId = writer.Allocate();
        int infoId = writer.Allocate();
        var fontIds = _fonts.ToDictionary(f => f, _ => writer.Allocate());
        var imageIds = _images.ToDictionary(i => i, _ => writer.Allocate());
        var pageIds = _pages.Select(_ => (Page: writer.Allocate(), Content: writer.Allocate())).ToList();

        writer.WriteObject(catalogId, $"<< /Type /Catalog /Pages {pagesId} 0 R >>");
        writer.WriteObject(pagesId,
            $"<< /Type /Pages /Kids [{string.Join(" ", pageIds.Select(p => $"{p.Page} 0 R"))}] /Count {_pages.Count} >>");

        for (int i = 0; i < _pages.Count; i++)
        {
            var page = _pages[i];
            var resources = new StringBuilder("<<");
            if (page.Canvas.Fonts.Count > 0)
                resources.Append(" /Font <<")
                         .AppendJoin("", page.Canvas.Fonts.Select(f => $" /{f.ResourceName} {fontIds[f]} 0 R"))
                         .Append(" >>");
            if (page.Canvas.Images.Count > 0)
                resources.Append(" /XObject <<")
                         .AppendJoin("", page.Canvas.Images.Select(img => $" /{img.ResourceName} {IdOf(imageIds, img)} 0 R"))
                         .Append(" >>");
            resources.Append(" >>");

            writer.WriteObject(pageIds[i].Page,
                $"<< /Type /Page /Parent {pagesId} 0 R /MediaBox [0 0 {PdfFormat.Number(page.Size.Width)} " +
                $"{PdfFormat.Number(page.Size.Height)}] /Resources {resources} /Contents {pageIds[i].Content} 0 R >>");
            writer.WriteStream(pageIds[i].Content, "", page.Canvas.ToBytes());
        }

        foreach (var (image, id) in imageIds) image.Write(writer, id);
        foreach (var (font, id) in fontIds) font.Write(writer, id);

        var info = new StringBuilder($"<< /Producer {PdfFormat.TextString(Producer)}");
        if (Title is not null) info.Append($" /Title {PdfFormat.TextString(Title)}");
        if (Author is not null) info.Append($" /Author {PdfFormat.TextString(Author)}");
        writer.WriteObject(infoId, info.Append(" >>").ToString());

        writer.Finish(catalogId, infoId);
    }

    private static int IdOf(Dictionary<PdfImage, int> ids, PdfImage image) =>
        ids.TryGetValue(image, out var id)
            ? id
            : throw new InvalidOperationException("Image was drawn without being added via PdfDocument.AddImage.");
}
