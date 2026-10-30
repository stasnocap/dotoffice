using DotOffice.Core;
using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Layout;

namespace DotOffice.Converter;

public enum DocumentFormat { Docx, Xlsx, Pptx, Pdf }

public sealed class ConvertOptions
{
    /// <summary>Fonts consulted before system fonts (e.g. a <see cref="FontCollection"/> of bundled fonts).</summary>
    public IFontProvider? Fonts { get; init; }

    /// <summary>Also look up fonts installed on the machine. Disable for fully reproducible output.</summary>
    public bool UseSystemFonts { get; init; } = true;

    /// <summary>Receives warnings about unsupported content and font substitutions.</summary>
    public IConversionWarningSink? Warnings { get; init; }

    /// <summary>Embed (subsetted) fonts in the PDF. Without embedding, viewers must have the fonts installed.</summary>
    public bool EmbedFonts { get; init; } = true;

    /// <summary>Upper bound on bytes decompressed from the input package (zip bomb guard).</summary>
    public long MaxUncompressedSize { get; init; } = 512L * 1024 * 1024;
}

public static class DocumentConverter
{
    public static async Task ConvertAsync(string inputPath, string outputPath, ConvertOptions? options = null, CancellationToken ct = default)
    {
        var from = FormatFromPath(inputPath);
        var to = FormatFromPath(outputPath);
        EnsureSupported(from, to);
        await using var input = File.OpenRead(inputPath);
        using var buffer = new MemoryStream();
        await ConvertAsync(input, buffer, from, to, options, ct).ConfigureAwait(false);
        // Only create the output once conversion succeeded.
        await using var output = File.Create(outputPath);
        buffer.Position = 0;
        await buffer.CopyToAsync(output, ct).ConfigureAwait(false);
    }

    public static async Task ConvertAsync(Stream input, Stream output, DocumentFormat from, DocumentFormat to,
        ConvertOptions? options = null, CancellationToken ct = default)
    {
        EnsureSupported(from, to);
        options ??= new ConvertOptions();

        Stream source = input;
        if (!input.CanSeek)
        {
            var copy = new MemoryStream();
            await input.CopyToAsync(copy, ct).ConfigureAwait(false);
            copy.Position = 0;
            source = copy;
        }

        ct.ThrowIfCancellationRequested();
        var document = DocxReader.Read(source, new DocxReaderOptions
        {
            Warnings = options.Warnings,
            MaxTotalUncompressedSize = options.MaxUncompressedSize,
        });

        ct.ThrowIfCancellationRequested();
        var resolver = (options.Fonts, options.UseSystemFonts) switch
        {
            (null, true) => FontResolver.CreateDefault(),
            (not null, true) => FontResolver.CreateDefault(options.Fonts),
            (not null, false) => new FontResolver(options.Fonts),
            _ => throw new ArgumentException("UseSystemFonts is false but no Fonts provider was given.", nameof(options)),
        };
        var layout = new LayoutEngine(new LayoutOptions { Fonts = resolver, Warnings = options.Warnings }).Layout(document, ct);

        ct.ThrowIfCancellationRequested();
        using var pdf = new MemoryStream();
        new PdfRenderer(options.EmbedFonts, options.Warnings).Render(layout, pdf);
        pdf.Position = 0;
        await pdf.CopyToAsync(output, ct).ConfigureAwait(false);
    }

    private static void EnsureSupported(DocumentFormat from, DocumentFormat to)
    {
        if (from != DocumentFormat.Docx || to != DocumentFormat.Pdf)
            throw new NotSupportedException($"Conversion {from} -> {to} is not implemented yet.");
    }

    internal static DocumentFormat FormatFromPath(string path) =>
        Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".docx" => DocumentFormat.Docx,
            ".xlsx" => DocumentFormat.Xlsx,
            ".pptx" => DocumentFormat.Pptx,
            ".pdf" => DocumentFormat.Pdf,
            var ext => throw new NotSupportedException($"Unknown format '{ext}'."),
        };
}