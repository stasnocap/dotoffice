namespace DotOffice.Converter;

public enum DocumentFormat { Docx, Xlsx, Pptx, Pdf }

public static class DocumentConverter
{
    public static async Task ConvertAsync(string inputPath, string outputPath, CancellationToken ct = default)
    {
        await using var input = File.OpenRead(inputPath);
        await using var output = File.Create(outputPath);
        await ConvertAsync(input, output, FormatFromPath(inputPath), FormatFromPath(outputPath), ct);
    }

    public static Task ConvertAsync(Stream input, Stream output, DocumentFormat from, DocumentFormat to, CancellationToken ct = default) =>
        throw new NotSupportedException($"Conversion {from} -> {to} is not implemented yet.");

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
