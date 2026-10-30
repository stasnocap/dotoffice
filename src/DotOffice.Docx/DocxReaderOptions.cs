using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>Reader options. Defaults are safe for untrusted input.</summary>
public sealed class DocxReaderOptions
{
    /// <summary>Receives warnings about unsupported or malformed content.</summary>
    public IConversionWarningSink? Warnings { get; init; }

    /// <summary>Maximum total bytes decompressed from the package (zip bomb guard).</summary>
    public long MaxTotalUncompressedSize { get; init; } = 512L * 1024 * 1024;

    /// <summary>Maximum decompressed size of any single part.</summary>
    public long MaxPartSize { get; init; } = 128L * 1024 * 1024;

    public int MaxPartCount { get; init; } = 10_000;

    /// <summary>Maximum XML element nesting depth; deeper content is skipped with a warning.</summary>
    public int MaxXmlDepth { get; init; } = 256;

    /// <summary>Maximum nesting of tables inside table cells.</summary>
    public int MaxTableNesting { get; init; } = 16;

    /// <summary>Upper bound for any page or image dimension in points (larger values are clamped).</summary>
    public double MaxDimension { get; init; } = 14_400; // 200 inches
}