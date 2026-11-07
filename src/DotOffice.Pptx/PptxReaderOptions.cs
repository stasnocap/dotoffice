using DotOffice.Core;

namespace DotOffice.Pptx;

/// <summary>Options and resource limits for <see cref="PptxReader"/>.</summary>
public sealed class PptxReaderOptions
{
    /// <summary>Receives warnings about unsupported or malformed content.</summary>
    public IConversionWarningSink? Warnings { get; init; }

    /// <summary>Maximum total bytes decompressed from the package (zip bomb guard).</summary>
    public long MaxTotalUncompressedSize { get; init; } = 512L * 1024 * 1024;

    /// <summary>Maximum decompressed size of any single part.</summary>
    public long MaxPartSize { get; init; } = 128L * 1024 * 1024;

    public int MaxPartCount { get; init; } = 10_000;
}
