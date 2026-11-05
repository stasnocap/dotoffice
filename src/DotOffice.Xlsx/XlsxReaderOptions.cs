using DotOffice.Core;

namespace DotOffice.Xlsx;

/// <summary>Options and resource limits for <see cref="XlsxReader"/>.</summary>
public sealed class XlsxReaderOptions
{
    /// <summary>Receives warnings about unsupported or malformed content.</summary>
    public IConversionWarningSink? Warnings { get; init; }

    /// <summary>
    /// The date that "today" means for rules that depend on it (conditional formatting by time period); the current date when null.
    /// </summary>
    public DateTime? Today { get; init; }

    /// <summary>Maximum total bytes decompressed from the package (zip bomb guard).</summary>
    public long MaxTotalUncompressedSize { get; init; } = 512L * 1024 * 1024;

    /// <summary>Maximum decompressed size of any single part.</summary>
    public long MaxPartSize { get; init; } = 128L * 1024 * 1024;

    public int MaxPartCount { get; init; } = 10_000;

    /// <summary>Maximum number of non-empty cells read from the whole workbook.</summary>
    public long MaxCellCount { get; init; } = 5_000_000;
}
