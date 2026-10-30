namespace DotOffice.Core;

/// <summary>A non-fatal problem found during conversion (unsupported content, substituted font…).</summary>
public sealed record ConversionWarning(string Code, string Message)
{
    public override string ToString() => $"{Code}: {Message}";
}

/// <summary>Receives conversion warnings. Implementations must be cheap; warnings may be frequent.</summary>
public interface IConversionWarningSink
{
    void Warn(ConversionWarning warning);
}

/// <summary>Collects warnings in memory, de-duplicating identical messages.</summary>
public sealed class ConversionWarningList : IConversionWarningSink
{
    private readonly List<ConversionWarning> _warnings = [];
    private readonly HashSet<ConversionWarning> _seen = [];

    public IReadOnlyList<ConversionWarning> Warnings => _warnings;

    public void Warn(ConversionWarning warning)
    {
        lock (_warnings)
            if (_seen.Add(warning)) _warnings.Add(warning);
    }
}

public static class ConversionWarningExtensions
{
    public static void Warn(this IConversionWarningSink? sink, string code, string message) =>
        sink?.Warn(new ConversionWarning(code, message));
}