namespace DotOffice.Fonts;

/// <summary>A font chosen for a request, with flags telling the renderer to fake missing styles.</summary>
public sealed record ResolvedFont(TrueTypeFont Font, bool SyntheticBold, bool SyntheticItalic, bool IsSubstitute);

public sealed class FontNotFoundException(string message) : Exception(message);

/// <summary>
/// Resolves family + style requests to concrete fonts. Lookup order: every provider (user fonts first,
/// then system fonts) for the requested family, then metric-compatible substitutes (Calibri → Carlito,
/// Arial → Liberation Sans…), then a generic serif/sans/mono family. Missing bold/italic faces are
/// synthesized. <see cref="ResolveFallback"/> provides per-character fallback.
/// </summary>
public sealed class FontResolver
{
    private static readonly Dictionary<string, string[]> Substitutes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Calibri"] = ["Carlito"],
        ["Calibri Light"] = ["Carlito"],
        ["Carlito"] = ["Calibri"],
        ["Cambria"] = ["Caladea"],
        ["Cambria Math"] = ["Caladea", "Cambria"],
        ["Caladea"] = ["Cambria"],
        ["Arial"] = ["Liberation Sans", "Arimo", "Helvetica"],
        ["Helvetica"] = ["Liberation Sans", "Arimo", "Arial"],
        ["Liberation Sans"] = ["Arimo", "Arial"],
        ["Arial Narrow"] = ["Liberation Sans Narrow"],
        ["Times New Roman"] = ["Liberation Serif", "Tinos", "Times"],
        ["Times"] = ["Liberation Serif", "Tinos", "Times New Roman"],
        ["Liberation Serif"] = ["Tinos", "Times New Roman"],
        ["Courier New"] = ["Liberation Mono", "Cousine", "Courier"],
        ["Courier"] = ["Liberation Mono", "Cousine", "Courier New"],
        ["Liberation Mono"] = ["Cousine", "Courier New"],
        ["Consolas"] = ["Inconsolata", "Liberation Mono", "DejaVu Sans Mono"],
        ["Georgia"] = ["Gelasio", "DejaVu Serif"],
        ["Verdana"] = ["DejaVu Sans"],
        ["Tahoma"] = ["DejaVu Sans"],
        ["Segoe UI"] = ["Selawik", "Open Sans", "Noto Sans", "DejaVu Sans"],
        ["Aptos"] = ["Liberation Sans", "Arial"],
        ["Garamond"] = ["EB Garamond", "Liberation Serif"],
        ["Book Antiqua"] = ["TeX Gyre Pagella", "P052", "Palatino Linotype"],
        ["Palatino Linotype"] = ["TeX Gyre Pagella", "P052", "Book Antiqua"],
        ["Century Gothic"] = ["TeX Gyre Adventor", "URW Gothic"],
        ["Comic Sans MS"] = ["Comic Neue"],
    };

    private static readonly string[] SansFamilies =
        ["Liberation Sans", "Arimo", "Arial", "DejaVu Sans", "Noto Sans", "Helvetica", "Verdana", "Segoe UI"];

    private static readonly string[] SerifFamilies =
        ["Liberation Serif", "Tinos", "Times New Roman", "DejaVu Serif", "Noto Serif", "Georgia"];

    private static readonly string[] MonoFamilies =
        ["Liberation Mono", "Cousine", "Courier New", "DejaVu Sans Mono", "Noto Sans Mono", "Consolas", "Menlo"];

    private static readonly string[] CharacterFallbackFamilies =
    [
        "Noto Sans", "DejaVu Sans", "Liberation Sans", "Arial", "Segoe UI", "Segoe UI Symbol", "Arial Unicode MS",
        "Noto Sans Symbols", "Noto Sans Symbols 2", "Noto Sans Math", "DejaVu Serif", "FreeSerif", "FreeSans",
        "Noto Sans Arabic", "Noto Sans Hebrew", "Noto Sans Devanagari", "Noto Sans Bengali", "Noto Sans Thai",
        "Noto Sans Armenian", "Noto Sans Georgian", "Noto Sans Ethiopic", "Noto Sans Tamil",
        "Noto Sans SC", "Noto Sans JP", "Noto Sans KR", "Noto Sans TC", "Droid Sans Fallback",
        "WenQuanYi Zen Hei", "Microsoft YaHei", "SimSun", "MS Gothic", "Malgun Gothic", "Nirmala UI", "Leelawadee UI",
    ];

    private readonly IFontProvider[] _providers;
    private readonly Dictionary<(string, bool, bool), ResolvedFont> _cache = [];
    private readonly Dictionary<(int, bool, bool), ResolvedFont?> _fallbackCache = [];
    private readonly Dictionary<(string, bool, bool), TrueTypeFont?> _lookups = [];

    public FontResolver(params IFontProvider[] providers) => _providers = providers;

    /// <summary>User fonts (optional) followed by the shared system font index.</summary>
    public static FontResolver CreateDefault(IFontProvider? userFonts = null) =>
        userFonts is null ? new FontResolver(SystemFontProvider.Shared) : new FontResolver(userFonts, SystemFontProvider.Shared);

    public ResolvedFont Resolve(string family, bool bold, bool italic)
    {
        var key = (family, bold, italic);
        lock (_cache)
        {
            if (_cache.TryGetValue(key, out var cached)) return cached;
            var resolved = ResolveCore(family, bold, italic);
            _cache[key] = resolved;
            return resolved;
        }
    }

    /// <summary>
    /// Finds a font containing <paramref name="codePoint"/>: user-registered fonts first, then a list of
    /// broad-coverage families. Returns null when no known font covers the character.
    /// </summary>
    public ResolvedFont? ResolveFallback(int codePoint, bool bold, bool italic)
    {
        var key = (codePoint, bold, italic);
        lock (_cache)
        {
            if (_fallbackCache.TryGetValue(key, out var cached)) return cached;
            ResolvedFont? result = null;
            var families = _providers.SelectMany(p => p.GetFallbackFamilies()).Concat(CharacterFallbackFamilies);
            foreach (var family in families.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var font = Find(family, bold, italic);
                if (font is null || !font.HasGlyph(codePoint)) continue;
                result = WithSynthesis(font, bold, italic, isSubstitute: true);
                break;
            }
            _fallbackCache[key] = result;
            return result;
        }
    }

    private ResolvedFont ResolveCore(string family, bool bold, bool italic)
    {
        var candidates = new List<string> { family };
        if (Substitutes.TryGetValue(family, out var subs)) candidates.AddRange(subs);
        candidates.AddRange(GenericFamilies(family));

        // The first family that exists wins, even if it has to synthesize the style: a metric-compatible
        // regular face with fake bold beats a real bold face of an unrelated family.
        foreach (var name in candidates)
            if (Find(name, bold, italic) is { } font)
                return WithSynthesis(font, bold, italic, !string.Equals(name, family, StringComparison.OrdinalIgnoreCase));

        foreach (var provider in _providers)
            foreach (var name in provider.GetFallbackFamilies())
                if (provider.FindFont(name, bold, italic) is { } any)
                    return WithSynthesis(any, bold, italic, isSubstitute: true);

        throw new FontNotFoundException(
            $"No font available for '{family}' and no fallback font could be found. Register fonts via a FontCollection.");
    }

    private TrueTypeFont? Find(string family, bool bold, bool italic)
    {
        var key = (family, bold, italic);
        if (_lookups.TryGetValue(key, out var found)) return found;
        TrueTypeFont? nearest = null;
        foreach (var provider in _providers)
        {
            var font = provider.FindFont(family, bold, italic);
            if (font is null) continue;
            if (FaceMatcher.Matches(font, bold, italic)) { nearest = font; break; }
            nearest ??= font;
        }
        _lookups[key] = nearest;
        return nearest;
    }

    private static ResolvedFont WithSynthesis(TrueTypeFont font, bool bold, bool italic, bool isSubstitute) =>
        new(font, bold && !FaceMatcher.IsBold(font), italic && !font.IsItalic, isSubstitute);

    private static string[] GenericFamilies(string family)
    {
        var f = family.ToLowerInvariant();
        if (f.Contains("mono") || f.Contains("courier") || f.Contains("consol") || f.Contains("code"))
            return MonoFamilies;
        if (f.Contains("serif") && !f.Contains("sans") || f.Contains("times") || f.Contains("roman")
            || f.Contains("cambria") || f.Contains("georgia") || f.Contains("garamond") || f.Contains("book")
            || f.Contains("palatino") || f.Contains("caladea"))
            return SerifFamilies;
        return SansFamilies;
    }
}