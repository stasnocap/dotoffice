using DotOffice.Core;
using DotOffice.Fonts;

namespace DotOffice.Layout;

/// <summary>Vertical metrics of a font at a size, following Office's line-height model.</summary>
internal readonly record struct FontMetrics(double Ascent, double Descent)
{
    public double Height => Ascent + Descent;

    /// <summary>
    /// Ascent/descent from OS/2 win metrics, plus any hhea line gap not already covered by them
    /// ("external leading", added below the baseline). This matches Word's single line spacing.
    /// </summary>
    public static FontMetrics Of(TrueTypeFont font, double size)
    {
        double scale = size / font.UnitsPerEm;
        double hhea = font.Ascender - font.Descender + Math.Max((short)0, font.LineGap);
        if (font.WinAscent + font.WinDescent == 0)
            return new FontMetrics(font.Ascender * scale, (-font.Descender + Math.Max((short)0, font.LineGap)) * scale);
        double win = font.WinAscent + font.WinDescent;
        return new FontMetrics(font.WinAscent * scale, (font.WinDescent + Math.Max(0, hhea - win)) * scale);
    }
}

/// <summary>Picks concrete fonts for runs (per character slot, with fallback) and reports substitutions.</summary>
internal sealed class FontSelector(FontResolver resolver, IConversionWarningSink? warnings)
{
    private readonly HashSet<string> _reported = new(StringComparer.OrdinalIgnoreCase);

    public ResolvedFont Resolve(string family, bool bold, bool italic)
    {
        var resolved = resolver.Resolve(family, bold, italic);
        if (resolved.IsSubstitute && _reported.Add(family))
            warnings.Warn("font.substituted", $"Font '{family}' is not available; using '{resolved.Font.FamilyName}'.");
        return resolved;
    }

    /// <summary>The font a run uses for a character, before glyph-coverage fallback.</summary>
    public ResolvedFont ForCharacter(RunProperties p, int codePoint)
    {
        string family = codePoint < 0x80
            ? p.FontFamily
            : IsEastAsian(codePoint) ? p.EastAsianFontFamily ?? p.HighAnsiFontFamily ?? p.FontFamily : p.HighAnsiFontFamily ?? p.FontFamily;
        return Resolve(family, p.Bold, p.Italic);
    }

    /// <summary>Like <see cref="ForCharacter"/>, then falls back to another font if the glyph is missing.</summary>
    public ResolvedFont ForCharacterWithFallback(RunProperties p, int codePoint)
    {
        var font = ForCharacter(p, codePoint);
        if (font.Font.HasGlyph(codePoint) || codePoint < 0x20 || IsInvisible(codePoint)) return font;
        var fallback = resolver.ResolveFallback(codePoint, p.Bold, p.Italic);
        if (fallback is null)
        {
            if (_reported.Add($"U+{codePoint:X4}"))
                warnings.Warn("font.missing-glyph", $"No available font contains U+{codePoint:X4}.");
            return font;
        }
        return fallback;
    }

    public static bool IsInvisible(int cp) => cp is 0x200B or 0x200C or 0x200D or 0x2060 or 0xFEFF or 0x00AD;

    public static bool IsEastAsian(int cp) =>
        cp is >= 0x1100 and <= 0x11FF        // Hangul Jamo
            or >= 0x2E80 and <= 0x303F       // CJK radicals, punctuation
            or >= 0x3040 and <= 0x33FF       // Kana, compatibility
            or >= 0x3400 and <= 0x4DBF       // CJK ext A
            or >= 0x4E00 and <= 0x9FFF       // CJK unified
            or >= 0xA960 and <= 0xA97F
            or >= 0xAC00 and <= 0xD7FF       // Hangul syllables
            or >= 0xF900 and <= 0xFAFF       // CJK compatibility ideographs
            or >= 0xFE30 and <= 0xFE4F
            or >= 0xFF00 and <= 0xFFEF       // full/half-width forms
            or >= 0x20000 and <= 0x3FFFF;
}