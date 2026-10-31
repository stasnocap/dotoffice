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
        if (UsesComplexScript(p, codePoint))
            return Resolve(p.ComplexScriptFontFamily ?? p.FontFamily, p.ComplexScriptBold ?? p.Bold, p.ComplexScriptItalic ?? p.Italic);
        string family = codePoint < 0x80
            ? p.FontFamily
            : IsEastAsian(codePoint) ? p.EastAsianFontFamily ?? p.HighAnsiFontFamily ?? p.FontFamily : p.HighAnsiFontFamily ?? p.FontFamily;
        return Resolve(family, p.Bold, p.Italic);
    }

    /// <summary>Font size of a character: complex-script text uses its own size (w:szCs).</summary>
    public static double SizeFor(RunProperties p, int codePoint) =>
        UsesComplexScript(p, codePoint) ? p.ComplexScriptFontSize ?? p.FontSize : p.FontSize;

    /// <summary>
    /// Whether a character takes the complex-script formatting (font, size, bold, italic): complex-script
    /// letters always do, and every character of a run marked right-to-left (w:rtl) or complex (w:cs).
    /// </summary>
    public static bool UsesComplexScript(RunProperties p, int codePoint) => p.ComplexScript || p.RightToLeft || IsComplexScript(codePoint);

    public static bool IsComplexScript(int cp) =>
        cp is >= 0x0590 and <= 0x08FF         // Hebrew, Arabic, Syriac, Thaana, NKo, Samaritan…
            or >= 0x0900 and <= 0x0DFF        // Indic
            or >= 0x0E00 and <= 0x0EFF        // Thai, Lao
            or >= 0xFB1D and <= 0xFDFF        // Hebrew and Arabic presentation forms
            or >= 0xFE70 and <= 0xFEFF;

    /// <summary>Like <see cref="ForCharacter"/>, then falls back to another font if the glyph is missing.</summary>
    public ResolvedFont ForCharacterWithFallback(RunProperties p, int codePoint)
    {
        var font = ForCharacter(p, codePoint);
        if (font.Font.HasGlyph(codePoint) || codePoint < 0x20 || IsInvisible(codePoint)) return font;
        bool cs = UsesComplexScript(p, codePoint);
        var fallback = resolver.ResolveFallback(codePoint, cs ? p.ComplexScriptBold ?? p.Bold : p.Bold, cs ? p.ComplexScriptItalic ?? p.Italic : p.Italic);
        if (fallback is null)
        {
            if (_reported.Add($"U+{codePoint:X4}"))
                warnings.Warn("font.missing-glyph", $"No available font contains U+{codePoint:X4}.");
            return font;
        }
        return fallback;
    }

    /// <summary>Reports a problem once per <paramref name="code"/>.</summary>
    public void WarnOnce(string code, string message)
    {
        if (_reported.Add(code)) warnings.Warn(code, message);
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