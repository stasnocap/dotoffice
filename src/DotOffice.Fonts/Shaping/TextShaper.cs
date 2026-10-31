using System.Globalization;

namespace DotOffice.Fonts.Shaping;

/// <summary>
/// A shaped glyph in logical order. <see cref="Cluster"/> is the UTF-16 index of the first character the
/// glyph represents and <see cref="Text"/> the characters themselves (several for a ligature, empty for the
/// extra glyphs of a decomposition); advance and offsets are in font units.
/// </summary>
public readonly record struct ShapedGlyph(ushort GlyphId, int Cluster, string Text, int Advance, int XOffset, int YOffset);

internal enum JoiningType : byte { NotListed, U, R, L, D, C, T }

/// <summary>
/// Basic OpenType shaping, after HarfBuzz's default, Arabic and Hebrew shapers: normalization, Arabic-style
/// joining forms (isol/fina/medi/init), ligatures, kerning and mark positioning from the font's GSUB/GPOS tables.
/// Indic scripts go to the <see cref="IndicShaper"/>; Sinhala, Tibetan, Myanmar and Khmer are not handled (their
/// glyphs are mapped one to one).
/// </summary>
public static class TextShaper
{
    private const uint Isol = 1, Fina = 2, Medi = 4, Init = 8, Common = 16, Kern = 32, Marks = 64;

    /// <summary>
    /// True when <paramref name="text"/> needs more than a cmap lookup in <paramref name="font"/>: complex-script
    /// letters, combining marks, glyphs a default substitution (ligatures such as "fi") may start at, or (with
    /// <paramref name="kerning"/>) pair kerning.
    /// </summary>
    public static bool NeedsShaping(TrueTypeFont font, string text, bool kerning)
    {
        if (kerning) return true;
        var layout = font.OpenTypeLayout;
        int previous = -1;
        foreach (var rune in text.EnumerateRunes())
        {
            if (ScriptOf(rune.Value) is "arab" or "syrc" or "nko " or "hebr" || IsMark(rune.Value) || IndicShaper.IsIndic(rune.Value))
                return true;
            int glyph = font.GetGlyphId(rune.Value);
            if (layout.HasSubstitutions && previous >= 0 && layout.StartsDefaultSubstitution(previous, glyph)) return true;
            previous = glyph;
        }
        return layout.HasSubstitutions && previous >= 0 && layout.StartsDefaultSubstitution(previous, -1);
    }

    /// <summary>True for scripts whose shaping (reordering, conjuncts) is not implemented: Sinhala, Tibetan, Myanmar, Khmer.</summary>
    public static bool IsUnsupportedComplexScript(int codePoint) =>
        codePoint is >= 0x0D80 and <= 0x0DFF or >= 0x0F00 and <= 0x0FFF or >= 0x1000 and <= 0x109F or >= 0x1780 and <= 0x17FF;

    /// <summary>
    /// Shapes one run of text in one font. <paramref name="rightToLeft"/> only affects how mark offsets are
    /// expressed (glyphs are returned in logical order either way). Without <paramref name="ligatures"/> the
    /// optional ligatures (liga, clig) are off, as LibreOffice does for text with character spacing.
    /// </summary>
    public static ShapedGlyph[] Shape(TrueTypeFont font, string text, bool rightToLeft, bool kerning, bool ligatures = true)
    {
        if (!rightToLeft && text.EnumerateRunes().Any(r => IndicShaper.IsIndic(r.Value)))
        {
            try
            {
                if (IndicShaper.Shape(font, text, ligatures) is { } indic)
                    return indic.Select(g => new ShapedGlyph(g.Glyph, g.Cluster, g.Text, g.XAdvance, g.XOffset, g.YOffset)).ToArray();
            }
            catch (Exception e) when (e is ArgumentOutOfRangeException or IndexOutOfRangeException)
            {
                // malformed layout tables: fall through to nominal shaping
            }
        }
        var codePoints = new List<(int CodePoint, int Cluster)>(text.Length);
        for (int i = 0; i < text.Length; i += char.IsSurrogatePair(text, i) ? 2 : 1)
            codePoints.Add((char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i], i));
        string script = "DFLT";
        foreach (var (cp, _) in codePoints)
            if (ScriptOf(cp) is { } s) { script = s; if (s != "latn") break; }
        bool joining = script is "arab" or "syrc" or "nko ";
        bool complex = joining || script == "hebr";

        var layout = font.OpenTypeLayout;
        var chars = Normalizer.Normalize(font, codePoints,
            joining ? Normalizer.Shaper.Arabic : script == "hebr" ? Normalizer.Shaper.Hebrew : Normalizer.Shaper.Default,
            hasGposMark: layout.Lookups(true, script, [("mark", Marks)]).Count > 0);
        var glyphs = Nominal(font, chars, Common | Kern | Marks);
        if (joining) SetJoiningForms(chars.Select(c => c.CodePoint).ToList(), glyphs);

        try
        {
            // HarfBuzz's stages: the Arabic shaper pauses after rlig and (unless the font has rclt) after calt;
            // the default shaper applies all its features together.
            string[][] stages = !joining ? [["ccmp", "locl", "rlig", "calt", "clig", "liga", "rclt"]]
                : layout.Lookups(false, script, [Feature("rclt")]).Count > 0
                    ? [["ccmp", "locl"], ["isol"], ["fina"], ["medi"], ["init"], ["rlig"], ["calt", "clig", "liga", "rclt"]]
                    : [["ccmp", "locl"], ["isol"], ["fina"], ["medi"], ["init"], ["rlig"], ["calt"], ["clig", "liga"]];
            if (layout.HasSubstitutions)
                foreach (var stage in stages)
                    layout.Apply(false, glyphs, layout.Lookups(false, script,
                        stage.Where(f => ligatures || f is not ("liga" or "clig")).Select(Feature).ToList()));

            for (int i = 0; i < glyphs.Count; i++)
                glyphs[i] = glyphs[i] with { XAdvance = font.GetAdvanceWidth(glyphs[i].Glyph) };
            if (layout.HasPositioning)
            {
                // Marks do not advance the pen; GPOS positions them relative to their base.
                for (int i = 0; i < glyphs.Count; i++)
                    if (layout.GlyphClass(glyphs[i].Glyph) == 3) glyphs[i] = glyphs[i] with { XAdvance = 0 };
                if (kerning || complex)
                    layout.Apply(true, glyphs, layout.Lookups(true, script, joining ? [("kern", Kern), ("dist", Kern), ("curs", Common)] : [("kern", Kern)]), rightToLeft);
                layout.Apply(true, glyphs, layout.Lookups(true, script,
                    [("mark", Marks | OpenTypeLayout.ManualJoiners), ("mkmk", Marks | OpenTypeLayout.ManualJoiners)]));
                ResolveAttachments(glyphs, rightToLeft);
            }
        }
        catch (Exception e) when (e is ArgumentOutOfRangeException or IndexOutOfRangeException)
        {
            glyphs = Nominal(font, chars, 0); // malformed layout tables: fall back to nominal glyphs
        }
        HideIgnorables(font, glyphs);
        return glyphs.Select(g => new ShapedGlyph(g.Glyph, g.Cluster, g.Text, g.XAdvance, g.XOffset, g.YOffset)).ToArray();
    }

    private static List<GlyphInfo> Nominal(TrueTypeFont font, List<NormalizedChar> chars, uint mask) =>
        chars.Select(c =>
        {
            var gid = font.GetGlyphId(c.CodePoint);
            return new GlyphInfo
            {
                Glyph = gid, Cluster = c.Cluster, Text = c.Text, Mask = mask,
                XAdvance = font.GetAdvanceWidth(gid), AttachTo = -1, CursiveTo = -1, CodePoint = c.CodePoint,
                Ignorable = IsDefaultIgnorable(c.CodePoint),
            };
        }).ToList();

    private static (string, uint) Feature(string tag) => tag switch
    {
        "isol" => (tag, Isol),
        "fina" => (tag, Fina),
        "medi" => (tag, Medi),
        "init" => (tag, Init),
        _ => (tag, Common),
    };

    /// <summary>ZWJ, ZWNJ and other default ignorables become invisible: the space glyph without advance (or nothing).</summary>
    internal static void HideIgnorables(TrueTypeFont font, List<GlyphInfo> glyphs)
    {
        int space = font.GetGlyphId(' ');
        for (int i = glyphs.Count - 1; i >= 0; i--)
        {
            if (!glyphs[i].Ignorable) continue;
            if (space == 0) glyphs.RemoveAt(i);
            else glyphs[i] = glyphs[i] with { Glyph = (ushort)space, XAdvance = 0, XOffset = 0, YOffset = 0 };
        }
    }

    internal static bool IsDefaultIgnorable(int cp) =>
        cp is 0x00AD or 0x034F or 0x061C or >= 0x115F and <= 0x1160 or >= 0x17B4 and <= 0x17B5 or >= 0x180B and <= 0x180F
            or >= 0x200B and <= 0x200F or >= 0x202A and <= 0x202E or >= 0x2060 and <= 0x206F or 0x3164
            or >= 0xFE00 and <= 0xFE0F or 0xFEFF or 0xFFA0 or >= 0xFFF0 and <= 0xFFF8;

    /// <summary>Turns cursive and mark attachments into offsets relative to each glyph's own pen position.</summary>
    internal static void ResolveAttachments(List<GlyphInfo> glyphs, bool rightToLeft)
    {
        // Cursive chains run forward or (with the RightToLeft lookup flag) backward: resolve parents first.
        for (int i = 0; i < glyphs.Count; i++) ResolveCursive(glyphs, i, forward: true);
        for (int i = glyphs.Count - 1; i >= 0; i--) ResolveCursive(glyphs, i, forward: false);

        for (int i = 0; i < glyphs.Count; i++)
        {
            var g = glyphs[i];
            int parent = g.AttachTo;
            if (parent < 0 || parent >= i) continue;
            // Same convention as HarfBuzz: offsets are measured from the mark's pen position in drawing order.
            if (rightToLeft)
                for (int j = parent + 1; j <= i; j++) g.XOffset += glyphs[j].XAdvance;
            else
                for (int j = parent; j < i; j++) g.XOffset -= glyphs[j].XAdvance;
            g.XOffset += glyphs[parent].XOffset; // the vertical offset already includes the base's (see OpenTypeLayout)
            g.AttachTo = -1;
            glyphs[i] = g;
        }
    }

    private static void ResolveCursive(List<GlyphInfo> glyphs, int i, bool forward)
    {
        var g = glyphs[i];
        int parent = g.CursiveTo;
        if (parent < 0 || (forward ? parent >= i : parent <= i)) return;
        g.YOffset += glyphs[parent].YOffset;
        g.CursiveTo = -1;
        glyphs[i] = g;
    }

    /// <summary>Arabic joining: picks the isolated, final, medial or initial form of each letter.</summary>
    private static void SetJoiningForms(List<int> codePoints, List<GlyphInfo> glyphs)
    {
        var types = codePoints.Select(Joining).ToArray();
        for (int i = 0; i < types.Length; i++)
        {
            var type = types[i];
            if (type is JoiningType.T or JoiningType.U) continue;
            int prev = i - 1, next = i + 1;
            while (prev >= 0 && types[prev] == JoiningType.T) prev--;
            while (next < types.Length && types[next] == JoiningType.T) next++;
            bool joinsPrev = prev >= 0 && type is JoiningType.R or JoiningType.D or JoiningType.C
                             && types[prev] is JoiningType.L or JoiningType.D or JoiningType.C;
            bool joinsNext = next < types.Length && type is JoiningType.L or JoiningType.D or JoiningType.C
                             && types[next] is JoiningType.R or JoiningType.D or JoiningType.C;
            uint form = (joinsPrev, joinsNext) switch { (true, true) => Medi, (true, false) => Fina, (false, true) => Init, _ => Isol };
            glyphs[i] = glyphs[i] with { Mask = glyphs[i].Mask | form };
        }
    }

    private static JoiningType Joining(int cp)
    {
        var starts = JoiningData.RangeStarts;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= cp) lo = mid; else hi = mid - 1;
        }
        var type = (JoiningType)JoiningData.RangeValues[lo];
        if (type != JoiningType.NotListed) return type;
        return CharUnicodeInfo.GetUnicodeCategory(cp) is UnicodeCategory.NonSpacingMark or UnicodeCategory.EnclosingMark
            or UnicodeCategory.Format ? JoiningType.T : JoiningType.U;
    }

    internal static bool IsMark(int cp) =>
        CharUnicodeInfo.GetUnicodeCategory(cp) is UnicodeCategory.NonSpacingMark or UnicodeCategory.EnclosingMark
            or UnicodeCategory.SpacingCombiningMark;

    /// <summary>OpenType script tag for a letter, or null for common characters (digits, punctuation…).</summary>
    private static string? ScriptOf(int cp) => cp switch
    {
        >= 0x0600 and <= 0x06FF or >= 0x0750 and <= 0x077F or >= 0x0870 and <= 0x08FF
            or >= 0xFB50 and <= 0xFDFF or >= 0xFE70 and <= 0xFEFF => cp is >= 0x0660 and <= 0x0669 ? null : "arab",
        >= 0x0590 and <= 0x05FF or >= 0xFB1D and <= 0xFB4F => "hebr",
        >= 0x0700 and <= 0x074F => "syrc",
        >= 0x07C0 and <= 0x07FF => "nko ",
        >= 0x0370 and <= 0x03FF => "grek",
        >= 0x0400 and <= 0x052F => "cyrl",
        >= 'A' and <= 'Z' or >= 'a' and <= 'z' or >= 0x00C0 and <= 0x024F => "latn",
        _ => null,
    };
}
