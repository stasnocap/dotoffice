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
/// Basic OpenType shaping: Arabic-style joining forms (isol/fina/medi/init), required and contextual
/// ligatures, kerning and mark positioning from the font's GSUB/GPOS tables. Indic and other scripts
/// needing reordering are not handled (their glyphs are mapped one to one).
/// </summary>
public static class TextShaper
{
    private const uint Isol = 1, Fina = 2, Medi = 4, Init = 8, Common = 16, Kern = 32, Marks = 64;

    /// <summary>
    /// True when <paramref name="text"/> needs more than a cmap lookup: complex-script letters, combining
    /// marks or (with <paramref name="kerning"/>) pair kerning.
    /// </summary>
    public static bool NeedsShaping(string text, bool kerning)
    {
        if (kerning) return true;
        foreach (var rune in text.EnumerateRunes())
            if (ScriptOf(rune.Value) is "arab" or "syrc" or "nko " or "hebr" || IsMark(rune.Value)) return true;
        return false;
    }

    /// <summary>True for scripts whose shaping (reordering, conjuncts) is not implemented, e.g. Indic.</summary>
    public static bool IsUnsupportedComplexScript(int codePoint) =>
        codePoint is >= 0x0900 and <= 0x0DFF or >= 0x0F00 and <= 0x0FFF or >= 0x1000 and <= 0x109F
            or >= 0x1780 and <= 0x17FF or >= 0xA8E0 and <= 0xA8FF;

    /// <summary>
    /// Shapes one run of text in one font. <paramref name="rightToLeft"/> only affects how mark offsets are
    /// expressed (glyphs are returned in logical order either way).
    /// </summary>
    public static ShapedGlyph[] Shape(TrueTypeFont font, string text, bool rightToLeft, bool kerning)
    {
        var codePoints = new List<(int CodePoint, int Cluster)>(text.Length);
        for (int i = 0; i < text.Length; i += char.IsSurrogatePair(text, i) ? 2 : 1)
            codePoints.Add((char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i], i));
        ReorderMarks(codePoints);
        var glyphs = Nominal(font, codePoints, Common | Kern | Marks);

        string script = "DFLT";
        foreach (var (cp, _) in codePoints)
            if (ScriptOf(cp) is { } s) { script = s; if (s != "latn") break; }
        bool joining = script is "arab" or "syrc" or "nko ";
        bool complex = joining || script == "hebr";
        if (joining) SetJoiningForms(codePoints.Select(c => c.CodePoint).ToList(), glyphs);

        var layout = font.OpenTypeLayout;
        try
        {
            string[][] stages = joining
                ? [["ccmp", "locl"], ["isol"], ["fina"], ["medi"], ["init"], ["rlig"], ["calt", "liga"]]
                : [["ccmp", "locl"], ["rlig"]];
            if (layout.HasSubstitutions)
                foreach (var stage in stages)
                    layout.Apply(false, glyphs, layout.Lookups(false, script, stage.Select(Feature).ToList()));

            for (int i = 0; i < glyphs.Count; i++)
                glyphs[i] = glyphs[i] with { XAdvance = font.GetAdvanceWidth(glyphs[i].Glyph) };
            if (layout.HasPositioning)
            {
                // Marks do not advance the pen; GPOS positions them relative to their base.
                for (int i = 0; i < glyphs.Count; i++)
                    if (layout.GlyphClass(glyphs[i].Glyph) == 3) glyphs[i] = glyphs[i] with { XAdvance = 0 };
                if (kerning || complex) layout.Apply(true, glyphs, layout.Lookups(true, script, [("kern", Kern)]));
                layout.Apply(true, glyphs, layout.Lookups(true, script, [("mark", Marks), ("mkmk", Marks)]));
                ResolveAttachments(glyphs, rightToLeft);
            }
        }
        catch (Exception e) when (e is ArgumentOutOfRangeException or IndexOutOfRangeException)
        {
            glyphs = Nominal(font, codePoints, 0); // malformed layout tables: fall back to nominal glyphs
        }
        return glyphs.Select(g => new ShapedGlyph(g.Glyph, g.Cluster, g.Text, g.XAdvance, g.XOffset, g.YOffset)).ToArray();
    }

    private static List<GlyphInfo> Nominal(TrueTypeFont font, List<(int CodePoint, int Cluster)> codePoints, uint mask) =>
        codePoints.Select(c =>
        {
            var gid = font.GetGlyphId(c.CodePoint);
            return new GlyphInfo
            {
                Glyph = gid, Cluster = c.Cluster, Text = char.ConvertFromUtf32(c.CodePoint), Mask = mask,
                XAdvance = font.GetAdvanceWidth(gid), AttachTo = -1,
            };
        }).ToList();

    /// <summary>
    /// Canonical ordering of combining marks (stable sort of each mark sequence by combining class), so mark
    /// stacking does not depend on typing order. Like HarfBuzz, Hebrew shin/sin dots sort right after the base
    /// and the Arabic shadda before the vowel marks.
    /// </summary>
    private static void ReorderMarks(List<(int CodePoint, int Cluster)> codePoints)
    {
        for (int i = 0; i < codePoints.Count; i++)
        {
            if (CombiningClass(codePoints[i].CodePoint) == 0) continue;
            int end = i;
            while (end < codePoints.Count && CombiningClass(codePoints[end].CodePoint) != 0) end++;
            if (end - i > 1)
            {
                var sorted = codePoints.GetRange(i, end - i).OrderBy(c => CombiningClass(c.CodePoint)).ToList(); // OrderBy is stable
                for (int k = 0; k < sorted.Count; k++) codePoints[i + k] = sorted[k];
            }
            i = end;
        }
    }

    private static int CombiningClass(int cp)
    {
        if (cp is 0x05C1 or 0x05C2) return 1; // shin dot, sin dot
        if (cp == 0x0651) return 27; // shadda
        var starts = CombiningClassData.RangeStarts;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= cp) lo = mid; else hi = mid - 1;
        }
        return CombiningClassData.RangeValues[lo];
    }

    private static (string, uint) Feature(string tag) => tag switch
    {
        "isol" => (tag, Isol),
        "fina" => (tag, Fina),
        "medi" => (tag, Medi),
        "init" => (tag, Init),
        _ => (tag, Common),
    };

    /// <summary>Turns mark attachments into offsets relative to each mark's own pen position.</summary>
    private static void ResolveAttachments(List<GlyphInfo> glyphs, bool rightToLeft)
    {
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
            g.XOffset += glyphs[parent].XOffset;
            g.YOffset += glyphs[parent].YOffset;
            g.AttachTo = -1;
            glyphs[i] = g;
        }
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

    private static bool IsMark(int cp) =>
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
