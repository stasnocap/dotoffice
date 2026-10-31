namespace DotOffice.Fonts.Shaping;

/// <summary>
/// A character after normalization. <see cref="Text"/> is the text it stands for: the original character on the
/// first part of a decomposition (empty on the others), the letter and its marks after a composition.
/// </summary>
internal readonly record struct NormalizedChar(int CodePoint, int Cluster, string Text);

/// <summary>
/// HarfBuzz's normalization before shaping (hb-ot-shape-normalize.cc): characters the font lacks are decomposed
/// into parts it has, combining marks are put in canonical order, then each letter is composed with following
/// marks when the font has the composite.
/// </summary>
internal static class Normalizer
{
    /// <summary>The shaper whose normalization hooks apply (HarfBuzz's decompose, compose and reorder_marks).</summary>
    public enum Shaper { Default, Arabic, Hebrew, Indic }

    private const int MaxCombiningMarks = 32;

    private struct Item
    {
        public int CodePoint, Cluster, Class;
        public string Text;
    }

    private static readonly Dictionary<int, (int First, int Second)> Decompositions = new();
    private static readonly Dictionary<(int, int), int> Compositions = new();

    static Normalizer()
    {
        var data = DecompositionData.Decompositions;
        for (int i = 0; i < data.Length; i += 4)
        {
            Decompositions[data[i]] = (data[i + 1], data[i + 2]);
            if (data[i + 3] != 0) Compositions[(data[i + 1], data[i + 2])] = data[i];
        }
    }

    /// <summary>
    /// Normalizes <paramref name="text"/> for <paramref name="font"/>. Indic text is fully decomposed before
    /// recomposing (HarfBuzz's COMPOSED_DIACRITICS_NO_SHORT_CIRCUIT); other text keeps characters the font has.
    /// <paramref name="hasGposMark"/> (the font positions marks) turns off the Hebrew presentation-form fallback.
    /// </summary>
    public static List<NormalizedChar> Normalize(TrueTypeFont font, IReadOnlyList<(int CodePoint, int Cluster)> text,
        Shaper shaper, bool hasGposMark = true)
    {
        bool shortCircuit = shaper != Shaper.Indic;
        var items = new List<Item>(text.Count);

        // First round: decompose.
        bool allSimple = true;
        int count = text.Count, idx = 0;
        while (idx < count)
        {
            int end = idx + 1;
            while (end < count && !TextShaper.IsMark(text[end].CodePoint)) end++;
            if (end < count) end--; // leave one base for the marks to cluster with
            for (; idx < end; idx++) DecomposeCurrent(font, text[idx], shortCircuit, shaper, items);
            if (idx == count) break;

            allSimple = false;
            for (end = idx + 1; end < count; end++)
                if (!TextShaper.IsMark(text[end].CodePoint)) break;
            bool variationSelector = false;
            for (int i = idx; i < end; i++) variationSelector |= IsVariationSelector(text[i].CodePoint);
            for (; idx < end; idx++)
            {
                // Clusters with a variation selector are left alone.
                if (variationSelector) items.Add(NewItem(text[idx].CodePoint, text[idx].Cluster, Original(text[idx].CodePoint)));
                else DecomposeCurrent(font, text[idx], shortest: false, shaper, items);
            }
        }
        if (allSimple) return items.Select(c => new NormalizedChar(c.CodePoint, c.Cluster, c.Text)).ToList();

        // Second round: reorder marks (stable sort by modified combining class, then the shaper's own rules).
        for (int i = 0; i < items.Count; i++)
        {
            if (items[i].Class == 0) continue;
            int end = i + 1;
            while (end < items.Count && items[end].Class != 0) end++;
            if (end - i <= MaxCombiningMarks)
            {
                var sorted = items.GetRange(i, end - i).OrderBy(c => c.Class).ToList(); // OrderBy is stable
                for (int k = 0; k < sorted.Count; k++) items[i + k] = sorted[k];
                if (shaper == Shaper.Arabic) ReorderArabicMarks(items, i, end);
                else if (shaper == Shaper.Hebrew) ReorderHebrewMarks(items, i, end);
            }
            i = end;
        }

        // Third round: recompose marks with their starter when not blocked and the font has the composite.
        var output = new List<Item>(items.Count) { items[0] };
        int starter = 0;
        for (int i = 1; i < items.Count; i++)
        {
            var current = items[i];
            if (TextShaper.IsMark(current.CodePoint)
                && (starter == output.Count - 1 || output[^1].Class < current.Class)
                && Compose(output[starter].CodePoint, current.CodePoint, shaper, hasGposMark) is var composed and not 0
                && font.GetGlyphId(composed) != 0)
            {
                var s = output[starter];
                output[starter] = NewItem(composed, s.Cluster, s.Text + current.Text);
                continue;
            }
            output.Add(current);
            if (current.Class == 0) starter = output.Count - 1;
        }
        return output.Select(c => new NormalizedChar(c.CodePoint, c.Cluster, c.Text)).ToList();
    }

    private static string Original(int cp) => char.ConvertFromUtf32(cp);

    private static Item NewItem(int cp, int cluster, string text) => new()
    {
        CodePoint = cp, Cluster = cluster, Text = text, Class = TextShaper.IsMark(cp) ? CombiningClass(cp) : 0,
    };

    private static void DecomposeCurrent(TrueTypeFont font, (int CodePoint, int Cluster) c, bool shortest, Shaper shaper, List<Item> output)
    {
        int start = output.Count;
        if (!(shortest && font.GetGlyphId(c.CodePoint) != 0) && Decompose(font, shortest, c.CodePoint, c.Cluster, shaper, output) > 0)
        {
            output[start] = output[start] with { Text = Original(c.CodePoint) };
            return;
        }
        output.Add(NewItem(c.CodePoint, c.Cluster, Original(c.CodePoint)));
    }

    /// <summary>Appends the decomposition of <paramref name="ab"/> into parts the font has; returns their number (0 if none).</summary>
    private static int Decompose(TrueTypeFont font, bool shortest, int ab, int cluster, Shaper shaper, List<Item> output)
    {
        if (!DecomposeChar(ab, shaper, out int a, out int b) || b != 0 && font.GetGlyphId(b) == 0) return 0;
        bool hasA = font.GetGlyphId(a) != 0;
        if (!(shortest && hasA))
        {
            int done = Decompose(font, shortest, a, cluster, shaper, output);
            if (done > 0)
            {
                if (b == 0) return done;
                output.Add(NewItem(b, cluster, ""));
                return done + 1;
            }
            if (!hasA) return 0;
        }
        output.Add(NewItem(a, cluster, ""));
        if (b == 0) return 1;
        output.Add(NewItem(b, cluster, ""));
        return 2;
    }

    /// <summary>
    /// The glyph for a character left whole by normalization: its own, else (HarfBuzz's decompose_current_character)
    /// the space glyph for a Unicode space or the hyphen's glyph for U+2011 NON-BREAKING HYPHEN.
    /// </summary>
    public static ushort NominalGlyph(TrueTypeFont font, int cp)
    {
        int glyph = font.GetGlyphId(cp);
        if (glyph != 0) return (ushort)glyph;
        if (SpaceFallback(cp) != 0) return (ushort)font.GetGlyphId(0x20);
        return (ushort)(cp == 0x2011 ? font.GetGlyphId(0x2010) : 0);
    }

    /// <summary>True when shaping draws <paramref name="cp"/> with a glyph of the font, directly or through normalization.</summary>
    public static bool Covers(TrueTypeFont font, int cp, Shaper shaper = Shaper.Default) =>
        NominalGlyph(font, cp) != 0
        || DecomposeChar(cp, shaper, out int a, out int b) && Covers(font, a, shaper) && (b == 0 || font.GetGlyphId(b) != 0);

    /// <summary>
    /// For a space HarfBuzz may draw with the font's space glyph, its width as a fraction of the em (n for 1/n),
    /// or a negative code: -1 the space's own width, -2 4/18 em, -3 a digit, -4 '.' or ',', -5 half a space. 0 for
    /// other characters.
    /// </summary>
    public static int SpaceFallback(int cp) => cp switch
    {
        0x0020 or 0x00A0 => -1,
        0x2001 or 0x2003 or 0x3000 => 1,
        0x2000 or 0x2002 => 2,
        0x2004 => 3,
        0x2005 => 4,
        0x2009 => 5,
        0x2006 => 6,
        0x200A => 16,
        0x205F => -2,
        0x2007 => -3,
        0x2008 => -4,
        0x202F => -5,
        _ => 0,
    };

    private const int SBase = 0xAC00, LBase = 0x1100, VBase = 0x1161, TBase = 0x11A7;
    private const int LCount = 19, VCount = 21, TCount = 28, NCount = VCount * TCount, SCount = LCount * NCount;

    /// <summary>One step of canonical decomposition (Hangul syllables algorithmically); <paramref name="b"/> is 0 for singletons.</summary>
    private static bool DecomposeChar(int ab, Shaper shaper, out int a, out int b)
    {
        a = b = 0;
        // Indic: RRA, Bengali RRA/RHA and Tamil AU are kept whole.
        if (shaper == Shaper.Indic && ab is 0x0931 or 0x09DC or 0x09DD or 0x0B94) return false;
        int si = ab - SBase;
        if (si is >= 0 and < SCount)
        {
            (a, b) = si % TCount != 0 ? (SBase + si / TCount * TCount, TBase + si % TCount)
                : (LBase + si / NCount, VBase + si % NCount / TCount);
            return true;
        }
        if (!Decompositions.TryGetValue(ab, out var parts)) return false;
        (a, b) = parts;
        return true;
    }

    /// <summary>The primary composite of <paramref name="a"/> and <paramref name="b"/> under the shaper's rules, or 0.</summary>
    private static int Compose(int a, int b, Shaper shaper, bool hasGposMark)
    {
        if (shaper == Shaper.Indic)
        {
            if (TextShaper.IsMark(a)) return 0; // do not recompose split matras
            if (a == 0x09AF && b == 0x09BC) return 0x09DF;
        }
        if (a is >= SBase and < SBase + SCount && b is > TBase and < TBase + TCount && (a - SBase) % TCount == 0)
            return a + b - TBase;
        if (a is >= LBase and < LBase + LCount && b is >= VBase and < VBase + VCount)
            return SBase + (a - LBase) * NCount + (b - VBase) * TCount;
        if (Compositions.TryGetValue((a, b), out int composite)) return composite;
        return shaper == Shaper.Hebrew && !hasGposMark ? HebrewPresentationForm(a, b) : 0;
    }

    /// <summary>Hebrew presentation forms excluded from normalization but wanted for old fonts without mark positioning.</summary>
    private static int HebrewPresentationForm(int a, int b) => (b, a) switch
    {
        (0x05B4, 0x05D9) => 0xFB1D,
        (0x05B7, 0x05F2) => 0xFB1F,
        (0x05B7, 0x05D0) => 0xFB2E,
        (0x05B8, 0x05D0) => 0xFB2F,
        (0x05B9, 0x05D5) => 0xFB4B,
        (0x05BC, >= 0x05D0 and <= 0x05EA) => DageshForms[a - 0x05D0],
        (0x05BC, 0xFB2A) => 0xFB2C,
        (0x05BC, 0xFB2B) => 0xFB2D,
        (0x05BF, 0x05D1) => 0xFB4C,
        (0x05BF, 0x05DB) => 0xFB4D,
        (0x05BF, 0x05E4) => 0xFB4E,
        (0x05C1, 0x05E9) => 0xFB2A,
        (0x05C1, 0xFB49) => 0xFB2C,
        (0x05C2, 0x05E9) => 0xFB2B,
        (0x05C2, 0xFB49) => 0xFB2D,
        _ => 0,
    };

    private static readonly int[] DageshForms =
    [
        0xFB30, 0xFB31, 0xFB32, 0xFB33, 0xFB34, 0xFB35, 0xFB36, 0, 0xFB38, 0xFB39, 0xFB3A, 0xFB3B, 0xFB3C, 0, 0xFB3E,
        0, 0xFB40, 0xFB41, 0, 0xFB43, 0xFB44, 0, 0xFB46, 0xFB47, 0xFB48, 0xFB49, 0xFB4A,
    ];

    /// <summary>Arabic modifier combining marks (hamza above/below and the like), which go before other marks of their class.</summary>
    private static bool IsModifierCombiningMark(int cp) =>
        cp is 0x0654 or 0x0655 or 0x0658 or 0x06DC or 0x06E3 or 0x06E7 or 0x06E8 or 0x08CA or 0x08CB or 0x08CD or 0x08CE
            or 0x08CF or 0x08D3 or 0x08F3;

    /// <summary>Moves Arabic modifier combining marks of class 220 and 230 to the start of the mark sequence.</summary>
    private static void ReorderArabicMarks(List<Item> items, int start, int end)
    {
        int i = start;
        for (int cc = 220; cc <= 230; cc += 10)
        {
            while (i < end && items[i].Class < cc) i++;
            if (i == end) break;
            if (items[i].Class > cc) continue;
            int j = i;
            while (j < end && items[j].Class == cc && IsModifierCombiningMark(items[j].CodePoint)) j++;
            if (i == j) continue;

            var moved = items.GetRange(i, j - i);
            items.RemoveRange(i, j - i);
            items.InsertRange(start, moved);
            // Renumber so the sequence stays sorted (HarfBuzz's modified classes 25 and 26, below all Arabic ones).
            int newStart = start + j - i;
            for (; start < newStart; start++) items[start] = items[start] with { Class = cc == 220 ? 25 : 26 };
            i = j;
        }
    }

    /// <summary>Patah or qamats, then sheva or hiriq, then meteg or a below mark: the last two swap (meteg goes first).</summary>
    private static void ReorderHebrewMarks(List<Item> items, int start, int end)
    {
        for (int i = start + 2; i < end; i++)
        {
            int c0 = items[i - 2].Class, c1 = items[i - 1].Class, c2 = items[i].Class;
            if (c0 is 20 or 21 && c1 is 22 or 23 && c2 is 25 or 220)
            {
                (items[i - 1], items[i]) = (items[i], items[i - 1]);
                break;
            }
        }
    }

    private static bool IsVariationSelector(int cp) => cp is >= 0x180B and <= 0x180D or 0x180F or >= 0xFE00 and <= 0xFE0F
        or >= 0xE0100 and <= 0xE01EF;

    /// <summary>
    /// HarfBuzz's modified combining classes: Hebrew and Arabic points in a typographic order (shin/sin dot and
    /// shadda first), Telugu length marks before the virama, and a few Thai, Lao and Tibetan vowels moved.
    /// </summary>
    private static readonly Dictionary<int, int> ModifiedClasses = new()
    {
        [10] = 22, [11] = 15, [12] = 16, [13] = 17, [14] = 23, [15] = 18, [16] = 19, [17] = 20, [18] = 21, [19] = 14,
        [20] = 24, [21] = 12, [22] = 25, [23] = 13, [24] = 10, [25] = 11, [26] = 26,
        [27] = 28, [28] = 29, [29] = 30, [30] = 31, [31] = 32, [32] = 33, [33] = 27, [34] = 34, [35] = 35, [36] = 36,
        [84] = 4, [91] = 5, [103] = 3, [130] = 132, [132] = 131,
    };

    internal static int CombiningClass(int cp)
    {
        if (cp is 0x1A60 or 0x0FC6) return 254;
        if (cp == 0x0F39) return 127;
        var starts = CombiningClassData.RangeStarts;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= cp) lo = mid; else hi = mid - 1;
        }
        int ccc = CombiningClassData.RangeValues[lo];
        return ModifiedClasses.GetValueOrDefault(ccc, ccc);
    }
}
