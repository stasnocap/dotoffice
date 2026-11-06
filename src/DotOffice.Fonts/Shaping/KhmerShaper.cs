using System.Runtime.InteropServices;
using static DotOffice.Fonts.Shaping.SyllablePatterns;

namespace DotOffice.Fonts.Shaping;

/// <summary>
/// Shaping category of a Khmer character: the numbers of the "export" lines of HarfBuzz's
/// hb-ot-shaper-khmer-machine.rl (they are the values of <see cref="KhmerData"/>). Coeng is category H.
/// </summary>
internal enum KhmerCategory : byte
{
    X = 0, C = 1, V = 2, H = 4, ZWNJ = 5, ZWJ = 6, Placeholder = 10, DottedCircle = 11, Ra = 15,
    VAbv = 20, VBlw = 21, VPre = 22, VPst = 23, Robatic = 25, Xgroup = 26, Ygroup = 27,
}

/// <summary>
/// The Khmer shaper, as HarfBuzz implements it (hb-ot-shaper-khmer.cc and its syllable machine, MIT license; see
/// THIRD-PARTY-NOTICES.md): HarfBuzz, and so LibreOffice, shapes Khmer with it whatever the font. The rules come from
/// Microsoft's Khmer script development page; the grammar "was experimentally extracted from what Uniscribe allows".
/// </summary>
internal static class KhmerShaper
{
    private enum SyllableType { Consonant, Broken, NonKhmer }

    // Feature masks: Global applies to every glyph; the others to the glyphs the reordering picks.
    private const uint Global = 1, Pref = 2, Blwf = 4, Abvf = 8, Pstf = 16, Cfar = 32;

    public static bool IsKhmer(int cp) => cp is >= 0x1780 and <= 0x17FF;

    public static List<GlyphInfo>? Shape(TrueTypeFont font, string text, string script, bool ligatures = true)
    {
        if (script != "khmr") return null;
        var layout = font.OpenTypeLayout;

        var glyphs = Nominal(font, Normalize(font, text));
        FindSyllables(glyphs);

        uint perSyllable = OpenTypeLayout.PerSyllable, manual = OpenTypeLayout.ManualJoiners;
        uint MaskOf(string feature, uint mask) => layout.Lookups(false, script, [(feature, mask)]).Count > 0 ? mask : 0;
        // The reordering comes before every feature (map->add_gsub_pause (reorder_khmer) precedes them in
        // collect_features_khmer).
        Reorder(font, glyphs, MaskOf("pref", Pref), MaskOf("blwf", Blwf) | MaskOf("abvf", Abvf) | MaskOf("pstf", Pstf), MaskOf("cfar", Cfar));

        // "Testing suggests that Uniscribe does NOT pause between basic features", so these form one stage.
        layout.Apply(false, glyphs, layout.Lookups(false, script, [
            ("locl", Global | perSyllable), ("ccmp", Global | perSyllable), ("pref", Pref | perSyllable | manual),
            ("blwf", Blwf | perSyllable | manual), ("abvf", Abvf | perSyllable | manual), ("pstf", Pstf | perSyllable | manual),
            ("cfar", Cfar | perSyllable | manual)]));
        // The syllables are cleared; pres, abvs, blws, psts apply with the common features (clig on, liga off).
        var last = layout.Lookups(false, script, new[] { "pres", "abvs", "blws", "psts" }.Select(f => (f, Global | manual)).ToArray())
            .Concat(layout.Lookups(false, script, (ligatures ? new[] { "rlig", "calt", "clig", "rclt" } : ["rlig", "calt", "rclt"])
                .Select(f => (f, Global)).ToArray()))
            .GroupBy(l => l.Item1).Select(g => (g.Key, g.Aggregate(0u, (m, l) => m | l.Item2))).OrderBy(l => l.Key).ToList();
        layout.Apply(false, glyphs, last);

        for (int i = 0; i < glyphs.Count; i++) glyphs[i] = glyphs[i] with { XAdvance = font.GetAdvanceWidth(glyphs[i].Glyph) };
        TextShaper.ApplySpaceFallback(font, glyphs);
        if (layout.HasPositioning)
        {
            // ZERO_WIDTH_MARKS_NONE: Khmer marks keep their advance; mark and mkmk see ZWJ/ZWNJ as in HarfBuzz.
            const uint withJoiners = Global | OpenTypeLayout.ManualJoiners;
            layout.Apply(true, glyphs, layout.Lookups(true, script,
                [("abvm", Global), ("blwm", Global), ("mark", withJoiners), ("mkmk", withJoiners), ("curs", Global), ("dist", Global), ("kern", Global)]));
        }
        TextShaper.ZeroIgnorables(glyphs);
        if (layout.HasPositioning) TextShaper.ResolveAttachments(glyphs, rightToLeft: false);
        TextShaper.HideIgnorables(font, glyphs);
        IndicShaper.FormClusters(text, glyphs);
        return glyphs;
    }

    // ------------------------------------------------------------------ text preparation

    private static List<NormalizedChar> Normalize(TrueTypeFont font, string text)
    {
        var codePoints = new List<(int CodePoint, int Cluster)>(text.Length);
        for (int i = 0; i < text.Length; i += char.IsSurrogatePair(text, i) ? 2 : 1)
            codePoints.Add((char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i], i));
        return Normalizer.Normalize(font, codePoints, Normalizer.Shaper.Khmer);
    }

    private static List<GlyphInfo> Nominal(TrueTypeFont font, List<NormalizedChar> chars)
    {
        var glyphs = new List<GlyphInfo>(chars.Count);
        foreach (var (cp, cluster, text, hidden) in chars)
            glyphs.Add(new GlyphInfo
            {
                Glyph = Normalizer.NominalGlyph(font, cp), Cluster = cluster, Text = text, Mask = Global, XAdvance = 0,
                AttachTo = -1, CursiveTo = -1, CodePoint = cp, ShaperCategory = (byte)CategoryOf(cp),
                Ignorable = TextShaper.IsDefaultIgnorable(cp), Hidden = hidden, Guess = TextShaper.GuessClass(cp),
            });
        return glyphs;
    }

    private static KhmerCategory CategoryOf(int cp)
    {
        var starts = KhmerData.RangeStarts;
        if (cp >= starts[^1]) return KhmerCategory.X;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= cp) lo = mid; else hi = mid - 1;
        }
        return (KhmerCategory)KhmerData.RangeValues[lo];
    }

    // ------------------------------------------------------------------ syllables

    // hb-ot-shaper-khmer-machine.rl:
    //   c = (C | Ra | V);  cn = c.((ZWJ|ZWNJ)?.Robatic)?;  joiner = (ZWJ | ZWNJ);
    //   xgroup = (joiner*.Xgroup)*;  ygroup = Ygroup*;
    //   matra_group = VPre? xgroup VBlw? xgroup (joiner?.VAbv)? xgroup VPst?;
    //   syllable_tail = xgroup matra_group xgroup (H.c)? ygroup;
    //   broken_cluster = Robatic? (H.cn)* (H | syllable_tail);
    //   consonant_syllable = (cn|PLACEHOLDER|DOTTEDCIRCLE) broken_cluster;
    private static Pattern Cat(params KhmerCategory[] c) => S(c.Select(x => (byte)x).ToArray());

    private static readonly Pattern Joiner = Cat(KhmerCategory.ZWJ, KhmerCategory.ZWNJ);
    private static readonly Pattern Consonant = Cat(KhmerCategory.C, KhmerCategory.Ra, KhmerCategory.V);
    private static readonly Pattern Cn = Seq(Consonant, Opt(Seq(Opt(Joiner), Cat(KhmerCategory.Robatic))));
    private static readonly Pattern Xgroup = Star(Seq(Star(Joiner), Cat(KhmerCategory.Xgroup)));
    private static readonly Pattern Ygroup = Star(Cat(KhmerCategory.Ygroup));
    private static readonly Pattern MatraGroup = Seq(Opt(Cat(KhmerCategory.VPre)), Xgroup, Opt(Cat(KhmerCategory.VBlw)), Xgroup,
        Opt(Seq(Opt(Joiner), Cat(KhmerCategory.VAbv))), Xgroup, Opt(Cat(KhmerCategory.VPst)));
    private static readonly Pattern SyllableTail = Seq(Xgroup, MatraGroup, Xgroup, Opt(Seq(Cat(KhmerCategory.H), Consonant)), Ygroup);
    private static readonly Pattern BrokenCluster = Seq(Opt(Cat(KhmerCategory.Robatic)), Star(Seq(Cat(KhmerCategory.H), Cn)),
        Or(Cat(KhmerCategory.H), SyllableTail));
    private static readonly Pattern ConsonantSyllable = Seq(Or(Cn, Cat(KhmerCategory.Placeholder, KhmerCategory.DottedCircle)), BrokenCluster);

    private static readonly (Pattern Rule, SyllableType Type)[] Rules =
    [
        (ConsonantSyllable, SyllableType.Consonant),
        (BrokenCluster, SyllableType.Broken),
    ];

    /// <summary>Numbers the syllables (Syllable = serial * 16 + type): longest match, first rule on ties.</summary>
    private static void FindSyllables(List<GlyphInfo> glyphs)
    {
        var categories = glyphs.Select(g => g.ShaperCategory).ToArray();
        int serial = 1;
        for (int p = 0; p < categories.Length;)
        {
            int best = 0;
            var type = SyllableType.NonKhmer;
            foreach (var (rule, ruleType) in Rules)
            {
                int end = rule.Longest(categories, p);
                if (end - p > best) (best, type) = (end - p, ruleType);
            }
            if (best == 0) best = 1; // "other"
            for (int k = p; k < p + best; k++) glyphs[k] = glyphs[k] with { Syllable = serial * 16 + (int)type };
            serial++;
            p += best;
        }
    }

    // ------------------------------------------------------------------ reordering

    private static KhmerCategory Cat(in GlyphInfo g) => (KhmerCategory)g.ShaperCategory;

    private static void Reorder(TrueTypeFont font, List<GlyphInfo> glyphs, uint prefMask, uint postBaseMask, uint cfarMask)
    {
        DottedCircles.Insert(font, glyphs, (int)SyllableType.Broken, (byte)KhmerCategory.DottedCircle);
        var g = CollectionsMarshal.AsSpan(glyphs);
        foreach (var (start, end) in IndicShaper.Syllables(glyphs).ToList())
            if ((SyllableType)(g[start].Syllable & 0xF) != SyllableType.NonKhmer)
                ReorderConsonantSyllable(g, start, end, prefMask, postBaseMask, cfarMask);
    }

    /// <summary>
    /// hb-ot-shaper-khmer.cc reorder_consonant_syllable. From the Microsoft Khmer page: "Subscript Type 2 - The COENG
    /// + RO characters are reordered to immediately before the base glyph. Then the COENG + RO characters are
    /// assigned to have the 'pref' OpenType feature applied to them." The left matra piece (VPre) moves to the start.
    /// </summary>
    private static void ReorderConsonantSyllable(Span<GlyphInfo> g, int start, int end, uint prefMask, uint postBaseMask, uint cfarMask)
    {
        for (int i = start + 1; i < end; i++) g[i].Mask |= postBaseMask;

        int coengs = 0;
        for (int i = start + 1; i < end; i++)
        {
            if (Cat(g[i]) == KhmerCategory.H && coengs <= 2 && i + 1 < end)
            {
                coengs++;
                if (Cat(g[i + 1]) != KhmerCategory.Ra) continue;
                g[i].Mask |= prefMask;
                g[i + 1].Mask |= prefMask;
                // Move the Coeng,Ro sequence to the start.
                var t0 = g[i];
                var t1 = g[i + 1];
                g[start..i].CopyTo(g[(start + 2)..]);
                g[start] = t0;
                g[start + 1] = t1;
                // Mark the subsequent stuff with 'cfar': it tells U+1784,U+17D2,U+179A,U+17D2,U+1782 from
                // U+1784,U+17D2,U+1782,U+17D2,U+179A in MS Khmer fonts.
                if (cfarMask != 0)
                    for (int j = i + 2; j < end; j++) g[j].Mask |= cfarMask;
                coengs = 2; // Done.
            }
            else if (Cat(g[i]) == KhmerCategory.VPre)
            {
                // Reorder the left matra piece: move to the start.
                var t = g[i];
                g[start..i].CopyTo(g[(start + 1)..]);
                g[start] = t;
            }
        }
    }
}
