using System.Runtime.InteropServices;
using static DotOffice.Fonts.Shaping.SyllablePatterns;

namespace DotOffice.Fonts.Shaping;

/// <summary>
/// Shaping category of a Myanmar character: the numbers of the "export" lines of HarfBuzz's
/// hb-ot-shaper-myanmar-machine.rl (they are the values of <see cref="MyanmarData"/>).
/// </summary>
internal enum MyanmarCategory : byte
{
    X = 0, C = 1, IV = 2, DB = 3, H = 4, ZWNJ = 5, ZWJ = 6, SM = 8, A = 9, GB = 10, DottedCircle = 11, Ra = 15, CS = 18,
    VAbv = 20, VBlw = 21, VPre = 22, VPst = 23, As = 32, MH = 35, MR = 36, MW = 37, MY = 38, PT = 39, VS = 40, ML = 41,
    SMPst = 57,
}

/// <summary>
/// The Myanmar shaper, as HarfBuzz implements it (hb-ot-shaper-myanmar.cc and its syllable machine, MIT license; see
/// THIRD-PARTY-NOTICES.md). HarfBuzz uses it when the font has a GSUB script of its own for Myanmar ("mym2"), and
/// the default shaper when the font was designed for DFLT, latn or the older "mymr" tag. The rules come from
/// Microsoft's Myanmar script development page.
/// </summary>
internal static class MyanmarShaper
{
    private enum SyllableType { Consonant, Broken, NonMyanmar }

    // Feature masks: Global applies to every glyph.
    private const uint Global = 1;

    public static bool IsMyanmar(int cp) => cp is >= 0x1000 and <= 0x109F or >= 0xA9E0 and <= 0xA9FF or >= 0xAA60 and <= 0xAA7F;

    /// <summary>Whether HarfBuzz shapes Myanmar with this shaper for a font with this layout (else with the default one).</summary>
    public static bool Applies(OpenTypeLayout layout) =>
        layout.HasScript("mym2") || !(layout.HasScript("DFLT") || layout.HasScript("latn") || layout.HasScript("mymr"));

    public static List<GlyphInfo>? Shape(TrueTypeFont font, string text, string script, bool ligatures = true)
    {
        if (script != "mymr") return null;
        var layout = font.OpenTypeLayout;
        if (!Applies(layout)) return null;
        const string tag = "mym2";

        var glyphs = Nominal(font, Normalize(font, text));
        FindSyllables(glyphs);

        uint perSyllable = OpenTypeLayout.PerSyllable, manualZwj = OpenTypeLayout.ManualZwj;
        void Stage(params (string Feature, uint Mask)[] features) => layout.Apply(false, glyphs, layout.Lookups(false, tag, features));

        // The syllables come first, then locl and ccmp, then the reordering (collect_features_myanmar).
        Stage(("locl", Global | perSyllable), ("ccmp", Global | perSyllable));
        Reorder(font, glyphs);
        // The basic features apply one at a time, each its own stage.
        foreach (var feature in new[] { "rphf", "pref", "blwf", "pstf" })
            Stage((feature, Global | perSyllable | manualZwj));
        var last = layout.Lookups(false, tag, new[] { "pres", "abvs", "blws", "psts" }.Select(f => (f, Global | manualZwj)).ToArray())
            .Concat(layout.Lookups(false, tag, (ligatures ? new[] { "rlig", "calt", "clig", "liga", "rclt" } : ["rlig", "calt", "rclt"])
                .Select(f => (f, Global)).ToArray()))
            .GroupBy(l => l.Item1).Select(g => (g.Key, g.Aggregate(0u, (m, l) => m | l.Item2))).OrderBy(l => l.Key).ToList();
        layout.Apply(false, glyphs, last);

        for (int i = 0; i < glyphs.Count; i++) glyphs[i] = glyphs[i] with { XAdvance = font.GetAdvanceWidth(glyphs[i].Glyph) };
        TextShaper.ApplySpaceFallback(font, glyphs);
        // ZERO_WIDTH_MARKS_BY_GDEF_EARLY: marks lose their advance before GPOS positions them.
        FallbackPositioning.ZeroMarkWidths(layout, glyphs, rightToLeft: false, adjustOffsets: !layout.HasPositioning);
        if (layout.HasPositioning)
        {
            const uint withJoiners = Global | OpenTypeLayout.ManualJoiners;
            layout.Apply(true, glyphs, layout.Lookups(true, tag,
                [("abvm", Global), ("blwm", Global), ("mark", withJoiners), ("mkmk", withJoiners), ("curs", Global), ("dist", Global), ("kern", Global)]));
        }
        TextShaper.ZeroIgnorables(glyphs);
        if (layout.HasPositioning) TextShaper.ResolveAttachments(glyphs, rightToLeft: false);
        TextShaper.HideIgnorables(font, glyphs);
        IndicShaper.AssignSyllableText(text, glyphs);
        return glyphs;
    }

    // ------------------------------------------------------------------ text preparation

    private static List<NormalizedChar> Normalize(TrueTypeFont font, string text)
    {
        var codePoints = new List<(int CodePoint, int Cluster)>(text.Length);
        for (int i = 0; i < text.Length; i += char.IsSurrogatePair(text, i) ? 2 : 1)
            codePoints.Add((char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i], i));
        return Normalizer.Normalize(font, codePoints, Normalizer.Shaper.Myanmar);
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

    private static MyanmarCategory CategoryOf(int cp)
    {
        var starts = MyanmarData.RangeStarts;
        if (cp >= starts[^1]) return MyanmarCategory.X;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= cp) lo = mid; else hi = mid - 1;
        }
        return (MyanmarCategory)MyanmarData.RangeValues[lo];
    }

    // ------------------------------------------------------------------ syllables

    // hb-ot-shaper-myanmar-machine.rl:
    //   j = ZWJ|ZWNJ;  k = (Ra As H);  sm = SM | SMPst;  c = C|Ra;
    //   medial_group = MY? As? MR? ((MW MH? ML? | MH ML? | ML) As?)?;
    //   main_vowel_group = (VPre.VS?)* VAbv* VBlw* A* (DB As?)?;
    //   post_vowel_group = VPst MH? ML? As* VAbv* A* (DB As?)?;
    //   tone_group = sm | PT A* DB? As?;
    //   complex_syllable_tail = As* medial_group main_vowel_group post_vowel_group* tone_group* j?;
    //   syllable_tail = (H (c|IV).VS?)* (H | complex_syllable_tail);
    //   consonant_syllable = (k|CS)? (c|IV|GB|DOTTEDCIRCLE).VS? syllable_tail;
    //   broken_cluster = k? VS? syllable_tail;
    // and the scanner: consonant_syllable; j | SMPst; broken_cluster; any.
    private static Pattern Cat(params MyanmarCategory[] c) => S(c.Select(x => (byte)x).ToArray());
    private static Pattern One(MyanmarCategory c) => Cat(c);
    private static Pattern Plus(Pattern p) => Seq(p, Star(p));

    private static readonly Pattern Joiner = Cat(MyanmarCategory.ZWJ, MyanmarCategory.ZWNJ);
    private static readonly Pattern Kinzi = Seq(One(MyanmarCategory.Ra), One(MyanmarCategory.As), One(MyanmarCategory.H));
    private static readonly Pattern Sm = Cat(MyanmarCategory.SM, MyanmarCategory.SMPst);
    private static readonly Pattern Consonant = Cat(MyanmarCategory.C, MyanmarCategory.Ra);
    private static readonly Pattern Vs = One(MyanmarCategory.VS);
    private static readonly Pattern DbAs = Opt(Seq(One(MyanmarCategory.DB), Opt(One(MyanmarCategory.As))));

    private static readonly Pattern MedialGroup = Seq(Opt(One(MyanmarCategory.MY)), Opt(One(MyanmarCategory.As)), Opt(One(MyanmarCategory.MR)),
        Opt(Seq(Or(
            Seq(One(MyanmarCategory.MW), Opt(One(MyanmarCategory.MH)), Opt(One(MyanmarCategory.ML))),
            Seq(One(MyanmarCategory.MH), Opt(One(MyanmarCategory.ML))),
            One(MyanmarCategory.ML)), Opt(One(MyanmarCategory.As)))));
    private static readonly Pattern MainVowelGroup = Seq(Star(Seq(One(MyanmarCategory.VPre), Opt(Vs))), Star(One(MyanmarCategory.VAbv)),
        Star(One(MyanmarCategory.VBlw)), Star(One(MyanmarCategory.A)), DbAs);
    private static readonly Pattern PostVowelGroup = Seq(One(MyanmarCategory.VPst), Opt(One(MyanmarCategory.MH)), Opt(One(MyanmarCategory.ML)),
        Star(One(MyanmarCategory.As)), Star(One(MyanmarCategory.VAbv)), Star(One(MyanmarCategory.A)), DbAs);
    private static readonly Pattern ToneGroup = Or(Sm, Seq(One(MyanmarCategory.PT), Star(One(MyanmarCategory.A)), Opt(One(MyanmarCategory.DB)),
        Opt(One(MyanmarCategory.As))));
    private static readonly Pattern ComplexSyllableTail = Seq(Star(One(MyanmarCategory.As)), MedialGroup, MainVowelGroup,
        Star(PostVowelGroup), Star(ToneGroup), Opt(Joiner));
    private static readonly Pattern SyllableTail = Seq(
        Star(Seq(One(MyanmarCategory.H), Or(Consonant, One(MyanmarCategory.IV)), Opt(Vs))),
        Or(One(MyanmarCategory.H), ComplexSyllableTail));
    private static readonly Pattern ConsonantSyllable = Seq(Opt(Or(Kinzi, One(MyanmarCategory.CS))),
        Or(Consonant, Cat(MyanmarCategory.IV, MyanmarCategory.GB, MyanmarCategory.DottedCircle)), Opt(Vs), SyllableTail);
    private static readonly Pattern BrokenCluster = Seq(Opt(Kinzi), Opt(Vs), SyllableTail);

    private static readonly (Pattern Rule, SyllableType Type)[] Rules =
    [
        (ConsonantSyllable, SyllableType.Consonant),
        (Or(Joiner, One(MyanmarCategory.SMPst)), SyllableType.NonMyanmar),
        (BrokenCluster, SyllableType.Broken),
    ];

    private static void FindSyllables(List<GlyphInfo> glyphs)
    {
        var categories = glyphs.Select(g => g.ShaperCategory).ToArray();
        int serial = 1;
        for (int p = 0; p < categories.Length;)
        {
            int best = 0;
            var type = SyllableType.NonMyanmar;
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

    private static MyanmarCategory Cat(in GlyphInfo g) => (MyanmarCategory)g.ShaperCategory;
    private static IndicPosition Pos(in GlyphInfo g) => (IndicPosition)g.ShaperPosition;
    private static void SetPos(ref GlyphInfo g, IndicPosition p) => g.ShaperPosition = (byte)p;

    /// <summary>CONSONANT_FLAGS_MYANMAR: vowels and placeholders count as consonants; a ligated glyph is nothing.</summary>
    private static bool IsConsonant(in GlyphInfo g) => (g.Flags & GlyphFlags.Ligated) == 0
        && Cat(g) is MyanmarCategory.C or MyanmarCategory.CS or MyanmarCategory.Ra or MyanmarCategory.IV or MyanmarCategory.GB
            or MyanmarCategory.DottedCircle;

    private static void Reorder(TrueTypeFont font, List<GlyphInfo> glyphs)
    {
        DottedCircles.Insert(font, glyphs, (int)SyllableType.Broken, (byte)MyanmarCategory.DottedCircle);
        var g = CollectionsMarshal.AsSpan(glyphs);
        foreach (var (start, end) in IndicShaper.Syllables(glyphs).ToList())
            if ((SyllableType)(g[start].Syllable & 0xF) != SyllableType.NonMyanmar)
                ReorderConsonantSyllable(g, start, end);
    }

    /// <summary>
    /// hb-ot-shaper-myanmar.cc initial_reordering_consonant_syllable. Medial ra and the kinzi go before the base, the
    /// left vowel sign to the very start (several of them in reverse order); the rest keeps its order, below-base signs
    /// first, then those after the subjoined ones. The glyphs are sorted (stably) by these positions.
    /// </summary>
    private static void ReorderConsonantSyllable(Span<GlyphInfo> g, int start, int end)
    {
        int baseIndex = end;
        bool hasReph = false;
        int limit = start;
        if (start + 3 <= end && Cat(g[start]) == MyanmarCategory.Ra && Cat(g[start + 1]) == MyanmarCategory.As
            && Cat(g[start + 2]) == MyanmarCategory.H)
        {
            limit += 3;
            baseIndex = start;
            hasReph = true;
        }
        if (!hasReph) baseIndex = limit;
        for (int i = limit; i < end; i++)
            if (IsConsonant(g[i]))
            {
                baseIndex = i;
                break;
            }

        {
            int i = start;
            for (; i < start + (hasReph ? 3 : 0); i++) SetPos(ref g[i], IndicPosition.AfterMain);
            for (; i < baseIndex; i++) SetPos(ref g[i], IndicPosition.PreC);
            if (i < end)
            {
                SetPos(ref g[i], IndicPosition.BaseC);
                i++;
            }
            var pos = IndicPosition.AfterMain;
            for (; i < end; i++)
            {
                var cat = Cat(g[i]);
                if (cat == MyanmarCategory.MR) { SetPos(ref g[i], IndicPosition.PreC); continue; } // pre-base reordering
                if (cat == MyanmarCategory.VPre) { SetPos(ref g[i], IndicPosition.PreM); continue; } // left matra
                if (cat == MyanmarCategory.VS) { SetPos(ref g[i], Pos(g[i - 1])); continue; }
                if (pos == IndicPosition.AfterMain && cat == MyanmarCategory.VBlw)
                {
                    pos = IndicPosition.BelowC;
                    SetPos(ref g[i], pos);
                    continue;
                }
                if (pos == IndicPosition.BelowC && cat == MyanmarCategory.A) { SetPos(ref g[i], IndicPosition.BeforeSub); continue; }
                if (pos == IndicPosition.BelowC && cat == MyanmarCategory.VBlw) { SetPos(ref g[i], pos); continue; }
                if (pos == IndicPosition.BelowC && cat != MyanmarCategory.A)
                {
                    pos = IndicPosition.AfterSub;
                    SetPos(ref g[i], pos);
                    continue;
                }
                SetPos(ref g[i], pos);
            }
        }

        // Sit tight, rock 'n roll! (a stable sort by position)
        var sorted = g[start..end].ToArray().OrderBy(x => x.ShaperPosition).ToArray();
        sorted.CopyTo(g[start..end]);

        // Flip the left-matra sequence (HarfBuzz issue 3863).
        int firstLeft = end, lastLeft = end;
        for (int i = start; i < end; i++)
            if (Pos(g[i]) == IndicPosition.PreM)
            {
                if (firstLeft == end) firstLeft = i;
                lastLeft = i;
            }
        if (firstLeft < lastLeft)
        {
            g[firstLeft..(lastLeft + 1)].Reverse();
            // Reverse back each left vowel sign with the variation selector that follows it.
            int i = firstLeft;
            for (int j = i; j <= lastLeft; j++)
                if (Cat(g[j]) == MyanmarCategory.VPre)
                {
                    g[i..(j + 1)].Reverse();
                    i = j + 1;
                }
        }
    }
}
