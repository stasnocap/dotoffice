using System.Runtime.InteropServices;
using static DotOffice.Fonts.Shaping.SyllablePatterns;
using static DotOffice.Fonts.Shaping.UseCategory;

namespace DotOffice.Fonts.Shaping;

/// <summary>
/// Shaping category of a character in the Universal Shaping Engine: the numbers are those of the "export" lines of
/// HarfBuzz's hb-ot-shaper-use-machine.rl (they are the values of <see cref="UseData"/>).
/// </summary>
internal enum UseCategory : byte
{
    O = 0, B = 1, N = 4, GB = 5, CGJ = 6, SUB = 11, H = 12, HN = 13, ZWNJ = 14, WJ = 16, R = 18,
    VPre = 22, VMPre = 23, FAbv = 24, FBlw = 25, FPst = 26, MAbv = 27, MBlw = 28, MPst = 29, MPre = 30,
    CMAbv = 31, CMBlw = 32, VAbv = 33, VBlw = 34, VPst = 35, VMAbv = 37, VMBlw = 38, VMPst = 39,
    SMAbv = 41, SMBlw = 42, CS = 43, IS = 44, FMAbv = 45, FMBlw = 46, FMPst = 47, Sk = 48, G = 49, J = 50,
    SB = 51, SE = 52, HVM = 53, HM = 54, HR = 55, RK = 56,
}

/// <summary>
/// The Universal Shaping Engine (https://learn.microsoft.com/en-us/typography/script-development/use), as HarfBuzz
/// implements it in hb-ot-shaper-use.cc (MIT license; see THIRD-PARTY-NOTICES.md). HarfBuzz, and so LibreOffice,
/// uses it for Sinhala and Tibetan when the font has a GSUB script of its own for them (otherwise the default
/// shaper applies). It shares the syllable scanner machinery, normalization and feature application with the
/// Indic shaper.
/// </summary>
internal static class UseShaper
{
    private enum SyllableType { Virama, Sakot, Standard, NumberJoiner, Numeral, Symbol, Hieroglyph, Broken, NonCluster }

    // Feature masks: Global applies to every glyph; the others to the glyphs the shaper picks.
    private const uint Global = 1, Rphf = 2, Isol = 4, Init = 8, Medi = 16, Fina = 32;

    /// <summary>True for the characters of the scripts shaped here (Sinhala and Tibetan).</summary>
    public static bool IsUse(int cp) => cp is >= 0x0D80 and <= 0x0DFF or >= 0x0F00 and <= 0x0FFF;

    /// <summary>
    /// Shapes <paramref name="text"/>, a run of the script with OpenType tag <paramref name="script"/>; returns null
    /// when that script does not go to USE or the font has no GSUB script of its own for it (HarfBuzz then uses its
    /// default shaper).
    /// </summary>
    public static List<GlyphInfo>? Shape(TrueTypeFont font, string text, string script, bool ligatures = true)
    {
        if (script is not ("sinh" or "tibt")) return null;
        var layout = font.OpenTypeLayout;
        if (!layout.HasScript(script)) return null;

        var glyphs = Nominal(font, Normalize(font, text, script));
        FindSyllables(glyphs);

        List<(int, uint)> Lookups(params (string Feature, uint Mask)[] features) => layout.Lookups(false, script, features);
        uint perSyllable = OpenTypeLayout.PerSyllable, manual = OpenTypeLayout.ManualJoiners;
        void Stage(params (string Feature, uint Mask)[] features) => layout.Apply(false, glyphs, Lookups(features));

        uint rphfMask = Lookups(("rphf", Rphf)).Count > 0 ? Rphf : 0;
        SetupRphfMask(glyphs, rphfMask);
        SetupTopographicalMasks(glyphs, Lookups(("isol", Isol)).Count > 0, Lookups(("init", Init)).Count > 0,
            Lookups(("medi", Medi)).Count > 0, Lookups(("fina", Fina)).Count > 0);

        // "Default glyph pre-processing group" (hb-ot-shaper-use.cc collect_features_use).
        Stage(("locl", Global | perSyllable), ("ccmp", Global | perSyllable), ("nukt", Global | perSyllable),
            ("akhn", Global | perSyllable | manual));
        // "Reordering group": rphf, then pref, each followed by a pause that records what formed.
        ClearSubstituted(glyphs);
        Stage(("rphf", Rphf | perSyllable | manual));
        RecordRphf(glyphs, rphfMask);
        ClearSubstituted(glyphs);
        Stage(("pref", Global | perSyllable | manual));
        RecordPref(glyphs);
        // "Orthographic unit shaping group".
        Stage(new[] { "rkrf", "abvf", "blwf", "half", "pstf", "vatu", "cjct" }.Select(f => (f, Global | perSyllable | manual)).ToArray());
        Reorder(font, glyphs);
        // "Topographical features", then "Standard typographic presentation" with the common and horizontal ones.
        Stage(("isol", Isol), ("init", Init), ("medi", Medi), ("fina", Fina));
        var last = Lookups(new[] { "abvs", "blws", "haln", "pres", "psts" }.Select(f => (f, Global | manual)).ToArray())
            .Concat(Lookups((ligatures ? new[] { "rlig", "calt", "clig", "liga", "rclt" } : ["rlig", "calt", "rclt"])
                .Select(f => (f, Global)).ToArray()))
            .GroupBy(l => l.Item1).Select(g => (g.Key, g.Aggregate(0u, (m, l) => m | l.Item2))).OrderBy(l => l.Key).ToList();
        layout.Apply(false, glyphs, last);

        for (int i = 0; i < glyphs.Count; i++) glyphs[i] = glyphs[i] with { XAdvance = font.GetAdvanceWidth(glyphs[i].Glyph) };
        TextShaper.ApplySpaceFallback(font, glyphs);
        // ZERO_WIDTH_MARKS_BY_GDEF_EARLY: marks lose their advance before GPOS positions them (and, without GPOS,
        // hang over the previous glyph); USE has no fallback mark positioning.
        FallbackPositioning.ZeroMarkWidths(layout, glyphs, rightToLeft: false, adjustOffsets: !layout.HasPositioning);
        if (layout.HasPositioning)
        {
            const uint withJoiners = Global | OpenTypeLayout.ManualJoiners;
            layout.Apply(true, glyphs, layout.Lookups(true, script,
                [("abvm", Global), ("blwm", Global), ("mark", withJoiners), ("mkmk", withJoiners), ("curs", Global), ("dist", Global), ("kern", Global)]));
            TextShaper.ZeroIgnorables(glyphs);
            TextShaper.ResolveAttachments(glyphs, rightToLeft: false);
        }
        else TextShaper.ZeroIgnorables(glyphs);
        TextShaper.HideIgnorables(font, glyphs);
        IndicShaper.AssignSyllableText(text, glyphs);
        return glyphs;
    }

    // ------------------------------------------------------------------ text preparation

    private static List<NormalizedChar> Normalize(TrueTypeFont font, string text, string script)
    {
        var codePoints = new List<(int CodePoint, int Cluster)>(text.Length);
        for (int i = 0; i < text.Length; i += char.IsSurrogatePair(text, i) ? 2 : 1)
            codePoints.Add((char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i], i));
        if (script == "sinh") codePoints = ApplySinhalaVowelConstraints(codePoints);
        return Normalizer.Normalize(font, codePoints, Normalizer.Shaper.Use);
    }

    /// <summary>
    /// A vowel followed by a sign that makes it look like another vowel gets a dotted circle in between (HarfBuzz's
    /// _hb_preprocess_text_vowel_constraints, from the USE script development spec; the Sinhala case).
    /// </summary>
    private static List<(int CodePoint, int Cluster)> ApplySinhalaVowelConstraints(List<(int CodePoint, int Cluster)> cps)
    {
        var result = new List<(int, int)>(cps.Count);
        for (int i = 0; i < cps.Count; i++)
        {
            result.Add(cps[i]);
            if (i + 1 >= cps.Count) continue;
            int next = cps[i + 1].CodePoint;
            bool matched = cps[i].CodePoint switch
            {
                0x0D85 => next is 0x0DCF or 0x0DD0 or 0x0DD1,
                0x0D8B or 0x0D8F or 0x0D94 => next == 0x0DDF,
                0x0D8D => next == 0x0DD8,
                0x0D91 => next is 0x0DCA or 0x0DD9 or 0x0DDA or 0x0DDC or 0x0DDD or 0x0DDE,
                _ => false,
            };
            if (matched) result.Add((0x25CC, cps[i + 1].Cluster));
        }
        return result;
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

    private static UseCategory CategoryOf(int cp)
    {
        var starts = UseData.RangeStarts;
        if (cp >= starts[^1]) return O;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= cp) lo = mid; else hi = mid - 1;
        }
        return (UseCategory)UseData.RangeValues[lo];
    }

    // ------------------------------------------------------------------ syllables

    // hb-ot-shaper-use-machine.rl, "main := |* ... *|": the grammar over USE categories.
    private static Pattern Cat(params UseCategory[] c) => S(c.Select(x => (byte)x).ToArray());
    private static Pattern Rep(Pattern p, int min) => min == 0 ? Star(p) : Seq(p, Star(p));
    private static Pattern Plus(Pattern p) => Rep(p, 1);
    private static Pattern Star(UseCategory c) => SyllablePatterns.Star(Cat(c));
    private static Pattern Star(Pattern p) => SyllablePatterns.Star(p);

    private static readonly Pattern HalantLike = Cat(H, HVM, IS, Sk);
    private static readonly Pattern ConsonantModifiers = Seq(Star(CMAbv), Star(CMBlw),
        Star(Seq(Or(Seq(HalantLike, Cat(B)), Cat(SUB)), Star(CMAbv), Star(CMBlw))));
    private static readonly Pattern MedialConsonants = Seq(Opt(Cat(MPre)), Opt(Cat(MAbv)), Opt(Cat(MBlw)), Opt(Cat(MPst)));
    private static readonly Pattern DependentVowels = Or(Seq(Star(VPre), Star(VAbv), Star(VBlw), Star(VPst)), Cat(H));
    private static readonly Pattern VowelModifiers = Seq(Opt(Cat(HVM)), Star(VMPre), Star(VMAbv), Star(VMBlw), Star(VMPst));
    private static readonly Pattern FinalConsonants = Seq(Star(FAbv), Star(FBlw), Star(FPst));
    private static readonly Pattern FinalModifiers = Or(Seq(Star(FMAbv), Star(FMBlw)), Opt(Cat(FMPst)));
    private static readonly Pattern ComplexSyllableStart = Seq(Opt(Cat(R, CS)), Cat(B, GB));
    private static readonly Pattern ComplexSyllableMiddle =
        Seq(ConsonantModifiers, MedialConsonants, DependentVowels, VowelModifiers, Star(Seq(Cat(Sk), Cat(B))));
    private static readonly Pattern ComplexSyllableTail = Seq(ComplexSyllableMiddle, FinalConsonants, FinalModifiers);
    private static readonly Pattern NumberJoinerTerminatedTail = Seq(Star(Seq(Cat(HN), Cat(N))), Cat(HN));
    private static readonly Pattern NumeralTail = Plus(Seq(Cat(HN), Cat(N)));
    private static readonly Pattern SymbolTail = Or(Seq(Plus(Cat(SMAbv)), Star(SMBlw)), Plus(Cat(SMBlw)));
    private static readonly Pattern ViramaTerminatedTail = Seq(ConsonantModifiers, Cat(IS, RK));
    private static readonly Pattern SakotTerminatedTail = Seq(ComplexSyllableMiddle, Cat(Sk));
    private static readonly Pattern AnyTail = Or(ComplexSyllableTail, SakotTerminatedTail, SymbolTail, ViramaTerminatedTail);
    private static readonly Pattern HieroglyphUnit = Seq(Star(SB), Cat(G), Opt(Cat(HR)), Opt(Cat(HM)), Star(SE));
    // After a hieroglyph joiner the next unit is G HR? HM? SE* (the leading SB* is consumed beside the joiner).
    private static readonly Pattern HieroglyphUnitNoLead = Seq(Cat(G), Opt(Cat(HR)), Opt(Cat(HM)), Star(SE));
    private static readonly Pattern Zwnj = Opt(Cat(ZWNJ));

    /// <summary>The scanner's rules in priority order (an earlier rule wins a tie).</summary>
    private static readonly (Pattern Rule, SyllableType Type)[] Rules =
    [
        (Seq(ComplexSyllableStart, ViramaTerminatedTail, Zwnj), SyllableType.Virama),
        (Seq(ComplexSyllableStart, SakotTerminatedTail, Zwnj), SyllableType.Sakot),
        (Seq(ComplexSyllableStart, ComplexSyllableTail, Zwnj), SyllableType.Standard),
        (Seq(Cat(N), NumberJoinerTerminatedTail, Zwnj), SyllableType.NumberJoiner),
        (Seq(Cat(N), Opt(NumeralTail), Zwnj), SyllableType.Numeral),
        (Seq(Cat(O, GB, SB), Opt(AnyTail), Zwnj), SyllableType.Symbol),
        (Seq(HieroglyphUnit, Star(Seq(Cat(J), Star(SB), Opt(HieroglyphUnitNoLead))), Zwnj), SyllableType.Hieroglyph),
        (Cat(FMPst), SyllableType.NonCluster),
        (Seq(Opt(Cat(R)), Or(AnyTail, NumberJoinerTerminatedTail, NumeralTail), Zwnj), SyllableType.Broken),
    ];

    /// <summary>
    /// Numbers the syllables (Syllable = serial * 16 + type). Like HarfBuzz, the scanner does not see CGJ-like
    /// characters (CGJ, variation selectors, ZWJ) nor a ZWNJ that is followed by a combining mark; they belong to the
    /// syllable they are inside of or follow.
    /// </summary>
    private static void FindSyllables(List<GlyphInfo> glyphs)
    {
        var kept = new List<int>(glyphs.Count);
        for (int i = 0; i < glyphs.Count; i++)
        {
            var category = (UseCategory)glyphs[i].ShaperCategory;
            if (category == CGJ) continue;
            if (category == ZWNJ)
            {
                int j = i + 1;
                while (j < glyphs.Count && (UseCategory)glyphs[j].ShaperCategory == CGJ) j++;
                if (j < glyphs.Count && TextShaper.IsMark(glyphs[j].CodePoint)) continue;
            }
            kept.Add(i);
        }
        var categories = kept.Select(i => glyphs[i].ShaperCategory).ToArray();
        int serial = 1;
        if (kept.Count > 0 && kept[0] > 0) // leading characters the scanner skips
        {
            for (int k = 0; k < kept[0]; k++) glyphs[k] = glyphs[k] with { Syllable = serial * 16 + (int)SyllableType.NonCluster };
            serial++;
        }
        for (int p = 0; p < categories.Length;)
        {
            int best = 0;
            var type = SyllableType.NonCluster;
            foreach (var (rule, ruleType) in Rules)
            {
                int end = rule.Longest(categories, p);
                if (end - p > best) (best, type) = (end - p, ruleType);
            }
            if (best == 0) best = 1; // "other"
            int from = kept[p], to = p + best < kept.Count ? kept[p + best] : glyphs.Count;
            for (int k = from; k < to; k++) glyphs[k] = glyphs[k] with { Syllable = serial * 16 + (int)type };
            serial++;
            p += best;
        }
    }

    private static SyllableType TypeOf(in GlyphInfo g) => (SyllableType)(g.Syllable & 0xF);
    private static UseCategory Cat(in GlyphInfo g) => (UseCategory)g.ShaperCategory;

    // ------------------------------------------------------------------ masks and pauses

    /// <summary>The first three glyphs of a syllable (only the repha for a syllable that starts with one) may form a reph.</summary>
    private static void SetupRphfMask(List<GlyphInfo> glyphs, uint mask)
    {
        if (mask == 0) return;
        foreach (var (start, end) in IndicShaper.Syllables(glyphs))
        {
            int limit = Cat(glyphs[start]) == R ? 1 : Math.Min(3, end - start);
            for (int i = start; i < start + limit; i++) glyphs[i] = glyphs[i] with { Mask = glyphs[i].Mask | mask };
        }
    }

    /// <summary>
    /// Topographical forms (isol, init, medi, fina) are assigned per syllable: a syllable after one that is
    /// isolated or final joins it (hb-ot-shaper-use.cc setup_topographical_masks).
    /// </summary>
    private static void SetupTopographicalMasks(List<GlyphInfo> glyphs, bool hasIsol, bool hasInit, bool hasMedi, bool hasFina)
    {
        uint isol = hasIsol ? Isol : 0, init = hasInit ? Init : 0, medi = hasMedi ? Medi : 0, fina = hasFina ? Fina : 0;
        uint all = isol | init | medi | fina;
        if (all == 0) return;
        uint other = ~all;
        // Indexed like the form masks of HarfBuzz: ISOL, INIT, MEDI, FINA.
        uint[] masks = [isol, init, medi, fina];
        const int IsolForm = 0, InitForm = 1, MediForm = 2, FinaForm = 3, NoForm = 4;
        int lastForm = NoForm, lastStart = 0;
        void SetForm(int from, int to, int form)
        {
            for (int i = from; i < to; i++) glyphs[i] = glyphs[i] with { Mask = (glyphs[i].Mask & other) | masks[form] };
        }
        foreach (var (start, end) in IndicShaper.Syllables(glyphs).ToList())
        {
            if (TypeOf(glyphs[start]) is SyllableType.Hieroglyph or SyllableType.NonCluster)
            {
                lastForm = NoForm; // these do not join
            }
            else
            {
                bool join = lastForm is FinaForm or IsolForm;
                if (join)
                {
                    lastForm = lastForm == FinaForm ? MediForm : InitForm;
                    SetForm(lastStart, start, lastForm);
                }
                lastForm = join ? FinaForm : IsolForm;
                SetForm(start, end, lastForm);
            }
            lastStart = start;
        }
    }

    private static void ClearSubstituted(List<GlyphInfo> glyphs)
    {
        for (int i = 0; i < glyphs.Count; i++) glyphs[i] = glyphs[i] with { Flags = glyphs[i].Flags & ~GlyphFlags.Substituted };
    }

    /// <summary>A substituted glyph in the reph position is a repha (record_rphf_use).</summary>
    private static void RecordRphf(List<GlyphInfo> glyphs, uint mask)
    {
        if (mask == 0) return;
        foreach (var (start, end) in IndicShaper.Syllables(glyphs).ToList())
            for (int i = start; i < end && (glyphs[i].Mask & mask) != 0; i++)
                if ((glyphs[i].Flags & GlyphFlags.Substituted) != 0)
                {
                    glyphs[i] = glyphs[i] with { ShaperCategory = (byte)R };
                    break;
                }
    }

    /// <summary>A substituted pref glyph behaves like a pre-base vowel sign (record_pref_use).</summary>
    private static void RecordPref(List<GlyphInfo> glyphs)
    {
        foreach (var (start, end) in IndicShaper.Syllables(glyphs).ToList())
            for (int i = start; i < end; i++)
                if ((glyphs[i].Flags & GlyphFlags.Substituted) != 0)
                {
                    glyphs[i] = glyphs[i] with { ShaperCategory = (byte)VPre };
                    break;
                }
    }

    // ------------------------------------------------------------------ reordering

    private static bool IsHalant(in GlyphInfo g) =>
        Cat(g) is H or HVM or IS && (g.Flags & GlyphFlags.Ligated) == 0;

    private static bool IsPostBase(UseCategory c) => c is FAbv or FBlw or FPst or FMAbv or FMBlw or FMPst or MAbv or MBlw
        or MPst or MPre or VAbv or VBlw or VPst or VPre or VMAbv or VMBlw or VMPst or VMPre;

    private static void Reorder(TrueTypeFont font, List<GlyphInfo> glyphs)
    {
        InsertDottedCircles(font, glyphs);
        var g = CollectionsMarshal.AsSpan(glyphs);
        foreach (var (start, end) in IndicShaper.Syllables(glyphs).ToList())
            ReorderSyllable(g, start, end);
    }

    /// <summary>A broken cluster gets a dotted circle (after a leading repha) for its marks to sit on (hb_syllabic_insert_dotted_circles).</summary>
    private static void InsertDottedCircles(TrueTypeFont font, List<GlyphInfo> glyphs)
    {
        int circle = font.GetGlyphId(0x25CC);
        if (circle == 0 || !glyphs.Exists(g => TypeOf(g) == SyllableType.Broken)) return;
        int last = 0;
        for (int i = 0; i < glyphs.Count; i++)
        {
            var g = glyphs[i];
            if (g.Syllable == last || TypeOf(g) != SyllableType.Broken) continue;
            last = g.Syllable;
            int at = i;
            while (at < glyphs.Count && glyphs[at].Syllable == last && Cat(glyphs[at]) == R) at++;
            glyphs.Insert(at, g with
            {
                Glyph = (ushort)circle, Text = "", CodePoint = 0x25CC, ShaperCategory = (byte)B, Flags = GlyphFlags.None,
                Ignorable = false, Guess = 1, LigatureId = 0, LigatureComponent = 0,
            });
        }
    }

    /// <summary>hb-ot-shaper-use.cc reorder_syllable_use: the repha goes after the base, pre-base vowels before it.</summary>
    private static void ReorderSyllable(Span<GlyphInfo> g, int start, int end)
    {
        if (TypeOf(g[start]) is not (SyllableType.Virama or SyllableType.Sakot or SyllableType.Standard or SyllableType.Symbol
            or SyllableType.Broken)) return;

        // Move things forward: a repha goes towards the end, but before the first post-base glyph.
        if (Cat(g[start]) == R && end - start > 1)
            for (int i = start + 1; i < end; i++)
            {
                bool postBase = IsPostBase(Cat(g[i])) || IsHalant(g[i]);
                if (postBase || i == end - 1)
                {
                    if (postBase) i--;
                    Move(g, start, i);
                    break;
                }
            }

        // Move things back: pre-base vowels go to the beginning, or after the last halant before them.
        int j = start;
        for (int i = start; i < end; i++)
        {
            if (IsHalant(g[i])) j = i + 1;
            else if (Cat(g[i]) is VPre or VMPre && g[i].LigatureComponent == 0 && j < i) // only the first component of a multiple substitution
            {
                Move(g, i, j);
            }
        }
    }

    /// <summary>Moves the glyph at <paramref name="from"/> to <paramref name="to"/>, shifting those between.</summary>
    private static void Move(Span<GlyphInfo> g, int from, int to)
    {
        var moved = g[from];
        if (from < to) g[(from + 1)..(to + 1)].CopyTo(g[from..]);
        else g[to..from].CopyTo(g[(to + 1)..]);
        g[to] = moved;
    }
}
