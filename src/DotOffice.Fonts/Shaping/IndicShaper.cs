using System.Globalization;
using System.Runtime.InteropServices;

namespace DotOffice.Fonts.Shaping;

/// <summary>Shaping category of a character (HarfBuzz's indic_syllable_machine categories; VD is the same as A).</summary>
internal enum IndicCategory : byte
{
    X, C, V, N, H, ZWNJ, ZWJ, M, SM, A, Placeholder, DottedCircle, RS, MPst, Repha, Ra, CM, Symbol, CS, SMPst,
}

/// <summary>Visual position of a character in its syllable, from left to right (HarfBuzz's ot_position_t).</summary>
internal enum IndicPosition : byte
{
    Start, RaToBecomeReph, PreM, PreC, BaseC, AfterMain, AboveC, BeforeSub, BelowC, AfterSub, BeforePost, PostC,
    AfterPost, Smvd, End,
}

/// <summary>
/// Shaper for the nine Indic scripts (Devanagari, Bengali, Gurmukhi, Gujarati, Oriya, Tamil, Telugu, Kannada,
/// Malayalam): syllables, reordering and the Indic GSUB feature stages. A port of HarfBuzz's Indic shaper
/// (hb-ot-shaper-indic.cc, MIT license; see THIRD-PARTY-NOTICES.md), which LibreOffice uses.
/// </summary>
internal static class IndicShaper
{
    private enum Script { Devanagari, Bengali, Gurmukhi, Gujarati, Oriya, Tamil, Telugu, Kannada, Malayalam }

    private enum RephMode { Implicit, Explicit, LogicalRepha }

    private sealed record Config(Script Script, string NewTag, string OldTag, int Virama, IndicPosition RephPosition,
        RephMode RephMode, bool BlwfPreAndPost);

    private static readonly Config[] Configs =
    [
        new(Script.Devanagari, "dev2", "deva", 0x094D, IndicPosition.BeforePost, RephMode.Implicit, true),
        new(Script.Bengali, "bng2", "beng", 0x09CD, IndicPosition.AfterSub, RephMode.Implicit, true),
        new(Script.Gurmukhi, "gur2", "guru", 0x0A4D, IndicPosition.BeforeSub, RephMode.Implicit, true),
        new(Script.Gujarati, "gjr2", "gujr", 0x0ACD, IndicPosition.BeforePost, RephMode.Implicit, true),
        new(Script.Oriya, "ory2", "orya", 0x0B4D, IndicPosition.AfterMain, RephMode.Implicit, true),
        new(Script.Tamil, "tml2", "taml", 0x0BCD, IndicPosition.AfterPost, RephMode.Implicit, true),
        new(Script.Telugu, "tel2", "telu", 0x0C4D, IndicPosition.AfterPost, RephMode.Explicit, false),
        new(Script.Kannada, "knd2", "knda", 0x0CCD, IndicPosition.AfterPost, RephMode.Implicit, false),
        new(Script.Malayalam, "mlm2", "mlym", 0x0D4D, IndicPosition.AfterMain, RephMode.LogicalRepha, true),
    ];

    // Feature masks; features absent from the font get 0, which disables the logic that depends on them.
    private const uint Global = 1, Rphf = 2, Pref = 4, Blwf = 8, Abvf = 16, Half = 32, Pstf = 64, Init = 128;

    private const uint ConsonantFlags = 1u << (int)IndicCategory.C | 1u << (int)IndicCategory.CS | 1u << (int)IndicCategory.Ra
        | 1u << (int)IndicCategory.CM | 1u << (int)IndicCategory.V | 1u << (int)IndicCategory.Placeholder
        | 1u << (int)IndicCategory.DottedCircle;
    private const uint JoinerFlags = 1u << (int)IndicCategory.ZWJ | 1u << (int)IndicCategory.ZWNJ;
    private const uint MatraFlags = 1u << (int)IndicCategory.M | 1u << (int)IndicCategory.MPst;

    /// <summary>True for characters of the scripts this shaper handles.</summary>
    public static bool IsIndic(int cp) => ScriptOf(cp) is not null;

    /// <summary>The script of a character (by its Unicode Script property: the Vedic stress signs are Inherited, the dandas Common).</summary>
    private static Config? ScriptOf(int cp) => cp switch
    {
        >= 0x0951 and <= 0x0954 or 0x0964 or 0x0965 => null,
        >= 0x0900 and <= 0x097F or >= 0xA8E0 and <= 0xA8FF => Configs[0],
        >= 0x0980 and <= 0x09FF => Configs[1],
        >= 0x0A00 and <= 0x0A7F => Configs[2],
        >= 0x0A80 and <= 0x0AFF => Configs[3],
        >= 0x0B00 and <= 0x0B7F => Configs[4],
        >= 0x0B80 and <= 0x0BFF => Configs[5],
        >= 0x0C00 and <= 0x0C7F => Configs[6],
        >= 0x0C80 and <= 0x0CFF => Configs[7],
        >= 0x0D00 and <= 0x0D7F => Configs[8],
        _ => null,
    };

    /// <summary>
    /// Shapes <paramref name="text"/>, a run of the script with the OpenType tag <paramref name="script"/>; returns
    /// null when that is not an Indic script or the font was not designed for it (no Indic script tag in GSUB), in
    /// which case the default shaping applies, as in HarfBuzz. Without <paramref name="ligatures"/> clig is off.
    /// </summary>
    public static List<GlyphInfo>? Shape(TrueTypeFont font, string text, string script, bool ligatures = true)
    {
        var config = Array.Find(Configs, c => c.OldTag == script);
        var layout = font.OpenTypeLayout;
        if (config is null) return null;
        string tag;
        if (layout.HasScript(config.NewTag)) tag = config.NewTag;
        else if (layout.HasScript(config.OldTag)) tag = config.OldTag;
        else return null;
        var plan = new Plan(font, layout, config, tag);

        var glyphs = Nominal(font, Normalize(font, text));
        FindSyllables(glyphs);

        plan.Apply(glyphs, ["locl", "ccmp"], Global, manualJoiners: false);
        InitialReordering(plan, glyphs);
        foreach (var (feature, mask) in new[]
                 {
                     ("nukt", Global), ("akhn", Global), ("rphf", Rphf), ("rkrf", Global), ("pref", Pref), ("blwf", Blwf),
                     ("abvf", Abvf), ("half", Half), ("pstf", Pstf), ("vatu", Global), ("cjct", Global),
                 })
            plan.Apply(glyphs, [feature], mask);
        FinalReordering(plan, glyphs);
        // The remaining Indic features apply together, per syllable, with the common ones (liga is off for Indic).
        var last = plan.Lookups(["init"], Init, perSyllable: true, manualJoiners: true)
            .Concat(plan.Lookups(["pres", "abvs", "blws", "psts", "haln"], Global, perSyllable: true, manualJoiners: true))
            .Concat(plan.Lookups(ligatures ? ["rlig", "calt", "clig", "rclt"] : ["rlig", "calt", "rclt"], Global,
                perSyllable: false, manualJoiners: false))
            .GroupBy(l => l.Lookup).Select(g => (g.Key, g.Aggregate(0u, (m, l) => m | l.Mask))).OrderBy(l => l.Key).ToList();
        layout.Apply(false, glyphs, last);

        for (int i = 0; i < glyphs.Count; i++) glyphs[i] = glyphs[i] with { XAdvance = font.GetAdvanceWidth(glyphs[i].Glyph) };
        TextShaper.ApplySpaceFallback(font, glyphs);
        if (layout.HasPositioning)
        {
            // Indic marks keep their advance (HarfBuzz's ZERO_WIDTH_MARKS_NONE).
            var gposTag = layout.HasScript(config.NewTag, gpos: true) ? config.NewTag : config.OldTag;
            const uint manual = Global | OpenTypeLayout.ManualJoiners; // mark and mkmk see ZWJ/ZWNJ, as in HarfBuzz
            layout.Apply(true, glyphs, layout.Lookups(true, gposTag,
                [("abvm", Global), ("blwm", Global), ("mark", manual), ("mkmk", manual), ("curs", Global), ("dist", Global), ("kern", Global)]));
            TextShaper.ResolveAttachments(glyphs, rightToLeft: false);
        }
        TextShaper.HideIgnorables(font, glyphs);
        AssignSyllableText(text, glyphs);
        return glyphs;
    }

    // ------------------------------------------------------------------ plan

    private sealed class Plan
    {
        public Plan(TrueTypeFont font, OpenTypeLayout layout, Config config, string tag)
        {
            Layout = layout;
            Config = config;
            Tag = tag;
            OldSpec = !tag.EndsWith('2');
            ViramaGlyph = font.GetGlyphId(config.Virama);
            DottedCircleGlyph = font.GetGlyphId(0x25CC);
            // Zero-context matching for the new spec, except Malayalam (as in HarfBuzz; see its comments).
            ZeroContext = !OldSpec && config.Script != Script.Malayalam;
            uint Mask(string feature, uint bit) => Lookups([feature], bit, false, false).Count > 0 ? bit : 0;
            RphfMask = Mask("rphf", Rphf);
            PrefMask = Mask("pref", Pref);
            BlwfMask = Mask("blwf", Blwf);
            AbvfMask = Mask("abvf", Abvf);
            HalfMask = Mask("half", Half);
            PstfMask = Mask("pstf", Pstf);
            InitMask = Mask("init", Init);
        }

        public OpenTypeLayout Layout { get; }
        public Config Config { get; }
        public string Tag { get; }
        public bool OldSpec { get; }
        public int ViramaGlyph { get; }
        public int DottedCircleGlyph { get; }
        public bool ZeroContext { get; }
        public uint RphfMask { get; }
        public uint PrefMask { get; }
        public uint BlwfMask { get; }
        public uint AbvfMask { get; }
        public uint HalfMask { get; }
        public uint PstfMask { get; }
        public uint InitMask { get; }
        public Script Script => Config.Script;

        public List<(int Lookup, uint Mask)> Lookups(string[] features, uint mask, bool perSyllable, bool manualJoiners) =>
            Layout.Lookups(false, Tag, features.Select(f => (f, mask | (perSyllable ? OpenTypeLayout.PerSyllable : 0)
                                                                     | (manualJoiners ? OpenTypeLayout.ManualJoiners : 0))).ToList());

        /// <summary>Applies one stage of per-syllable features (the Indic ones handle joiners themselves).</summary>
        public void Apply(List<GlyphInfo> glyphs, string[] features, uint mask, bool manualJoiners = true)
        {
            if (mask != 0) Layout.Apply(false, glyphs, Lookups(features, mask, perSyllable: true, manualJoiners));
        }

        public bool WouldSubstitute(string feature, ReadOnlySpan<int> glyphs)
        {
            foreach (var (lookup, _) in Lookups([feature], Global, false, false))
                if (Layout.WouldSubstitute(lookup, glyphs, ZeroContext)) return true;
            return false;
        }
    }

    // ------------------------------------------------------------------ text preparation

    /// <summary>
    /// HarfBuzz's COMPOSED_DIACRITICS_NO_SHORT_CIRCUIT normalization: decompose (when the font has the parts), order
    /// marks canonically, then recompose a letter and following mark when the font has the composite.
    /// </summary>
    private static List<NormalizedChar> Normalize(TrueTypeFont font, string text)
    {
        var codePoints = new List<(int CodePoint, int Cluster)>(text.Length);
        for (int i = 0; i < text.Length; i += char.IsSurrogatePair(text, i) ? 2 : 1)
            codePoints.Add((char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i], i));
        return Normalizer.Normalize(font, ApplyVowelConstraints(codePoints).ToList(), Normalizer.Shaper.Indic);
    }

    /// <summary>
    /// Independent vowel + matra pairs that look like another vowel (HarfBuzz's vowel constraints, from the USE
    /// script development spec): the matra gets a dotted circle of its own.
    /// </summary>
    private static readonly Dictionary<int, int[]> VowelConstraints = new()
    {
        [0x0905] = [0x093A, 0x093B, 0x093E, 0x0945, 0x0946, 0x0949, 0x094A, 0x094B, 0x094C, 0x094F, 0x0956, 0x0957],
        [0x0906] = [0x093A, 0x0945, 0x0946, 0x0947, 0x0948],
        [0x0909] = [0x0941],
        [0x090F] = [0x0945, 0x0946, 0x0947],
        [0x0985] = [0x09BE], [0x098B] = [0x09C3], [0x098C] = [0x09E2],
        [0x0A05] = [0x0A3E, 0x0A48, 0x0A4C], [0x0A72] = [0x0A3F, 0x0A40, 0x0A47], [0x0A73] = [0x0A41, 0x0A42, 0x0A4B],
        [0x0A85] = [0x0ABE, 0x0AC5, 0x0AC7, 0x0AC8, 0x0AC9, 0x0ACB, 0x0ACC], [0x0AC5] = [0x0ABE],
        [0x0B05] = [0x0B3E], [0x0B0F] = [0x0B57], [0x0B13] = [0x0B57],
        [0x0B85] = [0x0BC2],
        [0x0C12] = [0x0C4C, 0x0C55], [0x0C3F] = [0x0C55], [0x0C46] = [0x0C55], [0x0C4A] = [0x0C55],
        [0x0C89] = [0x0CBE], [0x0C8B] = [0x0CBE], [0x0C92] = [0x0CCC],
        [0x0D07] = [0x0D57], [0x0D09] = [0x0D57], [0x0D0E] = [0x0D46], [0x0D12] = [0x0D3E, 0x0D57],
    };

    private static IEnumerable<(int CodePoint, int Cluster)> ApplyVowelConstraints(List<(int CodePoint, int Cluster)> cps)
    {
        for (int i = 0; i < cps.Count; i++)
        {
            yield return cps[i];
            if (i + 1 >= cps.Count) continue;
            int cp = cps[i].CodePoint, next = cps[i + 1].CodePoint;
            bool matched = VowelConstraints.TryGetValue(cp, out var seconds) && Array.IndexOf(seconds, next) >= 0;
            if (cp == 0x0930 && next == 0x094D && i + 2 < cps.Count && cps[i + 2].CodePoint == 0x0907)
            {
                yield return cps[++i]; // Ra,Halant,I: the circle goes before the I
                matched = true;
            }
            if (!matched) continue;
            yield return (0x25CC, cps[i + 1].Cluster);
            yield return cps[++i];
        }
    }

    private static List<GlyphInfo> Nominal(TrueTypeFont font, List<NormalizedChar> chars)
    {
        var glyphs = new List<GlyphInfo>(chars.Count);
        foreach (var (cp, cluster, text) in chars)
        {
            var (category, position) = Properties(cp);
            glyphs.Add(new GlyphInfo
            {
                Glyph = Normalizer.NominalGlyph(font, cp), Cluster = cluster, Text = text, Mask = Global,
                XAdvance = 0, AttachTo = -1, CursiveTo = -1, CodePoint = cp, ShaperCategory = (byte)category,
                ShaperPosition = (byte)position, Ignorable = TextShaper.IsDefaultIgnorable(cp),
            });
        }
        return glyphs;
    }

    private static (IndicCategory, IndicPosition) Properties(int cp)
    {
        var starts = IndicData.RangeStarts;
        int lo = 0, hi = starts.Length - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) >> 1;
            if (starts[mid] <= cp) lo = mid; else hi = mid - 1;
        }
        int value = cp >= starts[^1] ? 0 | (int)IndicPosition.End << 8 : IndicData.RangeValues[lo];
        return ((IndicCategory)(value & 0xFF), (IndicPosition)(value >> 8));
    }

    // ------------------------------------------------------------------ syllables

    private enum SyllableType { Consonant, Vowel, Standalone, Symbol, Broken, NonIndic }

    /// <summary>Numbers the syllables (Syllable = serial * 16 + type), as HarfBuzz's Ragel scanner does: longest match, first rule on ties.</summary>
    private static void FindSyllables(List<GlyphInfo> glyphs)
    {
        var categories = glyphs.Select(g => (IndicCategory)g.ShaperCategory).ToArray();
        int serial = 1;
        for (int p = 0; p < categories.Length;)
        {
            int best = 0;
            var type = SyllableType.NonIndic;
            foreach (var (rule, ruleType) in SyllableGrammar.RuleList)
            {
                int end = rule.Longest(categories, p);
                if (end - p > best) (best, type) = (end - p, (SyllableType)ruleType);
            }
            if (best == 0) best = 1; // "other"
            for (int k = p; k < p + best; k++) glyphs[k] = glyphs[k] with { Syllable = serial * 16 + (int)type };
            serial++;
            p += best;
        }
    }

    private static SyllableType TypeOf(GlyphInfo g) => (SyllableType)(g.Syllable & 0xF);

    private static IEnumerable<(int Start, int End)> Syllables(List<GlyphInfo> glyphs)
    {
        for (int start = 0; start < glyphs.Count;)
        {
            int end = start + 1;
            while (end < glyphs.Count && glyphs[end].Syllable == glyphs[start].Syllable) end++;
            yield return (start, end);
            start = end;
        }
    }

    // ------------------------------------------------------------------ helpers

    private static IndicCategory Cat(in GlyphInfo g) => (IndicCategory)g.ShaperCategory;
    private static IndicPosition Pos(in GlyphInfo g) => (IndicPosition)g.ShaperPosition;
    private static void SetPos(ref GlyphInfo g, IndicPosition p) => g.ShaperPosition = (byte)p;
    private static uint Flag(IndicCategory c) => 1u << (int)c;
    private static bool Has(in GlyphInfo g, uint flags) => (Flag(Cat(g)) & flags) != 0;

    /// <summary>Category test that fails for ligated glyphs ("if it ligated, all bets are off").</summary>
    private static bool IsOneOf(in GlyphInfo g, uint flags) => (g.Flags & GlyphFlags.Ligated) == 0 && Has(g, flags);
    private static bool IsConsonant(in GlyphInfo g) => IsOneOf(g, ConsonantFlags);
    private static bool IsJoiner(in GlyphInfo g) => IsOneOf(g, JoinerFlags);
    private static bool IsHalant(in GlyphInfo g) => IsOneOf(g, Flag(IndicCategory.H));
    private static bool LigatedAndDidntMultiply(in GlyphInfo g) =>
        (g.Flags & (GlyphFlags.Ligated | GlyphFlags.Multiplied)) == GlyphFlags.Ligated;

    /// <summary>Moves the glyph at <paramref name="from"/> to <paramref name="to"/>, shifting those between.</summary>
    private static void Move(Span<GlyphInfo> g, int from, int to)
    {
        var moved = g[from];
        if (from < to) g[(from + 1)..(to + 1)].CopyTo(g[from..]);
        else g[to..from].CopyTo(g[(to + 1)..]);
        g[to] = moved;
    }

    // ------------------------------------------------------------------ initial reordering

    private static void InitialReordering(Plan plan, List<GlyphInfo> glyphs)
    {
        if (plan.ViramaGlyph != 0)
        {
            var g = CollectionsMarshal.AsSpan(glyphs);
            for (int i = 0; i < g.Length; i++)
                if (Pos(g[i]) == IndicPosition.BaseC) SetPos(ref g[i], ConsonantPosition(plan, g[i].Glyph));
        }
        InsertDottedCircles(plan, glyphs);
        var span = CollectionsMarshal.AsSpan(glyphs);
        foreach (var (start, end) in Syllables(glyphs).ToList())
            if (TypeOf(glyphs[start]) is SyllableType.Consonant or SyllableType.Vowel or SyllableType.Standalone or SyllableType.Broken)
                ReorderConsonantSyllable(plan, span, start, end);
    }

    private static IndicPosition ConsonantPosition(Plan plan, int consonant)
    {
        // Both orders: some fonts kept old-spec (consonant, virama) lookups in their new-spec tables.
        int[] glyphs = [plan.ViramaGlyph, consonant, plan.ViramaGlyph];
        bool Would(string feature) => plan.WouldSubstitute(feature, glyphs.AsSpan(0, 2)) || plan.WouldSubstitute(feature, glyphs.AsSpan(1));
        if (Would("blwf") || Would("vatu")) return IndicPosition.BelowC;
        if (Would("pstf") || Would("pref")) return IndicPosition.PostC;
        return IndicPosition.BaseC;
    }

    /// <summary>Inserts a dotted circle into broken clusters (after a leading repha), so marks have a base to sit on.</summary>
    private static void InsertDottedCircles(Plan plan, List<GlyphInfo> glyphs)
    {
        if (!glyphs.Exists(g => TypeOf(g) == SyllableType.Broken)) return;
        int circle = plan.DottedCircleGlyph;
        if (circle == 0) return;
        int last = 0;
        for (int i = 0; i < glyphs.Count; i++)
        {
            var g = glyphs[i];
            if (g.Syllable == last || TypeOf(g) != SyllableType.Broken) continue;
            last = g.Syllable;
            int at = i;
            while (at < glyphs.Count && glyphs[at].Syllable == last && Cat(glyphs[at]) == IndicCategory.Repha) at++;
            glyphs.Insert(at, g with
            {
                Glyph = (ushort)circle, Text = "", CodePoint = 0x25CC, ShaperCategory = (byte)IndicCategory.DottedCircle,
                ShaperPosition = (byte)IndicPosition.End, Flags = GlyphFlags.None, Ignorable = false,
            });
        }
    }

    private static void ReorderConsonantSyllable(Plan plan, Span<GlyphInfo> g, int start, int end)
    {
        var script = plan.Script;
        // Kannada: Ra,H,ZWJ behaves like Ra,ZWJ,H.
        if (script == Script.Kannada && start + 3 <= end && IsOneOf(g[start], Flag(IndicCategory.Ra))
            && IsOneOf(g[start + 1], Flag(IndicCategory.H)) && IsOneOf(g[start + 2], Flag(IndicCategory.ZWJ)))
            (g[start + 1], g[start + 2]) = (g[start + 2], g[start + 1]);

        // 1. Find the base consonant.
        int baseIndex = end;
        bool hasReph = false;
        int limit = start;
        var rephMode = plan.Config.RephMode;
        if (plan.RphfMask != 0 && start + 3 <= end
            && ((rephMode == RephMode.Implicit && !IsJoiner(g[start + 2]))
                || (rephMode == RephMode.Explicit && Cat(g[start + 2]) == IndicCategory.ZWJ)))
        {
            ReadOnlySpan<int> reph = [g[start].Glyph, g[start + 1].Glyph, rephMode == RephMode.Explicit ? g[start + 2].Glyph : 0];
            if (plan.WouldSubstitute("rphf", reph[..2]) || (rephMode == RephMode.Explicit && plan.WouldSubstitute("rphf", reph)))
            {
                limit += 2;
                while (limit < end && IsJoiner(g[limit])) limit++;
                baseIndex = start;
                hasReph = true;
            }
        }
        else if (rephMode == RephMode.LogicalRepha && Cat(g[start]) == IndicCategory.Repha)
        {
            limit += 1;
            while (limit < end && IsJoiner(g[limit])) limit++;
            baseIndex = start;
            hasReph = true;
        }

        {
            int i = end;
            bool seenBelow = false;
            do
            {
                i--;
                if (IsConsonant(g[i]))
                {
                    // A consonant without a below-base or post-base form (post-base forms follow below-base ones).
                    if (Pos(g[i]) != IndicPosition.BelowC && (Pos(g[i]) != IndicPosition.PostC || seenBelow))
                    {
                        baseIndex = i;
                        break;
                    }
                    if (Pos(g[i]) == IndicPosition.BelowC) seenBelow = true;
                    baseIndex = i;
                }
                else if (start < i && Cat(g[i]) == IndicCategory.ZWJ && Cat(g[i - 1]) == IndicCategory.H)
                    break; // a ZWJ after a halant requests a half form and stops the search
            } while (i > limit);
        }
        if (hasReph && baseIndex == start && limit - baseIndex <= 2) hasReph = false; // Ra is the only consonant

        // Positions before and at the base; reph.
        for (int i = start; i < baseIndex; i++)
            if (Pos(g[i]) > IndicPosition.PreC) SetPos(ref g[i], IndicPosition.PreC);
        if (baseIndex < end) SetPos(ref g[baseIndex], IndicPosition.BaseC);
        if (hasReph) SetPos(ref g[start], IndicPosition.RaToBecomeReph);

        // Old spec: the first post-base halant moves after the last consonant.
        if (plan.OldSpec)
        {
            bool disallowDoubleHalants = script == Script.Kannada;
            for (int i = baseIndex + 1; i < end; i++)
                if (Cat(g[i]) == IndicCategory.H)
                {
                    int j;
                    for (j = end - 1; j > i; j--)
                        if (IsConsonant(g[j]) || (disallowDoubleHalants && Cat(g[j]) == IndicCategory.H)) break;
                    if (Cat(g[j]) != IndicCategory.H && j > i) Move(g, i, j);
                    break;
                }
        }

        // Marks move with the character before them.
        {
            var lastPos = IndicPosition.Start;
            for (int i = start; i < end; i++)
            {
                if (Has(g[i], JoinerFlags | Flag(IndicCategory.N) | Flag(IndicCategory.RS) | Flag(IndicCategory.CM) | Flag(IndicCategory.H)))
                {
                    SetPos(ref g[i], lastPos);
                    if (Cat(g[i]) == IndicCategory.H && Pos(g[i]) == IndicPosition.PreM)
                        // Uniscribe does not move the halant with a left matra.
                        for (int j = i; j > start; j--)
                            if (Pos(g[j - 1]) != IndicPosition.PreM)
                            {
                                SetPos(ref g[i], Pos(g[j - 1]));
                                break;
                            }
                }
                else if (Pos(g[i]) != IndicPosition.Smvd)
                {
                    if (Cat(g[i]) == IndicCategory.MPst && i > start && Cat(g[i - 1]) == IndicCategory.SM)
                        SetPos(ref g[i - 1], Pos(g[i]));
                    lastPos = Pos(g[i]);
                }
            }
        }
        // Post-base consonants own what comes before them since the last consonant or matra.
        {
            int last = baseIndex;
            for (int i = baseIndex + 1; i < end; i++)
                if (IsConsonant(g[i]))
                {
                    for (int j = last + 1; j < i; j++)
                        if (Pos(g[j]) < IndicPosition.Smvd) SetPos(ref g[j], Pos(g[i]));
                    last = i;
                }
                else if (Has(g[i], MatraFlags)) last = i;
        }

        // Sort by position (stable), then find the base again and flip a sequence of left matras.
        var sorted = g[start..end].ToArray().Select((x, k) => (x, k)).OrderBy(p => p.x.ShaperPosition).ThenBy(p => p.k)
            .Select(p => p.x).ToArray();
        sorted.CopyTo(g[start..end]);
        int firstLeftMatra = end, lastLeftMatra = end;
        baseIndex = end;
        for (int i = start; i < end; i++)
        {
            if (Pos(g[i]) == IndicPosition.BaseC)
            {
                baseIndex = i;
                break;
            }
            if (Pos(g[i]) == IndicPosition.PreM)
            {
                if (firstLeftMatra == end) firstLeftMatra = i;
                lastLeftMatra = i;
            }
        }
        if (firstLeftMatra < lastLeftMatra)
        {
            g[firstLeftMatra..(lastLeftMatra + 1)].Reverse();
            int i = firstLeftMatra;
            for (int j = i; j <= lastLeftMatra; j++)
                if (Has(g[j], MatraFlags))
                {
                    g[i..(j + 1)].Reverse(); // nuktas etc. go back after their matra
                    i = j + 1;
                }
        }

        // Feature masks.
        for (int i = start; i < end && Pos(g[i]) == IndicPosition.RaToBecomeReph; i++) g[i].Mask |= plan.RphfMask;
        uint preBase = plan.HalfMask | (!plan.OldSpec && plan.Config.BlwfPreAndPost ? plan.BlwfMask : 0);
        for (int i = start; i < baseIndex; i++) g[i].Mask |= preBase;
        uint postBase = plan.BlwfMask | plan.AbvfMask | plan.PstfMask;
        for (int i = baseIndex + 1; i < end; i++) g[i].Mask |= postBase;
        if (plan.OldSpec && script == Script.Devanagari)
            // Old-spec eyelash Ra: blwf also applies to Ra,H before the base (unless Ra,H,ZWJ asks for the eyelash form).
            for (int i = start; i + 1 < baseIndex; i++)
                if (Cat(g[i]) == IndicCategory.Ra && Cat(g[i + 1]) == IndicCategory.H
                    && (i + 2 == baseIndex || Cat(g[i + 2]) != IndicCategory.ZWJ))
                {
                    g[i].Mask |= plan.BlwfMask;
                    g[i + 1].Mask |= plan.BlwfMask;
                }
        if (plan.PrefMask != 0 && baseIndex + 2 < end)
            // A Halant,Ra sequence that forms a pre-base-reordering Ra.
            for (int i = baseIndex + 1; i + 1 < end; i++)
                if (plan.WouldSubstitute("pref", [g[i].Glyph, g[i + 1].Glyph]))
                {
                    g[i].Mask |= plan.PrefMask;
                    g[i + 1].Mask |= plan.PrefMask;
                    break;
                }
        // A ZWNJ disables half forms before it (ZWJ/ZWNJ disable cjct by being there).
        for (int i = start + 1; i < end; i++)
            if (IsJoiner(g[i]))
            {
                bool nonJoiner = Cat(g[i]) == IndicCategory.ZWNJ;
                int j = i;
                do
                {
                    j--;
                    if (nonJoiner) g[j].Mask &= ~plan.HalfMask;
                } while (j > start && !IsConsonant(g[j]));
            }
    }

    // ------------------------------------------------------------------ final reordering

    private static void FinalReordering(Plan plan, List<GlyphInfo> glyphs)
    {
        var g = CollectionsMarshal.AsSpan(glyphs);
        foreach (var (start, end) in Syllables(glyphs).ToList())
            FinalReorderSyllable(plan, g, start, end);
    }

    private static void FinalReorderSyllable(Plan plan, Span<GlyphInfo> g, int start, int end)
    {
        var script = plan.Script;
        // Recover the halant category of a virama that went through a ligature and back.
        if (plan.ViramaGlyph != 0)
            for (int i = start; i < end; i++)
                if (g[i].Glyph == plan.ViramaGlyph && (g[i].Flags & (GlyphFlags.Ligated | GlyphFlags.Multiplied)) == (GlyphFlags.Ligated | GlyphFlags.Multiplied))
                {
                    g[i].ShaperCategory = (byte)IndicCategory.H;
                    g[i].Flags &= ~(GlyphFlags.Ligated | GlyphFlags.Multiplied);
                }

        bool tryPref = plan.PrefMask != 0;
        int baseIndex;
        for (baseIndex = start; baseIndex < end; baseIndex++)
            if (Pos(g[baseIndex]) >= IndicPosition.BaseC)
            {
                if (tryPref && baseIndex + 1 < end)
                {
                    for (int i = baseIndex + 1; i < end; i++)
                        if ((g[i].Mask & plan.PrefMask) != 0)
                        {
                            if (!((g[i].Flags & GlyphFlags.Substituted) != 0 && LigatedAndDidntMultiply(g[i])))
                            {
                                // A pref candidate that did not form: the base is around here.
                                baseIndex = i;
                                while (baseIndex < end && IsHalant(g[baseIndex])) baseIndex++;
                                if (baseIndex < end) SetPos(ref g[baseIndex], IndicPosition.BaseC);
                                tryPref = false;
                            }
                            break;
                        }
                    if (baseIndex == end) break;
                }
                if (script == Script.Malayalam)
                    // Skip over unformed below (but not post) forms.
                    for (int i = baseIndex + 1; i < end; i++)
                    {
                        while (i < end && IsJoiner(g[i])) i++;
                        if (i == end || !IsHalant(g[i])) break;
                        i++;
                        while (i < end && IsJoiner(g[i])) i++;
                        if (i < end && IsConsonant(g[i]) && Pos(g[i]) == IndicPosition.BelowC)
                        {
                            baseIndex = i;
                            SetPos(ref g[baseIndex], IndicPosition.BaseC);
                        }
                    }
                if (start < baseIndex && Pos(g[baseIndex]) > IndicPosition.BaseC) baseIndex--;
                break;
            }
        if (baseIndex == end && start < baseIndex && IsOneOf(g[baseIndex - 1], Flag(IndicCategory.ZWJ))) baseIndex--;
        if (baseIndex < end)
            while (start < baseIndex && IsOneOf(g[baseIndex], Flag(IndicCategory.N) | Flag(IndicCategory.H))) baseIndex--;

        // Pre-base matras move to after the last standalone halant before the base (not past a halant followed by ZWJ).
        if (start + 1 < end && start < baseIndex)
        {
            int newPos = baseIndex == end ? baseIndex - 2 : baseIndex - 1;
            if (script is not (Script.Malayalam or Script.Tamil))
            {
                while (true)
                {
                    while (newPos > start && !IsOneOf(g[newPos], MatraFlags | Flag(IndicCategory.H))) newPos--;
                    if (IsHalant(g[newPos]) && Pos(g[newPos]) != IndicPosition.PreM)
                    {
                        if (newPos + 1 < end && Cat(g[newPos + 1]) == IndicCategory.ZWJ && newPos > start)
                        {
                            newPos--;
                            continue;
                        }
                    }
                    else newPos = start;
                    break;
                }
            }
            if (start < newPos && Pos(g[newPos]) != IndicPosition.PreM)
                for (int i = newPos; i > start; i--)
                    if (Pos(g[i - 1]) == IndicPosition.PreM)
                    {
                        int oldPos = i - 1;
                        if (oldPos < baseIndex && baseIndex <= newPos) baseIndex--;
                        Move(g, oldPos, newPos);
                        newPos--;
                    }
        }

        // Reph: move it only if Ra,H ligated to a reph (or, for an encoded repha, if it did not ligate).
        if (start + 1 < end && Pos(g[start]) == IndicPosition.RaToBecomeReph
            && ((Cat(g[start]) == IndicCategory.Repha) ^ LigatedAndDidntMultiply(g[start])))
        {
            var rephPos = plan.Config.RephPosition;
            int newReph = -1;
            if (rephPos != IndicPosition.AfterPost) newReph = AfterFirstHalant(g, start, baseIndex); // step 2
            if (newReph < 0 && rephPos == IndicPosition.AfterMain) // step 3
            {
                newReph = baseIndex;
                while (newReph + 1 < end && Pos(g[newReph + 1]) <= IndicPosition.AfterMain) newReph++;
                if (newReph >= end) newReph = -1;
            }
            if (newReph < 0 && rephPos == IndicPosition.AfterSub) // step 4
            {
                newReph = baseIndex;
                while (newReph + 1 < end && Pos(g[newReph + 1]) is not (IndicPosition.PostC or IndicPosition.AfterPost or IndicPosition.Smvd))
                    newReph++;
                if (newReph >= end) newReph = -1;
            }
            if (newReph < 0) newReph = AfterFirstHalant(g, start, baseIndex); // step 5
            if (newReph < 0) // step 6: the end of the syllable, before syllable modifiers
            {
                newReph = end - 1;
                while (newReph > start && Pos(g[newReph]) == IndicPosition.Smvd) newReph--;
                // After a Matra,Halant sequence, go before the halant so the reph can interact with the matra.
                if (IsHalant(g[newReph]))
                    for (int i = baseIndex + 1; i < newReph; i++)
                        if (Has(g[i], MatraFlags)) newReph--;
            }
            Move(g, start, newReph);
            if (start < baseIndex && baseIndex <= newReph) baseIndex--;
        }

        // Pre-base-reordering consonants (pref) that formed go before the base, like pre-base matras.
        if (tryPref && baseIndex + 1 < end)
            for (int i = baseIndex + 1; i < end; i++)
                if ((g[i].Mask & plan.PrefMask) != 0)
                {
                    if (LigatedAndDidntMultiply(g[i]))
                    {
                        int newPos = baseIndex;
                        if (script is not (Script.Malayalam or Script.Tamil))
                            while (newPos > start && !IsOneOf(g[newPos - 1], MatraFlags | Flag(IndicCategory.H))) newPos--;
                        if (newPos > start && IsHalant(g[newPos - 1]) && newPos < end && IsJoiner(g[newPos])) newPos++;
                        Move(g, i, newPos);
                        if (newPos <= baseIndex && baseIndex < i) baseIndex++;
                    }
                    break;
                }

        // 'init' applies to a left matra at the start of a word.
        if (Pos(g[start]) == IndicPosition.PreM && (start == 0 || !IsLetterOrMark(g[start - 1].CodePoint)))
            g[start].Mask |= plan.InitMask;
    }

    /// <summary>Reph target after the first explicit halant before the base (and a joiner after it), or -1.</summary>
    private static int AfterFirstHalant(Span<GlyphInfo> g, int start, int baseIndex)
    {
        int p = start + 1;
        while (p < baseIndex && !IsHalant(g[p])) p++;
        if (p >= baseIndex) return -1;
        if (p + 1 < baseIndex && IsJoiner(g[p + 1])) p++;
        return p;
    }

    private static bool IsLetterOrMark(int cp) => cp != 0 && CharUnicodeInfo.GetUnicodeCategory(cp) is
        UnicodeCategory.Format or UnicodeCategory.OtherNotAssigned or UnicodeCategory.PrivateUse or UnicodeCategory.Surrogate
        or UnicodeCategory.LowercaseLetter or UnicodeCategory.ModifierLetter or UnicodeCategory.OtherLetter
        or UnicodeCategory.TitlecaseLetter or UnicodeCategory.UppercaseLetter or UnicodeCategory.SpacingCombiningMark
        or UnicodeCategory.EnclosingMark or UnicodeCategory.NonSpacingMark;

    // ------------------------------------------------------------------ output

    /// <summary>
    /// Reordering breaks the link between glyph order and text, so each syllable's text goes on its first glyph
    /// (HarfBuzz merges such clusters too); text extraction then gets the characters in logical order.
    /// </summary>
    private static void AssignSyllableText(string text, List<GlyphInfo> glyphs)
    {
        foreach (var (start, end) in Syllables(glyphs).ToList())
        {
            if (end - start == 1 && glyphs[start].Text.Length > 0) continue;
            int from = int.MaxValue, to = 0;
            for (int i = start; i < end; i++)
            {
                if (glyphs[i].CodePoint == 0x25CC && glyphs[i].Text.Length == 0) continue; // inserted dotted circle
                from = Math.Min(from, glyphs[i].Cluster);
                to = Math.Max(to, glyphs[i].Cluster + Math.Max(1, char.ConvertFromUtf32(Math.Max(glyphs[i].CodePoint, 1)).Length));
            }
            if (from == int.MaxValue) continue;
            to = Math.Min(to, text.Length);
            for (int i = start; i < end; i++)
                glyphs[i] = glyphs[i] with { Cluster = from, Text = i == start ? text[from..to] : "" };
        }
    }
}
