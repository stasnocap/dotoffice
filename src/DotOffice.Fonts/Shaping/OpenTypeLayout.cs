using System.Buffers.Binary;

namespace DotOffice.Fonts.Shaping;

/// <summary>A glyph being shaped. Positions are in font units.</summary>
internal struct GlyphInfo
{
    public ushort Glyph;

    /// <summary>Index of the first UTF-16 code unit of the text this glyph came from.</summary>
    public int Cluster;

    /// <summary>The characters this glyph stands for (several for a ligature, empty for extra glyphs of a decomposition).</summary>
    public string Text;

    /// <summary>Features enabled for this glyph (bit masks chosen by the shaper).</summary>
    public uint Mask;

    public int XAdvance, XOffset, YOffset;

    /// <summary>Index of the glyph a mark is attached to (relative position resolved at the end), or -1.</summary>
    public int AttachTo;

    /// <summary>
    /// Index of the glyph this one is cursively attached to, or -1. <see cref="YOffset"/> is then relative to
    /// that glyph's (resolved at the end, along the chain).
    /// </summary>
    public int CursiveTo;

    /// <summary>What GSUB did to this glyph (HarfBuzz's glyph props), used by the Indic shaper.</summary>
    public GlyphFlags Flags;

    /// <summary>Syllable number (non-zero for complex scripts); lookups of per-syllable features do not cross syllables.</summary>
    public int Syllable;

    /// <summary>The character the glyph was made from (the first one for ligatures; 0 for inserted glyphs).</summary>
    public int CodePoint;

    /// <summary>Shaper-specific character category and position (Indic shaping).</summary>
    public byte ShaperCategory, ShaperPosition;

    /// <summary>Made from a default-ignorable character (ZWJ, ZWNJ…), hidden after shaping.</summary>
    public bool Ignorable;

    /// <summary>A default ignorable GSUB must not skip (see <see cref="NormalizedChar.Hidden"/>).</summary>
    public bool Hidden;

    /// <summary>
    /// Ligature membership (HarfBuzz's lig_props): a ligature and the marks inside it share a non-zero id; a mark's
    /// component is the 1-based ligature component it belongs to (0 for the ligature itself).
    /// </summary>
    public int LigatureId, LigatureComponent;
}

[Flags]
internal enum GlyphFlags : byte
{
    None = 0,
    Substituted = 1,
    Ligated = 2,
    Multiplied = 4,
}

/// <summary>
/// GSUB/GPOS/GDEF interpreter: selects the lookups of a script's features and applies them to a glyph
/// buffer. Supports GSUB types 1–7 and GPOS types 1–9; anything else
/// (and malformed data, by throwing, which the shaper catches) leaves the glyphs unchanged.
/// </summary>
internal sealed class OpenTypeLayout
{
    /// <summary>Mask bit marking lookups of features that apply per syllable (never set on glyphs).</summary>
    public const uint PerSyllable = 0x8000_0000;

    /// <summary>
    /// Mask bit marking lookups of features that handle ZWJ/ZWNJ themselves (HarfBuzz's manual joiners) instead of
    /// skipping them; never set on glyphs.
    /// </summary>
    public const uint ManualJoiners = 0x4000_0000;

    private const int MaxNesting = 6;
    private const int MaxOperations = 100_000;

    private readonly byte[] _gsub, _gpos, _gdef;

    public OpenTypeLayout(TrueTypeFont font)
    {
        _gsub = font.GetTableData("GSUB").ToArray();
        _gpos = font.GetTableData("GPOS").ToArray();
        _gdef = font.GetTableData("GDEF").ToArray();
    }

    public bool HasSubstitutions => _gsub.Length > 0;
    public bool HasPositioning => _gpos.Length > 0;

    // ------------------------------------------------------------------ table helpers

    private static int U16(byte[] t, int o) => BinaryPrimitives.ReadUInt16BigEndian(t.AsSpan(o, 2));
    private static int S16(byte[] t, int o) => BinaryPrimitives.ReadInt16BigEndian(t.AsSpan(o, 2));
    private static int U32(byte[] t, int o) => (int)BinaryPrimitives.ReadUInt32BigEndian(t.AsSpan(o, 4));

    private static string Tag(byte[] t, int o) => new([(char)t[o], (char)t[o + 1], (char)t[o + 2], (char)t[o + 3]]);

    /// <summary>Coverage index of <paramref name="glyph"/>, or -1.</summary>
    private static int Coverage(byte[] t, int offset, int glyph)
    {
        int format = U16(t, offset);
        int count = U16(t, offset + 2);
        int lo = 0, hi = count - 1;
        if (format == 1)
        {
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1;
                int g = U16(t, offset + 4 + mid * 2);
                if (g == glyph) return mid;
                if (g < glyph) lo = mid + 1; else hi = mid - 1;
            }
        }
        else if (format == 2)
        {
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1;
                int rec = offset + 4 + mid * 6;
                int start = U16(t, rec), end = U16(t, rec + 2);
                if (glyph < start) hi = mid - 1;
                else if (glyph > end) lo = mid + 1;
                else return U16(t, rec + 4) + glyph - start;
            }
        }
        return -1;
    }

    /// <summary>All glyphs listed in a coverage table.</summary>
    private static IEnumerable<int> CoveredGlyphs(byte[] t, int offset)
    {
        int format = U16(t, offset), count = U16(t, offset + 2);
        for (int i = 0; i < count; i++)
        {
            if (format == 1) yield return U16(t, offset + 4 + i * 2);
            else if (format == 2)
                for (int g = U16(t, offset + 4 + i * 6); g <= U16(t, offset + 6 + i * 6); g++) yield return g;
        }
    }

    private (HashSet<int> Glyphs, HashSet<(int, int)> Pairs)? _defaultSubstitutionStarts;

    /// <summary>
    /// True when a lookup of the default GSUB features (ccmp, locl, rlig, calt, clig, liga, rclt) in the default
    /// language system of any script may start at <paramref name="glyph"/> followed by <paramref name="next"/> (-1
    /// at the end), so text with them must be shaped (e.g. for "fi" ligatures). Ligatures need their second
    /// component next; other substitutions only their first glyph.
    /// </summary>
    public bool StartsDefaultSubstitution(int glyph, int next)
    {
        var (glyphs, pairs) = _defaultSubstitutionStarts ??= DefaultSubstitutionStarts();
        return glyphs.Contains(glyph) || pairs.Contains((glyph, next));
    }

    private (HashSet<int>, HashSet<(int, int)>) DefaultSubstitutionStarts()
    {
        var glyphs = new HashSet<int>();
        var pairs = new HashSet<(int, int)>();
        var t = _gsub;
        if (t.Length < 10) return (glyphs, pairs);
        try
        {
            int scriptList = U16(t, 4), featureList = U16(t, 6), lookupList = U16(t, 8);
            var features = new SortedSet<int>();
            for (int i = 0; i < U16(t, scriptList); i++)
            {
                int script = scriptList + U16(t, scriptList + 2 + i * 6 + 4);
                if (U16(t, script) == 0) continue;
                int langSys = script + U16(t, script);
                if (U16(t, langSys + 2) != 0xFFFF) features.Add(U16(t, langSys + 2));
                for (int k = 0; k < U16(t, langSys + 4); k++) features.Add(U16(t, langSys + 6 + k * 2));
            }
            foreach (int f in features)
            {
                if (f >= U16(t, featureList)) continue;
                int rec = featureList + 2 + f * 6;
                if (Tag(t, rec) is not ("ccmp" or "locl" or "rlig" or "calt" or "clig" or "liga" or "rclt")) continue;
                int feature = featureList + U16(t, rec + 4);
                for (int k = 0; k < U16(t, feature + 2); k++)
                {
                    int lookup = U16(t, feature + 4 + k * 2);
                    if (lookup >= U16(t, lookupList)) continue;
                    int table = lookupList + U16(t, lookupList + 2 + lookup * 2);
                    int type = U16(t, table);
                    for (int sIndex = 0; sIndex < U16(t, table + 4); sIndex++)
                    {
                        int sub = table + U16(t, table + 6 + sIndex * 2), subType = type;
                        if (type == 7)
                        {
                            subType = U16(t, sub + 2);
                            sub += U32(t, sub + 4);
                        }
                        int coverage = subType is 5 or 6 && U16(t, sub) == 3
                            ? sub + U16(t, sub + (subType == 5 ? 6 : 4 + U16(t, sub + 2) * 2 + 2))
                            : sub + U16(t, sub + 2);
                        if (subType != 4)
                        {
                            glyphs.UnionWith(CoveredGlyphs(t, coverage));
                            continue;
                        }
                        int index = 0;
                        foreach (int first in CoveredGlyphs(t, coverage))
                        {
                            int set = sub + U16(t, sub + 6 + index++ * 2);
                            for (int l = 0; l < U16(t, set); l++)
                            {
                                int ligature = set + U16(t, set + 2 + l * 2);
                                if (U16(t, ligature + 2) < 2) glyphs.Add(first);
                                else pairs.Add((first, U16(t, ligature + 4)));
                            }
                        }
                    }
                }
            }
        }
        catch (ArgumentOutOfRangeException)
        {
            // malformed table: shape everything it covered so far
        }
        return (glyphs, pairs);
    }

    private static int ClassOf(byte[] t, int offset, int glyph)
    {
        if (offset <= 0 || offset >= t.Length) return 0;
        int format = U16(t, offset);
        if (format == 1)
        {
            int start = U16(t, offset + 2), count = U16(t, offset + 4);
            return glyph >= start && glyph < start + count ? U16(t, offset + 6 + (glyph - start) * 2) : 0;
        }
        if (format == 2)
        {
            int count = U16(t, offset + 2);
            int lo = 0, hi = count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1;
                int rec = offset + 4 + mid * 6;
                int start = U16(t, rec), end = U16(t, rec + 2);
                if (glyph < start) hi = mid - 1;
                else if (glyph > end) lo = mid + 1;
                else return U16(t, rec + 4);
            }
        }
        return 0;
    }

    /// <summary>GDEF glyph class: 1 base, 2 ligature, 3 mark, 4 component, 0 unknown.</summary>
    /// <summary>True when GDEF has a glyph class definition.</summary>
    public bool HasGlyphClasses => _gdef.Length >= 12 && U16(_gdef, 4) != 0;

    public int GlyphClass(int glyph) => _gdef.Length >= 12 ? ClassOf(_gdef, U16(_gdef, 4), glyph) : 0;

    private int MarkAttachClass(int glyph) => _gdef.Length >= 12 ? ClassOf(_gdef, U16(_gdef, 10), glyph) : 0;

    private bool InMarkSet(int set, int glyph)
    {
        // GDEF 1.2+: MarkGlyphSetsDef offset at 12.
        if (_gdef.Length < 14 || U16(_gdef, 2) < 2) return false;
        int sets = U16(_gdef, 12);
        if (sets == 0 || set >= U16(_gdef, sets + 2)) return false;
        return Coverage(_gdef, sets + U32(_gdef, sets + 4 + set * 4), glyph) >= 0;
    }

    // ------------------------------------------------------------------ feature selection

    /// <summary>Lookup indices of <paramref name="features"/> for a script (falling back to DFLT/latn), in lookup order.</summary>
    public List<(int Lookup, uint Mask)> Lookups(bool gpos, string script, IReadOnlyList<(string Tag, uint Mask)> features)
    {
        var t = gpos ? _gpos : _gsub;
        var result = new List<(int, uint)>();
        if (t.Length < 10) return result;
        int scriptList = U16(t, 4), featureList = U16(t, 6);
        int langSys = -1;
        foreach (var wanted in new[] { script, "DFLT", "latn" })
        {
            int count = U16(t, scriptList);
            for (int i = 0; i < count && langSys < 0; i++)
            {
                int rec = scriptList + 2 + i * 6;
                if (Tag(t, rec) != wanted) continue;
                int scriptTable = scriptList + U16(t, rec + 4);
                int defaultLangSys = U16(t, scriptTable);
                if (defaultLangSys != 0) langSys = scriptTable + defaultLangSys;
                else if (U16(t, scriptTable + 2) > 0) langSys = scriptTable + U16(t, scriptTable + 4 + 4);
            }
            if (langSys >= 0) break;
        }
        if (langSys < 0) return result;

        var masks = new SortedDictionary<int, uint>();
        int required = U16(t, langSys + 2);
        int featureCount = U16(t, langSys + 4);
        var indices = Enumerable.Range(0, featureCount).Select(i => U16(t, langSys + 6 + i * 2)).ToList();
        if (required != 0xFFFF) indices.Add(required);
        foreach (int index in indices)
        {
            if (index >= U16(t, featureList)) continue;
            int rec = featureList + 2 + index * 6;
            var tag = Tag(t, rec);
            foreach (var (featureTag, mask) in features)
            {
                if (featureTag != tag) continue;
                int feature = featureList + U16(t, rec + 4);
                int lookups = U16(t, feature + 2);
                for (int k = 0; k < lookups; k++)
                {
                    int lookup = U16(t, feature + 4 + k * 2);
                    masks[lookup] = masks.GetValueOrDefault(lookup) | mask;
                }
            }
        }
        foreach (var (lookup, mask) in masks) result.Add((lookup, mask));
        return result;
    }

    /// <summary>True when the GSUB (or GPOS) table has the script <paramref name="tag"/> (e.g. "dev2").</summary>
    public bool HasScript(string tag, bool gpos = false)
    {
        var t = gpos ? _gpos : _gsub;
        if (t.Length < 10) return false;
        int scriptList = U16(t, 4);
        for (int i = 0; i < U16(t, scriptList); i++)
            if (Tag(t, scriptList + 2 + i * 6) == tag) return true;
        return false;
    }

    /// <summary>
    /// Whether a GSUB lookup would substitute exactly the glyph sequence <paramref name="glyphs"/> (HarfBuzz's
    /// would_substitute). With <paramref name="zeroContext"/>, chained rules that need backtrack or lookahead
    /// glyphs do not count.
    /// </summary>
    public bool WouldSubstitute(int lookup, ReadOnlySpan<int> glyphs, bool zeroContext)
    {
        var t = _gsub;
        if (t.Length < 10 || glyphs.Length == 0) return false;
        int list = U16(t, 8);
        if (lookup >= U16(t, list)) return false;
        int table = list + U16(t, list + 2 + lookup * 2);
        int type = U16(t, table);
        for (int s = 0; s < U16(t, table + 4); s++)
        {
            int sub = table + U16(t, table + 6 + s * 2);
            int subType = type;
            if (type == 7)
            {
                subType = U16(t, sub + 2);
                sub += U32(t, sub + 4);
            }
            if (WouldSubstitute(t, subType, sub, glyphs, zeroContext)) return true;
        }
        return false;
    }

    private static bool WouldSubstitute(byte[] t, int type, int sub, ReadOnlySpan<int> glyphs, bool zeroContext)
    {
        int format = U16(t, sub);
        int coverage = Coverage(t, sub + U16(t, sub + 2), glyphs[0]);
        switch (type)
        {
            case 1 or 2 or 3:
                return glyphs.Length == 1 && coverage >= 0;
            case 4:
            {
                if (coverage < 0) return false;
                int set = sub + U16(t, sub + 6 + coverage * 2);
                for (int l = 0; l < U16(t, set); l++)
                {
                    int lig = set + U16(t, set + 2 + l * 2);
                    int components = U16(t, lig + 2);
                    if (components != glyphs.Length) continue;
                    bool all = true;
                    for (int c = 1; c < components && all; c++) all = U16(t, lig + 4 + (c - 1) * 2) == glyphs[c];
                    if (all) return true;
                }
                return false;
            }
            case 5 or 6:
                return WouldSubstituteContext(t, type == 6, format, sub, glyphs, zeroContext);
        }
        return false;
    }

    private static bool WouldSubstituteContext(byte[] t, bool chained, int format, int sub, ReadOnlySpan<int> glyphs, bool zeroContext)
    {
        if (format == 3)
        {
            int p = sub + 2;
            int backtrack = 0;
            if (chained)
            {
                backtrack = U16(t, p);
                p += 2 + backtrack * 2;
            }
            int inputCount = U16(t, p);
            int inputs = chained ? p + 2 : sub + 6;
            if (!chained) inputCount = U16(t, sub + 2);
            if (inputCount != glyphs.Length) return false;
            for (int k = 0; k < inputCount; k++)
                if (Coverage(t, sub + U16(t, inputs + k * 2), glyphs[k]) < 0) return false;
            if (!chained || !zeroContext) return true;
            int lookahead = U16(t, inputs + inputCount * 2);
            return backtrack == 0 && lookahead == 0;
        }

        int coverage = Coverage(t, sub + U16(t, sub + 2), glyphs[0]);
        if (coverage < 0 || format is not (1 or 2)) return false;
        int inputClasses = 0, sets;
        if (format == 2)
        {
            inputClasses = sub + U16(t, sub + (chained ? 6 : 4));
            sets = sub + (chained ? 10 : 6);
        }
        else sets = sub + 4;
        int setIndex = format == 2 ? ClassOf(t, inputClasses, glyphs[0]) : coverage;
        if (setIndex >= U16(t, sets) || U16(t, sets + 2 + setIndex * 2) == 0) return false;
        int set = sub + U16(t, sets + 2 + setIndex * 2);
        for (int r = 0; r < U16(t, set); r++)
        {
            int rule = set + U16(t, set + 2 + r * 2), q = rule;
            int backtrack = 0;
            if (chained)
            {
                backtrack = U16(t, q);
                q += 2 + backtrack * 2;
            }
            int inputCount = U16(t, q);
            int inputStart = chained ? q + 2 : q + 4;
            if (inputCount != glyphs.Length) continue;
            bool all = true;
            for (int k = 1; k < inputCount && all; k++)
            {
                int value = U16(t, inputStart + (k - 1) * 2);
                all = format == 1 ? glyphs[k] == value : ClassOf(t, inputClasses, glyphs[k]) == value;
            }
            if (!all) continue;
            if (!chained || !zeroContext) return true;
            int lookahead = U16(t, inputStart + (inputCount - 1) * 2);
            if (backtrack == 0 && lookahead == 0) return true;
        }
        return false;
    }

    // ------------------------------------------------------------------ lookup application

    private sealed class Context(byte[] table, List<GlyphInfo> glyphs, bool gpos, bool rightToLeft)
    {
        public byte[] Table { get; } = table;
        public List<GlyphInfo> Glyphs { get; } = glyphs;
        public bool Gpos { get; } = gpos;
        public bool RightToLeft { get; } = rightToLeft;
        public int Flags { get; set; }
        public int MarkSet { get; set; }
        public uint Mask { get; set; }
        public bool PerSyllable { get; set; }
        public bool ManualJoiners { get; set; }
        public int Nesting { get; set; }

        /// <summary>Lookup applications so far; bounds the work on malicious fonts.</summary>
        public int Operations { get; set; }
    }

    /// <summary>Applies <paramref name="lookups"/>; <paramref name="rightToLeft"/> is the text direction (for cursive attachment).</summary>
    public void Apply(bool gpos, List<GlyphInfo> glyphs, List<(int Lookup, uint Mask)> lookups, bool rightToLeft = false)
    {
        var t = gpos ? _gpos : _gsub;
        if (t.Length < 10) return;
        var ctx = new Context(t, glyphs, gpos, rightToLeft);
        foreach (var (lookup, mask) in lookups)
        {
            ctx.Mask = mask & ~(PerSyllable | ManualJoiners);
            ctx.PerSyllable = (mask & PerSyllable) != 0;
            ctx.ManualJoiners = (mask & ManualJoiners) != 0;
            if (!SetLookup(ctx, lookup)) continue;
            for (int i = 0; i < glyphs.Count;)
            {
                if (++ctx.Operations > MaxOperations) return;
                if ((glyphs[i].Mask & ctx.Mask) == 0 || Ignored(ctx, glyphs[i].Glyph)) { i++; continue; }
                i = ApplyLookup(ctx, lookup, i) is { } next && next > i ? next : i + 1;
            }
        }
    }

    /// <summary>Loads lookup flags into the context; false for a lookup index out of range.</summary>
    private static bool SetLookup(Context ctx, int lookup)
    {
        var t = ctx.Table;
        int list = U16(t, 8);
        if (lookup >= U16(t, list)) return false;
        int table = list + U16(t, list + 2 + lookup * 2);
        ctx.Flags = U16(t, table + 2);
        ctx.MarkSet = (ctx.Flags & 0x10) != 0 ? U16(t, table + 6 + U16(t, table + 4) * 2) : -1;
        return true;
    }

    private bool Ignored(Context ctx, int glyph)
    {
        int flags = ctx.Flags;
        if ((flags & 0x0E) == 0 && (flags & 0xFF10) == 0) return false;
        int cls = GlyphClass(glyph);
        if (cls == 1 && (flags & 0x2) != 0) return true;
        if (cls == 2 && (flags & 0x4) != 0) return true;
        if (cls != 3) return false;
        if ((flags & 0x8) != 0) return true;
        if ((flags & 0x10) != 0 && ctx.MarkSet >= 0) return !InMarkSet(ctx.MarkSet, glyph);
        if ((flags & 0xFF00) != 0) return MarkAttachClass(glyph) != flags >> 8;
        return false;
    }

    private enum Step { Match, NoMatch, Skip }

    /// <summary>
    /// The next glyph after <paramref name="i"/> that matches, or -1, following HarfBuzz's skippy iterator: glyphs
    /// the lookup flags ignore are skipped; an <paramref name="input"/> glyph (ligature component, context input)
    /// must carry the lookup's feature mask; per-syllable lookups stay in the syllable; default ignorables such as
    /// ZWJ are matched when <paramref name="matches"/> accepts them and skipped otherwise.
    /// </summary>
    private int Next(Context ctx, int i, bool input = false, Func<int, bool>? matches = null)
    {
        int start = i;
        for (i++; i < ctx.Glyphs.Count; i++)
        {
            var step = Test(ctx, start, i, input, matches);
            if (step != Step.Skip) return step == Step.Match ? i : -1;
        }
        return -1;
    }

    /// <summary>The previous glyph that matches, or -1; <paramref name="input"/> as for <see cref="Next"/> (false for backtrack context).</summary>
    private int Prev(Context ctx, int i, Func<int, bool>? matches = null, bool input = false)
    {
        int start = i;
        for (i--; i >= 0; i--)
        {
            var step = Test(ctx, start, i, input, matches);
            if (step != Step.Skip) return step == Step.Match ? i : -1;
        }
        return -1;
    }

    private Step Test(Context ctx, int from, int i, bool input, Func<int, bool>? matches)
    {
        var g = ctx.Glyphs[i];
        if (Ignored(ctx, g.Glyph)) return Step.Skip;
        bool maybeSkip = MaySkip(ctx, g, context: !input);
        bool? match = (input && (g.Mask & ctx.Mask) == 0) || (ctx.PerSyllable && !ctx.Gpos && g.Syllable != ctx.Glyphs[from].Syllable)
            ? false
            : matches?.Invoke(i);
        if (match == true || (match is null && !maybeSkip)) return Step.Match;
        return maybeSkip ? Step.Skip : Step.NoMatch;
    }

    /// <summary>
    /// Whether a default-ignorable glyph may be skipped: ZWNJ in GPOS (or in GSUB context with automatic joiners),
    /// ZWJ in context or with automatic joiners, hidden ones in GPOS, others always (HarfBuzz's may_skip).
    /// </summary>
    private static bool MaySkip(Context ctx, in GlyphInfo g, bool context)
    {
        if (!g.Ignorable || g.Hidden && !ctx.Gpos) return false;
        bool autoJoiners = !ctx.ManualJoiners;
        bool ignoreZwnj = ctx.Gpos || (context && autoJoiners), ignoreZwj = context || autoJoiners;
        return (ignoreZwnj || g.CodePoint != 0x200C) && (ignoreZwj || g.CodePoint != 0x200D);
    }

    /// <summary>Applies one lookup at position <paramref name="i"/>; returns where to continue, or null when nothing applied.</summary>
    private int? ApplyLookup(Context ctx, int lookup, int i)
    {
        var t = ctx.Table;
        int list = U16(t, 8);
        if (lookup >= U16(t, list)) return null;
        int table = list + U16(t, list + 2 + lookup * 2);
        int type = U16(t, table);
        int subtables = U16(t, table + 4);
        for (int s = 0; s < subtables; s++)
        {
            int sub = table + U16(t, table + 6 + s * 2);
            int subType = type;
            if ((ctx.Gpos && type == 9) || (!ctx.Gpos && type == 7))
            {
                subType = U16(t, sub + 2);
                sub += U32(t, sub + 4);
            }
            var next = ctx.Gpos ? ApplyPositioning(ctx, subType, sub, i) : ApplySubstitution(ctx, subType, sub, i);
            if (next is not null) return next;
        }
        return null;
    }

    private int? ApplySubstitution(Context ctx, int type, int sub, int i)
    {
        var t = ctx.Table;
        var glyphs = ctx.Glyphs;
        int glyph = glyphs[i].Glyph;
        int format = U16(t, sub);
        switch (type)
        {
            case 1:
            {
                int index = Coverage(t, sub + U16(t, sub + 2), glyph);
                if (index < 0) return null;
                int replacement = format == 1 ? (glyph + S16(t, sub + 4)) & 0xFFFF : U16(t, sub + 6 + index * 2);
                Replace(glyphs, i, (ushort)replacement, GlyphFlags.None);
                return i + 1;
            }
            case 2:
            {
                int index = Coverage(t, sub + U16(t, sub + 2), glyph);
                if (index < 0) return null;
                int sequence = sub + U16(t, sub + 6 + index * 2);
                int count = U16(t, sequence);
                if (count == 0) { glyphs.RemoveAt(i); return i; }
                var info = glyphs[i];
                var added = count > 1 ? GlyphFlags.Multiplied : GlyphFlags.None;
                Replace(glyphs, i, (ushort)U16(t, sequence + 2), added);
                for (int k = 1; k < count; k++)
                    glyphs.Insert(i + k, glyphs[i] with { Glyph = (ushort)U16(t, sequence + 2 + k * 2), Text = "" });
                return i + count;
            }
            case 3:
            {
                // An enabled feature (value 1) selects the first alternate, as in HarfBuzz.
                int index = Coverage(t, sub + U16(t, sub + 2), glyph);
                if (index < 0) return null;
                int set = sub + U16(t, sub + 6 + index * 2);
                if (U16(t, set) == 0) return null;
                Replace(glyphs, i, (ushort)U16(t, set + 2), GlyphFlags.None);
                return i + 1;
            }
            case 4:
            {
                int index = Coverage(t, sub + U16(t, sub + 2), glyph);
                if (index < 0) return null;
                int set = sub + U16(t, sub + 6 + index * 2);
                int ligatures = U16(t, set);
                for (int l = 0; l < ligatures; l++)
                {
                    int lig = set + U16(t, set + 2 + l * 2);
                    int components = U16(t, lig + 2);
                    var matched = new List<int>(components) { i };
                    int pos = i;
                    for (int c = 1; c < components; c++)
                    {
                        int component = U16(t, lig + 4 + (c - 1) * 2);
                        pos = Next(ctx, pos, input: true, j => glyphs[j].Glyph == component);
                        if (pos < 0) { matched = null; break; }
                        matched.Add(pos);
                    }
                    if (matched is null) continue;
                    // Marks skipped inside the match belong to the component before them.
                    int ligatureId = glyphs.Max(g => g.LigatureId) + 1;
                    for (int c = 1; c < matched.Count; c++)
                        for (int k = matched[c - 1] + 1; k < matched[c]; k++)
                            if (GlyphClass(glyphs[k].Glyph) == 3)
                                glyphs[k] = glyphs[k] with { LigatureId = ligatureId, LigatureComponent = c };
                    var ligature = glyphs[i];
                    ligature.LigatureId = ligatureId;
                    ligature.LigatureComponent = 0;
                    ligature.Glyph = (ushort)U16(t, lig);
                    ligature.Flags = (ligature.Flags | GlyphFlags.Substituted | GlyphFlags.Ligated) & ~GlyphFlags.Multiplied;
                    ligature.Text = string.Concat(matched.Select(k => glyphs[k].Text));
                    glyphs[i] = ligature;
                    for (int k = matched.Count - 1; k > 0; k--) glyphs.RemoveAt(matched[k]);
                    return i + 1;
                }
                return null;
            }
            case 5:
            case 6:
                return ApplyContext(ctx, type == 6, sub, i);
        }
        return null;
    }

    private static void Replace(List<GlyphInfo> glyphs, int i, ushort glyph, GlyphFlags added)
    {
        var info = glyphs[i];
        info.Glyph = glyph;
        info.Flags |= GlyphFlags.Substituted | added;
        glyphs[i] = info;
    }

    private int? ApplyPositioning(Context ctx, int type, int sub, int i)
    {
        var t = ctx.Table;
        var glyphs = ctx.Glyphs;
        int glyph = glyphs[i].Glyph;
        int format = U16(t, sub);
        switch (type)
        {
            case 1:
            {
                int index = Coverage(t, sub + U16(t, sub + 2), glyph);
                if (index < 0) return null;
                int valueFormat = U16(t, sub + 4);
                int value = format == 1 ? sub + 6 : sub + 8 + index * ValueSize(valueFormat);
                glyphs[i] = Adjust(glyphs[i], t, value, valueFormat);
                return i + 1;
            }
            case 2:
            {
                int index = Coverage(t, sub + U16(t, sub + 2), glyph);
                if (index < 0) return null;
                int second = Next(ctx, i);
                if (second < 0) return null;
                int vf1 = U16(t, sub + 4), vf2 = U16(t, sub + 6);
                int size1 = ValueSize(vf1), size2 = ValueSize(vf2);
                int record = -1;
                if (format == 1)
                {
                    int set = sub + U16(t, sub + 10 + index * 2);
                    int count = U16(t, set), stride = 2 + size1 + size2;
                    int lo = 0, hi = count - 1;
                    while (lo <= hi)
                    {
                        int mid = (lo + hi) >> 1;
                        int g = U16(t, set + 2 + mid * stride);
                        if (g == glyphs[second].Glyph) { record = set + 2 + mid * stride + 2; break; }
                        if (g < glyphs[second].Glyph) lo = mid + 1; else hi = mid - 1;
                    }
                }
                else if (format == 2)
                {
                    int class1 = ClassOf(t, sub + U16(t, sub + 8), glyph);
                    int class2 = ClassOf(t, sub + U16(t, sub + 10), glyphs[second].Glyph);
                    int count1 = U16(t, sub + 12), count2 = U16(t, sub + 14);
                    if (class1 < count1 && class2 < count2) record = sub + 16 + (class1 * count2 + class2) * (size1 + size2);
                }
                if (record < 0) return null;
                glyphs[i] = Adjust(glyphs[i], t, record, vf1);
                glyphs[second] = Adjust(glyphs[second], t, record + size1, vf2);
                return vf2 != 0 ? second + 1 : second;
            }
            case 3:
            {
                // Cursive attachment: this glyph's entry anchor meets the previous glyph's exit anchor (as in HarfBuzz).
                if (format != 1) return null;
                int coverage = sub + U16(t, sub + 2);
                int index = Coverage(t, coverage, glyph);
                if (index < 0) return null;
                int entry = U16(t, sub + 6 + index * 4);
                int prev = Prev(ctx, i, input: true);
                if (entry == 0 || prev < 0) return null;
                int prevIndex = Coverage(t, coverage, glyphs[prev].Glyph);
                if (prevIndex < 0) return null;
                int exit = U16(t, sub + 6 + prevIndex * 4 + 2);
                if (exit == 0) return null;
                int exitX = S16(t, sub + exit + 2), exitY = S16(t, sub + exit + 4);
                int entryX = S16(t, sub + entry + 2), entryY = S16(t, sub + entry + 4);

                GlyphInfo a = glyphs[prev], b = glyphs[i];
                if (ctx.RightToLeft)
                {
                    int d = exitX + a.XOffset;
                    a.XAdvance -= d;
                    a.XOffset -= d;
                    b.XAdvance = entryX + b.XOffset;
                }
                else
                {
                    a.XAdvance = exitX + a.XOffset;
                    int d = entryX + b.XOffset;
                    b.XAdvance -= d;
                    b.XOffset -= d;
                }
                // The RightToLeft lookup flag makes the last glyph of a chain sit on the baseline instead of the first.
                if ((ctx.Flags & 0x1) != 0) (a.CursiveTo, a.YOffset) = (i, entryY - exitY);
                else (b.CursiveTo, b.YOffset) = (prev, exitY - entryY);
                glyphs[prev] = a;
                glyphs[i] = b;
                return i + 1;
            }
            case 4:
            case 5:
            case 6:
            {
                int markIndex = Coverage(t, sub + U16(t, sub + 2), glyph);
                if (markIndex < 0) return null;
                // The attachment target: the previous base (4), ligature (5) or mark (6). As in HarfBuzz, the search
                // ignores all marks for 4/5 and only the lookup's mark filtering for 6; default ignorables are skipped.
                int flags = ctx.Flags;
                ctx.Flags = type != 6 ? 0x8 : flags & ~0x0E;
                int target;
                try
                {
                    target = Prev(ctx, i, input: true);
                }
                finally
                {
                    ctx.Flags = flags;
                }
                if (target < 0) return null;
                if (type == 6)
                {
                    if (GlyphClass(glyphs[target].Glyph) != 3) return null;
                    // Only marks on the same base, or the same ligature component, attach to each other.
                    var (a, b) = (glyphs[i], glyphs[target]);
                    bool good = a.LigatureId == b.LigatureId
                        ? a.LigatureId == 0 || a.LigatureComponent == b.LigatureComponent
                        : (a.LigatureId > 0 && a.LigatureComponent == 0) || (b.LigatureId > 0 && b.LigatureComponent == 0);
                    if (!good) return null;
                }
                int baseIndex = Coverage(t, sub + U16(t, sub + 4), glyphs[target].Glyph);
                if (baseIndex < 0) return null;
                int classCount = U16(t, sub + 6);
                int markArray = sub + U16(t, sub + 8);
                int baseArray = sub + U16(t, sub + 10);
                int markClass = U16(t, markArray + 2 + markIndex * 4);
                int markAnchor = markArray + U16(t, markArray + 2 + markIndex * 4 + 2);
                if (markClass >= classCount) return null;
                int baseAnchor;
                if (type == 5)
                {
                    // Attach to the mark's own component, or else the last one.
                    int attach = baseArray + U16(t, baseArray + 2 + baseIndex * 2);
                    int components = U16(t, attach);
                    if (components == 0) return null;
                    var mark = glyphs[i];
                    int component = mark.LigatureId != 0 && mark.LigatureId == glyphs[target].LigatureId && mark.LigatureComponent > 0
                        ? Math.Min(components, mark.LigatureComponent) - 1
                        : components - 1;
                    int offset = U16(t, attach + 2 + (component * classCount + markClass) * 2);
                    if (offset == 0) return null;
                    baseAnchor = attach + offset;
                }
                else
                {
                    int offset = U16(t, baseArray + 2 + (baseIndex * classCount + markClass) * 2);
                    if (offset == 0) return null;
                    baseAnchor = baseArray + offset;
                }
                var info = glyphs[i];
                info.XOffset = S16(t, baseAnchor + 2) - S16(t, markAnchor + 2);
                // The base's own vertical offset (along its cursive chain) is taken now, as in HarfBuzz; only the
                // horizontal offset is propagated when attachments are resolved.
                int baseOffset = glyphs[target].YOffset;
                for (int b = target, n = 1; glyphs[b].CursiveTo >= 0 && n < glyphs.Count; n++)
                {
                    b = glyphs[b].CursiveTo;
                    baseOffset += glyphs[b].YOffset;
                }
                info.YOffset = S16(t, baseAnchor + 4) - S16(t, markAnchor + 4) + baseOffset;
                info.AttachTo = target;
                glyphs[i] = info;
                return i + 1;
            }
            case 7:
            case 8:
                return ApplyContext(ctx, type == 8, sub, i);
        }
        return null;
    }

    private static int ValueSize(int format) => System.Numerics.BitOperations.PopCount((uint)format & 0xFF) * 2;

    private static GlyphInfo Adjust(GlyphInfo info, byte[] t, int offset, int format)
    {
        if ((format & 0x1) != 0) { info.XOffset += S16(t, offset); offset += 2; }
        if ((format & 0x2) != 0) { info.YOffset += S16(t, offset); offset += 2; }
        if ((format & 0x4) != 0) info.XAdvance += S16(t, offset);
        return info;
    }

    // ------------------------------------------------------------------ contextual lookups

    /// <summary>Sequence/chained-sequence context (GSUB 5/6, GPOS 7/8), formats 1–3.</summary>
    private int? ApplyContext(Context ctx, bool chained, int sub, int i)
    {
        var t = ctx.Table;
        int format = U16(t, sub);
        int glyph = ctx.Glyphs[i].Glyph;
        if (format == 3) return chained ? ChainedFormat3(ctx, sub, i) : Format3(ctx, sub, i);

        int coverage = Coverage(t, sub + U16(t, sub + 2), glyph);
        if (coverage < 0) return null;
        // Format 1 matches glyph ids, format 2 classes.
        int backtrackClasses = 0, inputClasses = 0, lookaheadClasses = 0, sets;
        if (format == 2)
        {
            if (chained)
            {
                backtrackClasses = sub + U16(t, sub + 4);
                inputClasses = sub + U16(t, sub + 6);
                lookaheadClasses = sub + U16(t, sub + 8);
                sets = sub + 10;
            }
            else
            {
                inputClasses = sub + U16(t, sub + 4);
                sets = sub + 6;
            }
        }
        else if (format == 1) sets = sub + 4;
        else return null;

        int setIndex = format == 2 ? ClassOf(t, inputClasses, glyph) : coverage;
        if (setIndex >= U16(t, sets)) return null;
        int setOffset = U16(t, sets + 2 + setIndex * 2);
        if (setOffset == 0) return null;
        int set = sub + setOffset;
        int rules = U16(t, set);
        for (int r = 0; r < rules; r++)
        {
            int rule = set + U16(t, set + 2 + r * 2);
            bool Match(int pos, int value, int classDef) =>
                format == 1 ? ctx.Glyphs[pos].Glyph == value : ClassOf(t, classDef, ctx.Glyphs[pos].Glyph) == value;

            int p = rule;
            if (chained)
            {
                int backtrack = U16(t, p);
                int pos = i;
                bool ok = true;
                for (int k = 0; k < backtrack && ok; k++)
                {
                    int value = U16(t, p + 2 + k * 2);
                    pos = Prev(ctx, pos, j => Match(j, value, backtrackClasses));
                    ok = pos >= 0;
                }
                if (!ok) continue;
                p += 2 + backtrack * 2;
            }
            int inputCount = U16(t, p);
            int records = chained ? 0 : U16(t, p + 2);
            int inputStart = chained ? p + 2 : p + 4;
            var matched = MatchInput(ctx, i, inputCount, checkFirst: false, (pos, k) => Match(pos, U16(t, inputStart + (k - 1) * 2), inputClasses));
            if (matched is null) continue;
            p = inputStart + (inputCount - 1) * 2;
            if (chained)
            {
                int lookahead = U16(t, p);
                int pos = matched[^1];
                bool ok = true;
                for (int k = 0; k < lookahead && ok; k++)
                {
                    int value = U16(t, p + 2 + k * 2);
                    pos = Next(ctx, pos, matches: j => Match(j, value, lookaheadClasses));
                    ok = pos >= 0;
                }
                if (!ok) continue;
                p += 2 + lookahead * 2;
                records = U16(t, p);
                p += 2;
            }
            return ApplyRecords(ctx, t, p, records, matched);
        }
        return null;
    }

    private int? Format3(Context ctx, int sub, int i)
    {
        var t = ctx.Table;
        int inputCount = U16(t, sub + 2), records = U16(t, sub + 4);
        var matched = MatchInput(ctx, i, inputCount, checkFirst: true, (pos, k) => Coverage(t, sub + U16(t, sub + 6 + k * 2), ctx.Glyphs[pos].Glyph) >= 0);
        if (matched is null) return null;
        return ApplyRecords(ctx, t, sub + 6 + inputCount * 2, records, matched);
    }

    private int? ChainedFormat3(Context ctx, int sub, int i)
    {
        var t = ctx.Table;
        int p = sub + 2;
        int backtrack = U16(t, p);
        int pos = i;
        for (int k = 0; k < backtrack; k++)
        {
            int coverage = sub + U16(t, p + 2 + k * 2);
            pos = Prev(ctx, pos, j => Coverage(t, coverage, ctx.Glyphs[j].Glyph) >= 0);
            if (pos < 0) return null;
        }
        p += 2 + backtrack * 2;
        int inputCount = U16(t, p);
        int inputs = p + 2;
        var matched = MatchInput(ctx, i, inputCount, checkFirst: true, (q, k) => Coverage(t, sub + U16(t, inputs + k * 2), ctx.Glyphs[q].Glyph) >= 0);
        if (matched is null) return null;
        p = inputs + inputCount * 2;
        int lookahead = U16(t, p);
        pos = matched[^1];
        for (int k = 0; k < lookahead; k++)
        {
            int coverage = sub + U16(t, p + 2 + k * 2);
            pos = Next(ctx, pos, matches: j => Coverage(t, coverage, ctx.Glyphs[j].Glyph) >= 0);
            if (pos < 0) return null;
        }
        p += 2 + lookahead * 2;
        return ApplyRecords(ctx, t, p + 2, U16(t, p), matched);
    }

    /// <summary>
    /// Matches <paramref name="count"/> input glyphs starting at <paramref name="i"/>. Formats 1/2 already matched the
    /// first glyph through the coverage table; format 3 checks it too (<paramref name="checkFirst"/>, k = 0).
    /// </summary>
    private List<int>? MatchInput(Context ctx, int i, int count, bool checkFirst, Func<int, int, bool> matches)
    {
        if (count == 0) return null;
        if (checkFirst && !matches(i, 0)) return null;
        var matched = new List<int>(count) { i };
        int pos = i;
        for (int k = 1; k < count; k++)
        {
            int index = k;
            pos = Next(ctx, pos, input: true, j => matches(j, index));
            if (pos < 0) return null;
            matched.Add(pos);
        }
        return matched;
    }

    /// <summary>Applies SequenceLookupRecords to the matched glyphs; returns the position after the match.</summary>
    private int? ApplyRecords(Context ctx, byte[] t, int records, int count, List<int> matched)
    {
        int end = matched[^1] + 1;
        if (ctx.Nesting >= MaxNesting) return end;
        int flags = ctx.Flags, markSet = ctx.MarkSet;
        ctx.Nesting++;
        try
        {
            for (int r = 0; r < count; r++)
            {
                int sequenceIndex = U16(t, records + r * 4), lookup = U16(t, records + r * 4 + 2);
                if (sequenceIndex >= matched.Count || !SetLookup(ctx, lookup)) continue;
                int before = ctx.Glyphs.Count;
                ApplyLookup(ctx, lookup, matched[sequenceIndex]);
                int delta = ctx.Glyphs.Count - before;
                if (delta == 0) continue;
                // The glyph count changed (as in HarfBuzz's apply_lookup): glyphs added by a multiple substitution
                // become matched positions right after it, glyphs consumed by a ligature drop out, later ones shift.
                end += delta;
                int next = sequenceIndex + 1;
                if (delta > 0)
                    matched.InsertRange(next, Enumerable.Range(matched[sequenceIndex] + 1, delta));
                else
                {
                    int removed = Math.Min(-delta, matched.Count - next);
                    matched.RemoveRange(next, removed);
                    delta = -removed;
                }
                for (int k = next + Math.Max(delta, 0); k < matched.Count; k++) matched[k] += delta;
            }
        }
        finally
        {
            ctx.Nesting--;
            ctx.Flags = flags;
            ctx.MarkSet = markSet;
        }
        return Math.Max(end, matched[0] + 1);
    }
}
