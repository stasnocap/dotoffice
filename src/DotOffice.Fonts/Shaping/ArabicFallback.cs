namespace DotOffice.Fonts.Shaping;

/// <summary>
/// HarfBuzz's Arabic fallback shaping (hb-ot-shaper-arabic-fallback.cc): for a font with no joining or rlig GSUB
/// features, a substitution is synthesized from the Arabic Presentation Forms characters the font maps in its cmap.
/// </summary>
internal sealed class ArabicFallback
{
    private readonly (ushort[] Map, uint Mask)[] _singles; // glyph -> presentation glyph, per joining form
    private readonly (Dictionary<int, List<(ushort Ligature, ushort[] Components)>> Sets, bool IgnoreMarks)[] _ligatures;

    private ArabicFallback((ushort[], uint)[] singles, (Dictionary<int, List<(ushort, ushort[])>>, bool)[] ligatures)
    {
        _singles = singles;
        _ligatures = ligatures;
    }

    /// <summary>Builds the fallback plan, or null when the font has none of the presentation forms.</summary>
    public static ArabicFallback? Create(TrueTypeFont font, uint isol, uint fina, uint medi, uint init)
    {
        var singles = new List<(ushort[], uint)>();
        // HarfBuzz order of shaping_table columns: init, medi, fina, isol.
        uint[] masks = [init, medi, fina, isol];
        for (int column = 0; column < 4; column++)
        {
            var map = new ushort[65536];
            bool any = false;
            for (int u = ArabicFallbackData.First; u <= ArabicFallbackData.Last; u++)
            {
                int s = ArabicFallbackData.Shaping[u - ArabicFallbackData.First][column];
                if (s == 0) continue;
                int ug = Normalizer.NominalGlyph(font, u), sg = Normalizer.NominalGlyph(font, s);
                if (ug == 0 || sg == 0 || ug == sg) continue;
                map[ug] = (ushort)sg;
                any = true;
            }
            if (any) singles.Add((map, masks[column]));
        }

        var ligatures = new List<(Dictionary<int, List<(ushort, ushort[])>>, bool)>();
        foreach (var (table, ignoreMarks) in new[]
                 {
                     (ArabicFallbackData.Ligatures3, true), (ArabicFallbackData.Ligatures2, true),
                     (ArabicFallbackData.LigaturesMark, false),
                 })
        {
            var sets = new Dictionary<int, List<(ushort, ushort[])>>();
            foreach (var (first, entries) in table)
            {
                int fg = Normalizer.NominalGlyph(font, first);
                if (fg == 0) continue;
                var list = new List<(ushort, ushort[])>();
                foreach (var (ligature, components) in entries)
                {
                    int lg = Normalizer.NominalGlyph(font, ligature);
                    if (lg == 0) continue;
                    var comps = new ushort[components.Length];
                    bool ok = true;
                    for (int k = 0; k < comps.Length && ok; k++)
                    {
                        int cg = Normalizer.NominalGlyph(font, components[k]);
                        ok = cg != 0;
                        comps[k] = (ushort)cg;
                    }
                    if (ok) list.Add(((ushort)lg, comps));
                }
                if (list.Count > 0) sets[fg] = list;
            }
            if (sets.Count > 0) ligatures.Add((sets, ignoreMarks));
        }
        return singles.Count == 0 && ligatures.Count == 0 ? null : new ArabicFallback(singles.ToArray(), ligatures.ToArray());
    }

    public void Apply(OpenTypeLayout layout, List<GlyphInfo> glyphs, uint ligatureMask)
    {
        foreach (var (map, mask) in _singles)
            for (int i = 0; i < glyphs.Count; i++)
            {
                var g = glyphs[i];
                if ((g.Mask & mask) == 0 || map[g.Glyph] == 0) continue;
                g.Glyph = map[g.Glyph];
                g.Flags |= GlyphFlags.Substituted;
                glyphs[i] = g;
            }

        foreach (var (sets, ignoreMarks) in _ligatures)
            for (int i = 0; i < glyphs.Count; i++)
            {
                if ((glyphs[i].Mask & ligatureMask) == 0 || !sets.TryGetValue(glyphs[i].Glyph, out var set)) continue;
                foreach (var (ligature, components) in set)
                {
                    var matched = new List<int> { i };
                    int pos = i;
                    foreach (var component in components)
                    {
                        pos++;
                        while (pos < glyphs.Count && (glyphs[pos].Ignorable || ignoreMarks && layout.ClassOf(glyphs[pos]) == 3)) pos++;
                        if (pos >= glyphs.Count || glyphs[pos].Glyph != component || (glyphs[pos].Mask & ligatureMask) == 0) { matched = null; break; }
                        matched.Add(pos);
                    }
                    if (matched is null) continue;
                    var info = glyphs[i];
                    info.Glyph = ligature;
                    info.Guess = 2;
                    info.Flags = (info.Flags | GlyphFlags.Substituted | GlyphFlags.Ligated) & ~GlyphFlags.Multiplied;
                    info.Text = string.Concat(matched.Select(k => glyphs[k].Text));
                    glyphs[i] = info;
                    for (int k = matched.Count - 1; k > 0; k--) glyphs.RemoveAt(matched[k]);
                    break;
                }
            }
    }
}
