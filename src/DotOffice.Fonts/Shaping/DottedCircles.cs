namespace DotOffice.Fonts.Shaping;

/// <summary>hb_syllabic_insert_dotted_circles, shared by the shapers that scan syllables.</summary>
internal static class DottedCircles
{
    /// <summary>
    /// A broken cluster (syllable of type <paramref name="brokenType"/>) gets a dotted circle for its marks to sit
    /// on: at its start, or after the leading glyphs of category <paramref name="rephaCategory"/> (-1: none). Does
    /// nothing when the font has no dotted circle.
    /// </summary>
    public static void Insert(TrueTypeFont font, List<GlyphInfo> glyphs, int brokenType, byte circleCategory, int rephaCategory = -1)
    {
        int circle = font.GetGlyphId(0x25CC);
        if (circle == 0 || !glyphs.Exists(g => (g.Syllable & 0xF) == brokenType)) return;
        int last = 0;
        for (int i = 0; i < glyphs.Count; i++)
        {
            var g = glyphs[i];
            if (g.Syllable == last || (g.Syllable & 0xF) != brokenType) continue;
            last = g.Syllable;
            int at = i;
            while (at < glyphs.Count && glyphs[at].Syllable == last && glyphs[at].ShaperCategory == rephaCategory) at++;
            glyphs.Insert(at, g with
            {
                Glyph = (ushort)circle, Text = "", CodePoint = 0x25CC, ShaperCategory = circleCategory, Flags = GlyphFlags.None,
                Ignorable = false, Guess = 1, LigatureId = 0, LigatureComponent = 0,
            });
        }
    }
}
