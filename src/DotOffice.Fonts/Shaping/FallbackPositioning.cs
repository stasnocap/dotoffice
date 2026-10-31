using System.Globalization;

namespace DotOffice.Fonts.Shaping;

/// <summary>
/// Mark positioning for fonts without GPOS, ported from HarfBuzz's hb-ot-shape-fallback.cc: marks are placed
/// around their base's ink box by their (recategorized) combining class.
/// </summary>
internal static class FallbackPositioning
{
    private const int AttachedBelowLeft = 200, AttachedBelow = 202, AttachedAbove = 214, AttachedAboveRight = 216,
        BelowLeft = 218, Below = 220, BelowRight = 222, AboveLeft = 228, Above = 230, AboveRight = 232,
        DoubleBelow = 233, DoubleAbove = 234;

    private struct Extents
    {
        public int XBearing, YBearing, Width, Height;
    }

    /// <summary>
    /// Zeroes the advance of marks (by GDEF class, or Mn when the font has no glyph classes), moving them back by it
    /// in left-to-right text (HarfBuzz's zero_mark_widths_by_gdef with adjust_offsets_when_zeroing).
    /// </summary>
    public static void ZeroMarkWidths(OpenTypeLayout layout, List<GlyphInfo> glyphs, bool rightToLeft)
    {
        for (int i = 0; i < glyphs.Count; i++)
        {
            var g = glyphs[i];
            bool mark = layout.HasGlyphClasses ? layout.GlyphClass(g.Glyph) == 3
                : g.CodePoint > 0 && CharUnicodeInfo.GetUnicodeCategory(g.CodePoint) == UnicodeCategory.NonSpacingMark;
            if (!mark) continue;
            glyphs[i] = g with { XOffset = rightToLeft ? g.XOffset : g.XOffset - g.XAdvance, XAdvance = 0 };
        }
    }

    /// <summary>Positions the marks after each base (HarfBuzz's _hb_ot_shape_fallback_mark_position).</summary>
    public static void PositionMarks(TrueTypeFont font, List<GlyphInfo> glyphs, bool rightToLeft)
    {
        int start = 0;
        for (int i = 1; i < glyphs.Count; i++)
            if (!IsMark(glyphs[i]) && !glyphs[i].Hidden && !glyphs[i].Ignorable)
            {
                PositionCluster(font, glyphs, start, i, rightToLeft);
                start = i;
            }
        PositionCluster(font, glyphs, start, glyphs.Count, rightToLeft);
    }

    private static bool IsMark(in GlyphInfo g) => g.CodePoint > 0 && TextShaper.IsMark(g.CodePoint);

    private static void PositionCluster(TrueTypeFont font, List<GlyphInfo> glyphs, int start, int end, bool rightToLeft)
    {
        if (end - start < 2) return;
        for (int i = start; i < end; i++)
        {
            if (IsMark(glyphs[i])) continue;
            int j = i + 1;
            for (; j < end; j++)
            {
                if (glyphs[j].Hidden || glyphs[j].Ignorable) continue;
                if (!IsMark(glyphs[j])) break;
            }
            PositionAroundBase(font, glyphs, i, j, rightToLeft);
            i = j - 1;
        }
    }

    private static void PositionAroundBase(TrueTypeFont font, List<GlyphInfo> glyphs, int b, int end, bool rightToLeft)
    {
        var base_ = glyphs[b];
        if (font.GetGlyphExtents(base_.Glyph) is not { } e)
        {
            // Without extents, marks only lose their width.
            for (int i = b + 1; i < end; i++)
                if (glyphs[i].CodePoint > 0 && CharUnicodeInfo.GetUnicodeCategory(glyphs[i].CodePoint) == UnicodeCategory.NonSpacingMark)
                    glyphs[i] = glyphs[i] with { XOffset = rightToLeft ? glyphs[i].XOffset : glyphs[i].XOffset - glyphs[i].XAdvance, XAdvance = 0 };
            return;
        }
        // Horizontally the base's advance is used rather than its ink (also works for zero-ink glyphs).
        var baseExtents = new Extents { XBearing = 0, YBearing = e.YBearing + base_.YOffset, Width = font.GetAdvanceWidth(base_.Glyph), Height = e.Height };

        int ligatureId = base_.LigatureId;
        int components = (base_.Flags & GlyphFlags.Ligated) != 0
            ? Math.Max(1, base_.Text.EnumerateRunes().Count(r => !TextShaper.IsMark(r.Value))) : 1;

        int xOffset = rightToLeft ? 0 : -base_.XAdvance;
        var componentExtents = baseExtents;
        var clusterExtents = baseExtents;
        int lastComponent = -1, lastClass = 255;
        for (int i = b + 1; i < end; i++)
        {
            var g = glyphs[i];
            int cls = ModifiedClass(g);
            if (cls == 0)
            {
                xOffset += rightToLeft ? g.XAdvance : -g.XAdvance;
                continue;
            }
            if (components > 1)
            {
                int component = g.LigatureComponent - 1;
                if (ligatureId == 0 || ligatureId != g.LigatureId || component >= components) component = components - 1;
                if (lastComponent != component)
                {
                    lastComponent = component;
                    lastClass = 255;
                    componentExtents = baseExtents;
                    componentExtents.XBearing += (rightToLeft ? components - 1 - component : component) * componentExtents.Width / components;
                    componentExtents.Width /= components;
                }
            }
            if (lastClass != cls)
            {
                lastClass = cls;
                clusterExtents = componentExtents;
            }
            var (x, y) = PositionMark(font, g, ref clusterExtents, cls, rightToLeft) ?? (g.XOffset, g.YOffset);
            glyphs[i] = g with { XAdvance = 0, XOffset = x + xOffset, YOffset = y };
        }
    }

    private static (int X, int Y)? PositionMark(TrueTypeFont font, in GlyphInfo g, ref Extents base_, int cls, bool rightToLeft)
    {
        if (font.GetGlyphExtents(g.Glyph) is not { } mark) return null;
        int yGap = font.UnitsPerEm / 16;
        int x = 0, y = 0;

        // Left and right marks are not positioned.
        switch (cls)
        {
            case DoubleBelow or DoubleAbove when !rightToLeft:
                x = base_.XBearing + base_.Width - mark.Width / 2 - mark.XBearing;
                break;
            case DoubleBelow or DoubleAbove:
                x = base_.XBearing - mark.Width / 2 - mark.XBearing;
                break;
            case AttachedBelowLeft or BelowLeft or AboveLeft:
                x = base_.XBearing - mark.XBearing;
                break;
            case AttachedAboveRight or BelowRight or AboveRight:
                x = base_.XBearing + base_.Width - mark.Width - mark.XBearing;
                break;
            default:
                x = base_.XBearing + (base_.Width - mark.Width) / 2 - mark.XBearing;
                break;
        }

        switch (cls)
        {
            case DoubleBelow or BelowLeft or Below or BelowRight or AttachedBelowLeft or AttachedBelow:
                if (cls is not (AttachedBelowLeft or AttachedBelow)) base_.Height -= yGap;
                y = base_.YBearing + base_.Height - mark.YBearing;
                // Below marks are never shifted up.
                if (yGap > 0 == y > 0)
                {
                    base_.Height -= y;
                    y = 0;
                }
                base_.Height += mark.Height;
                break;
            case DoubleAbove or AboveLeft or Above or AboveRight or AttachedAbove or AttachedAboveRight:
                if (cls is not (AttachedAbove or AttachedAboveRight))
                {
                    base_.YBearing += yGap;
                    base_.Height -= yGap;
                }
                y = base_.YBearing - mark.YBearing - mark.Height;
                // Above marks are not shifted down too much.
                if (yGap > 0 != y > 0)
                {
                    int correction = -(y / 2);
                    base_.YBearing += correction;
                    base_.Height -= correction;
                    y += correction;
                }
                base_.YBearing -= mark.Height;
                base_.Height += mark.Height;
                break;
        }
        return (x, y);
    }

    /// <summary>The modified combining class of a mark, recategorized for positioning (HarfBuzz's recategorize_combining_class).</summary>
    private static int ModifiedClass(in GlyphInfo g)
    {
        if (!IsMark(g)) return 0;
        int cp = g.CodePoint, cls = Normalizer.CombiningClass(cp);
        if (CharUnicodeInfo.GetUnicodeCategory(cp) != UnicodeCategory.NonSpacingMark || cls >= 200) return cls;

        // Thai and Lao need some per-character work.
        if ((cp & ~0xFF) == 0x0E00)
        {
            if (cls == 0)
                cls = cp switch
                {
                    0x0E31 or 0x0E34 or 0x0E35 or 0x0E36 or 0x0E37 or 0x0E47 or 0x0E4C or 0x0E4D or 0x0E4E => AboveRight,
                    0x0EB1 or 0x0EB4 or 0x0EB5 or 0x0EB6 or 0x0EB7 or 0x0EBB or 0x0ECC or 0x0ECD => Above,
                    0x0EBC => Below,
                    _ => 0,
                };
            else if (cp == 0x0E3A) cls = BelowRight;
            if (cls >= 200) return cls;
        }

        return Normalizer.RawCombiningClass(cp) switch
        {
            // Hebrew: sheva … qubuts, meteg; rafe; shin dot; sin dot and holam; point varika (dagesh stays)
            >= 10 and <= 18 or 20 or 22 => Below,
            23 => AttachedAbove,
            24 => AboveRight,
            25 or 19 => AboveLeft,
            26 => Above,
            // Arabic and Syriac
            27 or 28 or 30 or 31 or 33 or 34 or 35 or 36 => Above,
            29 or 32 => Below,
            // Thai, Lao, Tibetan
            103 => BelowRight,
            107 => AboveRight,
            118 => Below,
            122 => Above,
            129 => Below,
            130 => Above,
            132 => Below,
            _ => cls,
        };
    }
}
