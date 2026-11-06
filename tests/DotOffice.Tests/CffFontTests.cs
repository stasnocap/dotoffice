using DotOffice.Fonts;

namespace DotOffice.Tests;

/// <summary>
/// CFF-outline (OpenType, 'OTTO') fonts. The test asset set has none, so these use the system's Noto Sans CJK when it is installed
/// and pass without checking anything when it is not.
/// </summary>
public class CffFontTests
{
    private static TrueTypeFont? Japanese()
    {
        foreach (var dir in new[] { "/usr/share/fonts/noto-cjk", "/usr/share/fonts/opentype/noto", "/usr/share/fonts/truetype/noto" })
        {
            if (!Directory.Exists(dir)) continue;
            foreach (var file in Directory.EnumerateFiles(dir, "NotoSansCJK-Regular.*"))
            {
                var faces = FontFaceInfo.TryReadAll(file);
                if (faces.FirstOrDefault(f => f.FamilyName == "Noto Sans CJK JP") is { } face) return TrueTypeFont.Load(file, face.FaceIndex);
            }
        }
        return null;
    }

    [Fact]
    public void A_CFF_font_loads_with_its_metrics_and_cmap()
    {
        if (Japanese() is not { } font) return;
        Assert.True(font.IsCff);
        Assert.NotEqual(0, font.GetGlyphId('日'));
        Assert.Equal(font.UnitsPerEm, font.GetAdvanceWidth(font.GetGlyphId('日')));
    }

    [Fact]
    public void The_subset_keeps_the_used_glyphs_and_the_character_ids_and_drops_the_subroutines()
    {
        if (Japanese() is not { } font) return;
        var used = new[] { font.GetGlyphId('日'), font.GetGlyphId('本') };
        var subset = font.SubsetCff(used);

        Assert.True(subset.Length < 1_000_000, $"subset is {subset.Length} bytes");
        var cut = new CffFont(subset, font.NumGlyphs);
        Assert.True(cut.IsCid);
        for (ushort g = 0; g < font.NumGlyphs; g += 97) Assert.Equal(font.GetCharacterId(g), cut.Cid(g));
        Assert.All(used, g => Assert.True(cut.CharStringLength(g) > 1));
        Assert.Equal(1, cut.CharStringLength(font.GetGlyphId('語')));
    }
}
