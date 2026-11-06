using DotOffice.Fonts;
using DotOffice.Layout;

namespace DotOffice.Tests.Layout;

public class FontMetricsTests
{
    private static IEnumerable<TrueTypeFont> Fonts() =>
        Directory.GetFiles(Path.GetDirectoryName(TestAssets.Path("NotoSans-Regular.ttf"))!, "*.ttf").Select(TrueTypeFont.Load);

    /// <summary>Noto Sans with its USE_TYPO_METRICS bit (OS/2 fsSelection bit 7) cleared.</summary>
    private static TrueTypeFont WithoutTypoFlag()
    {
        var data = File.ReadAllBytes(TestAssets.Path("NotoSans-Regular.ttf"));
        int tables = data[4] << 8 | data[5];
        for (int i = 0; i < tables; i++)
        {
            int record = 12 + 16 * i;
            if (System.Text.Encoding.ASCII.GetString(data, record, 4) != "OS/2") continue;
            int offset = data[record + 8] << 24 | data[record + 9] << 16 | data[record + 10] << 8 | data[record + 11];
            data[offset + 63] &= 0x7F; // fsSelection low byte
        }
        return TrueTypeFont.Load(data);
    }

    [Fact]
    public void A_font_without_USE_TYPO_METRICS_takes_its_win_metrics_and_external_leading()
    {
        var font = WithoutTypoFlag();
        Assert.False(font.UseTypoMetrics);
        Assert.True(font.WinAscent + font.WinDescent > 0);
        double scale = 100.0 / font.UnitsPerEm;
        var m = FontMetrics.Of(font, 100);
        double leading = Math.Max(0, font.LineGap - (font.WinAscent + font.WinDescent - (font.Ascender - font.Descender)));
        Assert.Equal(font.WinAscent * scale, m.Ascent, 0.05);
        Assert.Equal((font.WinDescent + leading) * scale, m.Descent, 0.1);
    }

    [Fact]
    public void USE_TYPO_METRICS_does_not_change_the_line_height()
    {
        var typo = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));
        Assert.True(typo.UseTypoMetrics);
        Assert.Equal(FontMetrics.Of(WithoutTypoFlag(), 40), FontMetrics.Of(typo, 40));
    }
}
