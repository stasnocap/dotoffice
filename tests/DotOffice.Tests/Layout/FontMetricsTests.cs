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
        Assert.Equal(font.WinAscent * scale, m.Ascent, 1e-9);
        Assert.Equal((font.WinDescent + leading) * scale, m.Descent, 1e-9);
    }

    /// <summary>
    /// Word 2016 does not round the metrics to twips: its single-spaced pitch over 40 lines is the unrounded sum (Calibri 11 13.4297 =
    /// 2500/2048 em, not 13.40; Times New Roman 10 11.4997 = 2355/2048 em, not 11.45).
    /// </summary>
    [Theory]
    [InlineData(10)]
    [InlineData(10.5)]
    [InlineData(11)]
    [InlineData(12)]
    public void Metrics_are_not_rounded_to_twips(double size)
    {
        foreach (var font in Fonts().Where(f => f.WinAscent + f.WinDescent > 0 && f.Ascender - f.Descender > 0))
        {
            double leading = Math.Max(0, font.LineGap - (font.WinAscent + font.WinDescent - (font.Ascender - font.Descender)));
            var m = FontMetrics.Of(font, size);
            Assert.Equal(font.WinAscent * size / font.UnitsPerEm, m.Ascent, 1e-9);
            Assert.Equal((font.WinAscent + font.WinDescent + leading) * size / font.UnitsPerEm, m.Height, 1e-9);
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(259 / 240.0)]
    [InlineData(2)]
    public void Line_pitch_is_the_unrounded_height_times_the_spacing_factor(double factor)
    {
        // 12 pt: the unrounded Noto Sans height is 18.228, where twip-rounded ascent and descent gave 18.25.
        var run = LayoutTestHelpers.Text with { FontSize = 12 };
        var props = DotOffice.Core.ParagraphProperties.Default with { LineSpacing = new DotOffice.Core.LineSpacing(factor, DotOffice.Core.LineSpacingRule.Auto) };
        var blocks = Enumerable.Range(1, 40).Select(i => (DotOffice.Core.Block)LayoutTestHelpers.Para($"s{i:00}", props, run)).ToArray();
        var section = new DotOffice.Core.Section { PageHeight = 3000 };
        var lines = LayoutTestHelpers.Lines(LayoutTestHelpers.Render(LayoutTestHelpers.Doc(section, blocks)).Pages.Single());
        double height = LayoutTestHelpers.NotoLineHeight(12);
        Assert.Equal(18.228, height, 1e-9);
        Assert.Equal(height * factor, (lines[39].Baseline - lines[0].Baseline) / 39, 0.001);   // baselines are grouped to 0.01
    }

    [Fact]
    public void USE_TYPO_METRICS_does_not_change_the_line_height()
    {
        var typo = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));
        Assert.True(typo.UseTypoMetrics);
        Assert.Equal(FontMetrics.Of(WithoutTypoFlag(), 40), FontMetrics.Of(typo, 40));
    }
}
