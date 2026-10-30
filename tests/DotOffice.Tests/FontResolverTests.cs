using DotOffice.Fonts;

namespace DotOffice.Tests;

public class FontResolverTests
{
    private static readonly TrueTypeFont Noto = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));

    /// <summary>Exposes a single font under arbitrary family names.</summary>
    private sealed class AliasProvider(TrueTypeFont font, params string[] families) : IFontProvider
    {
        public List<string> Requests { get; } = [];

        public TrueTypeFont? FindFont(string family, bool bold, bool italic)
        {
            Requests.Add(family);
            return families.Contains(family, StringComparer.OrdinalIgnoreCase) ? font : null;
        }
    }

    [Fact]
    public void Exact_family_is_returned_without_substitution()
    {
        var resolver = new FontResolver(new FontCollection().Add(Noto));
        var resolved = resolver.Resolve("noto sans", bold: false, italic: false);
        Assert.Same(Noto, resolved.Font);
        Assert.False(resolved.IsSubstitute);
        Assert.False(resolved.SyntheticBold);
    }

    [Fact]
    public void Missing_styles_are_synthesized()
    {
        var resolver = new FontResolver(new FontCollection().Add(Noto));
        var resolved = resolver.Resolve("Noto Sans", bold: true, italic: true);
        Assert.Same(Noto, resolved.Font);
        Assert.True(resolved.SyntheticBold);
        Assert.True(resolved.SyntheticItalic);
    }

    [Theory]
    [InlineData("Calibri", "Carlito")]
    [InlineData("Cambria", "Caladea")]
    [InlineData("Arial", "Liberation Sans")]
    [InlineData("Times New Roman", "Liberation Serif")]
    [InlineData("Courier New", "Liberation Mono")]
    public void Office_fonts_map_to_metric_compatible_substitutes(string requested, string substitute)
    {
        var provider = new AliasProvider(Noto, substitute);
        var resolved = new FontResolver(provider).Resolve(requested, false, false);
        Assert.Same(Noto, resolved.Font);
        Assert.True(resolved.IsSubstitute);
        Assert.Equal([requested, substitute], provider.Requests.Take(2));
    }

    [Fact]
    public void Unknown_family_falls_back_to_generic_family_then_any_registered_font()
    {
        var generic = new FontResolver(new AliasProvider(Noto, "DejaVu Serif"));
        Assert.Same(Noto, generic.Resolve("Some Serif Font", false, false).Font);

        var resolver = new FontResolver(new FontCollection().Add(Noto));
        var resolved = resolver.Resolve("Nonexistent Family", false, false);
        Assert.Same(Noto, resolved.Font);
        Assert.True(resolved.IsSubstitute);
    }

    [Fact]
    public void No_fonts_at_all_throws_a_clear_error()
    {
        var resolver = new FontResolver(new FontCollection());
        Assert.Throws<FontNotFoundException>(() => resolver.Resolve("Calibri", false, false));
    }

    [Fact]
    public void Per_character_fallback_finds_a_covering_font()
    {
        var resolver = new FontResolver(new FontCollection().Add(Noto));
        Assert.Same(Noto, resolver.ResolveFallback('Ж', false, false)?.Font);
        Assert.Null(resolver.ResolveFallback(0x1F600, false, false)); // emoji: no registered font has it
    }

    [Fact]
    public void Face_info_reads_names_and_style_from_headers_only()
    {
        var info = FontFaceInfo.TryRead(TestAssets.Path("NotoSans-Regular.ttf"));
        Assert.NotNull(info);
        Assert.Equal("Noto Sans", info.FamilyName);
        Assert.Equal("Regular", info.SubfamilyName);
        Assert.Equal(400, info.Weight);
        Assert.False(info.IsItalic);
        Assert.Null(FontFaceInfo.TryRead(TestAssets.Path("gray.jpg")));
    }

    [Fact]
    public void System_provider_indexes_a_directory()
    {
        var dir = Directory.CreateTempSubdirectory("dotoffice-fonts-");
        try
        {
            File.Copy(TestAssets.Path("NotoSans-Regular.ttf"), Path.Combine(dir.FullName, "a.ttf"));
            File.WriteAllBytes(Path.Combine(dir.FullName, "broken.ttf"), [1, 2, 3]);
            var provider = new SystemFontProvider([dir.FullName]);
            Assert.Contains("Noto Sans", provider.Families);
            var font = provider.FindFont("Noto Sans", bold: true, italic: false);
            Assert.NotNull(font);
            Assert.False(font.IsBold); // nearest available face
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void Decoration_metrics_are_read()
    {
        Assert.True(Noto.UnderlinePosition < 0);
        Assert.True(Noto.UnderlineThickness > 0);
        Assert.True(Noto.StrikeoutPosition > 0);
        Assert.Equal("Noto Sans", Noto.TypographicFamilyName);
    }
}