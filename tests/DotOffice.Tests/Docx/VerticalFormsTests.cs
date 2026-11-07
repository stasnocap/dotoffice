using DotOffice.Docx;
using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;
using DotOffice.Layout;
using static DotOffice.Tests.Docx.DocxBuilder;
using static DotOffice.Tests.Layout.LayoutTestHelpers;

namespace DotOffice.Tests.Docx;

/// <summary>Vertical punctuation forms (the <c>vert</c> feature). They use the system's Noto Sans CJK and pass without checking if it is absent.</summary>
public class VerticalFormsTests
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

    private const string Font = "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\" w:eastAsia=\"Noto Sans CJK JP\"/><w:sz w:val=\"28\"/>";

    private static IEnumerable<TextItem> Upright(string direction, TrueTypeFont cjk)
    {
        var body = P(R("日、本「語」ー", Font)) + Section(extra: $"<w:textDirection w:val=\"{direction}\"/>");
        var doc = DocxReader.Read(new DocxBuilder { Body = body }.BuildStream());
        var fonts = new FontResolver(new FontCollection().Add(Noto).Add(cjk));
        return Texts(new LayoutEngine(new LayoutOptions { Fonts = fonts }).Layout(doc).Pages[0]).Where(t => t.Font.IsCff);
    }

    [Fact]
    public void The_font_gives_punctuation_its_vertical_form()
    {
        if (Japanese() is not { } font) return;
        foreach (var c in "、。「」ー")
        {
            var vertical = TextShaper.VerticalGlyph(font, c);
            Assert.NotNull(vertical);
            Assert.NotEqual(font.GetGlyphId(c), vertical!.Value);
        }
        Assert.Null(TextShaper.VerticalGlyph(font, '日'));
    }

    [Fact]
    public void Punctuation_in_text_running_down_is_drawn_with_the_vertical_glyph_at_the_same_pitch()
    {
        if (Japanese() is not { } font) return;
        var texts = Upright("tbRl", font).ToList();
        foreach (var c in "、「」ー")
        {
            var item = texts.Single(t => t.Text == c.ToString());
            Assert.Equal(TextShaper.VerticalGlyph(font, c), item.Glyphs![0].GlyphId);
            Assert.Equal(font.GetAdvanceWidth(font.GetGlyphId(c)), item.Glyphs[0].Advance);
        }
        Assert.Null(texts.Single(t => t.Text == "日").Glyphs);
    }

    [Fact]
    public void Horizontal_text_keeps_the_horizontal_forms()
    {
        if (Japanese() is not { } font) return;
        Assert.DoesNotContain(Upright("", font), t => t.Glyphs is { } g && g.Any(x => x.GlyphId == TextShaper.VerticalGlyph(font, '、')));
    }
}
