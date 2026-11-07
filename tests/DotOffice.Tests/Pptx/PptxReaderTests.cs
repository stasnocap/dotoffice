using DotOffice.Core;
using DotOffice.Pptx;
using static DotOffice.Tests.Pptx.PptxBuilder;

namespace DotOffice.Tests.Pptx;

public class PptxReaderTests
{
    private static readonly byte[] Png = PngTestEncoder.Encode(2, 2, 2, 8, [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]);

    private static (Presentation Presentation, ConversionWarningList Warnings) Read(PptxBuilder builder)
    {
        var warnings = new ConversionWarningList();
        return (PptxReader.Read(builder.BuildStream(), new PptxReaderOptions { Warnings = warnings }), warnings);
    }

    private static Slide ReadSlide(PptxBuilder builder) => Assert.Single(Read(builder).Presentation.Slides);

    private const string Red = "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>";

    private const string TitleLayout = "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"t\"/><p:cNvSpPr/><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr>" +
                                       "<p:spPr><a:xfrm><a:off x=\"127000\" y=\"254000\"/><a:ext cx=\"1270000\" cy=\"635000\"/></a:xfrm></p:spPr></p:sp>";

    [Fact]
    public void Each_shown_slide_becomes_a_slide_of_the_presentation_size()
    {
        var builder = new PptxBuilder { Size = (12_192_000, 6_858_000) }
            .AddSlide(Sp(Xfrm(0, 0, 12700, 12700) + Red))
            .AddSlide(Sp(Xfrm(0, 0, 12700, 12700) + Red), attributes: "show=\"0\"")
            .AddSlide(Sp(Xfrm(0, 0, 12700, 12700) + Red));
        var (presentation, warnings) = Read(builder);
        Assert.Equal(2, presentation.Slides.Count);
        Assert.Equal((960.0, 540.0), (presentation.Slides[0].Width, presentation.Slides[0].Height));
        Assert.Empty(warnings.Warnings);
    }

    [Fact]
    public void Shapes_are_placed_absolutely_from_their_frame()
    {
        var slide = ReadSlide(new PptxBuilder().AddSlide(Sp(Xfrm(127000, 254000, 381000, 508000) + "<a:prstGeom prst=\"ellipse\"/>" + Red)));
        var shape = Assert.IsType<SheetShape>(Assert.Single(slide.Objects));
        Assert.Equal((ImageAnchorKind.Absolute, 10.0, 20.0, 30.0, 40.0), (shape.Anchor, shape.X, shape.Y, shape.Width, shape.Height));
        Assert.Equal((ShapeGeometry.Ellipse, new Color(255, 0, 0)), (shape.Geometry, shape.Fill!.Value));
    }

    [Fact]
    public void A_placeholder_takes_its_frame_from_the_layout_and_its_text_style_from_the_master()
    {
        var builder = new PptxBuilder
        {
            LayoutTree = TitleLayout,
            TextStyles = "<p:titleStyle><a:lvl1pPr algn=\"ctr\"><a:defRPr sz=\"4400\" b=\"1\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mj-lt\"/></a:defRPr></a:lvl1pPr></p:titleStyle>",
        }.AddSlide(Sp("", P("Hello"), "type=\"title\""));
        var shape = Assert.IsType<SheetShape>(Assert.Single(ReadSlide(builder).Objects));
        Assert.Equal((10.0, 20.0, 100.0, 50.0), (shape.X, shape.Y, shape.Width, shape.Height));
        var paragraph = Assert.Single(shape.Paragraphs);
        Assert.Equal(HorizontalAlignment.Center, paragraph.Alignment);
        var font = Assert.Single(paragraph.Runs).Font;
        Assert.Equal((44.0, true, "Major Face", Color.Black), (font.FontSize, font.Bold, font.FontFamily, font.Color!.Value));
    }

    [Fact]
    public void A_placeholder_with_only_an_index_matches_the_layout_placeholder_of_that_index()
    {
        var layout = "<p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"b\"/><p:cNvSpPr/><p:nvPr><p:ph idx=\"7\"/></p:nvPr></p:nvSpPr><p:spPr>" + Xfrm(0, 1270000, 2540000, 1270000) + "</p:spPr></p:sp>";
        var builder = new PptxBuilder { LayoutTree = TitleLayout + layout }.AddSlide(Sp("", P("Body"), "idx=\"7\""));
        var shape = Assert.IsType<SheetShape>(Assert.Single(ReadSlide(builder).Objects));
        Assert.Equal((0.0, 100.0, 200.0), (shape.X, shape.Y, shape.Width));
    }

    [Fact]
    public void Text_outside_placeholders_uses_the_presentation_default_style_and_the_minor_font()
    {
        var builder = new PptxBuilder { DefaultTextStyle = "<a:lvl1pPr><a:defRPr sz=\"2000\"/></a:lvl1pPr>" }
            .AddSlide(Sp(Xfrm(0, 0, 1270000, 635000), P("Box")));
        var font = Assert.Single(Assert.Single(Assert.IsType<SheetShape>(Assert.Single(ReadSlide(builder).Objects)).Paragraphs).Runs).Font;
        Assert.Equal((20.0, "Minor Face"), (font.FontSize, font.FontFamily));
    }

    [Fact]
    public void The_color_map_decides_what_text_and_background_colors_stand_for()
    {
        var builder = new PptxBuilder { ColorMap = "bg1=\"dk1\" tx1=\"lt1\" bg2=\"dk2\" tx2=\"lt2\"" }
            .AddSlide(Sp(Xfrm(0, 0, 1270000, 635000), P("Light", rPr: "") .Replace("<a:rPr lang=\"en-US\" />", "<a:rPr><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:rPr>")),
                head: "<p:bg><p:bgPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>");
        var slide = ReadSlide(builder);
        Assert.Equal(Color.Black, slide.Background);
        Assert.Equal(Color.White, Assert.Single(Assert.Single(Assert.IsType<SheetShape>(Assert.Single(slide.Objects)).Paragraphs).Runs).Font.Color);
    }

    [Theory]
    [InlineData("<a:schemeClr val=\"tx1\"><a:tint val=\"75000\"/></a:schemeClr>", "898989")]
    [InlineData("<a:srgbClr val=\"FFFFFF\"><a:shade val=\"50000\"/></a:srgbClr>", "BCBCBC")]
    public void Tint_and_shade_are_applied_in_linear_rgb(string color, string expected)
    {
        var run = P("Gray").Replace("<a:rPr lang=\"en-US\" />", $"<a:rPr><a:solidFill>{color}</a:solidFill></a:rPr>");
        var font = Assert.Single(Assert.Single(Assert.IsType<SheetShape>(Assert.Single(ReadSlide(new PptxBuilder().AddSlide(Sp(Xfrm(0, 0, 1270000, 635000), run))).Objects)).Paragraphs).Runs).Font;
        Assert.Equal(Color.TryParseHex(expected), font.Color);
    }

    [Fact]
    public void The_background_is_inherited_from_the_master()
    {
        var builder = new PptxBuilder { MasterBackground = "<p:bg><p:bgPr>" + Red + "<a:effectLst/></p:bgPr></p:bg>" }.AddSlide("");
        Assert.Equal(new Color(255, 0, 0), ReadSlide(builder).Background);
    }

    [Fact]
    public void Master_shapes_are_drawn_behind_the_slide_unless_the_layout_hides_them()
    {
        var master = Sp(Xfrm(0, 0, 12700, 12700) + "<a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill>");
        var slide = ReadSlide(new PptxBuilder { MasterTree = master + TitleLayout }.AddSlide(Sp(Xfrm(0, 0, 12700, 12700) + Red)));
        Assert.Equal([new Color(0, 255, 0), new Color(255, 0, 0)], slide.Objects.OrderBy(o => o.ZOrder).Select(o => ((SheetShape)o).Fill!.Value));

        var hidden = ReadSlide(new PptxBuilder { MasterTree = master, LayoutAttributes = "showMasterSp=\"0\"" }.AddSlide(Sp(Xfrm(0, 0, 12700, 12700) + Red)));
        Assert.Single(hidden.Objects);
    }

    [Fact]
    public void Bullets_become_leading_text_and_levels_indent()
    {
        var builder = new PptxBuilder
        {
            TextStyles = "<p:bodyStyle><a:lvl1pPr marL=\"342900\" indent=\"-342900\"><a:buChar char=\"•\"/><a:defRPr sz=\"3200\"/></a:lvl1pPr>" +
                         "<a:lvl2pPr marL=\"742950\" indent=\"-285750\"><a:buChar char=\"–\"/><a:defRPr sz=\"2800\"/></a:lvl2pPr></p:bodyStyle>",
        }.AddSlide(Sp(Xfrm(0, 0, 5080000, 2540000), P("One") + P("Two", "<a:pPr lvl=\"1\"/>") + P("No bullet", "<a:pPr><a:buNone/></a:pPr>"), "type=\"body\""));
        var paragraphs = Assert.IsType<SheetShape>(Assert.Single(ReadSlide(builder).Objects)).Paragraphs;
        Assert.Equal(["• One", "– Two", "No bullet"], paragraphs.Select(p => string.Concat(p.Runs.Select(r => r.Text))));
        Assert.Equal((0.0, 36.0), (paragraphs[0].Indent, paragraphs[1].Indent));
        Assert.Equal(28.0, paragraphs[1].Runs[0].Font.FontSize);
    }

    [Fact]
    public void Numbered_bullets_count_per_level()
    {
        var number = "<a:pPr><a:buAutoNum type=\"alphaLcParenR\"/></a:pPr>";
        var builder = new PptxBuilder().AddSlide(Sp(Xfrm(0, 0, 5080000, 2540000), P("a", number) + P("b", number) + P("c", "<a:pPr lvl=\"1\"><a:buAutoNum type=\"romanUcPeriod\"/></a:pPr>") + P("d", number)));
        var texts = Assert.IsType<SheetShape>(Assert.Single(ReadSlide(builder).Objects)).Paragraphs.Select(p => string.Concat(p.Runs.Select(r => r.Text)));
        Assert.Equal(["a) a", "b) b", "I. c", "c) d"], texts);
    }

    [Fact]
    public void Font_scale_of_a_shrunk_text_body_applies_to_every_run()
    {
        var body = "<p:txBody><a:bodyPr><a:normAutofit fontScale=\"50000\"/></a:bodyPr><a:lstStyle/>" + P("Small", rPr: "sz=\"4000\"") + "</p:txBody>";
        var sp = Sp(Xfrm(0, 0, 1270000, 635000)).Replace("</p:sp>", body + "</p:sp>");
        var font = Assert.Single(Assert.Single(Assert.IsType<SheetShape>(Assert.Single(ReadSlide(new PptxBuilder().AddSlide(sp)).Objects)).Paragraphs).Runs).Font;
        Assert.Equal(20.0, font.FontSize);
    }

    [Fact]
    public void Pictures_and_grouped_shapes_are_placed_inside_their_frames()
    {
        var pic = "<p:pic><p:nvPicPr><p:cNvPr id=\"4\" name=\"p\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rIdImg\"/><a:srcRect l=\"25000\"/></p:blipFill>" +
                  "<p:spPr>" + Xfrm(0, 0, 254000, 254000) + "</p:spPr></p:pic>";
        var group = "<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"5\" name=\"g\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"1270000\" y=\"0\"/><a:ext cx=\"1270000\" cy=\"1270000\"/>" +
                    "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"200\" cy=\"200\"/></a:xfrm></p:grpSpPr>" + Sp(Xfrm(100, 0, 100, 100) + Red) + "</p:grpSp>";
        var slide = ReadSlide(new PptxBuilder().AddImage(Png).AddSlide(pic + group));
        var image = Assert.IsType<SheetImage>(slide.Objects[0]);
        Assert.Equal((20.0, 0.25), (image.Width, image.CropLeft));
        var shape = Assert.IsType<SheetShape>(slide.Objects[1]);
        Assert.Equal((100.0, 100.0), (shape.X, shape.Width));
        Assert.Equal(new FrameFraction(0.5, 0, 0.5, 0.5), shape.Frame);
    }

    [Fact]
    public void Tables_charts_and_diagrams_are_reported()
    {
        var frame = "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"6\" name=\"t\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>" + Xfrm(0, 0, 12700, 12700).Replace("a:xfrm", "p:xfrm") +
                    "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl/></a:graphicData></a:graphic></p:graphicFrame>";
        var (presentation, warnings) = Read(new PptxBuilder().AddSlide(frame + frame));
        Assert.Empty(presentation.Slides[0].Objects);
        var warning = Assert.Single(warnings.Warnings);
        Assert.Equal("pptx.unsupported", warning.Code);
        Assert.Contains("2 tables", warning.Message);
    }
}
