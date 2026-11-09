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
        Assert.Equal(["•\tOne", "–\tTwo", "No bullet"], paragraphs.Select(p => string.Concat(p.Runs.Select(r => r.Text))));
        Assert.Equal((0.0, 27.0, 36.0, 58.5), (paragraphs[0].Indent, paragraphs[0].Margin!.Value, paragraphs[1].Indent, paragraphs[1].Margin!.Value));
        Assert.Equal(28.0, paragraphs[1].Runs[0].Font.FontSize);
    }

    [Fact]
    public void Paragraph_and_line_spacing_are_read_in_points_and_percent()
    {
        var builder = new PptxBuilder { TextStyles = "<p:bodyStyle><a:lvl1pPr><a:lnSpc><a:spcPct val=\"90000\"/></a:lnSpc><a:spcBef><a:spcPct val=\"20000\"/></a:spcBef><a:defRPr sz=\"2000\"/></a:lvl1pPr></p:bodyStyle>" }
            .AddSlide(Sp(Xfrm(0, 0, 5080000, 2540000),
                P("Inherited") + P("Own", "<a:pPr><a:lnSpc><a:spcPts val=\"3000\"/></a:lnSpc><a:spcAft><a:spcPts val=\"1200\"/></a:spcAft></a:pPr>"), "type=\"body\""));
        var paragraphs = Assert.IsType<SheetShape>(Assert.Single(ReadSlide(builder).Objects)).Paragraphs;
        Assert.Equal((0.9, null, 4.0, 0.0), (paragraphs[0].LineSpacing, paragraphs[0].LinePitch, paragraphs[0].SpaceBefore, paragraphs[0].SpaceAfter));
        Assert.Equal((30.0, 4.0, 12.0), (paragraphs[1].LinePitch!.Value, paragraphs[1].SpaceBefore, paragraphs[1].SpaceAfter));
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
    public void Embedded_objects_are_drawn_from_their_preview_picture_in_the_frame()
    {
        static string Frame(string inner) =>
            "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"7\" name=\"o\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>" + Xfrm(127000, 254000, 1270000, 635000).Replace("a:xfrm", "p:xfrm") +
            "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/presentationml/2006/ole\"><p:oleObj progId=\"Excel.Sheet.12\" r:id=\"rIdOle\">" + inner +
            "</p:oleObj></a:graphicData></a:graphic></p:graphicFrame>";
        var pic = "<p:embed/><p:pic><p:nvPicPr><p:cNvPr id=\"0\" name=\"\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rIdImg\"/></p:blipFill>" +
                  "<p:spPr>" + Xfrm(127000, 254000, 1270000, 635000) + "</p:spPr></p:pic>";
        var (presentation, warnings) = Read(new PptxBuilder().AddImage(Png).AddSlide(Frame(pic) + Frame("<p:embed/>")));
        var image = Assert.IsType<SheetImage>(Assert.Single(presentation.Slides[0].Objects));
        Assert.Equal((10.0, 20.0, 100.0, 50.0), (image.X, image.Y, image.Width, image.Height));
        Assert.Contains("1 embedded objects without a preview picture", Assert.Single(warnings.Warnings).Message);
    }

    [Fact]
    public void Charts_and_diagrams_are_reported()
    {
        var frame = "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"6\" name=\"c\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>" + Xfrm(0, 0, 12700, 12700).Replace("a:xfrm", "p:xfrm") +
                    "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/></a:graphic></p:graphicFrame>";
        var (presentation, warnings) = Read(new PptxBuilder().AddSlide(frame + frame));
        Assert.Empty(presentation.Slides[0].Objects);
        var warning = Assert.Single(warnings.Warnings);
        Assert.Equal("pptx.unsupported", warning.Code);
        Assert.Contains("2 SmartArt diagrams", warning.Message);
    }

    private const string DefaultStyle = "<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>";

    private static string Row(long h, params string[] cells) => $"<a:tr h=\"{h}\">{string.Concat(cells)}</a:tr>";

    private static string StyledTable(string tblPr, string rows) => Table([1270000, 2540000], tblPr, rows).Replace("<a:tblPr " + tblPr + "/>", $"<a:tblPr {tblPr}>{DefaultStyle}</a:tblPr>");

    [Fact]
    public void Table_cells_are_placed_by_the_grid_and_merged_cells_span_their_columns_and_rows()
    {
        var rows = Row(254000, Tc("A", "gridSpan=\"2\""), Tc("", "hMerge=\"1\"")) + Row(381000, Tc("B", "rowSpan=\"2\""), Tc("C")) + Row(254000, Tc("", "vMerge=\"1\""), Tc("D"));
        var (presentation, warnings) = Read(new PptxBuilder().AddSlide(Table([1270000, 2540000], "", rows)));
        var cells = presentation.Slides[0].Objects.OfType<SheetShape>().Where(s => s.Geometry != ShapeGeometry.Line).ToList();
        Assert.Equal(["A", "B", "C", "D"], cells.Select(c => c.Paragraphs[0].Runs[0].Text));
        var frames = cells.Select(c => c.Frame).ToList();
        Assert.Equal(new FrameFraction(0, 0, 1, 0.2), frames[0]);
        Assert.Equal(new FrameFraction(0, 0.2, 1 / 3.0, 0.5), frames[1]);
        Assert.Equal(new FrameFraction(1 / 3.0, 0.5, 2 / 3.0, 0.2), frames[3]);
        Assert.All(cells, c => Assert.Equal((0.0, 300.0, 100.0), (c.X, c.Width, c.Height)));
        Assert.Empty(warnings.Warnings);
    }

    [Fact]
    public void Table_cells_and_borders_carry_the_rows_they_cover()
    {
        var rows = Row(254000, Tc("A", "rowSpan=\"2\"", tcPr: "<a:lnB w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:lnB>"
            + "<a:lnR w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:lnR>")) + Row(381000, Tc("", "vMerge=\"1\""));
        var objects = ReadSlide(new PptxBuilder().AddSlide(Table([1270000], "", rows))).Objects.OfType<SheetShape>().ToList();
        var cell = objects.Single(s => s.Geometry != ShapeGeometry.Line);
        Assert.Equal((0, 2), (cell.TableRows!.First, cell.TableRows.Count));
        Assert.Equal([254000.0, 381000.0], cell.TableRows.RowHeights);
        var lines = objects.Where(s => s.Geometry == ShapeGeometry.Line).Select(l => (l.TableRows!.First, l.TableRows.Count)).ToList();
        Assert.Equal([(2, 0), (0, 2)], lines);
    }

    [Theory]
    [InlineData("vert", -90.0)]
    [InlineData("vert270", 90.0)]
    [InlineData("horz", 0.0)]
    public void Vertical_text_turns_the_text_of_shapes_and_table_cells(string vert, double turn)
    {
        var body = $"<a:bodyPr vert=\"{vert}\"/><a:lstStyle/>" + P("text");
        var shape = Assert.IsType<SheetShape>(Assert.Single(ReadSlide(new PptxBuilder().AddSlide(
            Sp(Xfrm(0, 0, 1270000, 127000) + "<a:prstGeom prst=\"rect\"/>", body).Replace("<p:txBody><a:bodyPr/><a:lstStyle/>", "<p:txBody>"))).Objects));
        Assert.Equal(turn, shape.TextRotation);
        var cell = Assert.IsType<SheetShape>(Assert.Single(ReadSlide(new PptxBuilder().AddSlide(Table([1270000], "", Row(254000, Tc("A", tcPrAttributes: $"vert=\"{vert}\""))))).Objects));
        Assert.Equal(turn, cell.TextRotation);
    }

    [Fact]
    public void Stacked_vertical_text_is_drawn_horizontally_with_a_note()
    {
        var body = "<a:bodyPr vert=\"eaVert\"/><a:lstStyle/>" + P("text");
        var (presentation, warnings) = Read(new PptxBuilder().AddSlide(
            Sp(Xfrm(0, 0, 1270000, 127000) + "<a:prstGeom prst=\"rect\"/>", body).Replace("<p:txBody><a:bodyPr/><a:lstStyle/>", "<p:txBody>")));
        Assert.Equal(0, Assert.IsType<SheetShape>(Assert.Single(presentation.Slides[0].Objects)).TextRotation);
        Assert.Contains(warnings.Warnings, w => w.Message.Contains("stacked vertical text"));
    }

    [Fact]
    public void A_shape_that_fits_its_size_to_the_text_is_marked_to_grow()
    {
        var body = "<a:bodyPr><a:spAutoFit/></a:bodyPr><a:lstStyle/>" + P("text");
        var shape = Assert.IsType<SheetShape>(Assert.Single(ReadSlide(new PptxBuilder().AddSlide(
            Sp(Xfrm(0, 0, 1270000, 127000) + "<a:prstGeom prst=\"rect\"/>", body).Replace("<p:txBody><a:bodyPr/><a:lstStyle/>", "<p:txBody>"))).Objects));
        Assert.True(shape.GrowToFit);
        Assert.Null(shape.TableRows);
    }

    [Fact]
    public void Cell_margins_anchor_and_fill_come_from_the_cell_properties()
    {
        var rows = Row(254000, Tc("A", tcPrAttributes: "marL=\"0\" marT=\"12700\" anchor=\"ctr\"", tcPr: Red));
        var cell = Assert.IsType<SheetShape>(Assert.Single(ReadSlide(new PptxBuilder().AddSlide(Table([1270000], "", rows))).Objects));
        Assert.Equal(new Thickness(1, 7.2, 3.6, 0), cell.TextInsets);
        Assert.Equal(ShapeTextAnchor.Middle, cell.TextAnchor);
        Assert.Equal(new Color(255, 0, 0), cell.Fill!.Value);
    }

    [Fact]
    public void The_default_table_style_fills_the_header_row_and_bands_the_rows()
    {
        var rows = Row(254000, Tc("H"), Tc("H")) + Row(254000, Tc("1"), Tc("1")) + Row(254000, Tc("2"), Tc("2", tcPr: Red));
        var objects = ReadSlide(new PptxBuilder().AddSlide(StyledTable("firstRow=\"1\" bandRow=\"1\"", rows))).Objects.OfType<SheetShape>().ToList();
        var cells = objects.Where(s => s.Geometry != ShapeGeometry.Line).ToList();
        var header = cells[0];
        Assert.Equal(new Color(0x4F, 0x81, 0xBD), header.Fill!.Value);
        var font = header.Paragraphs[0].Runs[0].Font;
        Assert.Equal((true, Color.White), (font.Bold, font.Color!.Value));
        Assert.NotEqual(cells[2].Fill, cells[4].Fill); // band 1 is darker than band 2 (the whole table's tint)
        Assert.Equal(Color.Black, cells[2].Paragraphs[0].Runs[0].Font.Color!.Value);
        Assert.Equal(new Color(255, 0, 0), cells[5].Fill!.Value);

        var lines = objects.Where(s => s.Geometry == ShapeGeometry.Line).ToList();
        Assert.Equal(6 * 4, lines.Count);
        Assert.All(lines, l => Assert.Equal(Color.White, l.LineColor!.Value));
        Assert.Contains(lines, l => l.LineWidth == 3); // the thick rule under the header row
        Assert.True(lines.Min(l => l.ZOrder) > cells.Max(c => c.ZOrder));
    }

    [Fact]
    public void A_table_style_from_the_presentation_and_cell_borders_are_drawn()
    {
        var style = "<a:tblStyle styleId=\"{X}\" styleName=\"Mine\"><a:wholeTbl><a:tcTxStyle i=\"on\"><a:srgbClr val=\"00FF00\"/></a:tcTxStyle>" +
                    "<a:tcStyle><a:fill><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle>";
        var rows = Row(254000, Tc("A", tcPr: "<a:lnB w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:lnB>"));
        var table = Table([1270000], "", rows).Replace("<a:tblPr />", "<a:tblPr><a:tableStyleId>{X}</a:tableStyleId></a:tblPr>");
        var objects = ReadSlide(new PptxBuilder { TableStyles = style }.AddSlide(table)).Objects.OfType<SheetShape>().ToList();
        Assert.Equal(2, objects.Count);
        Assert.Equal(new Color(0, 0, 255), objects[0].Fill!.Value);
        var font = objects[0].Paragraphs[0].Runs[0].Font;
        Assert.Equal((true, new Color(0, 255, 0)), (font.Italic, font.Color!.Value));
        var line = objects[1];
        Assert.Equal((ShapeGeometry.Line, new Color(255, 0, 0), 2.0), (line.Geometry, line.LineColor!.Value, line.LineWidth));
        Assert.Equal(new FrameFraction(0, 0.2, 1, 0), line.Frame);
    }

    private const string ColumnChart =
        "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:chart><c:plotArea>" +
        "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" +
        "<c:spPr><a:solidFill><a:schemeClr val=\"tx2\"/></a:solidFill></c:spPr>" +
        "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>East</c:v></c:pt><c:pt idx=\"1\"><c:v>West</c:v></c:pt></c:strCache></c:strRef></c:cat>" +
        "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>3</c:v></c:pt><c:pt idx=\"1\"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val>" +
        "</c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart><c:catAx><c:axId val=\"1\"/><c:crossAx val=\"2\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:crossAx val=\"1\"/></c:valAx>" +
        "</c:plotArea></c:chart></c:chartSpace>";

    [Fact]
    public void A_chart_is_read_from_its_cached_values_without_a_chart_area_frame()
    {
        var (presentation, warnings) = Read(new PptxBuilder().AddChart(ColumnChart).AddSlide(ChartFrame()));
        var chart = Assert.IsType<SheetChart>(Assert.Single(presentation.Slides[0].Objects));
        Assert.Equal((ImageAnchorKind.Absolute, 10.0, 20.0, 200.0, 100.0), (chart.Anchor, chart.X, chart.Y, chart.Width, chart.Height));
        Assert.Equal(ChartKind.Column, chart.Kind);
        Assert.Equal(["East", "West"], chart.Categories);
        var series = Assert.Single(chart.Series);
        Assert.Equal([3.0, 5.0], series.Values);
        Assert.Equal(new Color(0x1F, 0x49, 0x7D), series.Color); // tx2 maps to dk2 through the color map
        Assert.Equal((null, null), (chart.Background, chart.Border));
        Assert.Empty(warnings.Warnings);
    }

    private const string LineChartXml =
        "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:chart><c:plotArea>" +
        "<c:lineChart><c:grouping val=\"standard\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>MARKER<c:cat><c:strLit><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>a</c:v></c:pt><c:pt idx=\"1\"><c:v>b</c:v></c:pt></c:strLit></c:cat>" +
        "<c:val><c:numLit><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>1</c:v></c:pt><c:pt idx=\"1\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>" +
        "<c:catAx><c:axId val=\"1\"/><c:crossAx val=\"2\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:crossAx val=\"1\"/></c:valAx></c:plotArea></c:chart></c:chartSpace>";

    private static ChartSeries LineSeries(string marker) =>
        Assert.Single(Assert.IsType<SheetChart>(Assert.Single(ReadSlide(new PptxBuilder().AddChart(LineChartXml.Replace("MARKER", marker)).AddSlide(ChartFrame())).Objects)).Series);

    private static string Dpt(string idx, string marker) => $"<c:dPt><c:idx val=\"{idx}\"/>{marker}</c:dPt>";

    [Fact]
    public void A_data_point_marker_states_symbol_size_and_style_and_the_series_keeps_its_own()
    {
        var series = LineSeries("<c:marker><c:symbol val=\"circle\"/><c:size val=\"7\"/></c:marker>" +
            Dpt("1", "<c:marker><c:symbol val=\"diamond\"/><c:size val=\"11\"/><c:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></c:spPr></c:marker>"));
        Assert.Equal((ChartMarkerSymbol.Circle, 7.0), (series.MarkerSymbol, series.MarkerSize));
        var point = Assert.Single(series.PointMarkers!);
        Assert.Equal(1, point.Key);
        Assert.Equal((true, (ChartMarkerSymbol?)ChartMarkerSymbol.Diamond, (double?)11), (point.Value.Visible, point.Value.Symbol, point.Value.Size));
        Assert.Equal(new Color(0, 255, 0), point.Value.Style!.Fill);
    }

    [Fact]
    public void A_data_point_marker_states_only_what_it_says()
    {
        var size = Assert.Single(LineSeries(Dpt("0", "<c:marker><c:size val=\"9\"/></c:marker>")).PointMarkers!).Value;
        Assert.Equal(new ChartPointMarker(null, null, 9, null), size);
        var none = Assert.Single(LineSeries(Dpt("0", "<c:marker><c:symbol val=\"none\"/></c:marker>")).PointMarkers!).Value;
        Assert.Equal(new ChartPointMarker(false), none);
        var style = Assert.Single(LineSeries(Dpt("0", "<c:marker><c:spPr><a:ln w=\"25400\"/><a:noFill/></c:spPr></c:marker>")).PointMarkers!).Value;
        Assert.True(style.Style!.NoFill);
        Assert.Null(style.Symbol);
    }

    [Fact]
    public void A_data_point_marker_inherits_the_series_marker_and_a_point_symbol_shows_a_series_with_none()
    {
        var series = LineSeries("<c:marker><c:symbol val=\"none\"/></c:marker>" + Dpt("1", "<c:marker><c:symbol val=\"square\"/></c:marker>"));
        Assert.False(series.Marker);
        Assert.True(Assert.Single(series.PointMarkers!).Value.Visible);
    }

    [Theory]
    [InlineData("5")]
    [InlineData("-1")]
    [InlineData("abc")]
    [InlineData("1e400")]
    [InlineData("")]
    public void A_data_point_marker_with_a_bad_or_out_of_range_index_is_ignored(string idx) =>
        Assert.Null(LineSeries(Dpt(idx, "<c:marker><c:symbol val=\"square\"/></c:marker>")).PointMarkers);

    [Fact]
    public void A_data_point_without_idx_or_marker_and_malformed_marker_values_are_tolerated()
    {
        Assert.Null(LineSeries("<c:dPt><c:marker><c:symbol val=\"square\"/></c:marker></c:dPt>").PointMarkers);
        Assert.Null(LineSeries(Dpt("0", "<c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></c:spPr>")).PointMarkers);
        Assert.Null(LineSeries(Dpt("0", "<c:marker><c:size val=\"x\"/><c:spPr><a:ln w=\"abc\"/></c:spPr></c:marker>")).PointMarkers);
        var odd = LineSeries(Dpt("0", "<c:marker><c:symbol val=\"blob\"/><c:size val=\"500\"/></c:marker>"));
        Assert.Equal(new ChartPointMarker(true, ChartMarkerSymbol.Auto), Assert.Single(odd.PointMarkers!).Value);
    }

    [Fact]
    public void A_line_series_without_c_size_or_line_width_has_the_automatic_marker_of_5_points_for_powerpoints_1_5_pt_line()
    {
        var series = LineSeries("");
        Assert.True(series.Marker);
        Assert.Null(series.MarkerSize);
        Assert.Equal(5, series.MarkerDiameter);
        Assert.Equal((0.5, series.AutomaticMarker.Color), (series.AutomaticMarker.OutlineWidth, series.AutomaticMarker.Outline)); // PowerPoint 2016: 0.5 pt, the fill's own color
    }

    private static ChartSeries[] LineSeriesList(params string[] seriesProperties)
    {
        var sers = string.Concat(seriesProperties.Select((p, i) =>
        {
            var parts = p.Split('|');
            return $"<c:ser><c:idx val=\"{parts[0]}\"/><c:order val=\"{i}\"/>{(parts.Length > 1 ? parts[1] : "")}<c:val><c:numLit><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>";
        }));
        var xml = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:chart><c:plotArea><c:lineChart><c:grouping val=\"standard\"/>" + sers +
            "<c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart><c:catAx><c:axId val=\"1\"/><c:crossAx val=\"2\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:crossAx val=\"1\"/></c:valAx></c:plotArea></c:chart></c:chartSpace>";
        return Assert.IsType<SheetChart>(Assert.Single(ReadSlide(new PptxBuilder().AddChart(xml).AddSlide(ChartFrame())).Objects)).Series.ToArray();
    }

    [Theory]
    [InlineData("<c:spPr><a:ln w=\"12700\"/></c:spPr>", 5.0)]
    [InlineData("<c:spPr><a:ln w=\"28575\"/></c:spPr>", 7.0)]
    [InlineData("", 5.0)]
    public void The_automatic_marker_size_follows_the_line_width(string properties, double size) =>
        Assert.Equal(size, LineSeriesList("0|" + properties)[0].MarkerDiameter);

    [Fact]
    public void The_automatic_marker_color_of_14_presentation_series_is_the_themes_accents_shaded_plain_and_tinted()
    {
        // PowerPoint 2016 (rotation chart, 14 series): c:idx 0 shaded, 6 the plain accent, 12 tinted, not the explicit line color. The test deck's theme has accent 1 4F81BD (Excel's 2007 palette gave 40699C, AABAD7 there).
        var series = LineSeriesList(Enumerable.Range(0, 14).Select(i => i == 0 ? "0|<c:spPr><a:ln w=\"28575\"><a:solidFill><a:srgbClr val=\"7F7F7F\"/></a:solidFill></a:ln></c:spPr>" : i.ToString()).ToArray());
        Assert.Equal(new Color(0x7F, 0x7F, 0x7F), series[0].Color);
        Assert.InRange(series[0].AutomaticMarker.Color.R, 0x3F, 0x41);
        Assert.InRange(series[0].AutomaticMarker.Color.B, 0x9B, 0x9D);
        Assert.Equal(series[0].AutomaticMarker.Color, series[0].AutomaticMarker.Outline); // PowerPoint outlines in the fill's color
        Assert.Equal(new Color(0x4F, 0x81, 0xBD), series[6].AutomaticMarker.Color);
        Assert.InRange(series[12].AutomaticMarker.Color.R, 0xA9, 0xAB);
    }

    [Fact]
    public void A_presentation_series_takes_its_automatic_marker_from_c_idx_not_from_its_position()
    {
        var series = LineSeriesList("9", "1");
        Assert.Equal([9, 1], series.Select(s => s.AutomaticMarker.Index));
    }

    [Fact]
    public void An_explicit_c_size_wins_over_the_default_marker_size()
    {
        Assert.Equal(9, LineSeries("<c:marker><c:symbol val=\"circle\"/><c:size val=\"9\"/></c:marker>").MarkerDiameter);
    }

    [Theory]
    [InlineData("circle", ChartMarkerSymbol.Circle)]
    [InlineData("dash", ChartMarkerSymbol.Dash)]
    [InlineData("diamond", ChartMarkerSymbol.Diamond)]
    [InlineData("dot", ChartMarkerSymbol.Dot)]
    [InlineData("plus", ChartMarkerSymbol.Plus)]
    [InlineData("square", ChartMarkerSymbol.Square)]
    [InlineData("star", ChartMarkerSymbol.Star)]
    [InlineData("triangle", ChartMarkerSymbol.Triangle)]
    [InlineData("x", ChartMarkerSymbol.X)]
    [InlineData("auto", ChartMarkerSymbol.Auto)]
    [InlineData("picture", ChartMarkerSymbol.Auto)]
    [InlineData("hexagon", ChartMarkerSymbol.Auto)]
    [InlineData("", ChartMarkerSymbol.Auto)]
    public void A_marker_symbol_is_read_and_an_unknown_one_is_automatic(string symbol, ChartMarkerSymbol expected)
    {
        var series = LineSeries($"<c:marker><c:symbol val=\"{symbol}\"/></c:marker>");
        Assert.Equal(expected, series.MarkerSymbol);
        Assert.True(series.Marker);
    }

    [Fact]
    public void A_none_marker_symbol_draws_no_marker_and_no_symbol_is_automatic()
    {
        Assert.False(LineSeries("<c:marker><c:symbol val=\"none\"/></c:marker>").Marker);
        Assert.Equal(ChartMarkerSymbol.Auto, LineSeries("").MarkerSymbol);
    }

    private static ChartMarkerStyle? MarkerStyleOf(string spPr) =>
        LineSeries($"<c:marker><c:symbol val=\"square\"/><c:spPr>{spPr}</c:spPr></c:marker>").MarkerStyle;

    [Fact]
    public void A_marker_without_sp_pr_or_with_nothing_stated_has_no_style()
    {
        Assert.Null(LineSeries("<c:marker><c:symbol val=\"square\"/></c:marker>").MarkerStyle);
        Assert.Null(MarkerStyleOf(""));
        Assert.Null(MarkerStyleOf("<a:ln w=\"19050\"/>"));
    }

    [Fact]
    public void A_marker_solid_fill_and_scheme_color_are_resolved_like_a_series_color()
    {
        Assert.Equal(new Color(0x12, 0x34, 0x56), MarkerStyleOf("<a:solidFill><a:srgbClr val=\"123456\"/></a:solidFill>")!.Fill);
        var scheme = MarkerStyleOf("<a:solidFill><a:schemeClr val=\"tx2\"><a:alpha val=\"50000\"/></a:schemeClr></a:solidFill>")!;
        Assert.Equal(new Color(0x1F, 0x49, 0x7D), scheme.Fill); // tx2 maps to dk2 through the color map
        Assert.Equal(0.5, scheme.FillAlpha);
    }

    [Fact]
    public void A_marker_no_fill_and_outline_color_width_and_no_line_are_read()
    {
        var hollow = MarkerStyleOf("<a:noFill/><a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln>")!;
        Assert.True(hollow.NoFill);
        Assert.Null(hollow.Fill);
        Assert.Equal(new Color(255, 0, 0), hollow.Line);
        Assert.Equal(2.0, hollow.LineWidth);
        Assert.False(hollow.NoLine);
        var none = MarkerStyleOf("<a:noFill/><a:ln><a:noFill/></a:ln>")!;
        Assert.True(none.NoFill && none.NoLine);
    }

    [Theory]
    [InlineData("<a:solidFill><a:srgbClr val=\"zzz\"/></a:solidFill><a:ln w=\"abc\"/>")]
    [InlineData("<a:solidFill/><a:ln w=\"-5\"><a:solidFill/></a:ln>")]
    [InlineData("<a:ln w=\"99999999999999\"><a:solidFill><a:schemeClr val=\"nope\"/></a:solidFill></a:ln>")]
    public void A_malformed_marker_style_does_not_throw_and_states_nothing_it_cannot_read(string spPr)
    {
        var style = MarkerStyleOf(spPr);
        Assert.Null(style?.LineWidth);
    }

    [Fact]
    public void A_chart_automatic_title_is_larger_and_bold_and_its_text_takes_the_theme_text_color()
    {
        var titled = ColumnChart.Replace("<c:chart><c:plotArea>", "<c:chart><c:title/><c:plotArea>")
            .Replace("</c:chart>", "</c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"1800\"/></a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></c:txPr>");
        var chart = Assert.IsType<SheetChart>(Assert.Single(ReadSlide(new PptxBuilder().AddChart(titled).AddSlide(ChartFrame())).Objects));
        Assert.Equal((18.0, 21.6), (chart.FontSize, chart.TitleSize));
        Assert.True(chart.TextStyles.Title!.Bold);
        Assert.Equal(Color.Black, chart.TextColor);
    }

    private static string DiagramShape(string xfrm, string? txXfrm = null) =>
        "<dsp:sp xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\"><dsp:nvSpPr><dsp:cNvPr id=\"0\" name=\"\"/><dsp:cNvSpPr/></dsp:nvSpPr>" +
        $"<dsp:spPr>{xfrm}<a:prstGeom prst=\"ellipse\"/>{Red}</dsp:spPr>" +
        "<dsp:style><a:lnRef idx=\"2\"><a:schemeClr val=\"lt1\"/></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef>" +
        "<a:effectRef idx=\"0\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></dsp:style>" +
        $"<dsp:txBody><a:bodyPr anchor=\"ctr\"/><a:lstStyle/>{P("Step")}</dsp:txBody>{txXfrm}</dsp:sp>";

    [Fact]
    public void SmartArt_is_drawn_from_its_saved_drawing_inside_the_frame()
    {
        var shapes = DiagramShape(Xfrm(0, 0, 1270000, 1270000)) + DiagramShape(Xfrm(1270000, 0, 1270000, 1270000), Xfrm(1270000, 635000, 1270000, 635000).Replace("a:xfrm", "dsp:txXfrm"));
        var (presentation, warnings) = Read(new PptxBuilder().AddDiagram(shapes).AddSlide(DiagramFrame()));
        var objects = presentation.Slides[0].Objects.OfType<SheetShape>().ToList();
        Assert.Equal(3, objects.Count);
        Assert.All(objects, o => Assert.Equal((10.0, 20.0, 200.0, 100.0), (o.X, o.Y, o.Width, o.Height)));
        Assert.Equal(new FrameFraction(0, 0, 0.5, 1), objects[0].Frame);
        Assert.Equal((ShapeGeometry.Ellipse, new Color(255, 0, 0)), (objects[0].Geometry, objects[0].Fill!.Value));
        Assert.Equal("Step", objects[0].Paragraphs[0].Runs[0].Text);
        Assert.Equal(Color.White, objects[0].Paragraphs[0].Runs[0].Font.Color!.Value);

        // The second shape has its own text box: the shape without text, then the text over it, unfilled.
        Assert.Equal((new FrameFraction(0.5, 0, 0.5, 1), 0), (objects[1].Frame, objects[1].Paragraphs.Count));
        Assert.Equal(new FrameFraction(0.5, 0.5, 0.5, 0.5), objects[2].Frame);
        Assert.Equal((ShapeGeometry.Rectangle, (Color?)null, (Color?)null), (objects[2].Geometry, objects[2].Fill, objects[2].LineColor));
        Assert.Equal(Color.White, objects[2].Paragraphs[0].Runs[0].Font.Color!.Value);
        Assert.Empty(warnings.Warnings);
    }

    [Fact]
    public void SmartArt_without_a_saved_drawing_is_skipped_with_a_warning()
    {
        var (presentation, warnings) = Read(new PptxBuilder().AddSlide(DiagramFrame()));
        Assert.Empty(presentation.Slides[0].Objects);
        Assert.Contains(warnings.Warnings, w => w.Message.Contains("SmartArt diagrams without a saved drawing"));
    }

    [Fact]
    public void A_chart_whose_part_is_missing_is_skipped_with_a_warning()
    {
        var (presentation, warnings) = Read(new PptxBuilder().AddSlide(ChartFrame("rIdNone")));
        Assert.Empty(presentation.Slides[0].Objects);
        Assert.Equal("pptx.chart-missing", Assert.Single(warnings.Warnings).Code);
    }

    private static List<ChartSeries> ManyLineSeries(int count, string extra = "")
    {
        var one = LineChartXml.Substring(LineChartXml.IndexOf("<c:ser>"), LineChartXml.IndexOf("</c:ser>") + 8 - LineChartXml.IndexOf("<c:ser>"));
        var many = string.Concat(Enumerable.Range(0, count).Select(i => one.Replace("<c:idx val=\"0\"/><c:order val=\"0\"/>MARKER", $"<c:idx val=\"{i}\"/><c:order val=\"{i}\"/>{extra}")));
        var xml = LineChartXml.Replace(one, many);
        return Assert.IsType<SheetChart>(Assert.Single(ReadSlide(new PptxBuilder().AddChart(xml).AddSlide(ChartFrame())).Objects)).Series.ToList();
    }

    [Fact]
    public void A_line_chart_line_without_a_color_is_the_automatic_marker_fill_of_14_series_and_1_5_pt_wide_in_powerpoint()
    {
        // Measured in a PowerPoint 2016 PDF with 14 series: the line equals the marker fill.
        var series = ManyLineSeries(14);
        Assert.Equal(14, series.Count);
        Assert.All(series, s => Assert.Equal((s.AutomaticMarker.Color, 1.5), (s.StrokeColor, s.StrokeWidth)));
        Assert.NotEqual(series[0].Color, series[0].StrokeColor);
    }

    [Fact]
    public void A_stated_line_color_and_width_win_in_powerpoint()
    {
        var series = ManyLineSeries(1, "<c:spPr><a:ln w=\"28575\"><a:solidFill><a:srgbClr val=\"123456\"/></a:solidFill></a:ln></c:spPr>").Single();
        Assert.Equal((new Color(0x12, 0x34, 0x56), 2.25), (series.StrokeColor, series.StrokeWidth));
    }
}
