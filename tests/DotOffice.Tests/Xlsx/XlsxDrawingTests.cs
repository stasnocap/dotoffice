using DotOffice.Core;
using DotOffice.Xlsx;
using static DotOffice.Tests.Xlsx.XlsxBuilder;

namespace DotOffice.Tests.Xlsx;

public class XlsxDrawingTests
{
    private static readonly byte[] Png = PngTestEncoder.Encode(2, 2, 2, 8, [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]);

    private static (Worksheet Sheet, ConversionWarningList Warnings) Read(string anchors, params byte[][] media)
    {
        var warnings = new ConversionWarningList();
        var builder = new XlsxBuilder().AddSheet("S", Row(1, C("A1", "1")), tail: DrawingElement).AddDrawing(1, anchors, media.Length == 0 ? [Png] : media);
        var wb = XlsxReader.Read(builder.BuildStream(), new XlsxReaderOptions { Warnings = warnings });
        return (wb.Sheets[0], warnings);
    }

    [Fact]
    public void Reads_a_two_cell_anchored_picture_without_a_warning()
    {
        var (sheet, warnings) = Read(TwoCellAnchor(1, 2, 4, 6, Pic(), fromColumnOffset: 3, fromRowOffset: 4, toColumnOffset: 5, toRowOffset: 6));
        var image = Assert.Single(sheet.Images);
        Assert.Equal(ImageAnchorKind.TwoCell, image.Anchor);
        Assert.Equal(new AnchorPoint(1, 2, 3, 4), image.From);
        Assert.Equal(new AnchorPoint(4, 6, 5, 6), image.To);
        Assert.Equal(ImageFormat.Png, image.Image.Format);
        Assert.Empty(warnings.Warnings);
    }

    [Fact]
    public void Reads_one_cell_and_absolute_anchors_in_points()
    {
        var (sheet, _) = Read(OneCellAnchor(2, 3, 100, 50, Pic(), columnOffset: 7) + AbsoluteAnchor(36, 72, 20, 10, Pic()));
        Assert.Equal(2, sheet.Images.Count);
        var one = sheet.Images[0];
        Assert.Equal((ImageAnchorKind.OneCell, 100.0, 50.0), (one.Anchor, one.Width, one.Height));
        Assert.Equal(new AnchorPoint(2, 3, 7, 0), one.From);
        var absolute = sheet.Images[1];
        Assert.Equal((ImageAnchorKind.Absolute, 36.0, 72.0, 20.0, 10.0), (absolute.Anchor, absolute.X, absolute.Y, absolute.Width, absolute.Height));
    }

    [Fact]
    public void Reads_cropping_and_rotation_with_counter_clockwise_degrees()
    {
        var (sheet, _) = Read(TwoCellAnchor(0, 0, 2, 2, Pic(srcRect: "l=\"10000\" t=\"20000\" r=\"30000\" b=\"5000\"", rotation: 90 * 60000)));
        var image = sheet.Images.Single();
        Assert.Equal((0.1, 0.2, 0.3, 0.05), (image.CropLeft, image.CropTop, image.CropRight, image.CropBottom));
        Assert.Equal(-90.0, image.Rotation);
    }

    [Fact]
    public void The_same_media_part_is_shared_between_pictures()
    {
        var (sheet, _) = Read(OneCellAnchor(0, 0, 10, 10, Pic()) + OneCellAnchor(1, 1, 10, 10, Pic()));
        Assert.Same(sheet.Images[0].Image, sheet.Images[1].Image);
    }

    [Fact]
    public void Hidden_pictures_and_pictures_that_do_not_print_are_left_out_silently()
    {
        var (sheet, warnings) = Read(
            OneCellAnchor(0, 0, 10, 10, Pic(nameAttributes: " hidden=\"1\"")) +
            TwoCellAnchor(0, 0, 1, 1, Pic(), clientData: "<xdr:clientData fPrintsWithSheet=\"0\"/>") +
            OneCellAnchor(1, 1, 10, 10, Pic()));
        Assert.Single(sheet.Images);
        Assert.Empty(warnings.Warnings);
    }

    [Fact]
    public void Charts_and_shapes_still_warn_once_and_pictures_next_to_them_are_kept()
    {
        var (sheet, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Pic()) + OneCellAnchor(2, 2, 10, 10, CustomShape) + OneCellAnchor(4, 4, 10, 10, ChartFrame) + OneCellAnchor(6, 6, 10, 10, CustomShape));
        Assert.Single(sheet.Images);
        var warning = Assert.Single(warnings.Warnings);
        Assert.Equal("xlsx.drawing", warning.Code);
        Assert.Contains("chart", warning.Message);
        Assert.Contains("2 shapes", warning.Message);
    }

    [Fact]
    public void A_shape_that_paints_nothing_is_dropped_silently()
    {
        var (sheet, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Shape));
        Assert.Empty(sheet.Shapes);
        Assert.Empty(warnings.Warnings);
    }

    [Fact]
    public void A_picture_only_drawing_does_not_mention_charts_or_shapes()
    {
        var (_, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Pic()));
        Assert.DoesNotContain(warnings.Warnings, w => w.Code == "xlsx.drawing");
    }

    [Fact]
    public void A_picture_whose_image_is_missing_is_skipped_with_a_warning()
    {
        var (sheet, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Pic("rId9")));
        Assert.Empty(sheet.Images);
        Assert.Equal(["xlsx.image-missing"], warnings.Warnings.Select(w => w.Code));
    }

    [Fact]
    public void Reads_flipped_pictures()
    {
        var (sheet, warnings) = Read(OneCellAnchor(0, 0, 10, 10, Pic(flip: " flipH=\"1\"")) + OneCellAnchor(0, 0, 10, 10, Pic(flip: " flipV=\"true\"")));
        Assert.Equal([(true, false), (false, true)], sheet.Images.Select(i => (i.FlipH, i.FlipV)));
        Assert.Empty(warnings.Warnings);
    }
}
