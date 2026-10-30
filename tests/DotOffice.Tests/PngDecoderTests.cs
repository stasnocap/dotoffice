using System.Text;
using DotOffice.Pdf;

namespace DotOffice.Tests;

public class PngDecoderTests
{
    private const int W = 11, H = 9; // odd sizes exercise partial bytes and empty Adam7 passes

    private static int Sample(int x, int y, int ch, int max) => (x * 7 + y * 13 + ch * 5) % (max + 1);

    private static int[] Samples(int channels, int bitDepth)
    {
        int max = (1 << bitDepth) - 1;
        var s = new int[W * H * channels];
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
                for (int c = 0; c < channels; c++)
                    s[(y * W + x) * channels + c] = Sample(x, y, c, max) * (bitDepth == 16 ? 257 : 1);
        return s;
    }

    private static byte To8(int value, int bitDepth) =>
        bitDepth == 16 ? (byte)(value >> 8) : (byte)(value * 255 / ((1 << bitDepth) - 1));

    public static TheoryData<int, bool> GrayDepths() => new() { { 1, false }, { 2, false }, { 4, true }, { 8, false }, { 16, true }, { 1, true } };

    [Theory]
    [MemberData(nameof(GrayDepths))]
    public void Decodes_grayscale(int bitDepth, bool interlace)
    {
        var samples = Samples(1, bitDepth);
        var png = PngDecoder.Decode(PngTestEncoder.Encode(W, H, 0, bitDepth, samples, interlace: interlace));
        Assert.Equal((W, H, 1), (png.Width, png.Height, png.Channels));
        Assert.Null(png.Alpha);
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
            {
                byte v = To8(samples[y * W + x], bitDepth);
                Assert.Equal((v, v, v, (byte)255), png.GetPixel(x, y));
            }
    }

    [Theory]
    [InlineData(2, 8, false)]
    [InlineData(2, 16, true)]
    [InlineData(4, 8, true)]
    [InlineData(4, 16, false)]
    [InlineData(6, 8, false)]
    [InlineData(6, 8, true)]
    [InlineData(6, 16, false)]
    public void Decodes_truecolor_and_alpha(int colorType, int bitDepth, bool interlace)
    {
        int channels = colorType switch { 2 => 3, 4 => 2, _ => 4 };
        var s = Samples(channels, bitDepth);
        // Guarantee some translucency so alpha is kept.
        if (colorType != 2) s[channels - 1] = 0;
        var png = PngDecoder.Decode(PngTestEncoder.Encode(W, H, colorType, bitDepth, s, interlace: interlace));
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
            {
                int i = (y * W + x) * channels;
                var expected = colorType switch
                {
                    2 => (To8(s[i], bitDepth), To8(s[i + 1], bitDepth), To8(s[i + 2], bitDepth), (byte)255),
                    4 => (To8(s[i], bitDepth), To8(s[i], bitDepth), To8(s[i], bitDepth), To8(s[i + 1], bitDepth)),
                    _ => (To8(s[i], bitDepth), To8(s[i + 1], bitDepth), To8(s[i + 2], bitDepth), To8(s[i + 3], bitDepth)),
                };
                Assert.Equal(expected, png.GetPixel(x, y));
            }
        Assert.Equal(colorType != 2, png.Alpha is not null);
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(2, true)]
    [InlineData(4, false)]
    [InlineData(8, true)]
    public void Decodes_palette_with_transparency(int bitDepth, bool interlace)
    {
        int entries = 1 << bitDepth;
        var palette = new byte[entries * 3];
        for (int i = 0; i < palette.Length; i++) palette[i] = (byte)(i * 37 % 256);
        var trns = new byte[] { 0, 128 }; // entry 0 transparent, 1 half, rest opaque
        var s = Samples(1, bitDepth);
        var png = PngDecoder.Decode(PngTestEncoder.Encode(W, H, 3, bitDepth, s, palette, trns, interlace));
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
            {
                int idx = s[y * W + x];
                byte a = idx < trns.Length ? trns[idx] : (byte)255;
                Assert.Equal((palette[idx * 3], palette[idx * 3 + 1], palette[idx * 3 + 2], a), png.GetPixel(x, y));
            }
    }

    [Fact]
    public void Gray_and_rgb_transparency_keys_become_alpha()
    {
        var gray = PngDecoder.Decode(PngTestEncoder.Encode(2, 1, 0, 8, [10, 20], trns: [0, 20]));
        Assert.Equal((byte)255, gray.GetPixel(0, 0).A);
        Assert.Equal((byte)0, gray.GetPixel(1, 0).A);

        var rgb = PngDecoder.Decode(PngTestEncoder.Encode(2, 1, 2, 8, [1, 2, 3, 4, 5, 6], trns: [0, 4, 0, 5, 0, 6]));
        Assert.Equal((byte)255, rgb.GetPixel(0, 0).A);
        Assert.Equal((byte)0, rgb.GetPixel(1, 0).A);
    }

    [Fact]
    public void Opaque_alpha_channel_is_dropped()
    {
        var png = PngDecoder.Decode(PngTestEncoder.Encode(1, 1, 6, 8, [1, 2, 3, 255]));
        Assert.Null(png.Alpha);
    }

    [Fact]
    public void Rejects_invalid_and_oversized_images()
    {
        Assert.Throws<InvalidDataException>(() => PngDecoder.Decode([1, 2, 3]));
        var big = PngTestEncoder.Encode(4, 4, 0, 8, new int[16]);
        Assert.Throws<InvalidDataException>(() => PngDecoder.Decode(big, maxPixels: 15));
        var badDepth = PngTestEncoder.Encode(1, 1, 2, 8, [0, 0, 0]);
        badDepth[24] = 4; // RGB with 4-bit depth is invalid
        Assert.Throws<InvalidDataException>(() => PngDecoder.Decode(badDepth));
    }

    [Fact]
    public void Png_with_alpha_is_written_with_soft_mask()
    {
        var doc = new PdfDocument();
        var image = doc.AddImage(PdfImage.FromPng(PngTestEncoder.Encode(2, 2, 6, 8, [255, 0, 0, 0, 0, 255, 0, 128, 0, 0, 255, 255, 9, 9, 9, 9])));
        doc.AddPage(PageSize.A4).Canvas.DrawImage(image, 10, 10, 20, 20);
        using var ms = new MemoryStream();
        doc.Save(ms);
        var text = Encoding.Latin1.GetString(ms.ToArray());
        Assert.Contains("/SMask", text);
        Assert.Contains("/ColorSpace /DeviceRGB", text);
        Assert.Contains("/ColorSpace /DeviceGray", text);
    }
}