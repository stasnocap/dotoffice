using System.Text;
using DotOffice.Pdf;

namespace DotOffice.Tests;

public class RasterDecoderTests
{
    private static byte[] Bmp(int bpp, int w, int h, byte[] palette, byte[] rows)
    {
        var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        int offset = 14 + 40 + palette.Length;
        bw.Write("BM"u8.ToArray()); bw.Write(offset + rows.Length); bw.Write(0); bw.Write(offset);
        bw.Write(40); bw.Write(w); bw.Write(h); bw.Write((short)1); bw.Write((short)bpp);
        bw.Write(0); bw.Write(rows.Length); bw.Write(0); bw.Write(0); bw.Write(palette.Length / 4); bw.Write(0);
        bw.Write(palette); bw.Write(rows);
        return ms.ToArray();
    }

    [Fact]
    public void Bmp_24_bit_is_bottom_up_bgr_with_row_padding()
    {
        // 1x2: bottom row red, top row blue; each row is 3 bytes + 1 padding.
        var r = RasterDecoders.DecodeBmp(Bmp(24, 1, 2, [], [0, 0, 255, 0, 255, 0, 0, 0]));
        Assert.Equal([0, 0, 255, 255, 0, 0], r.Rgb);
        Assert.Equal((1, 2), (r.Width, r.Height));
    }

    [Fact]
    public void Bmp_4_bit_palette_and_16_bit_555()
    {
        var palette = new byte[] { 0, 0, 0, 0, 0, 0, 255, 0 }; // black, red (BGRA)
        var r = RasterDecoders.DecodeBmp(Bmp(4, 2, 1, palette, [0x01, 0, 0, 0]));
        Assert.Equal([0, 0, 0, 255, 0, 0], r.Rgb);
        var r16 = RasterDecoders.DecodeBmp(Bmp(16, 1, 1, [], [0x1F, 0x00, 0, 0])); // blue 31
        Assert.Equal([0, 0, 255], r16.Rgb);
    }

    [Fact]
    public void Bmp_with_unknown_compression_is_rejected_and_truncated_data_is_invalid()
    {
        var jpeg = Bmp(8, 1, 1, new byte[4], [0, 0, 0, 0]);
        jpeg[30] = 4;
        Assert.Throws<NotSupportedException>(() => RasterDecoders.DecodeBmp(jpeg));
        Assert.Throws<InvalidDataException>(() => RasterDecoders.DecodeBmp(Bmp(24, 5, 5, [], [0, 0])));
    }

    private static byte[] BmpRle(int bpp, int w, int h, byte[] palette, byte[] data, bool negativeHeight = false)
    {
        var bmp = Bmp(bpp, w, negativeHeight ? -h : h, palette, data);
        bmp[30] = (byte)(bpp == 8 ? 1 : 2);
        return bmp;
    }

    [Fact]
    public void Bmp_rle8_runs_absolute_runs_and_deltas()
    {
        var palette = new byte[] { 0, 0, 0, 0, 0, 0, 255, 0, 0, 255, 0, 0 }; // black, red, green (BGRA)
        // Row 0 (bottom): 3 x red, absolute run of green,black (padded), end of line;
        // row 1 (top): delta skips 1 pixel, then 1 x green; end of bitmap.
        byte[] data = [3, 1, 0, 2, 2, 0, 0, 0, 0, 2, 1, 0, 1, 2, 0, 1];
        var r = RasterDecoders.DecodeBmp(BmpRle(8, 3, 2, palette, data));
        Assert.Equal((3, 2), (r.Width, r.Height));
        // Top row: pixel 0 never written (transparent), pixel 1 skipped by the delta, pixel 2 = green? Delta (1,0) lands on x=1.
        Assert.NotNull(r.Alpha);
        Assert.Equal(new byte[] { 0, 255, 0, 255, 255, 255 }, r.Alpha);
        Assert.Equal([0, 255, 0], r.Rgb[3..6]); // top row, x = 1 is green
        Assert.Equal([255, 0, 0, 255, 0, 0, 255, 0, 0], r.Rgb[9..18]); // bottom row: three reds
    }

    [Fact]
    public void Bmp_rle4_uses_nibbles_and_top_down_order()
    {
        var palette = new byte[] { 0, 0, 0, 0, 0, 0, 255, 0 }; // black, red
        // 4 pixels alternating red,black from one byte 0x10: run of 4 -> 1,0,1,0; end of bitmap.
        var r = RasterDecoders.DecodeBmp(BmpRle(4, 4, 1, palette, [4, 0x10, 0, 1], negativeHeight: true));
        Assert.Equal([255, 0, 0, 0, 0, 0, 255, 0, 0, 0, 0, 0], r.Rgb);
        // Absolute run of 3 nibbles (0x12 0x30) -> 1,2->clamped to 0,3->0.
        var a = RasterDecoders.DecodeBmp(BmpRle(4, 3, 1, palette, [0, 3, 0x10, 0x20, 0, 1]));
        Assert.Equal([255, 0, 0, 0, 0, 0, 0, 0, 0], a.Rgb);
        Assert.Throws<InvalidDataException>(() => RasterDecoders.DecodeBmp(BmpRle(4, 3, 1, palette, [0, 8, 0x10])));
    }

    /// <summary>A GIF whose LZW stream emits a clear code before every pixel (valid, if inefficient).</summary>
    private static byte[] Gif(int w, int h, byte[] palette, int[] pixels, int transparent = -1, bool interlaced = false)
    {
        var ms = new MemoryStream();
        ms.Write("GIF89a"u8);
        ms.Write([(byte)w, 0, (byte)h, 0, 0x80 | 1, 0, 0]);
        ms.Write(palette);
        if (transparent >= 0) ms.Write([0x21, 0xF9, 4, 1, 0, 0, (byte)transparent, 0]);
        ms.Write([0x2C, 0, 0, 0, 0, (byte)w, 0, (byte)h, 0, (byte)(interlaced ? 0x40 : 0)]);
        ms.WriteByte(2);
        // min code size 2: clear = 4, end = 5, code size 3 bits; clear before each pixel keeps the table untouched.
        var bits = new List<bool>();
        void Put(int code) { for (int i = 0; i < 3; i++) bits.Add((code >> i & 1) != 0); }
        foreach (var p in pixels) { Put(4); Put(p); }
        Put(5);
        var bytes = new List<byte>();
        for (int i = 0; i < bits.Count; i += 8)
        {
            int b = 0;
            for (int k = 0; k < 8 && i + k < bits.Count; k++) if (bits[i + k]) b |= 1 << k;
            bytes.Add((byte)b);
        }
        for (int i = 0; i < bytes.Count; i += 255)
        {
            int n = Math.Min(255, bytes.Count - i);
            ms.WriteByte((byte)n);
            ms.Write(bytes.GetRange(i, n).ToArray());
        }
        ms.Write([0, 0x3B]);
        return ms.ToArray();
    }

    private static readonly byte[] Pal4 = [255, 0, 0, 0, 255, 0, 0, 0, 255, 9, 9, 9];

    [Fact]
    public void Gif_decodes_palette_pixels()
    {
        var r = RasterDecoders.DecodeGif(Gif(2, 2, Pal4, [0, 1, 2, 3]));
        Assert.Equal([255, 0, 0, 0, 255, 0, 0, 0, 255, 9, 9, 9], r.Rgb);
        Assert.Null(r.Alpha);
    }

    [Fact]
    public void Gif_transparent_index_becomes_alpha_and_interlacing_reorders_rows()
    {
        var r = RasterDecoders.DecodeGif(Gif(1, 2, Pal4, [0, 1], transparent: 1));
        Assert.Equal([255, 0], r.Alpha);
        // 5 rows, interlaced: stored order is rows 0, 4, 2, 1, 3.
        var il = RasterDecoders.DecodeGif(Gif(1, 5, Pal4, [0, 1, 2, 3, 0], interlaced: true));
        Assert.Equal([0, 3, 2, 0, 1], Enumerable.Range(0, 5).Select(y => (int)il.Rgb[y * 3 + 1] == 255 ? 1 : il.Rgb[y * 3] == 9 ? 3 : il.Rgb[y * 3 + 2] == 255 ? 2 : 0));
    }

    [Fact]
    public void Gif_lzw_with_table_growth_matches_a_run()
    {
        // Codes: clear, 0, 6 (= "00" once added), end - hand-built: "0" then code 6 would reference next=6.
        // clear(4) 0 6 ... table entry 6 is not defined until after the second code, so 6 is the KwKwK case.
        var ms = new MemoryStream();
        ms.Write("GIF89a"u8); ms.Write([3, 0, 1, 0, 0x80 | 1, 0, 0]); ms.Write(Pal4);
        ms.Write([0x2C, 0, 0, 0, 0, 3, 0, 1, 0, 0, 2]);
        var bits = new List<bool>();
        void Put(int code) { for (int i = 0; i < 3; i++) bits.Add((code >> i & 1) != 0); }
        Put(4); Put(1); Put(6); Put(5); // 1, then 6 = KwKwK -> "11", total 3 pixels of index 1
        var bytes = new byte[(bits.Count + 7) / 8];
        for (int i = 0; i < bits.Count; i++) if (bits[i]) bytes[i / 8] |= (byte)(1 << i % 8);
        ms.WriteByte((byte)bytes.Length); ms.Write(bytes); ms.Write([0, 0x3B]);
        var r = RasterDecoders.DecodeGif(ms.ToArray());
        Assert.Equal([0, 255, 0, 0, 255, 0, 0, 255, 0], r.Rgb);
    }

    private static byte[] Tiff(int w, int h, int photo, int spp, byte[] strip, int compression = 1, int bits = 8, ushort[]? map = null, bool bigEndian = false, int predictor = 1)
    {
        var entries = new List<(int Tag, int Type, int Count, long Value)>
        {
            (256, 3, 1, w), (257, 3, 1, h), (258, 3, 1, bits), (259, 3, 1, compression), (262, 3, 1, photo),
            (273, 4, 1, 0), (277, 3, 1, spp), (279, 4, 1, strip.Length),
        };
        if (predictor != 1) entries.Add((317, 3, 1, predictor));
        int ifdAt = 8 + strip.Length;
        int dirSize = 2 + (entries.Count + (map != null ? 1 : 0)) * 12 + 4;
        int mapAt = ifdAt + dirSize;
        var ms = new MemoryStream();
        void W16(int v) { if (bigEndian) { ms.WriteByte((byte)(v >> 8)); ms.WriteByte((byte)v); } else { ms.WriteByte((byte)v); ms.WriteByte((byte)(v >> 8)); } }
        void W32(long v) { if (bigEndian) { W16((int)(v >> 16)); W16((int)(v & 0xFFFF)); } else { W16((int)(v & 0xFFFF)); W16((int)(v >> 16)); } }
        ms.Write(Encoding.ASCII.GetBytes(bigEndian ? "MM" : "II")); W16(42); W32(ifdAt);
        ms.Write(strip);
        W16(entries.Count + (map != null ? 1 : 0));
        foreach (var e in entries.Select(e => e.Tag == 273 ? (273, 4, 1, 8L) : e))
        {
            W16(e.Item1); W16(e.Item2); W32(e.Item3);
            if (e.Item2 == 3) { W16((int)e.Item4); W16(0); } else W32(e.Item4);
        }
        if (map != null) { W16(320); W16(3); W32(map.Length); W32(mapAt); }
        W32(0);
        if (map != null) foreach (var m in map) W16(m);
        return ms.ToArray();
    }

    /// <summary>A little-endian TIFF with one JPEG strip (compression 7) and, when given, the JPEGTables tag.</summary>
    private static byte[] TiffWithJpeg(byte[] strip, byte[]? tables, int strips = 1)
    {
        var entries = new List<(int Tag, int Type, int Count, long Value)> { (256, 3, 1, 1), (257, 3, 1, 1), (258, 3, 1, 8), (259, 3, 1, 7), (262, 3, 1, 6), (273, 4, strips, 8), (277, 3, 1, 3), (279, 4, strips, strip.Length) };
        int ifdAt = 8 + strip.Length;
        int tablesAt = ifdAt + 2 + (entries.Count + (tables is null ? 0 : 1)) * 12 + 4;
        if (tables is not null) entries.Add((347, 7, tables.Length, tablesAt));
        var ms = new MemoryStream();
        void W16(int v) { ms.WriteByte((byte)v); ms.WriteByte((byte)(v >> 8)); }
        void W32(long v) { W16((int)(v & 0xFFFF)); W16((int)(v >> 16)); }
        ms.Write(Encoding.ASCII.GetBytes("II")); W16(42); W32(ifdAt);
        ms.Write(strip);
        W16(entries.Count);
        foreach (var e in entries.OrderBy(e => e.Tag))
        {
            W16(e.Tag); W16(e.Type); W32(e.Count);
            if (e.Type == 3) { W16((int)e.Value); W16(0); } else W32(e.Value);
        }
        W32(0);
        if (tables is not null) ms.Write(tables);
        return ms.ToArray();
    }

    [Fact]
    public void A_tiff_holding_one_jpeg_strip_gives_that_jpeg_with_its_shared_tables()
    {
        var jpeg = TestAssets.Bytes("gradient.jpg");
        int frame = Enumerable.Range(2, jpeg.Length - 3).First(i => jpeg[i] == 0xFF && jpeg[i + 1] is 0xC0 or 0xC2);
        byte[] tables = [.. jpeg[..frame], 0xFF, 0xD9];                 // SOI, the quantization and Huffman tables, EOI
        byte[] strip = [0xFF, 0xD8, .. jpeg[frame..]];                   // SOI, frame, scan
        Assert.Equal(jpeg, RasterDecoders.TiffJpeg(TiffWithJpeg(strip, tables)));
        Assert.Equal(jpeg, RasterDecoders.TiffJpeg(TiffWithJpeg(jpeg, null)));       // a complete JPEG needs no tables
        Assert.Null(RasterDecoders.TiffJpeg(Tiff(1, 1, 2, 3, [1, 2, 3])));            // not a JPEG TIFF
        var image = PdfImage.FromTiff(TiffWithJpeg(strip, tables));
        Assert.True(image.Width > 1);                                                 // the JPEG's own size, not the tag's
    }

    [Fact]
    public void A_tiff_with_several_jpeg_strips_is_joined_into_one_jpeg_with_restart_markers()
    {
        // Written by libtiff (tiffcp -c jpeg -r 16): 40x48, YCbCr 2x2, three strips of 16 rows, the tables shared in JPEGTables.
        var tiff = TestAssets.Bytes("three-strips-jpeg.tif");
        var jpeg = RasterDecoders.TiffJpeg(tiff)!;
        Assert.NotNull(jpeg);
        Assert.Equal((byte)0xFF, jpeg[0]);
        Assert.Equal((byte)0xD8, jpeg[1]);
        Assert.Equal([(byte)0xFF, (byte)0xD9], jpeg[^2..]);
        // The restart interval is one strip: 3 MCUs across (40 / 16, rounded up) by one MCU row (16 / 16).
        int dri = Enumerable.Range(2, jpeg.Length - 8).First(i => jpeg[i] == 0xFF && jpeg[i + 1] == 0xDD);
        Assert.Equal(3, jpeg[dri + 4] << 8 | jpeg[dri + 5]);
        // The frame has the full height, and the strips are chained by RST0 and RST1.
        int frame = Enumerable.Range(2, jpeg.Length - 8).First(i => jpeg[i] == 0xFF && jpeg[i + 1] == 0xC0);
        Assert.Equal((48, 40), (jpeg[frame + 5] << 8 | jpeg[frame + 6], jpeg[frame + 7] << 8 | jpeg[frame + 8]));
        Assert.Contains(Enumerable.Range(0, jpeg.Length - 1), i => jpeg[i] == 0xFF && jpeg[i + 1] == 0xD0);
        Assert.Contains(Enumerable.Range(0, jpeg.Length - 1), i => jpeg[i] == 0xFF && jpeg[i + 1] == 0xD1);
        Assert.DoesNotContain(Enumerable.Range(0, jpeg.Length - 1), i => jpeg[i] == 0xFF && jpeg[i + 1] == 0xD2);
        Assert.Equal((40, 48), (PdfImage.FromTiff(tiff).Width, PdfImage.FromTiff(tiff).Height));
    }

    /// <summary>The entropy-coded data of a JPEG: everything after its scan header up to the end marker.</summary>
    private static byte[] ScanData(byte[] jpeg)
    {
        int sos = Enumerable.Range(2, jpeg.Length - 4).First(i => jpeg[i] == 0xFF && jpeg[i + 1] == 0xDA);
        return jpeg[(sos + 2 + (jpeg[sos + 2] << 8 | jpeg[sos + 3]))..];
    }

    [Theory]
    [InlineData("tiled-jpeg-aligned.tif", "whole-jpeg-aligned.tif", 4)]                  // 64x48, YCbCr 2x2, 16x16 tiles
    [InlineData("tiled-jpeg-aligned-wide.tif", "whole-jpeg-aligned.tif", 4)]             // 32x16 tiles
    [InlineData("tiled-jpeg-aligned-gray.tif", "whole-jpeg-aligned-gray.tif", 4)]        // one component: an MCU is one block
    public void A_tiff_with_jpeg_tiles_is_joined_into_one_jpeg_with_the_same_coefficients(string tiled, string whole, int _)
    {
        // Written by libtiff (tiffcp -c jpeg:90 -t -w 16 -l 16, and -r 64 for the single strip). Both use the standard Huffman tables and the same
        // quantization, so re-coding the tiles' coefficients on the image's own grid gives the very bytes libtiff codes for the whole picture.
        var joined = RasterDecoders.TiffJpeg(TestAssets.Bytes(tiled))!;
        var reference = RasterDecoders.TiffJpeg(TestAssets.Bytes(whole))!;
        Assert.NotNull(joined);
        Assert.Equal(ScanData(reference), ScanData(joined));
        int frame = Enumerable.Range(2, joined.Length - 8).First(i => joined[i] == 0xFF && joined[i + 1] == 0xC0);
        Assert.Equal((48, 64), (joined[frame + 5] << 8 | joined[frame + 6], joined[frame + 7] << 8 | joined[frame + 8]));
        Assert.DoesNotContain(Enumerable.Range(0, joined.Length - 1), i => joined[i] == 0xFF && joined[i + 1] == 0xDD);   // no restart intervals
        Assert.Equal((64, 48), (PdfImage.FromTiff(TestAssets.Bytes(tiled)).Width, PdfImage.FromTiff(TestAssets.Bytes(tiled)).Height));
    }

    [Fact]
    public void Tiles_that_overhang_the_picture_are_cut_to_its_size()
    {
        // 70x50 in 16x16 tiles: the last column and row of tiles are partly outside. (libtiff pads the tiles differently from a whole-picture
        // encode, so the edge blocks cannot be compared with one; the frame must have the picture's size and the data must be one scan.)
        var tiff = TestAssets.Bytes("tiled-jpeg.tif");
        var joined = RasterDecoders.TiffJpeg(tiff)!;
        Assert.NotNull(joined);
        int frame = Enumerable.Range(2, joined.Length - 8).First(i => joined[i] == 0xFF && joined[i + 1] == 0xC0);
        Assert.Equal((50, 70), (joined[frame + 5] << 8 | joined[frame + 6], joined[frame + 7] << 8 | joined[frame + 8]));
        Assert.Equal(1, Enumerable.Range(0, joined.Length - 1).Count(i => joined[i] == 0xFF && joined[i + 1] == 0xDA));
        Assert.Equal([(byte)0xFF, (byte)0xD9], joined[^2..]);
        Assert.Equal((70, 50), (PdfImage.FromTiff(tiff).Width, PdfImage.FromTiff(tiff).Height));
    }

    [Fact]
    public void Jpeg_tiles_that_are_damaged_or_not_whole_mcus_are_not_joined()
    {
        var tiff = TestAssets.Bytes("tiled-jpeg.tif");
        var odd = (byte[])tiff.Clone();      // a tile width of 12 is not a whole MCU
        int ifd = BitConverter.ToInt32(odd, 4), count = BitConverter.ToUInt16(odd, ifd);
        for (int i = 0; i < count; i++)
        {
            int e = ifd + 2 + i * 12;
            if (BitConverter.ToUInt16(odd, e) == 322) odd[e + 8] = 12;
        }
        Assert.Null(RasterDecoders.TiffJpeg(odd));
        var broken = (byte[])tiff.Clone();   // the entropy-coded data of the first tile overwritten
        Array.Fill(broken, (byte)0xFE, 700, 120);
        var result = RasterDecoders.TiffJpeg(broken);
        Assert.True(result is null || result.Length > 100);   // never throws
    }

    [Fact]
    public void Jpeg_strips_that_cannot_be_chained_are_not_joined()
    {
        var tiff = TestAssets.Bytes("three-strips-jpeg.tif");
        // A rows-per-strip that is not a whole number of MCU rows (tag 278, value 16 -> 12) cannot be joined.
        var odd = (byte[])tiff.Clone();
        int ifd = BitConverter.ToInt32(odd, 4);
        int count = BitConverter.ToUInt16(odd, ifd);
        for (int i = 0; i < count; i++)
        {
            int e = ifd + 2 + i * 12;
            if (BitConverter.ToUInt16(odd, e) == 278) odd[e + 8] = 12;
        }
        Assert.Null(RasterDecoders.TiffJpeg(odd));
    }

    [Fact]
    public void A_tiffs_pixel_size_and_resolution_are_read_from_its_first_directory()
    {
        Assert.Equal((4, 2, 200.0), new DotOffice.Core.ImageData(TestAssets.Bytes("dpi-200.tif"), "a").PixelSize);        // pixels per inch
        var perCentimeter = new DotOffice.Core.ImageData(TestAssets.Bytes("dpcm-100.tif"), "b").PixelSize!.Value;        // 100 pixels per centimeter
        Assert.Equal((4, 2), (perCentimeter.Width, perCentimeter.Height));
        Assert.Equal(254, perCentimeter.Dpi, 6);
        Assert.Equal((40, 48, 96.0), new DotOffice.Core.ImageData(TestAssets.Bytes("three-strips-jpeg.tif"), "c").PixelSize);  // no resolution: 96
        Assert.Null(new DotOffice.Core.ImageData([(byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0], "d").PixelSize);                 // a directory that is cut off
    }

    [Fact]
    public void Tiff_rgb_gray_and_palette()
    {
        var rgb = RasterDecoders.DecodeTiff(Tiff(2, 1, 2, 3, [1, 2, 3, 4, 5, 6]));
        Assert.Equal([1, 2, 3, 4, 5, 6], rgb.Rgb);
        var rgbBe = RasterDecoders.DecodeTiff(Tiff(1, 1, 2, 3, [7, 8, 9], bigEndian: true));
        Assert.Equal([7, 8, 9], rgbBe.Rgb);
        var white0 = RasterDecoders.DecodeTiff(Tiff(2, 1, 0, 1, [0, 200])); // WhiteIsZero inverts
        Assert.Equal([255, 255, 255, 55, 55, 55], white0.Rgb);
        var map = new ushort[] { 0, 65535, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 }; // 2-bit: only index 1 has red
        var pal = RasterDecoders.DecodeTiff(Tiff(2, 1, 3, 1, [0b0100_0000], bits: 2, map: map.Take(12).ToArray()));
        Assert.Equal(new byte[] { 255, 0, 0, 0, 0, 0 }, pal.Rgb);
    }

    [Fact]
    public void Tiff_packbits_and_rgba()
    {
        // run of 4 x 0x10 (-3, 0x10) then literal of 2 (1, a, b)... 2 pixels RGB = 6 bytes: -2 -> 3 copies of 0x10; literal 3 bytes
        var r = RasterDecoders.DecodeTiff(Tiff(2, 1, 2, 3, [unchecked((byte)-2), 0x10, 2, 1, 2, 3], compression: 32773));
        Assert.Equal([0x10, 0x10, 0x10, 1, 2, 3], r.Rgb);
        var a = RasterDecoders.DecodeTiff(Tiff(1, 1, 2, 4, [1, 2, 3, 99]));
        Assert.Equal([99], a.Alpha);
        Assert.Throws<NotSupportedException>(() => RasterDecoders.DecodeTiff(Tiff(1, 1, 2, 3, [1, 2, 3], compression: 7)));
    }

    /// <summary>A TIFF LZW encoder (MSB-first, early change), as libtiff writes it.</summary>
    private static byte[] TiffLzwEncode(byte[] input)
    {
        var bits = new List<bool>();
        int size = 9, next = 258;
        void Put(int code) { for (int i = size - 1; i >= 0; i--) bits.Add((code >> i & 1) != 0); }
        var table = new Dictionary<(int, byte), int>();
        Put(256);
        int cur = input[0];
        for (int i = 1; i < input.Length; i++)
        {
            if (table.TryGetValue((cur, input[i]), out var code)) { cur = code; continue; }
            Put(cur);
            table[(cur, input[i])] = next++;
            if (next > (1 << size) - 1 && size < 12) size++;
            if (next == 4094) { Put(256); size = 9; next = 258; table.Clear(); }
            cur = input[i];
        }
        Put(cur);
        next++;
        if (next > (1 << size) - 1 && size < 12) size++;
        Put(257);
        var bytes = new byte[(bits.Count + 7) / 8];
        for (int i = 0; i < bits.Count; i++) if (bits[i]) bytes[i / 8] |= (byte)(0x80 >> i % 8);
        return bytes;
    }

    [Fact]
    public void Tiff_lzw_is_msb_first_with_early_change_and_survives_table_growth()
    {
        // 100 x 40 RGB of pseudo-random bytes: the table passes 511, 1023 and 2047 and is cleared once.
        var rng = new Random(7);
        var pixels = new byte[100 * 40 * 3];
        for (int i = 0; i < pixels.Length; i++) pixels[i] = (byte)(rng.Next(4) * 60);
        var r = RasterDecoders.DecodeTiff(Tiff(100, 40, 2, 3, TiffLzwEncode(pixels), compression: 5));
        Assert.Equal(pixels, r.Rgb);

        // The textbook case: codes "7 7 7 8 8 7 7 6 6" -> repeated runs use the not-yet-defined code (KwKwK).
        var small = new byte[] { 10, 10, 10, 10, 10, 10, 20, 20, 20 };
        Assert.Equal(small, RasterDecoders.DecodeTiff(Tiff(3, 1, 2, 3, TiffLzwEncode(small), compression: 5)).Rgb);
    }

    [Fact]
    public void Tiff_horizontal_predictor_adds_the_left_neighbour()
    {
        var r = RasterDecoders.DecodeTiff(Tiff(3, 1, 2, 3, [10, 20, 30, 1, 2, 3, 1, 2, 3], predictor: 2));
        Assert.Equal([10, 20, 30, 11, 22, 33, 12, 24, 36], r.Rgb);
        var lzw = RasterDecoders.DecodeTiff(Tiff(2, 1, 1, 1, TiffLzwEncode([100, 5]), compression: 5, predictor: 2));
        Assert.Equal([100, 100, 100, 105, 105, 105], lzw.Rgb);
    }

    [Fact]
    public void Decoded_pictures_are_written_as_pdf_images()
    {
        var doc = new PdfDocument();
        var image = doc.AddImage(PdfImage.FromGif(Gif(1, 2, Pal4, [0, 1], transparent: 1)));
        doc.AddPage(PageSize.A4).Canvas.DrawImage(image, 10, 10, 20, 20);
        doc.AddImage(PdfImage.FromBmp(Bmp(24, 1, 1, [], [1, 2, 3, 0])));
        doc.AddImage(PdfImage.FromTiff(Tiff(1, 1, 2, 3, [1, 2, 3])));
        using var ms = new MemoryStream();
        doc.Save(ms);
        Assert.Contains("/SMask", Encoding.Latin1.GetString(ms.ToArray()));
    }

    [Fact]
    public void Set_alpha_writes_a_graphics_state_resource_and_the_operator()
    {
        var doc = new PdfDocument();
        var canvas = doc.AddPage(PageSize.A4).Canvas;
        canvas.SaveState().SetAlpha(0.4).Rectangle(0, 0, 10, 10).Fill();
        canvas.RestoreState();
        using var ms = new MemoryStream();
        doc.Save(ms);
        var text = Encoding.Latin1.GetString(ms.ToArray());
        Assert.Contains("/ExtGState << /GSa40 << /ca 0.4 /CA 0.4 >> >>", text);
    }
}
