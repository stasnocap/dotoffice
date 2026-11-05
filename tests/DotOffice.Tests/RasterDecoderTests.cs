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
    public void Bmp_with_rle_is_rejected_and_truncated_data_is_invalid()
    {
        var rle = Bmp(8, 1, 1, new byte[4], [0, 0, 0, 0]);
        rle[30] = 1;
        Assert.Throws<NotSupportedException>(() => RasterDecoders.DecodeBmp(rle));
        Assert.Throws<InvalidDataException>(() => RasterDecoders.DecodeBmp(Bmp(24, 5, 5, [], [0, 0])));
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

    private static byte[] Tiff(int w, int h, int photo, int spp, byte[] strip, int compression = 1, int bits = 8, ushort[]? map = null, bool bigEndian = false)
    {
        var entries = new List<(int Tag, int Type, int Count, long Value)>
        {
            (256, 3, 1, w), (257, 3, 1, h), (258, 3, 1, bits), (259, 3, 1, compression), (262, 3, 1, photo),
            (273, 4, 1, 0), (277, 3, 1, spp), (279, 4, 1, strip.Length),
        };
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
        Assert.Throws<NotSupportedException>(() => RasterDecoders.DecodeTiff(Tiff(1, 1, 2, 3, [1, 2, 3], compression: 5)));
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
