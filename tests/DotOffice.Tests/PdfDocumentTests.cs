using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using DotOffice.Fonts;
using DotOffice.Pdf;

namespace DotOffice.Tests;

public class PdfDocumentTests
{
    private static readonly TrueTypeFont Font = TrueTypeFont.Load(TestAssets.Path("NotoSans-Regular.ttf"));

    private static byte[] BuildSample()
    {
        var doc = new PdfDocument { Title = "Hello Привет" };
        var font = doc.AddFont(Font);
        var image = doc.AddImage(PdfImage.FromJpeg(TestAssets.Bytes("gradient.jpg")));
        var page = doc.AddPage(PageSize.A4);
        page.Canvas
            .DrawText(font, 24, 72, 760, "Hello, world!")
            .SetStrokeColor(1, 0, 0).Rectangle(72, 700, 200, 40).Stroke()
            .DrawImage(image, 72, 550, 160, 100);
        doc.AddPage(PageSize.Letter).Canvas.DrawText(font, 12, 72, 700, "Page two (Ω)");

        using var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    [Fact]
    public void Xref_offsets_point_at_their_objects()
    {
        var pdf = BuildSample();
        var text = Encoding.Latin1.GetString(pdf);

        Assert.StartsWith("%PDF-1.7", text);
        Assert.EndsWith("%%EOF\n", text);

        int startXref = int.Parse(Regex.Match(text, @"startxref\n(\d+)").Groups[1].Value);
        Assert.StartsWith("xref", text[startXref..]);

        var entries = Regex.Matches(text[startXref..], @"(\d{10}) 00000 n ");
        Assert.NotEmpty(entries);
        for (int i = 0; i < entries.Count; i++)
        {
            int offset = int.Parse(entries[i].Groups[1].Value);
            Assert.StartsWith($"{i + 1} 0 obj", text[offset..]);
        }
    }

    [Fact]
    public void Pages_and_resources_are_declared()
    {
        var text = Encoding.Latin1.GetString(BuildSample());
        Assert.Contains("/Count 2", text);
        Assert.Contains("/MediaBox [0 0 612 792]", text);
        Assert.Contains("/Subtype /CIDFontType2", text);
        Assert.Contains("/Filter /DCTDecode", text);
        Assert.Matches(@"/BaseFont /[A-Z]{6}\+NotoSans-Regular", text);
        Assert.Contains("<FEFF", text); // non-ASCII title
    }

    [Fact]
    public void Content_stream_encodes_glyph_ids()
    {
        var text = Encoding.Latin1.GetString(BuildSample());
        var hGlyph = Font.GetGlyphId('H').ToString("X4");
        var contents = Inflated(BuildSample());
        Assert.Contains(contents, c => c.Contains($"Tf 72 760 Td <{hGlyph}") && c.Contains(" Tj ET"));
        Assert.Contains(contents, c => c.Contains("beginbfchar") && c.Contains($"<{hGlyph}> <0048>"));
    }

    [Fact]
    public void Saving_empty_document_throws()
    {
        Assert.Throws<InvalidOperationException>(() => new PdfDocument().Save(Stream.Null));
    }

    [Fact]
    public void Reads_jpeg_dimensions()
    {
        var color = PdfImage.FromJpeg(TestAssets.Bytes("gradient.jpg"));
        Assert.Equal((160, 100), (color.Width, color.Height));
        var gray = PdfImage.FromJpeg(TestAssets.Bytes("gray.jpg"));
        Assert.Equal((64, 64), (gray.Width, gray.Height));
        Assert.Throws<InvalidDataException>(() => PdfImage.FromJpeg([1, 2, 3, 4]));
    }

    private static List<string> Inflated(byte[] pdf)
    {
        var result = new List<string>();
        var text = Encoding.Latin1.GetString(pdf);
        foreach (Match m in Regex.Matches(text, @"/Length (\d+) /Filter /FlateDecode[^>]*>>\nstream\n"))
        {
            int start = m.Index + m.Length, length = int.Parse(m.Groups[1].Value);
            using var zlib = new ZLibStream(new MemoryStream(pdf, start, length), CompressionMode.Decompress);
            using var reader = new StreamReader(zlib, Encoding.Latin1);
            result.Add(reader.ReadToEnd());
        }
        return result;
    }
}
